"""
Template-driven PowerPoint generation API
=========================================

Architecture
------------
Copilot Studio agent -> Power Platform custom connector -> this Flask API
-> python-pptx -> corporate_template.pptx -> downloadable .pptx

Design rules (the template is the single source of truth)
---------------------------------------------------------
* Every deck is created from ``corporate_template.pptx``. There is NO
  ``Presentation()`` fallback. If the template is missing/invalid the API
  returns a controlled JSON error (HTTP 503, errorCode TEMPLATE_ERROR).
* Layouts are discovered dynamically by NAME (with a structural fallback based
  on the placeholders each layout contains), never by hard-coded index only.
* Template placeholders are populated wherever possible. New text boxes /
  shapes are created only when the chosen layout has no suitable placeholder.
* Fonts are never set by name (everything inherits the theme fonts) and
  colours are only ever *theme* colours (ACCENT_1..6, DARK_1, LIGHT_1), so
  recolouring the theme recolours generated shapes, charts and tables too.
* Backgrounds, logos, footers and slide size live on the template's master /
  layouts and are never touched.

Endpoints
---------
GET  /                      health + template status
GET  /health                same as /
GET  /template-info         template diagnostics (layouts, placeholders, roles)
POST /createppt             generate a presentation
GET  /download/<filename>   download a generated presentation
"""

from __future__ import annotations

import copy
import json
import logging
import math
import os
import re
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from flask import Flask, jsonify, request, send_from_directory
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.shapes.placeholder import ChartPlaceholder, TablePlaceholder
from pptx.util import Emu, Pt
from werkzeug.exceptions import HTTPException
from werkzeug.middleware.proxy_fix import ProxyFix

# --------------------------------------------------------------------------- #
# Configuration (all overridable with environment variables on Render)
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_FILENAME = os.environ.get("TEMPLATE_FILENAME", "corporate_template.pptx")
TEMPLATE_PATH = BASE_DIR / TEMPLATE_FILENAME
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", str(Path(tempfile.gettempdir()) / "generated_presentations")))
FILE_TTL_SECONDS = int(os.environ.get("FILE_TTL_SECONDS", "86400"))  # delete files after 24h
MAX_SLIDES = int(os.environ.get("MAX_SLIDES", "60"))
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")  # e.g. https://my-app.onrender.com
KEEP_FOOTERS = os.environ.get("KEEP_FOOTERS", "1") != "0"  # copy slide-number/footer placeholders to slides

EMU_PER_INCH = 914400
REFERENCE_WIDTH = 12192000  # 13.333in (16:9). Font/shape sizes scale relative to this.

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("ppt-service")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024  # 2 MB request cap
app.config["JSON_SORT_KEYS"] = False
# Render terminates TLS in front of the app; trust its forwarded headers so
# request.url_root is https://<your-host>/ and not http://.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)


# --------------------------------------------------------------------------- #
# Errors
# --------------------------------------------------------------------------- #
class ApiError(Exception):
    """An error whose message is safe to show to the end user / Copilot."""

    def __init__(self, message: str, status: int = 400, code: str = "BAD_REQUEST", details: Any = None):
        super().__init__(message)
        self.message = message
        self.status = status
        self.code = code
        self.details = details


class TemplateError(ApiError):
    """The corporate template is missing, unreadable or unusable."""

    def __init__(self, message: str, details: Any = None):
        super().__init__(message, status=503, code="TEMPLATE_ERROR", details=details)


def error_response(err: ApiError):
    body = {"status": "error", "errorCode": err.code, "message": err.message}
    if err.details:
        body["details"] = err.details
    return jsonify(body), err.status


# --------------------------------------------------------------------------- #
# Small generic helpers
# --------------------------------------------------------------------------- #
_CTRL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")  # illegal in XML


def clean(value: Any, limit: Optional[int] = None, multiline: bool = False) -> str:
    """Convert to a safe single-line (or multi-line) string; strip XML-illegal characters."""
    if value is None:
        return ""
    text = _CTRL_CHARS.sub("", str(value)).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not multiline:
        text = re.sub(r"\s*\n\s*", " ", text)
    if limit and len(text) > limit:
        text = text[: limit - 1].rstrip() + "…"
    return text


def norm_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def pick(d: Any, keys: Tuple[str, ...], default: Any = None) -> Any:
    """Return the first non-empty value among several alias keys."""
    if not isinstance(d, dict):
        return default
    for k in keys:
        v = d.get(k)
        if v not in (None, "", [], {}):
            return v
    return default


HEAD_KEYS = ("title", "heading", "name", "label", "header", "point")
BODY_KEYS = ("text", "description", "desc", "detail", "details", "body", "content", "subtitle")


def to_points(raw: Any, level: int = 0, out: Optional[List[Tuple[str, int]]] = None) -> List[Tuple[str, int]]:
    """
    Normalise bullet input into [(text, level)].
    Accepts: "a\\nb", ["a", "b"], ["a", ["sub1", "sub2"]], [{"text": "a", "children": [...]}],
    [{"title": "Head", "description": "Detail"}].
    """
    if out is None:
        out = []
    if raw is None:
        return out
    level = min(level, 4)
    if isinstance(raw, str):
        for line in raw.split("\n"):
            t = clean(re.sub(r"^\s*[•\-\*–·]\s+", "", line), 500)
            if t:
                out.append((t, level))
    elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
        out.append((clean(raw), level))
    elif isinstance(raw, dict):
        head = pick(raw, HEAD_KEYS)
        body = pick(raw, BODY_KEYS)
        children = pick(raw, ("children", "subPoints", "subpoints", "items", "bullets"))
        if isinstance(body, (list, dict)):  # nested list given as body
            children, body = body, None
        text = f"{clean(head)}: {clean(body)}" if head and body else clean(head or body)
        lvl = raw.get("level")
        lvl = level if not isinstance(lvl, int) else max(0, min(lvl, 4))
        if text:
            out.append((clean(text, 500), lvl))
        if children:
            to_points(children, lvl + 1, out)
    elif isinstance(raw, (list, tuple)):
        for item in raw:
            if isinstance(item, (list, tuple)):
                to_points(item, level + 1, out)
            else:
                to_points(item, level, out)
    return out


def split_title_text(s: str) -> Tuple[str, str]:
    """'Title: detail' or 'Title - detail' -> ('Title', 'detail'); otherwise (s, '')."""
    m = re.match(r"^(.{2,45}?)\s*(?::|\s[-–—]\s)\s*(.+)$", s)
    return (clean(m.group(1), 120), clean(m.group(2), 400)) if m else (clean(s, 400), "")


def norm_items(raw: Any, limit: int) -> List[Dict[str, str]]:
    """Normalise cards / rows / process steps into [{'title':..., 'text':...}]."""
    if raw is None:
        return []
    if isinstance(raw, str):
        raw = [ln for ln in raw.split("\n") if ln.strip()]
    if isinstance(raw, dict):
        raw = [{"title": k, "text": v} for k, v in raw.items()]
    items: List[Dict[str, str]] = []
    for it in raw if isinstance(raw, list) else [raw]:
        if isinstance(it, dict):
            title = clean(pick(it, HEAD_KEYS), 120)
            body = pick(it, BODY_KEYS)
            if isinstance(body, list):
                body = "; ".join(t for t, _ in to_points(body))
            text = clean(body, 400)
        else:
            title, text = split_title_text(clean(it, 500))
        if title or text:
            items.append({"title": title, "text": text})
    return items[:limit]


_STAT_RE = re.compile(r"^\s*([^\s:]*\d[^\s:]*)\s*[:\-–—]?\s*(.*)$")


def norm_stats(raw: Any, limit: int = 8) -> List[Dict[str, str]]:
    """Normalise statistics into [{'value': '85%', 'label': 'of leaders ...'}]."""
    if raw is None:
        return []
    if isinstance(raw, str):
        raw = [ln for ln in raw.split("\n") if ln.strip()]
    if isinstance(raw, dict):
        raw = [{"value": v, "label": k} for k, v in raw.items()]
    out: List[Dict[str, str]] = []
    for it in raw if isinstance(raw, list) else [raw]:
        if isinstance(it, dict):
            value = clean(pick(it, ("value", "number", "stat", "figure", "metric", "kpi")), 24)
            label = clean(pick(it, ("label", "title", "name", "description", "text", "caption")), 160)
        else:
            s = clean(it, 200)
            m = _STAT_RE.match(s)
            value, label = (clean(m.group(1), 24), clean(m.group(2), 160)) if m else ("", s)
        if value or label:
            out.append({"value": value, "label": label})
    return out[:limit]


def to_num(v: Any) -> Optional[float]:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = re.sub(r"[,%$€£₹\s]", "", v)
        try:
            return float(s)
        except ValueError:
            return None
    return None


# --------------------------------------------------------------------------- #
# Geometry helper
# --------------------------------------------------------------------------- #
@dataclass
class Rect:
    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    def inset(self, dx: int, dy: Optional[int] = None) -> "Rect":
        dy = dx if dy is None else dy
        return Rect(self.left + dx, self.top + dy, max(self.width - 2 * dx, 1), max(self.height - 2 * dy, 1))

    def split_h(self, ratio: float, gap: int) -> Tuple["Rect", "Rect"]:
        """Split horizontally into two rects; the left gets `ratio` of the width."""
        lw = int((self.width - gap) * ratio)
        rw = self.width - gap - lw
        return Rect(self.left, self.top, lw, self.height), Rect(self.left + lw + gap, self.top, rw, self.height)


def shape_rect(shape: Any) -> Optional[Rect]:
    """Effective (possibly inherited) geometry of a shape/placeholder, or None."""
    try:
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
    except Exception:
        return None
    if None in (l, t, w, h) or w <= 0 or h <= 0:
        return None
    return Rect(int(l), int(t), int(w), int(h))


def remove_shape(shape: Any) -> None:
    el = shape._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


# --------------------------------------------------------------------------- #
# Placeholder helpers
# --------------------------------------------------------------------------- #
_TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_BODY_TYPES = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT}
_PICTURE_TYPES = {PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BITMAP}
_FOOTER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.HEADER}


def ph_kind(ph: Any) -> str:
    """Classify a placeholder: title | subtitle | body | chart | table | picture | footer | other."""
    try:
        t = ph.placeholder_format.type
    except Exception:
        return "other"
    if t in _TITLE_TYPES:
        return "title"
    if t == PP_PLACEHOLDER.SUBTITLE:
        return "subtitle"
    if t in _BODY_TYPES:
        return "body"
    if t == PP_PLACEHOLDER.CHART:
        return "chart"
    if t == PP_PLACEHOLDER.TABLE:
        return "table"
    if t in _PICTURE_TYPES:
        return "picture"
    if t in _FOOTER_TYPES:
        return "footer"
    return "other"  # vertical text, media, org chart, ...


def ph_type_name(ph: Any) -> str:
    try:
        t = ph.placeholder_format.type
        return getattr(t, "name", str(t))
    except Exception:
        return "UNKNOWN"


def inherited_font_pt(ph: Any, kind: str, default: float) -> float:
    """
    Best-effort lookup of the template's own font size for a placeholder
    (slide -> layout -> master placeholder lstStyle, then master text styles).
    Used ONLY to decide whether text must be shrunk; never to override fonts.
    """
    try:
        node = ph
        for _ in range(3):
            if node is None:
                break
            vals = node._element.xpath("./p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz")
            if vals:
                return int(vals[0]) / 100.0
            node = getattr(node, "_base_placeholder", None)
        master = ph.part.slide_layout.slide_master
        style = "titleStyle" if kind == "title" else "bodyStyle"
        vals = master._element.xpath(f"./p:txStyles/p:{style}/a:lvl1pPr/a:defRPr/@sz")
        if vals:
            return int(vals[0]) / 100.0
    except Exception:
        pass
    return default


def estimate_font_pt(items: List[Tuple[str, int]], width_emu: int, height_emu: int,
                     max_pt: float, min_pt: float) -> float:
    """
    Largest font size (<= max_pt, >= min_pt) at which the text is estimated to
    fit the box. Heuristic (avg glyph width ~0.52em, line height 1.2em) because
    python-pptx cannot measure text. normAutofit is also enabled as a backstop.
    """
    min_pt = min(min_pt, max_pt)
    width_pt = max(width_emu / 12700.0 - 14 - 22, 40)  # insets + bullet indent
    height_pt = max(height_emu / 12700.0 - 8, 20)
    pt = float(max_pt)
    while True:
        total = 0.0
        for text, level in items:
            usable = max(width_pt - level * 20, 30)
            chars_per_line = max(int(usable / (pt * 0.52)), 1)
            total += math.ceil(max(len(text), 1) / chars_per_line) * pt * 1.2 + pt * 0.35
        if total <= height_pt or pt <= min_pt:
            return float(int(pt)) if pt > min_pt else float(min_pt)
        pt -= 0.5


def apply_font_size(tf: Any, pt: float) -> None:
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(pt)


# --------------------------------------------------------------------------- #
# Template loading & validation
# --------------------------------------------------------------------------- #
def validate_template_file() -> None:
    """Cheap pre-flight checks that give friendly messages before python-pptx runs."""
    if not TEMPLATE_PATH.exists():
        raise TemplateError(
            f"The corporate template '{TEMPLATE_FILENAME}' was not found on the server. "
            "Add it to the repository root next to app.py and redeploy.")
    if not TEMPLATE_PATH.is_file() or TEMPLATE_PATH.stat().st_size == 0:
        raise TemplateError(f"The corporate template '{TEMPLATE_FILENAME}' is empty or not a file.")
    if TEMPLATE_PATH.suffix.lower() not in (".pptx", ".potx"):
        raise TemplateError("The corporate template must be a .pptx file.")
    if not zipfile.is_zipfile(TEMPLATE_PATH):
        raise TemplateError(
            f"'{TEMPLATE_FILENAME}' is not a valid PowerPoint file (it may be corrupted, "
            "or a Git LFS pointer / HTML page was committed instead of the binary).")


def clear_template_slides(prs: Any) -> int:
    """
    Remove any sample slides stored inside the template so only generated
    slides remain. Master, layouts, theme, logo and size are untouched.
    """
    sld_id_lst = prs.slides._sldIdLst
    removed = 0
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.rId)  # orphaned slide parts are not written on save
        sld_id_lst.remove(sld_id)
        removed += 1
    return removed


def load_template() -> Any:
    """Open corporate_template.pptx (never Presentation() with no argument)."""
    validate_template_file()
    try:
        prs = Presentation(str(TEMPLATE_PATH))
    except Exception as exc:  # python-pptx raises many exception types for bad packages
        log.exception("Template could not be opened")
        raise TemplateError(
            f"The corporate template '{TEMPLATE_FILENAME}' could not be opened as a PowerPoint file.",
            details=str(exc)[:300]) from exc
    if len(prs.slide_layouts) == 0:
        raise TemplateError(f"The corporate template '{TEMPLATE_FILENAME}' contains no slide layouts.")
    clear_template_slides(prs)
    return prs


# --------------------------------------------------------------------------- #
# Dynamic layout discovery
# --------------------------------------------------------------------------- #
# role -> ([(name keyword, score)], [words that disqualify a layout name])
# Keywords are matched at word starts (so "end" does not match "legend").
ROLE_RULES: Dict[str, Tuple[List[Tuple[str, int]], List[str]]] = {
    "cover": ([("cover", 100), ("title slide", 95), ("opening", 70), ("front", 70), ("title", 20)],
              ["content", "section", "only", "chart", "table", "agenda", "closing", "comparison",
               "two", "thank", "end", "divider"]),
    "agenda": ([("agenda", 100), ("outline", 70), ("contents", 80), ("toc", 80), ("overview", 40)], []),
    "section": ([("section", 100), ("divider", 95), ("chapter", 90), ("header", 30)], []),
    "content": ([("title and content", 100), ("title & content", 100), ("bullet", 70), ("content", 60),
                 ("body", 50), ("text", 30)],
                ["two", "comparison", "section", "agenda", "chart", "table", "closing", "thank", "cover",
                 "picture", "caption", "column", "split", "half"]),
    "two_column": ([("two content", 100), ("two column", 100), ("2 column", 100), ("two-column", 100),
                    ("two col", 90), ("columns", 80), ("split", 70), ("side by side", 70), ("2 content", 100)], []),
    "comparison": ([("comparison", 100), ("compare", 90), ("versus", 80), ("vs", 75), ("before and after", 70)], []),
    "chart": ([("chart", 100), ("graph", 90), ("data", 30), ("insight", 30)], []),
    "table": ([("table", 100), ("matrix", 60), ("grid", 40)], []),
    "closing": ([("closing", 100), ("thank", 100), ("end", 70), ("contact", 60), ("questions", 60)], []),
    "title_only": ([("title only", 100), ("only title", 100), ("header only", 80)], []),
    "blank": ([("blank", 100)], []),
}
GLOBAL_EXCLUDE = ["vertical"]

# Placeholder-signature fallback when no layout *name* matches a role.
STRUCTURAL = {
    "cover": lambda s: s["title"] >= 1 and s["subtitle"] >= 1 and s["body"] == 0,
    "content": lambda s: s["title"] >= 1 and s["body"] == 1 and s["subtitle"] == 0,
    "two_column": lambda s: s["title"] >= 1 and s["body"] == 2,
    "comparison": lambda s: s["title"] >= 1 and s["body"] >= 4,
    "chart": lambda s: s["chart"] >= 1,
    "table": lambda s: s["table"] >= 1,
    "title_only": lambda s: (s["title"] >= 1 and s["body"] == 0 and s["subtitle"] == 0
                             and s["chart"] == 0 and s["table"] == 0 and s["picture"] == 0),
}
# If a role cannot be found by name or structure, borrow another role's layout.
FALLBACK_CHAIN = {
    "cover": [], "content": [], "title_only": ["content"],
    "agenda": ["content"], "section": ["title_only", "content"],
    "two_column": ["content"], "comparison": ["two_column", "content"],
    "chart": ["title_only", "content"], "table": ["title_only", "content"],
    "closing": ["cover"], "blank": ["title_only"],
}
ALL_ROLES = list(ROLE_RULES.keys())


@dataclass
class LayoutChoice:
    role: str
    layout: Any
    index: int
    how: str  # override | name | structure | fallback:<role> | last-resort


def layout_signature(layout: Any) -> Dict[str, int]:
    sig = {"title": 0, "subtitle": 0, "body": 0, "chart": 0, "table": 0, "picture": 0, "footer": 0}
    for ph in layout.placeholders:
        k = ph_kind(ph)
        if k in sig:
            sig[k] += 1
    return sig


def _name_score(role: str, name: str) -> int:
    n = name.lower().strip()
    includes, excludes = ROLE_RULES[role]
    if any(re.search(r"\b" + re.escape(w), n) for w in excludes + GLOBAL_EXCLUDE):
        return 0
    best = 0
    for kw, score in includes:
        if re.search(r"\b" + re.escape(kw), n):
            best = max(best, score)
    return best


def _env_layout_overrides() -> Dict[str, str]:
    """Optional: LAYOUT_MAP='{"cover":"Cover","closing":"Thank You"}' (role -> exact layout name)."""
    raw = os.environ.get("LAYOUT_MAP", "").strip()
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return {str(k): str(v) for k, v in data.items()} if isinstance(data, dict) else {}
    except ValueError:
        log.warning("LAYOUT_MAP is not valid JSON; ignoring it")
        return {}


def discover_layouts(prs: Any) -> Dict[str, LayoutChoice]:
    """Resolve every role (cover, agenda, section, ...) to a layout of the template."""
    layouts = list(prs.slide_layouts)
    sigs = [layout_signature(l) for l in layouts]
    overrides = _env_layout_overrides()
    resolved: Dict[str, LayoutChoice] = {}

    def by_override(role: str) -> Optional[LayoutChoice]:
        want = overrides.get(role)
        if want:
            for i, l in enumerate(layouts):
                if l.name.strip().lower() == want.strip().lower():
                    return LayoutChoice(role, l, i, "override")
        return None

    def by_name(role: str) -> Optional[LayoutChoice]:
        best_i, best = -1, 0
        for i, l in enumerate(layouts):
            s = _name_score(role, l.name or "")
            if s > best:  # strict '>' keeps the earliest layout on ties
                best_i, best = i, s
        return LayoutChoice(role, layouts[best_i], best_i, "name") if best_i >= 0 else None

    def by_structure(role: str) -> Optional[LayoutChoice]:
        pred = STRUCTURAL.get(role)
        if pred:
            for i, sig in enumerate(sigs):
                if pred(sig):
                    return LayoutChoice(role, layouts[i], i, "structure")
        return None

    # Pass 1: direct (override / name / structure)
    for role in ALL_ROLES:
        choice = by_override(role) or by_name(role) or by_structure(role)
        if choice:
            resolved[role] = choice

    # Pass 2: fallback chains, then a guaranteed last resort
    def last_resort(role: str) -> LayoutChoice:
        for i, sig in enumerate(sigs):
            if sig["title"] >= 1:
                return LayoutChoice(role, layouts[i], i, "last-resort")
        return LayoutChoice(role, layouts[0], 0, "last-resort")

    for role in ALL_ROLES:
        if role in resolved:
            continue
        for alt in FALLBACK_CHAIN.get(role, []):
            if alt in resolved:
                a = resolved[alt]
                resolved[role] = LayoutChoice(role, a.layout, a.index, f"fallback:{alt}")
                break
        else:
            resolved[role] = last_resort(role)
        # roles resolved later in the loop may still serve earlier roles; fixed by 2nd sweep below

    # Second sweep: re-run chains for roles that fell to last-resort, now that all roles exist.
    for role in ALL_ROLES:
        if resolved[role].how == "last-resort":
            for alt in FALLBACK_CHAIN.get(role, []):
                if alt in resolved and resolved[alt].how != "last-resort":
                    a = resolved[alt]
                    resolved[role] = LayoutChoice(role, a.layout, a.index, f"fallback:{alt}")
                    break
    return resolved


def layout_placeholders_info(layout: Any) -> List[Dict[str, Any]]:
    out = []
    for ph in layout.placeholders:
        r = shape_rect(ph)
        out.append({
            "idx": ph.placeholder_format.idx,
            "name": ph.name,
            "type": ph_type_name(ph),
            "kind": ph_kind(ph),
            "leftInches": round(r.left / EMU_PER_INCH, 2) if r else None,
            "topInches": round(r.top / EMU_PER_INCH, 2) if r else None,
            "widthInches": round(r.width / EMU_PER_INCH, 2) if r else None,
            "heightInches": round(r.height / EMU_PER_INCH, 2) if r else None,
        })
    return out


def template_info() -> Dict[str, Any]:
    prs = load_template()
    roles = discover_layouts(prs)
    w, h = int(prs.slide_width), int(prs.slide_height)
    warnings = []
    for role in ("cover", "section", "content", "two_column", "comparison", "chart", "table", "closing"):
        c = roles[role]
        if c.how not in ("override", "name"):
            warnings.append(f"Role '{role}' has no layout named for it; using '{c.layout.name}' ({c.how}). "
                            f"Consider adding a layout whose name contains '{ROLE_RULES[role][0][0][0]}'.")
    return {
        "status": "success",
        "templateFile": TEMPLATE_FILENAME,
        "slideWidthEmu": w,
        "slideHeightEmu": h,
        "slideWidthInches": round(w / EMU_PER_INCH, 3),
        "slideHeightInches": round(h / EMU_PER_INCH, 3),
        "masterCount": len(prs.slide_masters),
        "layouts": [
            {"index": i, "name": l.name, "placeholders": layout_placeholders_info(l)}
            for i, l in enumerate(prs.slide_layouts)
        ],
        "roleMapping": {
            role: {"layoutIndex": c.index, "layoutName": c.layout.name, "matchedBy": c.how}
            for role, c in roles.items()
        },
        "warnings": warnings,
    }


# --------------------------------------------------------------------------- #
# Slide-kind resolution (request -> which builder)
# --------------------------------------------------------------------------- #
KIND_ALIASES = {
    "content": "content", "bullets": "content", "bullet": "content", "bulletpoints": "content",
    "titleandcontent": "content", "text": "content", "default": "content", "standard": "content",
    "cards": "cards", "card": "cards", "cardgrid": "cards", "cardslide": "cards", "features": "cards",
    "rows": "rows", "row": "rows", "rowlist": "rows", "rowslide": "rows", "list": "rows",
    "process": "process", "steps": "process", "flow": "process", "processflow": "process",
    "workflow": "process", "roadmap": "process", "timeline": "process",
    "comparison": "comparison", "compare": "comparison", "vs": "comparison", "beforeafter": "comparison",
    "proscons": "comparison", "comparisonslide": "comparison",
    "twocolumn": "two_column", "twocolumns": "two_column", "twocontent": "two_column",
    "columns": "two_column", "split": "two_column", "twocolumncontent": "two_column",
    "stats": "stats", "statistics": "stats", "stat": "stats", "kpi": "stats", "kpis": "stats",
    "metrics": "stats", "numbers": "stats",
    "chart": "chart", "charts": "chart", "graph": "chart", "chartslide": "chart",
    "table": "table", "tables": "table", "tableslide": "table",
    "section": "section", "divider": "section", "sectionheader": "section", "sectionslide": "section",
    "chapter": "section",
    "agenda": "agenda", "contents": "agenda", "toc": "agenda", "agendaslide": "agenda",
    "closing": "closing", "closingslide": "closing", "thankyou": "closing", "thanks": "closing",
    "end": "closing", "final": "closing",
    "cover": "cover", "titleslide": "cover", "coverslide": "cover",
}

# slide kind -> layout role used to create it
ROLE_FOR_KIND = {
    "cover": "cover", "agenda": "agenda", "section": "section", "content": "content",
    "two_column": "two_column", "comparison": "comparison", "chart": "chart", "table": "table",
    "closing": "closing",
    # custom-drawn layouts sit on a title-only layout (or the content layout as fallback)
    "cards": "title_only", "rows": "title_only", "process": "title_only", "stats": "title_only",
}


def resolve_kind(spec: Dict[str, Any]) -> str:
    for key in ("layout", "type", "slideType", "slide_type", "template", "style", "kind"):
        v = spec.get(key)
        if isinstance(v, str) and v.strip():
            kind = KIND_ALIASES.get(norm_key(v))
            if kind:
                return kind
    # Infer from the data present
    if pick(spec, ("chart", "chartData", "chart_data")) or (spec.get("categories") and spec.get("series")):
        return "chart"
    if pick(spec, ("table", "tableData", "table_data")) or (spec.get("headers") and spec.get("rows")):
        return "table"
    if pick(spec, ("stats", "statistics", "metrics", "kpis")):
        return "stats"
    if pick(spec, ("cards",)):
        return "cards"
    if pick(spec, ("steps", "process")):
        return "process"
    if pick(spec, ("rows",)):
        return "rows"
    if spec.get("comparison") or (spec.get("leftTitle") and spec.get("rightTitle")):
        return "comparison"
    if spec.get("columns") or any(spec.get(k) for k in ("left", "right", "leftColumn", "rightColumn")):
        return "two_column"
    return "content"


# --------------------------------------------------------------------------- #
# Drawing primitives
# --------------------------------------------------------------------------- #
@dataclass
class Para:
    text: str
    size: float = 18
    bold: bool = False
    color: Optional[MSO_THEME_COLOR] = None
    align: Optional[PP_ALIGN] = None
    bullet: bool = False
    level: int = 0
    space_after: float = 4


ACCENTS = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3,
           MSO_THEME_COLOR.ACCENT_4, MSO_THEME_COLOR.ACCENT_5, MSO_THEME_COLOR.ACCENT_6]


def auto_cols(n: int, max_cols: int) -> int:
    """Balanced column count: 4 cards with max 3 -> 2x2 (not 3+1)."""
    if n <= max_cols:
        return max(n, 1)
    return math.ceil(n / math.ceil(n / max_cols))


# --------------------------------------------------------------------------- #
# The deck builder
# --------------------------------------------------------------------------- #
class DeckBuilder:
    def __init__(self, prs: Any, roles: Dict[str, LayoutChoice]):
        self.prs = prs
        self.roles = roles
        self.sw = int(prs.slide_width)
        self.sh = int(prs.slide_height)
        self.scale = self.sw / REFERENCE_WIDTH
        self.warnings: List[str] = []

    # ---- units ----------------------------------------------------------- #
    def emu_in(self, inches: float) -> int:
        """Inches at 16:9 reference size, scaled to this template's slide width."""
        return int(inches * EMU_PER_INCH * self.scale)

    def size(self, pt: float) -> float:
        return max(8.0, pt * self.scale)

    # ---- slide creation -------------------------------------------------- #
    def new_slide(self, kind: str) -> Tuple[Any, Any]:
        choice = self.roles[ROLE_FOR_KIND[kind]]
        slide = self.prs.slides.add_slide(choice.layout)
        if KEEP_FOOTERS:
            self._clone_footer_placeholders(slide, choice.layout)
        return slide, choice.layout

    def _clone_footer_placeholders(self, slide: Any, layout: Any) -> None:
        """
        python-pptx does not copy footer / date / slide-number placeholders to new
        slides, so numbers and footer text defined on the layout would vanish.
        Copy the ones that carry content (slide number always; footer/date if text).
        """
        try:
            for ph in layout.placeholders:
                t = ph.placeholder_format.type
                if t not in (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE):
                    continue
                if t != PP_PLACEHOLDER.SLIDE_NUMBER and not ph.text_frame.text.strip():
                    continue
                new_el = copy.deepcopy(ph._element)
                new_el.nvSpPr.cNvPr.set("id", str(slide.shapes._next_shape_id))
                slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
        except Exception:  # footers are cosmetic; never fail the deck because of them
            log.exception("Could not clone footer placeholders")

    # ---- placeholder access --------------------------------------------- #
    @staticmethod
    def body_phs(slide: Any) -> List[Any]:
        return [p for p in slide.placeholders if ph_kind(p) == "body"]

    @staticmethod
    def first_ph(slide: Any, kind: str) -> Optional[Any]:
        for p in slide.placeholders:
            if ph_kind(p) == kind:
                return p
        return None

    @staticmethod
    def _lt(ph: Any) -> Tuple[int, int]:
        return (ph.left or 0, ph.top or 0)

    # ---- title ----------------------------------------------------------- #
    def set_title(self, slide: Any, text: str) -> Any:
        text = clean(text, 200) or "Untitled"
        title = slide.shapes.title
        if title is None:
            rect = Rect(int(0.06 * self.sw), int(0.04 * self.sh), int(0.88 * self.sw), int(0.12 * self.sh))
            return self.add_text(slide, rect, [Para(text, size=self.size(30), bold=True)], anchor=MSO_ANCHOR.MIDDLE)
        title.text_frame.text = text
        self._shrink_to_fit(title, [(text, 0)], "title", 32)
        return title

    def _shrink_to_fit(self, ph: Any, items: List[Tuple[str, int]], kind: str, default_pt: float) -> None:
        """Only shrinks when the estimate says the template's own size would overflow."""
        rect = shape_rect(ph) or Rect(0, 0, int(0.8 * self.sw), int(0.5 * self.sh))
        base = inherited_font_pt(ph, kind, default_pt)
        pt = estimate_font_pt(items, rect.width, rect.height, base, max(10.0, base * 0.55))
        if pt < base:
            apply_font_size(ph.text_frame, pt)
        try:
            ph.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # PowerPoint backstop
        except Exception:
            pass

    # ---- regions --------------------------------------------------------- #
    def footer_limit(self, layout: Any) -> int:
        """Top edge of the lowest-zone footer placeholder on the layout (content must stay above it)."""
        tops = []
        for ph in layout.placeholders:
            if ph_kind(ph) == "footer":
                r = shape_rect(ph)
                if r and r.top > 0.6 * self.sh:
                    tops.append(r.top)
        return (min(tops) - int(0.02 * self.sh)) if tops else int(0.93 * self.sh)

    def default_region(self, slide: Any, layout: Any) -> Rect:
        """Free area below the title and above the footer zone."""
        title = slide.shapes.title
        tr = shape_rect(title) if title is not None else None
        if tr and tr.width >= 0.5 * self.sw:
            left, width, top = tr.left, tr.width, tr.bottom + int(0.03 * self.sh)
        else:
            left, width = int(0.06 * self.sw), int(0.88 * self.sw)
            top = tr.bottom + int(0.03 * self.sh) if tr else int(0.22 * self.sh)
        bottom = max(self.footer_limit(layout), top + int(0.3 * self.sh))
        return Rect(left, top, width, max(bottom - top, int(0.3 * self.sh)))

    def claim_region(self, slide: Any, layout: Any) -> Rect:
        """
        Use the template's own content-area geometry: take the largest body
        placeholder's rectangle (and remove that placeholder because we are about
        to draw a custom visual there). Falls back to the free area under the title.
        """
        bodies = self.body_phs(slide)
        if bodies:
            ph = max(bodies, key=lambda p: (p.width or 0) * (p.height or 0))
            r = shape_rect(ph)
            if r:
                remove_shape(ph)
                return r
        return self.default_region(slide, layout)

    # ---- shapes ---------------------------------------------------------- #
    def add_text(self, slide: Any, rect: Rect, paras: List[Para], anchor: Any = MSO_ANCHOR.TOP,
                 margin_in: float = 0.08) -> Any:
        tb = slide.shapes.add_textbox(Emu(rect.left), Emu(rect.top), Emu(rect.width), Emu(rect.height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.vertical_anchor = anchor
        m = self.emu_in(margin_in)
        tf.margin_left = tf.margin_right = Emu(m)
        tf.margin_top = tf.margin_bottom = Emu(m // 2)
        first = True
        for para in paras:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            if para.align is not None:
                p.alignment = para.align
            p.space_after = Pt(para.space_after)
            run = p.add_run()
            run.text = para.text
            run.font.size = Pt(para.size)
            run.font.bold = para.bold
            if para.color is not None:
                run.font.color.theme_color = para.color
            if para.bullet:
                self._apply_bullet(p, para.level)
        return tb

    def _apply_bullet(self, p: Any, level: int = 0) -> None:
        pPr = p._p.get_or_add_pPr()
        indent = self.emu_in(0.25)
        pPr.set("marL", str(indent * (level + 1)))
        pPr.set("indent", str(-indent))
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        bu = etree.SubElement(pPr, qn("a:buChar"))
        bu.set("char", "•")

    @staticmethod
    def _no_bullet(p: Any) -> None:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        etree.SubElement(pPr, qn("a:buNone"))

    def add_box(self, slide: Any, rect: Rect, color: MSO_THEME_COLOR = MSO_THEME_COLOR.ACCENT_1,
                tint: Optional[float] = None, shape: Any = MSO_SHAPE.RECTANGLE,
                rounded: Optional[float] = None) -> Any:
        """Filled shape using THEME colours only. `tint` lightens the colour (0..1)."""
        shp = slide.shapes.add_shape(shape, Emu(rect.left), Emu(rect.top), Emu(rect.width), Emu(rect.height))
        shp.fill.solid()
        shp.fill.fore_color.theme_color = color
        if tint is not None:
            shp.fill.fore_color.brightness = tint
        shp.line.fill.background()
        try:
            shp.shadow.inherit = False
        except Exception:
            pass
        if rounded is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            shp.adjustments[0] = rounded
        return shp

    # ---- text into placeholder OR rectangle ----------------------------- #
    def write_points(self, slide: Any, target: Any, points: List[Tuple[str, int]],
                     heading: Optional[str] = None) -> Any:
        """
        Write bullets (with optional bold heading) into a placeholder if given a
        shape, or into a new text box if given a Rect.
        """
        items = ([(heading, 0)] if heading else []) + points
        if isinstance(target, Rect):
            pt = estimate_font_pt(items, target.width, target.height, self.size(20), self.size(11))
            paras: List[Para] = []
            if heading:
                paras.append(Para(heading, size=pt + 2, bold=True, space_after=6))
            for text, lvl in points:
                paras.append(Para(text, size=pt, bullet=True, level=lvl, space_after=pt * 0.4))
            return self.add_text(slide, target, paras)

        ph = target
        tf = ph.text_frame
        tf.clear()
        first = True
        if heading:
            p = tf.paragraphs[0]
            first = False
            r = p.add_run()
            r.text = heading
            r.font.bold = True
            self._no_bullet(p)
        for text, lvl in points:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = min(lvl, 8)
            p.add_run().text = text  # bullets/indent/fonts inherited from the template
        if items:
            self._shrink_to_fit(ph, items, "body", 20)
        return ph

    # ---- finishing ------------------------------------------------------- #
    @staticmethod
    def prune_empty_placeholders(slide: Any) -> None:
        """Remove unfilled text placeholders ("Click to add text" boxes)."""
        for ph in list(slide.placeholders):
            if ph_kind(ph) == "footer":
                continue
            el = ph._element
            if el.tag != qn("p:sp"):  # filled chart/table/picture frames are not p:sp
                continue
            if ph.has_text_frame and ph.text_frame.text.strip():
                continue
            remove_shape(ph)

    @staticmethod
    def add_notes(slide: Any, spec: Dict[str, Any]) -> None:
        notes = clean(pick(spec, ("notes", "speakerNotes", "speaker_notes", "speakernotes")), 5000, multiline=True)
        if notes:
            tf = slide.notes_slide.notes_text_frame
            if tf is not None:
                tf.text = notes

    def finish(self, slide: Any, spec: Dict[str, Any]) -> None:
        self.prune_empty_placeholders(slide)
        self.add_notes(slide, spec)

    # ===================================================================== #
    # Slide builders
    # ===================================================================== #
    def build_cover(self, spec: Dict[str, Any], title: str) -> None:
        slide, _ = self.new_slide("cover")
        self.set_title(slide, title)
        sub = clean(pick(spec, ("subtitle", "tagline")), 200)
        if not sub:
            pts = to_points(pick(spec, ("content", "points")))
            sub = " | ".join(t for t, _ in pts[:3])
        if sub:
            target = self.first_ph(slide, "subtitle")
            if target is None:
                bodies = self.body_phs(slide)
                target = bodies[0] if bodies else None
            if target is not None:
                target.text_frame.text = sub
            else:  # layout has no subtitle placeholder -> text box below the title
                tr = shape_rect(slide.shapes.title) if slide.shapes.title is not None else None
                top = (tr.bottom + int(0.02 * self.sh)) if tr else int(0.55 * self.sh)
                left = tr.left if tr else int(0.08 * self.sw)
                width = tr.width if tr else int(0.84 * self.sw)
                self.add_text(slide, Rect(left, top, width, int(0.1 * self.sh)), [Para(sub, size=self.size(20))])
        self.finish(slide, spec)

    def build_agenda(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("agenda")
        self.set_title(slide, title or "Agenda")
        points = to_points(pick(spec, ("content", "items", "points", "agenda")))
        bodies = self.body_phs(slide)
        if bodies:
            self.write_points(slide, max(bodies, key=lambda p: (p.width or 0) * (p.height or 0)), points)
        elif points:
            self.write_points(slide, self.default_region(slide, layout), points)
        self.finish(slide, spec)

    def build_section(self, spec: Dict[str, Any], title: str) -> None:
        slide, _ = self.new_slide("section")
        self.set_title(slide, title)
        desc = clean(pick(spec, ("subtitle", "description", "text")), 300)
        if not desc:
            pts = to_points(pick(spec, ("content", "points")))
            desc = " ".join(t for t, _ in pts[:2])
        if desc:
            target = self.first_ph(slide, "subtitle") or (self.body_phs(slide) or [None])[0]
            if target is not None:
                target.text_frame.text = desc
            else:
                tr = shape_rect(slide.shapes.title) if slide.shapes.title is not None else None
                if tr:
                    self.add_text(slide, Rect(tr.left, tr.bottom + int(0.02 * self.sh), tr.width, int(0.12 * self.sh)),
                                  [Para(desc, size=self.size(18))])
        self.finish(slide, spec)

    def build_content(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("content")
        self.set_title(slide, title)
        points = to_points(pick(spec, ("content", "points", "bullets", "items", "text")))
        bodies = self.body_phs(slide)
        if bodies:
            self.write_points(slide, max(bodies, key=lambda p: (p.width or 0) * (p.height or 0)), points)
        elif points:  # chosen layout has no body placeholder -> text box in the free area
            self.write_points(slide, self.default_region(slide, layout), points)
        self.finish(slide, spec)

    # ---- two column / comparison ---------------------------------------- #
    @staticmethod
    def get_columns(spec: Dict[str, Any]) -> List[Tuple[str, List[Tuple[str, int]]]]:
        """Return exactly two (heading, points) tuples from many accepted shapes."""

        def col(v: Any, heading: str = "") -> Tuple[str, List[Tuple[str, int]]]:
            if isinstance(v, dict):
                h = clean(pick(v, HEAD_KEYS), 120) or heading
                return h, to_points(pick(v, ("points", "items", "content", "bullets", "text", "children")))
            return heading, to_points(v)

        cols = spec.get("columns") or spec.get("comparison")
        if isinstance(cols, dict):
            cols = [cols.get("left"), cols.get("right")]
        if isinstance(cols, list) and len(cols) >= 2:
            result = [col(cols[0]), col(cols[1])]
        else:
            left = pick(spec, ("left", "leftColumn", "leftContent", "leftPoints", "before", "pros"))
            right = pick(spec, ("right", "rightColumn", "rightContent", "rightPoints", "after", "cons"))
            if left is None and right is None:
                pts = to_points(pick(spec, ("content", "points", "bullets")))
                half = math.ceil(len(pts) / 2)
                result = [("", pts[:half]), ("", pts[half:])]
            else:
                result = [col(left, clean(pick(spec, ("leftTitle", "leftHeading", "leftHeader")), 120)),
                          col(right, clean(pick(spec, ("rightTitle", "rightHeading", "rightHeader")), 120))]
        return result

    def build_two_column(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("two_column")
        self.set_title(slide, title)
        (lh, lp), (rh, rp) = self.get_columns(spec)
        bodies = sorted(self.body_phs(slide), key=self._lt)
        if len(bodies) >= 2:
            self.write_points(slide, bodies[0], lp, lh or None)
            self.write_points(slide, bodies[1], rp, rh or None)
        else:  # layout is single-column -> split its content area in two
            region = self.claim_region(slide, layout)
            a, b = region.split_h(0.5, self.emu_in(0.3))
            self.write_points(slide, a, lp, lh or None)
            self.write_points(slide, b, rp, rh or None)
        self.finish(slide, spec)

    def build_comparison(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("comparison")
        self.set_title(slide, title)
        (lh, lp), (rh, rp) = self.get_columns(spec)
        bodies = self.body_phs(slide)
        mid = self.sw / 2
        left = sorted([b for b in bodies if (b.left or 0) + (b.width or 0) / 2 < mid], key=lambda p: p.top or 0)
        right = sorted([b for b in bodies if (b.left or 0) + (b.width or 0) / 2 >= mid], key=lambda p: p.top or 0)
        if len(left) >= 2 and len(right) >= 2:
            # Typical "Comparison" layout: [heading][content] per side. Smaller box = heading.
            for pair, head, pts in ((left[:2], lh, lp), (right[:2], rh, rp)):
                hdr, body = sorted(pair, key=lambda p: p.height or 0)
                hdr.text_frame.text = head
                self._shrink_to_fit(hdr, [(head, 0)], "body", 20)
                self.write_points(slide, body, pts)
        else:
            # Template has no comparison placeholders -> two panels with an accent header bar.
            region = self.claim_region(slide, layout)
            gap = self.emu_in(0.3)
            a, b = region.split_h(0.5, gap)
            bar_h = self.emu_in(0.6)
            for idx, (rect, head, pts) in enumerate(((a, lh, lp), (b, rh, rp))):
                accent = ACCENTS[idx % len(ACCENTS)]
                self.add_box(slide, Rect(rect.left, rect.top, rect.width, rect.height), accent, tint=0.88)
                bar = self.add_box(slide, Rect(rect.left, rect.top, rect.width, bar_h), accent)
                tf = bar.text_frame
                tf.text = head or ("Option A" if idx == 0 else "Option B")
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for r in tf.paragraphs[0].runs:
                    r.font.size = Pt(self.size(20))
                    r.font.bold = True
                    r.font.color.theme_color = MSO_THEME_COLOR.LIGHT_1
                body_rect = Rect(rect.left, rect.top + bar_h, rect.width, rect.height - bar_h).inset(self.emu_in(0.12))
                self.write_points_dark(slide, body_rect, pts)
        self.finish(slide, spec)

    def write_points_dark(self, slide: Any, rect: Rect, points: List[Tuple[str, int]]) -> None:
        """Bullets in a text box with an explicit dark colour (used on light tinted panels)."""
        pt = estimate_font_pt(points, rect.width, rect.height, self.size(18), self.size(11))
        self.add_text(slide, rect, [Para(t, size=pt, bullet=True, level=l, color=MSO_THEME_COLOR.DARK_1,
                                         space_after=pt * 0.4) for t, l in points])

    # ---- cards / rows / process / stats --------------------------------- #
    def build_cards(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("cards")
        self.set_title(slide, title)
        cards = norm_items(pick(spec, ("cards", "items", "points", "content")), 8)
        region = self.claim_region(slide, layout)
        if cards:
            n = len(cards)
            cols = auto_cols(n, 3)
            rows = math.ceil(n / cols)
            gap = self.emu_in(0.25)
            cw = (region.width - gap * (cols - 1)) // cols
            ch = min((region.height - gap * (rows - 1)) // rows, self.emu_in(3.4))
            strip = self.emu_in(0.09)
            for i, card in enumerate(cards):
                r, c = divmod(i, cols)
                rect = Rect(region.left + c * (cw + gap), region.top + r * (ch + gap), cw, ch)
                accent = ACCENTS[i % len(ACCENTS)]
                self.add_box(slide, rect, accent, tint=0.88)
                self.add_box(slide, Rect(rect.left, rect.top, rect.width, strip), accent)
                inner = Rect(rect.left, rect.top + strip, rect.width, rect.height - strip).inset(self.emu_in(0.1))
                items = [(card["title"], 0), (card["text"], 0)]
                pt = estimate_font_pt([x for x in items if x[0]], inner.width, inner.height, self.size(20), self.size(10))
                paras = []
                if card["title"]:
                    paras.append(Para(card["title"], size=pt + 3, bold=True, color=MSO_THEME_COLOR.DARK_1, space_after=6))
                if card["text"]:
                    paras.append(Para(card["text"], size=pt, color=MSO_THEME_COLOR.DARK_1))
                self.add_text(slide, inner, paras)
        self.finish(slide, spec)

    def build_rows(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("rows")
        self.set_title(slide, title)
        rows = norm_items(pick(spec, ("rows", "items", "points", "content")), 7)
        region = self.claim_region(slide, layout)
        if rows:
            n = len(rows)
            gap = self.emu_in(0.15)
            rh = min((region.height - gap * (n - 1)) // n, self.emu_in(1.1))
            bar_w = self.emu_in(0.14)
            for i, row in enumerate(rows):
                rect = Rect(region.left, region.top + i * (rh + gap), region.width, rh)
                accent = ACCENTS[i % len(ACCENTS)]
                self.add_box(slide, rect, accent, tint=0.9)
                self.add_box(slide, Rect(rect.left, rect.top, bar_w, rect.height), accent)
                inner = Rect(rect.left + bar_w, rect.top, rect.width - bar_w, rect.height).inset(self.emu_in(0.12), self.emu_in(0.04))
                if row["text"]:
                    a, b = inner.split_h(0.32, self.emu_in(0.15))
                    pt = estimate_font_pt([(row["text"], 0)], b.width, b.height, self.size(20), self.size(10))
                    self.add_text(slide, a, [Para(row["title"], size=min(self.size(22), pt + 3), bold=True,
                                                  color=MSO_THEME_COLOR.DARK_1)], anchor=MSO_ANCHOR.MIDDLE)
                    self.add_text(slide, b, [Para(row["text"], size=pt, color=MSO_THEME_COLOR.DARK_1)],
                                  anchor=MSO_ANCHOR.MIDDLE)
                else:
                    self.add_text(slide, inner, [Para(row["title"], size=self.size(18), bold=True,
                                                      color=MSO_THEME_COLOR.DARK_1)], anchor=MSO_ANCHOR.MIDDLE)
        self.finish(slide, spec)

    def build_process(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("process")
        self.set_title(slide, title)
        steps = norm_items(pick(spec, ("steps", "process", "items", "points", "content")), 6)
        region = self.claim_region(slide, layout)
        if steps:
            n = len(steps)
            step_w = region.width // n
            dia = min(self.emu_in(0.85), int(step_w * 0.55))
            cy = region.top + dia // 2 + self.emu_in(0.1)
            if n > 1:  # connector first so circles sit on top of it
                line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                                  Emu(region.left + step_w // 2), Emu(cy),
                                                  Emu(region.left + step_w * (n - 1) + step_w // 2), Emu(cy))
                line.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
                line.line.width = Pt(max(2.0, 3 * self.scale))
            for i, step in enumerate(steps):
                cx = region.left + i * step_w + step_w // 2
                accent = ACCENTS[i % len(ACCENTS)]
                circle = self.add_box(slide, Rect(cx - dia // 2, cy - dia // 2, dia, dia), accent, shape=MSO_SHAPE.OVAL)
                tf = circle.text_frame
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf.text = str(i + 1)
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(self.size(24))
                    r.font.bold = True
                    r.font.color.theme_color = MSO_THEME_COLOR.LIGHT_1
                top = cy + dia // 2 + self.emu_in(0.15)
                box = Rect(cx - step_w // 2, top, step_w, max(region.bottom - top, self.emu_in(0.8))).inset(self.emu_in(0.06), 0)
                items = [x for x in ((step["title"], 0), (step["text"], 0)) if x[0]]
                pt = estimate_font_pt(items, box.width, box.height, self.size(18), self.size(10))
                paras = []
                if step["title"]:
                    paras.append(Para(step["title"], size=pt + 2, bold=True, align=PP_ALIGN.CENTER, space_after=4))
                if step["text"]:
                    paras.append(Para(step["text"], size=pt, align=PP_ALIGN.CENTER))
                self.add_text(slide, box, paras)
        self.finish(slide, spec)

    def build_stats(self, spec: Dict[str, Any], title: str) -> None:
        slide, layout = self.new_slide("stats")
        self.set_title(slide, title)
        stats = norm_stats(pick(spec, ("stats", "statistics", "metrics", "kpis", "items", "content")))
        region = self.claim_region(slide, layout)
        if stats:
            n = len(stats)
            cols = auto_cols(n, 4)
            rows = math.ceil(n / cols)
            gap = self.emu_in(0.25)
            cw = (region.width - gap * (cols - 1)) // cols
            ch = min((region.height - gap * (rows - 1)) // rows, self.emu_in(2.6))
            strip = self.emu_in(0.09)
            for i, st in enumerate(stats):
                r, c = divmod(i, cols)
                rect = Rect(region.left + c * (cw + gap), region.top + r * (ch + gap), cw, ch)
                accent = ACCENTS[i % len(ACCENTS)]
                self.add_box(slide, rect, accent, tint=0.9)
                self.add_box(slide, Rect(rect.left, rect.top, rect.width, strip), accent)
                inner = Rect(rect.left, rect.top + strip, rect.width, rect.height - strip).inset(self.emu_in(0.1))
                vsize = self.size(44 if cols <= 3 else 36)
                # shrink the big number if it is very long (e.g. "$2.5 trillion")
                vsize = min(vsize, max(self.size(18), inner.width / 12700.0 / max(len(st["value"]), 1) / 0.6))
                paras = []
                if st["value"]:
                    paras.append(Para(st["value"], size=vsize, bold=True, color=accent, align=PP_ALIGN.CENTER, space_after=6))
                if st["label"]:
                    lp = estimate_font_pt([(st["label"], 0)], inner.width, int(inner.height * 0.45), self.size(20), self.size(10))
                    paras.append(Para(st["label"], size=lp, color=MSO_THEME_COLOR.DARK_1, align=PP_ALIGN.CENTER))
                self.add_text(slide, inner, paras, anchor=MSO_ANCHOR.MIDDLE)
        self.finish(slide, spec)

    # ---- chart ------------------------------------------------------------ #
    CHART_TYPES = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "columnclustered": XL_CHART_TYPE.COLUMN_CLUSTERED, "verticalbar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "stackedcolumn": XL_CHART_TYPE.COLUMN_STACKED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
        "stackedbar": XL_CHART_TYPE.COLUMN_STACKED,
        "horizontalbar": XL_CHART_TYPE.BAR_CLUSTERED, "barh": XL_CHART_TYPE.BAR_CLUSTERED,
        "horizontal": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS, "linemarkers": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT, "donut": XL_CHART_TYPE.DOUGHNUT,
        "area": XL_CHART_TYPE.AREA, "stackedarea": XL_CHART_TYPE.AREA_STACKED,
    }

    def parse_chart(self, spec: Dict[str, Any]) -> Tuple[CategoryChartData, Any, bool]:
        raw = pick(spec, ("chart", "chartData", "chart_data"))
        c: Dict[str, Any] = raw if isinstance(raw, dict) else spec
        ctype_name = norm_key(pick(c, ("type", "chartType", "chart_type", "kind"))
                              or pick(spec, ("chartType", "chart_type")) or "column")
        ctype = self.CHART_TYPES.get(ctype_name)
        if ctype is None:
            self.warnings.append(f"Unknown chart type '{ctype_name}'; used a column chart.")
            ctype = XL_CHART_TYPE.COLUMN_CLUSTERED

        categories = pick(c, ("categories", "labels", "x", "xAxis"))
        series_raw = pick(c, ("series", "datasets", "values", "data"))
        series: List[Tuple[str, List[Any]]] = []
        if isinstance(series_raw, dict):  # {"Q1": 10, "Q2": 12}  or  {"Revenue": [1,2,3]}
            vals = list(series_raw.values())
            if vals and all(not isinstance(v, (list, tuple)) for v in vals):
                categories = categories or list(series_raw.keys())
                series = [(clean(pick(c, ("seriesName", "title", "name")) or "Series 1", 60), vals)]
            else:
                series = [(clean(k, 60), list(v)) for k, v in series_raw.items() if isinstance(v, (list, tuple))]
        elif isinstance(series_raw, list) and series_raw:
            if all(isinstance(s, dict) for s in series_raw):
                for i, s in enumerate(series_raw):
                    name = clean(pick(s, ("name", "label", "title")) or f"Series {i + 1}", 60)
                    vals = pick(s, ("values", "data", "points"))
                    if isinstance(vals, list):
                        series.append((name, vals))
                    if categories is None and isinstance(pick(s, ("categories", "labels")), list):
                        categories = pick(s, ("categories", "labels"))
            else:  # flat list of numbers
                series = [(clean(pick(c, ("seriesName", "title", "name")) or "Series 1", 60), series_raw)]
        if not isinstance(categories, list) or not categories or not series:
            raise ApiError("A chart slide needs 'categories' and at least one series with 'values'. "
                           "Example: {\"chart\": {\"type\": \"column\", \"categories\": [\"Q1\", \"Q2\"], "
                           "\"series\": [{\"name\": \"Revenue\", \"values\": [10, 12]}]}}", code="INVALID_CHART")
        categories = [clean(x, 60) or " " for x in categories[:12]]
        cd = CategoryChartData()
        cd.categories = categories
        is_pie = ctype in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT)
        for name, vals in (series[:1] if is_pie else series[:6]):
            nums = [to_num(v) for v in vals[:len(categories)]]
            nums += [None] * (len(categories) - len(nums))  # pad short series
            cd.add_series(name, nums)
        return cd, ctype, is_pie

    def style_chart(self, chart: Any, is_pie: bool, n_series: int) -> None:
        chart.has_title = False  # the slide title already says it
        chart.has_legend = is_pie or n_series > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
        chart.font.size = Pt(self.size(12))  # size only: font family stays the theme font
        plot = chart.plots[0]
        if is_pie:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.number_format = "0%"
            dl.number_format_is_linked = False
        elif n_series == 1:
            plot.has_data_labels = True
            plot.data_labels.show_value = True

    def build_chart(self, spec: Dict[str, Any], title: str) -> None:
        cd, ctype, is_pie = self.parse_chart(spec)  # validate BEFORE creating the slide
        slide, layout = self.new_slide("chart")
        self.set_title(slide, title)
        bullets = to_points(pick(spec, ("bullets", "insights", "points", "content")))
        chart_ph = next((p for p in slide.placeholders if isinstance(p, ChartPlaceholder)), None)
        bodies = sorted(self.body_phs(slide), key=self._lt)
        text_target: Any = None
        chart_rect: Optional[Rect] = None
        n_series = len(cd._series) if hasattr(cd, "_series") else 1

        if chart_ph is not None:
            # Template provides a real chart placeholder: use it.
            if bodies:
                text_target = bodies[0]
            elif bullets:
                r = shape_rect(chart_ph)
                if r:
                    ra, rb = r.split_h(0.62, self.emu_in(0.25))
                    chart_ph.left, chart_ph.top, chart_ph.width, chart_ph.height = (
                        Emu(ra.left), Emu(ra.top), Emu(ra.width), Emu(ra.height))
                    text_target = rb
            gf = chart_ph.insert_chart(ctype, cd)
            chart = gf.chart
        else:
            if bullets and len(bodies) >= 2:  # e.g. "chart + commentary" layout
                chart_rect = shape_rect(bodies[0])
                text_target = bodies[1]
                remove_shape(bodies[0])
            else:
                region = self.claim_region(slide, layout)
                if bullets:
                    chart_rect, text_target = region.split_h(0.62, self.emu_in(0.25))
                else:
                    chart_rect = region
            if chart_rect is None:
                chart_rect = self.default_region(slide, layout)
            gf = slide.shapes.add_chart(ctype, Emu(chart_rect.left), Emu(chart_rect.top),
                                        Emu(chart_rect.width), Emu(chart_rect.height), cd)
            chart = gf.chart
        self.style_chart(chart, is_pie, n_series)
        if bullets and text_target is not None:
            self.write_points(slide, text_target, bullets)
        self.finish(slide, spec)

    # ---- table ------------------------------------------------------------ #
    def parse_table(self, spec: Dict[str, Any]) -> Tuple[List[str], List[List[str]]]:
        raw = pick(spec, ("table", "tableData", "table_data"))
        t: Any = raw if raw is not None else spec
        headers: List[str] = []
        rows_raw: Any = None
        if isinstance(t, dict):
            headers = [clean(h, 80) for h in (pick(t, ("headers", "columns", "header", "head")) or [])]
            rows_raw = pick(t, ("rows", "data", "body", "values"))
        elif isinstance(t, list):
            rows_raw = t
            if t and isinstance(t[0], list) and spec.get("hasHeader", True):
                headers, rows_raw = [clean(h, 80) for h in t[0]], t[1:]
        if not isinstance(rows_raw, list) or not rows_raw:
            raise ApiError("A table slide needs 'headers' and 'rows'. Example: {\"table\": {\"headers\": "
                           "[\"Area\", \"Impact\"], \"rows\": [[\"Finance\", \"High\"]]}}", code="INVALID_TABLE")
        rows: List[List[str]] = []
        for r in rows_raw:
            if isinstance(r, dict):
                if not headers:
                    headers = [clean(k, 80) for k in r.keys()]
                rows.append([clean(r.get(h), 300) for h in headers])
            elif isinstance(r, (list, tuple)):
                rows.append([clean(v, 300) for v in r])
            else:
                rows.append([clean(r, 300)])
        ncols = max([len(headers)] + [len(r) for r in rows])
        if ncols > 8:
            self.warnings.append("Table truncated to 8 columns.")
            ncols = 8
        if len(rows) > 12:
            self.warnings.append("Table truncated to 12 rows.")
            rows = rows[:12]
        headers = (headers + [""] * ncols)[:ncols] if headers else []
        rows = [(r + [""] * ncols)[:ncols] for r in rows]
        return headers, rows

    def build_table(self, spec: Dict[str, Any], title: str) -> None:
        headers, rows = self.parse_table(spec)  # validate BEFORE creating the slide
        slide, layout = self.new_slide("table")
        self.set_title(slide, title)
        nrows = len(rows) + (1 if headers else 0)
        ncols = len(rows[0])
        table_ph = next((p for p in slide.placeholders if isinstance(p, TablePlaceholder)), None)
        if table_ph is not None:
            gf = table_ph.insert_table(nrows, ncols)
            tbl = gf.table
        else:
            region = self.claim_region(slide, layout)
            row_h = min(region.height // nrows, self.emu_in(0.7))
            gf = slide.shapes.add_table(nrows, ncols, Emu(region.left), Emu(region.top),
                                        Emu(region.width), Emu(row_h * nrows))
            tbl = gf.table
            for r in tbl.rows:
                r.height = Emu(row_h)
        tbl.first_row = bool(headers)  # table style (theme accent) formats the header row
        font_pt = self.size(20 if nrows <= 6 else 16 if nrows <= 9 else 12)
        data = ([headers] if headers else []) + rows
        for ri, row in enumerate(data):
            for ci, text in enumerate(row):
                cell = tbl.cell(ri, ci)
                cell.text = text
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(font_pt)
        self.finish(slide, spec)

    # ---- closing ---------------------------------------------------------- #
    def build_closing(self, spec: Dict[str, Any], title: str) -> None:
        slide, _ = self.new_slide("closing")
        self.set_title(slide, title or "Thank You")
        msg = clean(pick(spec, ("subtitle", "message", "text")), 300)
        if not msg:
            pts = to_points(pick(spec, ("content", "points")))
            msg = " | ".join(t for t, _ in pts[:3])
        if msg:
            target = self.first_ph(slide, "subtitle") or (self.body_phs(slide) or [None])[0]
            if target is not None:
                target.text_frame.text = msg
            else:
                tr = shape_rect(slide.shapes.title) if slide.shapes.title is not None else None
                if tr:
                    self.add_text(slide, Rect(tr.left, tr.bottom + int(0.02 * self.sh), tr.width, int(0.1 * self.sh)),
                                  [Para(msg, size=self.size(20))])
        self.finish(slide, spec)

    # ---- dispatch --------------------------------------------------------- #
    def build(self, kind: str, spec: Dict[str, Any], title: str) -> None:
        getattr(self, f"build_{kind}")(spec, title)


# --------------------------------------------------------------------------- #
# Request handling
# --------------------------------------------------------------------------- #
def parse_payload(payload: Any) -> Tuple[str, List[Dict[str, Any]], Dict[str, Any]]:
    if not isinstance(payload, dict):
        raise ApiError("Request body must be a JSON object.", code="INVALID_JSON")
    if "slides" not in payload and isinstance(payload.get("body"), dict):  # tolerate wrapped bodies
        payload = payload["body"]
    title = clean(pick(payload, ("presentationTitle", "title", "presentation_title")), 200)
    if not title:
        raise ApiError("'presentationTitle' is required.", code="MISSING_TITLE")
    slides = payload.get("slides")
    if isinstance(slides, str):  # Copilot Studio sometimes sends arrays as JSON strings
        try:
            slides = json.loads(slides)
        except ValueError:
            raise ApiError("'slides' must be an array of slide objects.", code="INVALID_SLIDES")
    if not isinstance(slides, list) or not slides:
        raise ApiError("'slides' must be a non-empty array.", code="INVALID_SLIDES")
    if len(slides) > MAX_SLIDES:
        raise ApiError(f"Too many slides ({len(slides)}). The maximum is {MAX_SLIDES}.", code="TOO_MANY_SLIDES")
    cleaned: List[Dict[str, Any]] = []
    for i, s in enumerate(slides, 1):
        if isinstance(s, str):
            s = {"title": s}
        if not isinstance(s, dict):
            raise ApiError(f"Slide {i} must be an object with a 'title'.", code="INVALID_SLIDES")
        cleaned.append(s)
    return title, cleaned, payload


def truthy(v: Any, default: bool) -> bool:
    if v is None:
        return default
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "yes", "y")
    return bool(v)


def generate_presentation(payload: Any) -> Tuple[str, int, List[str]]:
    title, specs, payload = parse_payload(payload)
    prs = load_template()                 # raises TemplateError on a bad template
    roles = discover_layouts(prs)
    builder = DeckBuilder(prs, roles)

    plan: List[Tuple[str, Dict[str, Any], str]] = []
    for spec in specs:
        plan.append((resolve_kind(spec), spec, clean(pick(spec, ("title", "heading", "name")), 200)))

    # Cover: added automatically unless the caller supplied a cover slide first.
    if truthy(payload.get("includeCover"), True) and not (plan and plan[0][0] == "cover"):
        plan.insert(0, ("cover", {"subtitle": payload.get("subtitle")}, title))

    # Agenda: automatic only if requested; an explicit agenda slide without items is auto-filled.
    sections = [t for k, _, t in plan if k == "section" and t]
    topics = [t for k, _, t in plan if k not in ("cover", "agenda", "closing", "section") and t]
    agenda_items = (sections if len(sections) >= 2 else topics)[:8]
    if truthy(payload.get("includeAgenda"), False) and not any(k == "agenda" for k, _, _ in plan):
        plan.insert(1 if plan and plan[0][0] == "cover" else 0, ("agenda", {}, "Agenda"))
    plan = [(k, dict(s, content=agenda_items) if k == "agenda" and not pick(s, ("content", "items", "points", "agenda"))
             else s, t) for k, s, t in plan]

    # Closing: automatic unless the caller supplied one last.
    if truthy(payload.get("includeClosing"), True) and not (plan and plan[-1][0] == "closing"):
        plan.append(("closing", {"subtitle": payload.get("closingMessage") or "Questions & Discussion"},
                     clean(payload.get("closingTitle"), 100) or "Thank You"))

    for idx, (kind, spec, slide_title) in enumerate(plan, 1):
        try:
            builder.build(kind, spec, slide_title)
        except ApiError as exc:
            exc.message = f"Slide {idx} ('{slide_title}'): {exc.message}"
            raise
        except Exception as exc:
            log.exception("Failed building slide %s (%s)", idx, kind)
            raise ApiError(f"Slide {idx} ('{slide_title}') could not be built as a '{kind}' slide. "
                           "Check its fields and try again.", status=422, code="SLIDE_BUILD_FAILED",
                           details=str(exc)[:200]) from exc

    slug = re.sub(r"[^A-Za-z0-9]+", "-", title).strip("-")[:50] or "presentation"
    filename = f"{slug}-{uuid.uuid4().hex[:8]}.pptx"
    prs.save(str(OUTPUT_DIR / filename))
    return filename, len(prs.slides), builder.warnings


def cleanup_old_files() -> None:
    cutoff = time.time() - FILE_TTL_SECONDS
    for f in OUTPUT_DIR.glob("*.pptx"):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def base_url() -> str:
    if PUBLIC_BASE_URL:
        return PUBLIC_BASE_URL
    render_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")  # set automatically by Render
    if render_url:
        return render_url
    return request.url_root.rstrip("/")


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
@app.get("/")
@app.get("/health")
def health():
    try:
        validate_template_file()
        tmpl = "ok"
    except TemplateError as exc:
        tmpl = exc.message
    return jsonify({"status": "ok", "service": "ppt-generator", "template": tmpl, "templateFile": TEMPLATE_FILENAME})


@app.get("/template-info")
def get_template_info():
    return jsonify(template_info())


@app.post("/createppt")
def create_ppt():
    payload = request.get_json(silent=True)
    if payload is None:
        raise ApiError("Request body must be valid JSON with Content-Type: application/json.", code="INVALID_JSON")
    cleanup_old_files()
    filename, count, warnings = generate_presentation(payload)
    body: Dict[str, Any] = {
        "status": "success",
        "fileName": filename,
        "downloadUrl": f"{base_url()}/download/{filename}",
        "slideCount": count,
    }
    if warnings:  # only present when something was adjusted
        body["warnings"] = warnings
    log.info("Generated %s (%s slides)", filename, count)
    return jsonify(body), 200


@app.get("/download/<path:filename>")
def download(filename: str):
    if not re.fullmatch(r"[A-Za-z0-9._-]+\.pptx", filename):
        raise ApiError("Invalid file name.", status=400, code="INVALID_FILENAME")
    if not (OUTPUT_DIR / filename).is_file():
        raise ApiError("File not found. Generated files expire after 24 hours; please generate it again.",
                       status=404, code="FILE_NOT_FOUND")
    return send_from_directory(
        str(OUTPUT_DIR), filename, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


# --------------------------------------------------------------------------- #
# Error handlers: always JSON, never a stack trace
# --------------------------------------------------------------------------- #
@app.errorhandler(ApiError)
def handle_api_error(err: ApiError):
    return error_response(err)


@app.errorhandler(HTTPException)
def handle_http_error(err: HTTPException):
    code = {404: "NOT_FOUND", 405: "METHOD_NOT_ALLOWED", 413: "PAYLOAD_TOO_LARGE"}.get(err.code or 500, "HTTP_ERROR")
    return jsonify({"status": "error", "errorCode": code, "message": err.description}), err.code or 500


@app.errorhandler(Exception)
def handle_unexpected(err: Exception):
    log.exception("Unhandled error")
    return jsonify({"status": "error", "errorCode": "INTERNAL_ERROR",
                    "message": "Something went wrong while generating the presentation. Please try again."}), 500


if __name__ == "__main__":  # local development only; Render uses gunicorn
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
