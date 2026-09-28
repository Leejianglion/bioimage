"""Read a .pptx deck: every picture and concentration label with its absolute
position on the slide, plus chart / table / text content.

Mapping rule (the deck's own convention): the concentration label sits on top
and the picture directly below it is that concentration's confocal image.
A picture that does not satisfy the rule unambiguously gets a status other
than "ok" and is never silently assigned a concentration.
"""
from __future__ import annotations

import hashlib
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field

from lxml import etree
from pptx import Presentation
from pptx.shapes.group import GroupShape

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}
R_EMBED = "{%s}embed" % NS["r"]


# ---------------------------------------------------------------- geometry
@dataclass(frozen=True)
class Rect:
    left: float
    top: float
    width: float
    height: float

    @property
    def right(self):
        return self.left + self.width

    @property
    def bottom(self):
        return self.top + self.height

    @property
    def cx(self):
        return self.left + self.width / 2

    @property
    def cy(self):
        return self.top + self.height / 2


# (sx, tx, sy, ty): slide_x = sx * x + tx, slide_y = sy * y + ty
IDENTITY = (1.0, 0.0, 1.0, 0.0)


def _to_slide(tf, left, top, width, height):
    sx, tx, sy, ty = tf
    return Rect(sx * left + tx, sy * top + ty, sx * width, sy * height)


def _shape_rect(shape, tf):
    if None in (shape.left, shape.top, shape.width, shape.height):
        return None
    return _to_slide(tf, shape.left, shape.top, shape.width, shape.height)


def _group_transform(group, tf):
    """Compose the parent transform with a group's child-coordinate mapping."""
    xfrm = group._element.find("p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return tf

    def pair(tag, a, b):
        el = xfrm.find(tag, NS)
        return (int(el.get(a)), int(el.get(b))) if el is not None else (0, 0)

    ox, oy = pair("a:off", "x", "y")
    ex, ey = pair("a:ext", "cx", "cy")
    cox, coy = pair("a:chOff", "x", "y")
    cex, cey = pair("a:chExt", "cx", "cy")
    gx = ex / cex if cex else 1.0
    gy = ey / cey if cey else 1.0
    sx, tx, sy, ty = tf
    return (sx * gx, sx * (ox - cox * gx) + tx, sy * gy, sy * (oy - coy * gy) + ty)


def _walk(shapes, tf=IDENTITY):
    for shape in shapes:
        if isinstance(shape, GroupShape):
            yield from _walk(shape.shapes, _group_transform(shape, tf))
        else:
            yield shape, tf


# ------------------------------------------------------ concentration labels
@dataclass(frozen=True)
class Conc:
    value: float  # expressed in `unit`
    unit: str | None  # "ng/mL", "nM", or None (bare number / control word)
    raw: str
    kind: str  # "unit" | "bare-number" | "control-word"

    def __str__(self):
        if self.kind == "control-word":
            return f"0 ({self.raw})"
        return f"{self.value:g} {self.unit or '(no unit)'}"


_NUM = r"(\d+(?:\.\d+)?(?:\s*[x×*]\s*10\s*\^?\s*-?\d+|e-?\d+)?)"
_MASS = re.compile(_NUM + r"\s*(pg|ng|ug|mg)\s*(?:/\s*ml|ml\s*-\s*1|ml\b)", re.I)
_MOLAR = re.compile(_NUM + r"\s*([pnum]?)M(?![A-Za-z])")  # case-sensitive: "nm"/"um" are lengths
_MASS_TO_NG_ML = {"pg": 1e-3, "ng": 1.0, "ug": 1e3, "mg": 1e6}
_MOLAR_TO_NM = {"p": 1e-3, "n": 1.0, "u": 1e3, "m": 1e6, "": 1e9}
_CONTROL_WORDS = {"control", "ctrl", "blank", "对照", "空白", "空白对照", "对照组", "空白组", "阴性对照"}


def _normalize(text):
    t = unicodedata.normalize("NFKC", text)
    for a, b in (("μ", "u"), ("µ", "u"), ("−", "-"), ("–", "-"), ("·", " "), (" ", " ")):
        t = t.replace(a, b)
    return t


def _to_float(s):
    s = s.replace(" ", "")
    m = re.fullmatch(r"(\d+(?:\.\d+)?)[x×*]10\^?(-?\d+)", s)
    if m:
        return float(m.group(1)) * 10 ** int(m.group(2))
    return float(s)


def parse_concentrations(text):
    """All concentrations written in `text`, left to right."""
    t = _normalize(text)
    found = []
    for m in _MASS.finditer(t):
        value = _to_float(m.group(1)) * _MASS_TO_NG_ML[m.group(2).lower()]
        found.append((m.start(), Conc(value, "ng/mL", m.group(0).strip(), "unit")))
    for m in _MOLAR.finditer(t):
        value = _to_float(m.group(1)) * _MOLAR_TO_NM[m.group(2)]
        found.append((m.start(), Conc(value, "nM", m.group(0).strip(), "unit")))
    if found:
        return [c for _, c in sorted(found, key=lambda x: x[0])]
    s = t.strip().lower()
    if re.fullmatch(r"\d+(?:\.\d+)?", s):
        return [Conc(float(s), None, s, "bare-number")]
    if s in _CONTROL_WORDS:
        return [Conc(0.0, None, t.strip(), "control-word")]
    return []


# ------------------------------------------------------------ deck content
@dataclass
class Picture:
    slide: int
    order: int
    shape_name: str
    rect: Rect | None
    blob: bytes
    ext: str
    crop: tuple  # (left, top, right, bottom) fractions cut away in PowerPoint
    effects: list  # PowerPoint picture adjustments: displayed pixels != embedded pixels
    rotation: float
    flip: str
    cstate: str | None  # PowerPoint compression state, e.g. "email" / "screen"

    @property
    def key(self):
        return f"s{self.slide:02d}_p{self.order:02d}"

    @property
    def sha1(self):
        return hashlib.sha1(self.blob).hexdigest()


@dataclass
class Label:
    slide: int
    text: str
    rect: Rect | None
    concs: list


@dataclass
class Slide:
    number: int
    pictures: list = field(default_factory=list)
    labels: list = field(default_factory=list)
    texts: list = field(default_factory=list)  # (shape name, text)
    tables: list = field(default_factory=list)  # list of rows of cell text
    charts: list = field(default_factory=list)  # dicts, see _chart_data
    other: list = field(default_factory=list)  # shapes we could not read (SmartArt, OLE, ...)


@dataclass
class Deck:
    path: str
    width: int
    height: int
    slides: list


def _fmt_el(el):
    attrs = ", ".join(f"{k}={v}" for k, v in el.attrib.items())
    return f"{etree.QName(el).localname}({attrs})"


def _blip_effects(blip):
    out = []
    for child in blip:
        name = etree.QName(child).localname
        if name == "extLst":
            for layer in child.iter("{%s}imgLayer" % NS["a14"]):
                out.append("imgLayer(艺术效果：PPT 显示的是处理过的副本)")
                for eff in layer.iter("{%s}imgEffect" % NS["a14"]):
                    out.extend(_fmt_el(e) for e in eff)
            continue
        out.append(_fmt_el(child))
    return out


def _picture(shape, tf, slide_no, order):
    el = shape._element
    tag = etree.QName(el).localname
    if tag == "pic":
        blip_fill = el.find("p:blipFill", NS)
    elif tag == "sp":  # autoshape filled with a picture
        blip_fill = el.find("p:spPr/a:blipFill", NS)
    else:
        return None
    if blip_fill is None:
        return None
    blip = blip_fill.find("a:blip", NS)
    if blip is None or blip.get(R_EMBED) is None:
        return None
    part = shape.part.related_part(blip.get(R_EMBED))
    src = blip_fill.find("a:srcRect", NS)
    crop = tuple(int(src.get(k, 0)) / 100000 if src is not None else 0.0 for k in "ltrb")
    xfrm = el.find(".//a:xfrm", NS)
    flip = ""
    if xfrm is not None:
        flip = ("H" if xfrm.get("flipH") == "1" else "") + ("V" if xfrm.get("flipV") == "1" else "")
    return Picture(
        slide=slide_no,
        order=order,
        shape_name=shape.name,
        rect=_shape_rect(shape, tf),
        blob=part.blob,
        ext=str(part.partname).rsplit(".", 1)[-1].lower(),
        crop=crop,
        effects=_blip_effects(blip),
        rotation=float(getattr(shape, "rotation", 0.0) or 0.0),
        flip=flip,
        cstate=blip.get("cstate"),
    )


def _table_cells(shape, tf):
    """(text, slide rect) of every table cell; rows can render taller than stored."""
    table = shape.table
    base = _shape_rect(shape, tf)
    sx, _, sy, _ = tf
    col_x = [0]
    for col in table.columns:
        col_x.append(col_x[-1] + col.width)
    row_y = [0]
    for row in table.rows:
        row_y.append(row_y[-1] + row.height)
    cells, grid = [], []
    for r, row in enumerate(table.rows):
        grid_row = []
        for c, cell in enumerate(row.cells):
            grid_row.append(cell.text)
            if cell.is_spanned or base is None:
                continue
            w = col_x[min(c + cell.span_width, len(col_x) - 1)] - col_x[c]
            h = row_y[min(r + cell.span_height, len(row_y) - 1)] - row_y[r]
            cells.append((cell.text, Rect(base.left + sx * col_x[c], base.top + sy * row_y[r], sx * w, sy * h)))
        grid.append(grid_row)
    return grid, cells


def _pts(el):
    if el is None:
        return []
    pts = {}
    for pt in el.iter("{%s}pt" % NS["c"]):
        v = pt.find("c:v", NS)
        pts[int(pt.get("idx"))] = None if v is None else v.text
    count = el.find(".//c:ptCount", NS)
    n = int(count.get("val")) if count is not None else (max(pts) + 1 if pts else 0)
    return [pts.get(i) for i in range(n)]


def _num(values):
    out = []
    for v in values:
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            out.append(None)
    return out


def _text_of(el):
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter("{%s}t" % NS["a"]))


def _chart_data(shape, slide_no):
    chart_part = shape.chart.part
    root = etree.fromstring(chart_part.blob)
    plot_area = root.find(".//c:plotArea", NS)
    kinds = [etree.QName(e).localname for e in plot_area if etree.QName(e).localname.endswith("Chart")]
    axes = {}
    for ax in plot_area:
        name = etree.QName(ax).localname
        if name.endswith("Ax"):
            axes[name + ":" + ax.find("c:axId", NS).get("val")] = _text_of(ax.find("c:title", NS))
    series = []
    for ser in plot_area.iter("{%s}ser" % NS["c"]):
        tx = ser.find("c:tx", NS)
        name = "".join(v.text or "" for v in tx.iter("{%s}v" % NS["c"])) if tx is not None else ""
        if not name and tx is not None:
            name = _text_of(tx)
        errbars = []
        for eb in ser.findall("c:errBars", NS):

            def val(tag, eb=eb):
                e = eb.find(tag, NS)
                return None if e is None else e.get("val")

            errbars.append({
                "dir": val("c:errDir"), "type": val("c:errBarType"), "valtype": val("c:errValType"),
                "val": val("c:val"),
                "plus": _num(_pts(eb.find("c:plus", NS))), "minus": _num(_pts(eb.find("c:minus", NS))),
            })
        series.append({
            "name": name,
            "categories": _pts(ser.find("c:cat", NS)),
            "values": _num(_pts(ser.find("c:val", NS))),
            "x": _num(_pts(ser.find("c:xVal", NS))),
            "y": _num(_pts(ser.find("c:yVal", NS))),
            "errbars": errbars,
        })
    workbook = None
    wb = getattr(chart_part, "chart_workbook", None)
    if wb is not None and wb.xlsx_part is not None:
        workbook = wb.xlsx_part.blob
    return {
        "slide": slide_no, "shape": shape.name, "types": kinds,
        "title": _text_of(root.find(".//c:chart/c:title", NS)), "axes": axes,
        "series": series, "workbook": workbook,
    }


def read_deck(path):
    prs = Presentation(path)
    slides = []
    for number, sl in enumerate(prs.slides, start=1):
        slide = Slide(number)
        for shape, tf in _walk(sl.shapes):
            pic = _picture(shape, tf, number, len(slide.pictures) + 1)
            if pic is not None:
                slide.pictures.append(pic)
            if getattr(shape, "has_chart", False):
                slide.charts.append(_chart_data(shape, number))
            elif getattr(shape, "has_table", False):
                grid, cells = _table_cells(shape, tf)
                slide.tables.append(grid)
                for text, rect in cells:
                    concs = parse_concentrations(text)
                    if concs:
                        slide.labels.append(Label(number, text, rect, concs))
            elif getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                text = shape.text_frame.text
                slide.texts.append((shape.name, text))
                concs = parse_concentrations(text)
                if concs:
                    slide.labels.append(Label(number, text.strip(), _shape_rect(shape, tf), concs))
            elif pic is None and etree.QName(shape._element).localname == "graphicFrame":
                slide.other.append(shape.name)
        slides.append(slide)
    return Deck(str(path), prs.slide_width, prs.slide_height, slides)


# ----------------------------------------------------------------- mapping
@dataclass
class Assignment:
    picture: Picture
    conc: Conc | None
    label: str
    status: str  # ok | by-order | check | ambiguous | conflict | unlabeled
    note: str = ""

    @property
    def usable(self):
        return self.conc is not None and self.status in ("ok", "by-order", "check")


def _below(label, pics):
    """Pictures whose top edge is under the label's vertical centre."""
    return [p for p in pics if p.rect.top >= label.rect.cy - 0.05 * p.rect.height]


def _nearest_row(pics):
    top = min(p.rect.top for p in pics)
    return sorted((p for p in pics if p.rect.top - top < 0.5 * p.rect.height), key=lambda p: p.rect.left)


def map_slide(slide):
    """Assign concentrations to the slide's pictures; returns (assignments, unused labels)."""
    pics = [p for p in slide.pictures if p.rect is not None]
    if not slide.pictures:
        return [], []  # numbers on a text/chart-only slide are not image labels
    claims = defaultdict(list)  # picture key -> [(conc, label text, status, note)]
    unused = []
    for lab in slide.labels:
        if lab.rect is None:
            unused.append(lab)
            continue
        below = _below(lab, pics)
        if len(lab.concs) == 1:
            cand = [p for p in below
                    if p.rect.left - 0.05 * p.rect.width <= lab.rect.cx <= p.rect.right + 0.05 * p.rect.width]
            if not cand:
                unused.append(lab)
                continue
            row = _nearest_row(cand)
            if len(row) > 1:
                strict = [p for p in row if p.rect.left <= lab.rect.cx <= p.rect.right]
                row = strict if len(strict) == 1 else row
            if len(row) > 1:
                for p in row:
                    claims[p.key].append((None, lab.text, "ambiguous",
                                          f"标签“{lab.text}”正下方有 {len(row)} 张并排的图，无法确定是哪一张"))
                continue
            pic, conc = row[0], lab.concs[0]
            status, note = "ok", ""
            if conc.kind != "unit":
                status, note = "check", ("标签只有数字、没有单位" if conc.kind == "bare-number"
                                         else "标签是对照/空白字样，按 0 处理")
            if pic.rect.top - lab.rect.bottom > pic.rect.height:
                status, note = "check", "标签离图片较远（间距大于一张图的高度）"
            claims[pic.key].append((conc, lab.text, status, note))
        else:
            k = len(lab.concs)
            if below:
                half_w = sorted(p.rect.width for p in below)[len(below) // 2] / 2
                below = [p for p in below if lab.rect.left - half_w <= p.rect.cx <= lab.rect.right + half_w]
            if not below:
                unused.append(lab)
                continue
            row = _nearest_row(below)
            if len(row) != k:
                for p in row:
                    claims[p.key].append((None, lab.text, "ambiguous",
                                          f"一个文本框写了 {k} 个浓度，但下方这一行有 {len(row)} 张图，数量对不上"))
                continue
            for p, conc in zip(row, lab.concs):
                claims[p.key].append((conc, lab.text, "by-order",
                                      f"一个文本框写了 {k} 个浓度，按从左到右的顺序对应"))

    out = []
    for pic in slide.pictures:
        cl = claims.get(pic.key, [])
        labels = "; ".join(dict.fromkeys(c[1] for c in cl))
        if pic.rect is None:
            out.append(Assignment(pic, None, "", "unlabeled", "图片没有位置信息"))
        elif not cl:
            out.append(Assignment(pic, None, "", "unlabeled", "正上方没有浓度标签"))
        elif any(c[2] == "ambiguous" for c in cl):
            out.append(Assignment(pic, None, labels, "ambiguous", "; ".join(dict.fromkeys(c[3] for c in cl))))
        elif len({(c[0].value, c[0].unit) for c in cl}) > 1:
            out.append(Assignment(pic, None, labels, "conflict", "有多个不同的浓度标签都指向这张图"))
        else:
            rank = {"ok": 0, "by-order": 1, "check": 2}
            conc, _, status, note = max(cl, key=lambda c: rank[c[2]])
            if pic.rotation or pic.flip:
                status, note = "check", (note + "; " if note else "") + "图片在 PPT 中被旋转/翻转"
            out.append(Assignment(pic, conc, labels, status, note))
    return out, unused


def map_deck(deck, slides=None):
    assignments, unused = [], []
    for slide in deck.slides:
        if slides is not None and slide.number not in slides:
            continue
        a, u = map_slide(slide)
        assignments.extend(a)
        unused.extend(u)
    return assignments, unused
