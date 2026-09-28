"""Synthetic decks with known answers: label layout, PowerPoint traps and
images whose true FI follows F = 8 + 180*C/(C + 15) on top of background."""
import io

import numpy as np
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

REF_CONC = [10, 20, 50, 100, 200]
REF_NOISE = [0.6, -0.4, 0.5, -0.7, 0.3]


def ref_curve(c):
    return 20 + 100 * np.asarray(c, float) / (np.asarray(c, float) + 15)


def cell_png(conc, seed, n_cells=30, gain=1.0, size=256, scalebar=False):
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:size, 0:size]
    img = np.full((size, size), 5.0)
    peak = gain * (8 + 180 * conc / (conc + 15))
    for _ in range(n_cells):
        cx, cy = rng.uniform(24, size - 24, 2)
        s = rng.uniform(6, 10)
        img += peak * np.exp(-((xx - cx) ** 2 + (yy - cy) ** 2) / (2 * s * s))
    img += rng.normal(0, 2, img.shape)
    rgb = np.zeros((size, size, 3), np.uint8)
    rgb[..., 1] = np.clip(np.round(img), 0, 255).astype(np.uint8)
    if scalebar:
        rgb[215:222, 170:230] = 255
    buf = io.BytesIO()
    Image.fromarray(rgb).save(buf, format="PNG")
    return buf.getvalue()


def _label(slide, text, left, top, width=Inches(2.2)):
    tb = slide.shapes.add_textbox(left, top, width, Inches(0.45))
    tb.text_frame.text = text
    return tb


def _pic(slide_or_group, png, left, top, width=Inches(2.2), height=Inches(2.2)):
    return slide_or_group.shapes.add_picture(io.BytesIO(png), left, top, width, height)


def _add_lum(pic, bright, contrast):
    blip = pic._element.find(qn("p:blipFill")).find(qn("a:blip"))
    lum = etree.Element(qn("a:lum"))
    lum.set("bright", str(bright))
    lum.set("contrast", str(contrast))
    blip.insert(0, lum)


def _col(i):
    return Inches(0.4 + i * 2.55)


def _new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs, prs.slide_layouts[6]


def _reference_slides(prs, blank):
    s1 = prs.slides.add_slide(blank)
    _label(s1, "FI analysis (synthetic)", Inches(0.5), Inches(0.3), Inches(6))
    table = s1.shapes.add_table(2, 6, Inches(0.5), Inches(2), Inches(9), Inches(1)).table
    table.cell(0, 0).text, table.cell(1, 0).text = "浓度 (ng/mL)", "FI"
    for j, (c, e) in enumerate(zip(REF_CONC, REF_NOISE), start=1):
        table.cell(0, j).text, table.cell(1, j).text = str(c), f"{ref_curve(c) + e:.2f}"
    s2 = prs.slides.add_slide(blank)
    _label(s2, "≥10 ng/mL 结果", Inches(0.5), Inches(0.3), Inches(6))
    data = XyChartData()
    ser = data.add_series("FI")
    for c, e in zip(REF_CONC, REF_NOISE):
        ser.add_data_point(c, float(ref_curve(c) + e))
    s2.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(1), Inches(1.2), Inches(8), Inches(5), data)


def build_clean_deck(path):
    prs, blank = _new_deck()
    _reference_slides(prs, blank)
    # slide 3: >=10 ng/mL images, used to bridge onto the reference scale
    s3 = prs.slides.add_slide(blank)
    for i, c in enumerate([10, 50, 100]):
        _label(s3, f"{c} ng/mL", _col(i), Inches(0.9))
        _pic(s3, cell_png(c, seed=100 + i), _col(i), Inches(1.45))
    # slide 4: low concentrations; one picture in a scaled group, one with a
    # scale bar, one brightened in PowerPoint
    s4 = prs.slides.add_slide(blank)
    for i, (text, c) in enumerate([("Control", 0), ("0.1 ng/mL", 0.1), ("0.5 ng/mL", 0.5),
                                   ("1 ng/mL", 1), ("5 ng/mL", 5)]):
        _label(s4, text, _col(i), Inches(0.9))
        png = cell_png(c, seed=200 + i, scalebar=(c == 1))
        if c == 0.5:
            grp = s4.shapes.add_group_shape()
            _pic(grp, png, Inches(0.2), Inches(4.0), Inches(4.4), Inches(4.4))
            xfrm = grp._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
            xfrm.find(qn("a:off")).set("x", str(_col(i)))
            xfrm.find(qn("a:off")).set("y", str(Inches(1.45)))
            xfrm.find(qn("a:ext")).set("cx", str(Inches(2.2)))
            xfrm.find(qn("a:ext")).set("cy", str(Inches(2.2)))
        else:
            pic = _pic(s4, png, _col(i), Inches(1.45))
            if c == 5:
                _add_lum(pic, 40000, 20000)
    # slide 5: one text box listing three concentrations over three pictures;
    # the 5 ng/mL picture is the same file as on slide 4
    s5 = prs.slides.add_slide(blank)
    _label(s5, "0.1 ng/ml          1 ng/ml          5 ng/ml", _col(0), Inches(0.9), Inches(7.3))
    _pic(s5, cell_png(0.1, seed=300), _col(0), Inches(1.45))
    _pic(s5, cell_png(1, seed=301), _col(1), Inches(1.45))
    _pic(s5, cell_png(5, seed=204), _col(2), Inches(1.45))
    # slide 6: a label over two touching pictures, an unlabelled picture and a
    # picture claimed by two different labels
    s6 = prs.slides.add_slide(blank)
    _label(s6, "20 ng/mL", Inches(1.5), Inches(0.9))
    _pic(s6, cell_png(20, seed=400), Inches(0.4), Inches(1.45))
    _pic(s6, cell_png(20, seed=401), Inches(2.6), Inches(1.45))
    _pic(s6, cell_png(2, seed=402), Inches(5.5), Inches(1.45))
    _label(s6, "2 ng/mL", Inches(8.5), Inches(0.4))
    _label(s6, "3 ng/mL", Inches(8.5), Inches(0.9))
    _pic(s6, cell_png(3, seed=403), Inches(8.5), Inches(1.45))
    prs.save(path)
    return path


def build_problem_deck(path):
    """The inversions seen in practice, each produced by a known cause."""
    prs, blank = _new_deck()
    _reference_slides(prs, blank)
    s3 = prs.slides.add_slide(blank)
    specs = [
        ("0.5 ng/mL", 0.5, dict(seed=500)),
        ("1 ng/mL", 1, dict(seed=501)),
        ("5 ng/mL", 5, dict(seed=502, n_cells=5, gain=1.6)),  # few, bright cells
        ("50 ng/mL", 50, dict(seed=503)),
    ]
    for i, (text, c, kw) in enumerate(specs):
        _label(s3, text, _col(i), Inches(0.9))
        _pic(s3, cell_png(c, **kw), _col(i), Inches(1.45))
    s4 = prs.slides.add_slide(blank)
    _label(s4, "100 ng/mL", _col(0), Inches(0.9))
    pic = _pic(s4, cell_png(100, seed=504, gain=0.5), _col(0), Inches(1.45))  # low gain...
    _add_lum(pic, 30000, 40000)  # ...made to look bright in PowerPoint
    _label(s4, "200 ng/mL", _col(1), Inches(0.9))
    _pic(s4, cell_png(200, seed=505, gain=2.5), _col(1), Inches(1.45))  # saturated
    prs.save(path)
    return path
