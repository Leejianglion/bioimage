import csv

import numpy as np
import pytest
import tifffile
from pptx.util import Inches

from fi_analysis import curve
from fi_analysis.deck import map_deck, parse_concentrations, read_deck
from fi_analysis.pipeline import reference_candidates, run
from fi_analysis.report import render_slides, write_all
from synthetic import REF_CONC, REF_NOISE, ref_curve


def _one(text):
    concs = parse_concentrations(text)
    assert len(concs) == 1, (text, concs)
    return concs[0]


@pytest.mark.parametrize("text, value, unit", [
    ("0.1 ng/mL", 0.1, "ng/mL"),
    ("10ng/ml", 10, "ng/mL"),
    ("100 pg/mL", 0.1, "ng/mL"),
    ("1 μg/mL", 1000, "ng/mL"),
    ("1 µg/ml", 1000, "ng/mL"),
    ("5 ng mL-1", 5, "ng/mL"),
    ("5 ng·mL⁻¹", 5, "ng/mL"),
    ("1×10^2 ng/mL", 100, "ng/mL"),
    ("１０ ng／mL", 10, "ng/mL"),
    ("10 nM", 10, "nM"),
    ("2 μM", 2000, "nM"),
])
def test_parse_concentration_units(text, value, unit):
    c = _one(text)
    assert c.value == pytest.approx(value) and c.unit == unit and c.kind == "unit"


def test_parse_concentration_edge_cases():
    assert parse_concentrations("scale bar 50 μm") == []  # a length, not a concentration
    assert parse_concentrations("Ex 488 nm") == []
    assert _one("Control").kind == "control-word" and _one("对照").value == 0
    assert _one("10").kind == "bare-number"
    assert [c.value for c in parse_concentrations("0.1 ng/ml   1 ng/ml   5 ng/ml")] == [0.1, 1, 5]


def test_clean_deck_mapping(clean_deck):
    deck = read_deck(clean_deck)
    got = {a.picture.key: (a.conc.value if a.conc else None, a.status) for a in map_deck(deck)[0]}
    assert got == {
        "s03_p01": (10, "ok"), "s03_p02": (50, "ok"), "s03_p03": (100, "ok"),
        "s04_p01": (0, "check"), "s04_p02": (0.1, "ok"), "s04_p03": (0.5, "ok"),
        "s04_p04": (1, "ok"), "s04_p05": (5, "ok"),
        "s05_p01": (0.1, "by-order"), "s05_p02": (1, "by-order"), "s05_p03": (5, "by-order"),
        "s06_p01": (None, "ambiguous"), "s06_p02": (None, "ambiguous"),
        "s06_p03": (None, "unlabeled"), "s06_p04": (None, "conflict"),
    }
    pics = {p.key: p for s in deck.slides for p in s.pictures}
    grouped = pics["s04_p03"].rect  # lives in a group scaled by 0.5
    assert grouped.left == pytest.approx(Inches(0.4 + 2 * 2.55), abs=2)
    assert grouped.width == pytest.approx(Inches(2.2), abs=2)
    assert any(e.startswith("lum(") for e in pics["s04_p05"].effects)


def test_reference_candidates(clean_deck):
    deck = read_deck(clean_deck)
    (chart,) = reference_candidates(deck, {2, 3})
    assert [p[0] for p in chart["points"]] == REF_CONC
    assert [p[1] for p in chart["points"]] == pytest.approx(list(ref_curve(REF_CONC) + REF_NOISE))
    (table,) = reference_candidates(deck, {1})
    assert [p[0] for p in table["points"]] == REF_CONC


def test_clean_deck_run(clean_deck):
    res = run(deck_path=clean_deck)
    items = {it.key: it for it in res.items}
    assert res.channel == "G"
    f0, fmax, k = res.fit.params
    assert abs(f0 - 20) < 6 and abs(fmax - 100) < 8 and abs(k - 15) < 4
    assert res.calibration.method == "linear" and len(res.calibration.points) == 3
    assert res.calibration.r2 > 0.95
    for c in res.comparison:  # measured images, converted, follow the true curve
        assert c["converted"] == pytest.approx(float(ref_curve(c["conc"])), rel=0.1)
    assert res.inversions == [] and res.spearman == pytest.approx(1.0)
    assert sorted(res.groups) == [0, 0.1, 0.5, 1, 5, 10, 50, 100]
    assert res.groups[0.1]["n"] == 2 and res.groups[5]["n"] == 1  # identical file counted once
    assert items["s05_p03"].excluded
    assert items["s04_p04"].metrics["overlay_frac"] > 0.005  # scale bar excluded
    assert any("PPT 对这张图做了调整" in f for f in items["s04_p05"].flags)


def test_problem_deck_diagnosis(problem_deck):
    res = run(deck_path=problem_deck)
    inv = {(d["low"], d["high"]): " ".join(r for rs in d["reasons"].values() for r in rs) for d in res.inversions}
    assert set(inv) == {(1, 5), (50, 100)}
    assert "覆盖更少" in inv[(1, 5)]  # fewer but brighter cells
    assert "PPT 里显示的是调整后的图" in inv[(50, 100)]  # low gain, brightened in PowerPoint
    sat = next(it for it in res.items if it.conc and it.conc.value == 200)
    assert sat.metrics["sat_frac"] > 0.05 and any("饱和" in f for f in sat.flags)
    # the images cannot reproduce the reference at 50/100/200, so nothing is put on its scale
    assert not res.calibration.usable and res.comparison == []


def test_report_files(clean_deck, tmp_path):
    res = run(deck_path=clean_deck)
    write_all(res, tmp_path, "clean.pptx", render=False)
    for name in ["summary.md", "mapping.csv", "measurements.csv", "per_concentration.csv", "reference_used.csv",
                 "roi_sheet.png", "dose_response.png", "deck_content.md", "images/s04_p05.png"]:
        assert (tmp_path / name).exists(), name
    summary = (tmp_path / "summary.md").read_text(encoding="utf-8")
    assert "按顺序推定" in summary and "外推" in summary and "s06_p04（冲突）" in summary


def test_slide_render(clean_deck, tmp_path):
    pngs, err = render_slides(clean_deck, tmp_path / "slides", width=800)
    if err:
        pytest.skip(err)
    assert len(pngs) == 6


def test_folder_of_raw_tiffs(tmp_path):
    rng = np.random.default_rng(0)
    rows = []
    for i, c in enumerate([0.1, 1, 10]):
        img = np.clip(rng.normal(200 + 300 * c, 20, (128, 128)), 0, 4095).astype(np.uint16)
        img[:4, :4] = 4095  # a clipped corner, outside the default ROI
        tifffile.imwrite(tmp_path / f"img{i}.tif", img)
        rows.append({"file": f"img{i}.tif", "conc": c, "unit": "ng/mL"})
    with open(tmp_path / "map.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, ["file", "conc", "unit"])
        w.writeheader()
        w.writerows(rows)
    res = run(images_dir=tmp_path, map_csv=tmp_path / "map.csv")
    assert res.channel == "gray" and res.unit == "ng/mL"
    assert [it.plane.sat_value for it in res.items] == [4095] * 3  # 12-bit inferred
    assert [res.groups[c]["fi"] for c in (0.1, 1, 10)] == pytest.approx([230, 500, 3200], rel=0.02)


def test_band_widens_outside_the_data():
    c = np.array(REF_CONC, float)
    fit = curve.fit_saturation(c, ref_curve(c) + REF_NOISE, "langmuir")
    lo_in, hi_in = fit.band([50], prediction=True)
    lo_out, hi_out = fit.band([0.1], prediction=True)
    assert (hi_out - lo_out)[0] > (hi_in - lo_in)[0]


def test_calibrate_single_shared_point_is_a_ratio():
    cal = curve.calibrate({10: 60.0, 20: 77.0}, {10: 30.0, 1: 8.0})
    assert cal.method == "ratio" and cal.a == pytest.approx(2.0) and cal.b == 0


def test_near_duplicates_ignore_shared_illumination():
    import io

    from PIL import Image

    from fi_analysis import measure
    from synthetic import cell_png

    yy, xx = np.mgrid[0:256, 0:256]
    vignette = 1 + 2.0 * np.exp(-((xx - 128) ** 2 + (yy - 128) ** 2) / (2 * 60 ** 2))

    def dim_field(key, seed, gain=1.0):
        rgb = np.array(Image.open(io.BytesIO(cell_png(0.1, seed=seed, gain=gain, scalebar=True))))
        rgb[..., 1] = np.clip(rgb[..., 1] * vignette, 0, 254).astype(np.uint8)
        buf = io.BytesIO()
        Image.fromarray(rgb).save(buf, format="PNG")
        return measure.decode(buf.getvalue(), key)

    planes = [dim_field("a", 7), dim_field("b", 7, gain=1.5), dim_field("c", 8)]
    pairs = measure.near_duplicates(planes, "G", (0.1, 0.1, 0.8, 0.8))
    assert [(a, b) for a, b, _ in pairs] == [("a", "b")]  # same field reused; c is a different field
