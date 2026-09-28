"""End-to-end run: map images to concentrations, measure the same ROI on every
image, check the existing curve, and diagnose any "brighter but lower FI" pair."""
from __future__ import annotations

import csv
import hashlib
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

from . import curve, measure
from .deck import Conc, map_deck, parse_concentrations, read_deck

USABLE = ("ok", "by-order", "check")


@dataclass
class Item:
    key: str
    source: str
    conc: Conc | None
    status: str
    label: str = ""
    note: str = ""
    picture: object = None  # deck.Picture when the image comes from a deck
    plane: measure.Plane | None = None
    error: str | None = None
    metrics: dict | None = None
    flags: list = field(default_factory=list)
    excluded: str | None = None  # why the item is kept out of the curve

    @property
    def usable(self):
        return (self.conc is not None and self.status in USABLE and self.metrics is not None
                and self.excluded is None)


@dataclass
class Result:
    items: list
    channel: str
    roi: tuple
    unit: str | None
    groups: dict  # conc -> {"fi", "sd" (tile SD), "se", "sd_between", "n", "keys"}
    duplicates: list
    unused_labels: list
    reference: dict | None  # conc -> (fi, sd)
    reference_source: str
    reference_candidates: list
    fit: curve.Fit | None
    fit_error: str | None
    calibration: curve.Calibration | None
    comparison: list
    inversions: list
    spearman: float
    deck: object = None
    warnings: list = field(default_factory=list)


# ----------------------------------------------------------------- inputs
def _items_from_deck(deck):
    assignments, unused = map_deck(deck)
    items = []
    for a in assignments:
        p = a.picture
        items.append(Item(p.key, f"第 {p.slide} 页第 {p.order} 张图（{p.shape_name}）", a.conc, a.status,
                          a.label, a.note, picture=p))
    return items, unused


def _items_from_folder(folder, map_csv):
    items = []
    with open(map_csv, newline="", encoding="utf-8-sig") as fh:
        for i, row in enumerate(csv.DictReader(fh), start=1):
            path = Path(folder) / row["file"].strip()
            text = (row.get("conc") or "").strip() + " " + (row.get("unit") or "").strip()
            concs = parse_concentrations(text)
            conc = concs[0] if len(concs) == 1 else None
            items.append(Item(f"f{i:02d}", str(path), conc, "ok" if conc else "unlabeled", text.strip(),
                              "" if conc else "对照表里的浓度无法识别"))
    return items


def _float(text):
    try:
        return float(str(text).strip().replace(",", ""))
    except ValueError:
        return None


def _cell_conc(text):
    concs = parse_concentrations(str(text))
    return concs[0].value if len(concs) == 1 else None


def reference_candidates(deck, slides):
    """Numeric concentration/FI series found in charts and tables on `slides`."""
    out = []
    for s in deck.slides:
        if s.number not in slides:
            continue
        for ch in s.charts:
            for ser in ch["series"]:
                if any(v is not None for v in ser["x"]):
                    xs, ys = ser["x"], ser["y"]
                else:
                    xs, ys = [_cell_conc(c) if c is not None else None for c in ser["categories"]], ser["values"]
                sd = [None] * len(ys)
                for eb in ser["errbars"]:
                    if eb["valtype"] == "cust" and eb["dir"] in (None, "y") and eb["plus"]:
                        sd = eb["plus"] + [None] * (len(ys) - len(eb["plus"]))
                pts = [(x, y, e) for x, y, e in zip(xs, ys, sd) if x is not None and y is not None]
                if len(pts) >= 3:
                    out.append({"where": f"第 {s.number} 页图表“{ch['shape']}”系列“{ser['name']}”",
                                "points": pts})
        for t, grid in enumerate(s.tables, start=1):
            if not grid or len({len(r) for r in grid}) != 1:
                continue
            for orient, lines in (("column", [list(c) for c in zip(*grid)]), ("row", grid)):
                for i, line in enumerate(lines):
                    concs = [_cell_conc(c) for c in line]
                    if sum(c is not None for c in concs) < 3:
                        continue
                    for other in lines[i + 1:]:
                        pts = [(c, _float(v), None) for c, v in zip(concs, other)
                               if c is not None and _float(v) is not None]
                        if len(pts) >= 3:
                            where = f"第 {s.number} 页表格 {t} 的{'列' if orient == 'column' else '行'}“{other[0]}”"
                            out.append({"where": where, "points": pts})
    return out


def read_reference_csv(path):
    ref = {}
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            c = _cell_conc(row["conc"])
            fi = _float(row["fi"])
            if c is None or fi is None:
                continue
            ref[c] = (fi, _float(row.get("sd", "")) if row.get("sd") else float("nan"))
    return ref


# ------------------------------------------------------------------- QC
def _qc_flags(item, channel, overlay):
    m, p = item.metrics, item.plane
    flags = list(p.notes)
    if m["sat_frac"] > 0.001:
        flags.append(f"ROI 内 {m['sat_frac']:.2%} 像素处于饱和值 {p.sat_value:g}（被截平）")
    if m["zero_frac"] > 0.05:
        flags.append(f"ROI 内 {m['zero_frac']:.1%} 像素为 0（黑电平被截断，弱信号丢失）")
    if m["overlay_frac"] > 0.0005:
        flags.append(f"已排除叠加标注（比例尺/文字），占 ROI {m['overlay_frac']:.2%}")
    flags.extend(measure.channel_report(p, channel, overlay))
    if item.picture is not None:
        if item.picture.effects:
            flags.append("PPT 对这张图做了调整：" + "; ".join(item.picture.effects))
        if item.picture.cstate in ("email", "screen", "print"):
            flags.append(f"PPT 压缩过这张图（cstate={item.picture.cstate}）")
    if p.fmt == "JPEG":
        flags.append("JPEG（有损导出，不是原始数据）")
    return flags


def _file_hash(it):
    if it.picture is not None:
        return it.picture.sha1, it.picture.crop
    return hashlib.sha1(Path(it.source).read_bytes()).hexdigest(), None


def _duplicates(items, channel, roi):
    dups = []
    by_blob = defaultdict(list)
    for it in items:
        if it.plane is not None:
            by_blob[_file_hash(it)].append(it)
    for same in by_blob.values():
        if len(same) > 1:
            concs = {(i.conc.value, i.conc.unit) for i in same if i.conc is not None}
            dups.append({"keys": [i.key for i in same], "kind": "同一个文件",
                         "different_conc": len(concs) > 1})
    planes = [it.plane for it in items if it.plane is not None]
    exact = {frozenset(d["keys"]) for d in dups}
    by_key = {it.key: it for it in items}
    for a, b, r in measure.near_duplicates(planes, channel, roi):
        if any({a, b} <= s for s in exact):
            continue
        ca, cb = by_key[a].conc, by_key[b].conc
        dups.append({"keys": [a, b], "kind": f"几乎相同（r={r:.3f}）",
                     "different_conc": ca is not None and cb is not None and
                     (ca.value, ca.unit) != (cb.value, cb.unit)})
    return dups


# ------------------------------------------------------------- diagnosis
STATUS_CN = {"ok": "确定", "by-order": "按顺序推定", "check": "需核对", "ambiguous": "不确定",
             "conflict": "冲突", "unlabeled": "无标签"}


def diagnose(lo, hi, channel, roi):
    """Reasons why image `hi` (higher concentration) measures below image `lo`."""
    reasons = []
    for it in (lo, hi):
        if it.status != "ok":
            reasons.append(f"{it.key}：浓度来自“{STATUS_CN[it.status]}”的匹配（{it.note}），需要你确认")
    mh, ml = hi.metrics, lo.metrics
    if mh["sat_frac"] > 0.005:
        reasons.append(f"{hi.key}：ROI 内 {mh['sat_frac']:.1%} 像素饱和，高端信号被截平，FI 被低估")
    for it in (lo, hi):
        if it.picture is not None and it.picture.effects:
            reasons.append(f"{it.key}：PPT 里显示的是调整后的图（{'; '.join(it.picture.effects)}），"
                           "你看到的亮度不是被测量的嵌入像素")
        reasons.extend(f"{it.key}：{f}" for f in it.flags if f.startswith("通道 "))
    t, (area_l, _), (area_h, _) = measure.pair_coverage(lo.plane, hi.plane, channel, roi)
    if mh["p99"] > ml["p99"] and area_h < area_l:
        reasons.append(
            f"{hi.key}：亮的部分更亮（最亮 1% 像素 {mh['p99']:.1f} vs {ml['p99']:.1f}），但高于共同阈值 {t:.1f} "
            f"的面积更小（{area_h:.1%} vs {area_l:.1%}）——细胞更少/覆盖更少，把整块 ROI 的平均值拉低了；"
            "肉眼看的是亮细胞，平均 FI 算的是整块区域")
    se = math.hypot(ml["tile_se"], mh["tile_se"])
    drop = ml["mean"] - mh["mean"]
    if np.isfinite(se) and se > 0 and drop < 2 * se:
        reasons.append(f"差值 {drop:.2f} 不到取样误差（{se:.2f}，按 3×3 分块估计）的 2 倍：换个视野，"
                       "排序就可能反过来；需要每个浓度多个视野才能比较")
    pl, ph = lo.plane, hi.plane
    if (pl.fmt, pl.size) != (ph.fmt, ph.size):
        reasons.append(f"两张图导出方式不同：{lo.key} {pl.fmt} {pl.size[0]}x{pl.size[1]}，"
                       f"{hi.key} {ph.fmt} {ph.size[0]}x{ph.size[1]}")
    settings = "两张图的采集参数（激光功率、增益/HV、offset、针孔、扫描速度/平均次数）是否一致"
    if reasons:
        reasons.append(f"另外，导出图里看不出采集参数：请在原始文件里核对{settings}")
    else:
        reasons.append(f"上面这些常见原因都没找到：请在原始文件里核对{settings}，并再次确认标签没有放错")
    return reasons


# ------------------------------------------------------------------- run
def run(deck_path=None, images_dir=None, map_csv=None, roi=(0.1, 0.1, 0.8, 0.8), channel="auto",
        reference_csv=None, ref_slides=(2, 3), model="auto", tiles=3):
    deck = None
    warnings = []
    if deck_path:
        deck = read_deck(deck_path)
        items, unused = _items_from_deck(deck)
    else:
        items, unused = _items_from_folder(images_dir, map_csv), []

    for it in items:
        try:
            if it.picture is not None:
                it.plane = measure.decode(it.picture.blob, it.key, it.picture.crop)
            else:
                it.plane = measure.decode(it.source, it.key)
        except (measure.UnsupportedImage, FileNotFoundError) as exc:
            it.error = str(exc)

    decoded = [it for it in items if it.plane is not None]
    if not decoded:
        raise SystemExit("没有可读取的图片")
    target = [it for it in decoded if it.conc is not None and it.status in USABLE] or decoded
    channel = measure.choose_channel([it.plane for it in target], channel)
    if len({it.plane.dtype for it in target}) > 1:
        warnings.append("图片位深不一致：" +
                        ", ".join(sorted({f'{it.key}={it.plane.dtype}' for it in target})))
    for it in decoded:
        try:
            it.metrics = measure.measure(it.plane, channel, roi, tiles)
        except measure.UnsupportedImage as exc:
            it.error = str(exc)
            continue
        it.flags = _qc_flags(it, channel, measure.overlay_mask(it.plane))

    # identical files are certain and handled here; near-identical pairs are
    # only reported, because a match is a hint, not proof
    duplicates = _duplicates(items, channel, roi)
    by_key = {it.key: it for it in items}
    for d in duplicates:
        if d["kind"] != "同一个文件":
            continue
        if d["different_conc"]:
            for k in d["keys"]:
                by_key[k].excluded = "同一张图被标成了不同浓度"
        else:
            for k in d["keys"][1:]:
                by_key[k].excluded = f"与 {d['keys'][0]} 是同一张图（只计一次）"

    units = Counter(it.conc.unit for it in items if it.usable and it.conc.kind == "unit")
    unit = units.most_common(1)[0][0] if units else None
    if len(units) > 1:
        warnings.append(f"标签用了多种单位 {dict(units)}，曲线只用“{unit}”")
        for it in items:
            if it.usable and it.conc.kind == "unit" and it.conc.unit != unit:
                it.excluded = f"单位 {it.conc.unit} 与 {unit} 不同"

    per_conc = defaultdict(list)
    for it in items:
        if it.usable:
            per_conc[it.conc.value].append(it)
    groups = {}
    for c, its in sorted(per_conc.items()):
        means = [i.metrics["mean"] for i in its]
        groups[c] = {"fi": float(np.mean(means)),
                     "sd": float(np.mean([i.metrics["tile_sd"] for i in its])),
                     "se": float(np.sqrt(np.sum([i.metrics["tile_se"] ** 2 for i in its]))) / len(its),
                     "sd_between": float(np.std(means, ddof=1)) if len(means) > 1 else float("nan"),
                     "n": len(its), "keys": [i.key for i in its]}

    # existing analysis (reference curve)
    reference, ref_source, candidates = None, "", []
    if reference_csv:
        reference, ref_source = read_reference_csv(reference_csv), f"文件 {reference_csv}"
    elif deck is not None:
        candidates = reference_candidates(deck, set(ref_slides))
        if len(candidates) == 1:
            reference = {x: (y, e if e is not None else float("nan")) for x, y, e in candidates[0]["points"]}
            ref_source = f"自动读取自{candidates[0]['where']}——请核对数值"
    fit = fit_error = cal = None
    if reference and len(reference) >= 3:
        cs = sorted(reference)
        try:
            fit = curve.fit_saturation(cs, [reference[c][0] for c in cs], model)
        except (RuntimeError, ValueError) as exc:
            fit_error = str(exc)
    elif reference:
        fit_error = f"参考数据只有 {len(reference)} 个点"
    if reference and groups:
        cal = curve.calibrate({c: v[0] for c, v in reference.items()}, {c: g["fi"] for c, g in groups.items()})
        if cal is not None and cal.a <= 0:
            warnings.append("换算斜率不为正：图片测量没有复现参考数据的趋势")

    comparison = []
    if fit is not None and cal is not None and cal.usable:
        cs = np.array(sorted(groups))
        band = fit.band(cs, prediction=True)
        for i, c in enumerate(cs):
            conv = float(cal.apply(groups[c]["fi"]))
            lo, hi = (band[0][i], band[1][i]) if band is not None else (float("nan"), float("nan"))
            comparison.append({"conc": float(c), "measured": groups[c]["fi"], "converted": conv,
                               "predicted": float(fit.f(c)), "pi_low": float(lo), "pi_high": float(hi),
                               "extrapolated": bool(c < fit.x.min()),
                               "inside": bool(lo <= conv <= hi) if band is not None else None})

    inv, rho = curve.inversions({c: (g["fi"], g["se"]) for c, g in groups.items()})
    for d in inv:
        pairs = [(a, b) for a in per_conc[d["low"]] for b in per_conc[d["high"]]
                 if a.usable and b.usable and b.metrics["mean"] < a.metrics["mean"]]
        d["reasons"] = {f"{a.key} vs {b.key}": diagnose(a, b, channel, roi) for a, b in pairs}

    return Result(items, channel, tuple(roi), unit, groups, duplicates, unused, reference, ref_source,
                  candidates, fit, fit_error, cal, comparison, inv, rho, deck, warnings)


def parse_roi(text):
    vals = [float(v) for v in re.split(r"[,\s]+", text.strip()) if v]
    if len(vals) != 4 or min(vals) < 0 or vals[0] + vals[2] > 1 or vals[1] + vals[3] > 1:
        raise ValueError("ROI 应为 x,y,w,h（占图片宽高的比例，范围 0~1）")
    return tuple(vals)
