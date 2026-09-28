"""Write a run's outputs: CSV tables, the ROI contact sheet, slide mapping
overlays, the dose-response plot and a summary in Chinese."""
from __future__ import annotations

import csv
import math
import shutil
import subprocess
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

from . import measure
from .pipeline import STATUS_CN

FONT_FILES = [
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]
INK, INK2, MUTED, GRID, AXIS, SURFACE = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7", "#fcfcfb"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a"]
ROI_COLOR, OVERLAY_COLOR = (237, 161, 0), (232, 123, 164)
STATUS_COLOR = {"ok": "#0ca30c", "by-order": "#fab219", "check": "#fab219",
                "ambiguous": "#d03b3b", "conflict": "#d03b3b", "unlabeled": "#d03b3b"}


def _font(size):
    for f in FONT_FILES:
        if Path(f).exists():
            try:
                return ImageFont.truetype(f, size)
            except OSError:
                pass
    return ImageFont.load_default()


def _conc_text(it):
    return str(it.conc) if it.conc is not None else "—"


def _fmt(v, spec=".2f"):
    return "—" if v is None or (isinstance(v, float) and not math.isfinite(v)) else format(v, spec)


# ------------------------------------------------------------------ tables
def write_tables(res, out):
    with open(out / "mapping.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["image", "source", "label", "conc", "unit", "status", "note", "excluded"])
        for it in res.items:
            w.writerow([it.key, it.source, it.label, it.conc.value if it.conc else "",
                        (it.conc.unit or "") if it.conc else "", it.status, it.note, it.excluded or ""])
    cols = ["n_px", "mean", "median", "sd", "p99", "sat_frac", "zero_frac", "overlay_frac", "tile_sd", "tile_se",
            "tile_cv"]
    with open(out / "measurements.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["image", "conc", "unit", "status", "format", "width", "height", "dtype", "sat_value"]
                   + cols + ["flags", "excluded", "error"])
        for it in res.items:
            p, m = it.plane, it.metrics or {}
            w.writerow([it.key, it.conc.value if it.conc else "", (it.conc.unit or "") if it.conc else "",
                        it.status, p.fmt if p else "", p.size[0] if p else "", p.size[1] if p else "",
                        p.dtype if p else "", p.sat_value if p else ""]
                       + [m.get(c, "") for c in cols] + [" | ".join(it.flags), it.excluded or "", it.error or ""])
    comp = {c["conc"]: c for c in res.comparison}
    with open(out / "per_concentration.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["conc", "unit", "fi_mean", "spatial_sd", "between_image_sd", "n_images", "images",
                    "fi_on_reference_scale", "curve_prediction", "pi95_low", "pi95_high", "inside_pi",
                    "extrapolated"])
        for c, g in res.groups.items():
            k = comp.get(c, {})
            w.writerow([c, res.unit or "", g["fi"], g["sd"], g["sd_between"], g["n"], " ".join(g["keys"]),
                        k.get("converted", ""), k.get("predicted", ""), k.get("pi_low", ""),
                        k.get("pi_high", ""), k.get("inside", ""), k.get("extrapolated", "")])
    if res.reference:
        with open(out / "reference_used.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["conc", "fi", "sd"])
            for c in sorted(res.reference):
                w.writerow([c, *res.reference[c]])


def write_deck_content(deck, out):
    """Every slide's text, tables and chart data, for checking pages 2-3 by eye."""
    charts_dir = out / "charts"
    lines = [f"# PPT 内容：{Path(deck.path).name}", ""]
    for s in deck.slides:
        lines += [f"## 第 {s.number} 页", ""]
        lines += [f"- 图片 {len(s.pictures)} 张；识别到浓度标签 {len(s.labels)} 个"]
        for name, text in s.texts:
            lines.append(f"- 文本（{name}）：{' / '.join(t.strip() for t in text.splitlines() if t.strip())}")
        for t, grid in enumerate(s.tables, start=1):
            lines += ["", f"表格 {t}：", ""]
            if grid:
                lines.append("| " + " | ".join(c.replace("\n", " ") for c in grid[0]) + " |")
                lines.append("|" + "---|" * len(grid[0]))
                lines += ["| " + " | ".join(c.replace("\n", " ") for c in row) + " |" for row in grid[1:]]
        for ch in s.charts:
            lines += ["", f"图表“{ch['shape']}”（{', '.join(ch['types'])}）标题：{ch['title'] or '—'}；"
                          f"坐标轴：{'; '.join(v for v in ch['axes'].values() if v) or '—'}", ""]
            for ser in ch["series"]:
                xs = ser["x"] if any(v is not None for v in ser["x"]) else ser["categories"]
                ys = ser["y"] if any(v is not None for v in ser["y"]) else ser["values"]
                lines.append(f"- 系列“{ser['name']}”：" + ", ".join(f"({x}, {y})" for x, y in zip(xs, ys)))
                for eb in ser["errbars"]:
                    lines.append(f"  - 误差棒 {eb['valtype']} {eb['dir'] or ''}：+{eb['plus']} −{eb['minus']} "
                                 f"{eb['val'] or ''}")
            if ch["workbook"]:
                charts_dir.mkdir(exist_ok=True)
                name = f"slide{s.number:02d}_{ch['shape'].replace(' ', '_')}.xlsx"
                (charts_dir / name).write_bytes(ch["workbook"])
                lines.append(f"- 图表内嵌的 Excel 数据已导出：charts/{name}")
        if s.other:
            lines.append(f"- 无法读取的对象（SmartArt/OLE 等）：{', '.join(s.other)}")
        lines.append("")
    (out / "deck_content.md").write_text("\n".join(lines), encoding="utf-8")


def export_images(res, out):
    img_dir = out / "images"
    img_dir.mkdir(exist_ok=True)
    for it in res.items:
        if it.picture is not None:
            (img_dir / f"{it.key}.{it.picture.ext}").write_bytes(it.picture.blob)


# ----------------------------------------------------------------- figures
def _display_max(items):
    sats = {it.plane.sat_value for it in items}
    if len(sats) == 1 and math.isfinite(next(iter(sats))):
        return next(iter(sats))
    return max(float(np.percentile(measure.analysis_plane(it.plane, "gray"), 99.9)) for it in items)


def contact_sheet(res, path, thumb=220):
    """Every measured image, same fixed display range, ROI drawn, sorted by concentration."""
    items = [it for it in res.items if it.metrics is not None]
    if not items:
        return
    items.sort(key=lambda it: (not it.usable, it.conc.value if it.conc else math.inf, it.key))
    vmax = _display_max(items)
    cols = min(6, len(items))
    rows = math.ceil(len(items) / cols)
    aspect = max(it.plane.size[1] / it.plane.size[0] for it in items)
    th = int(thumb * aspect)
    cap, head = 70, 44
    sheet = Image.new("RGB", (cols * (thumb + 12) + 12, head + rows * (th + cap + 12)), SURFACE)
    draw = ImageDraw.Draw(sheet)
    draw.text((12, 10), f"固定显示范围 0–{vmax:g}（所有图相同，未做自动对比度）；黄框 = ROI；"
                        f"品红 = 已排除的标注像素；通道 {res.channel}", fill=INK, font=_font(15))
    f_small = _font(13)
    for n, it in enumerate(items):
        r, c = divmod(n, cols)
        x, y = 12 + c * (thumb + 12), head + r * (th + cap + 12)
        plane = measure.analysis_plane(it.plane, res.channel)
        gray = np.clip(plane / vmax * 255, 0, 255).astype(np.uint8)
        rgb = np.stack([gray] * 3, axis=-1)
        rgb[measure.overlay_mask(it.plane)] = OVERLAY_COLOR
        im = Image.fromarray(rgb)
        x0, y0, x1, y1 = measure.roi_box(plane.shape, res.roi)
        lw = max(2, round(im.width / thumb * 2))
        ImageDraw.Draw(im).rectangle([x0, y0, x1 - 1, y1 - 1], outline=ROI_COLOR, width=lw)
        im = im.resize((thumb, round(thumb * im.height / im.width)), Image.LANCZOS)
        sheet.paste(im, (x, y))
        status = STATUS_CN.get(it.status, it.status) + ("，未计入" if it.excluded else "")
        draw.text((x, y + im.height + 4), _conc_text(it), fill=INK, font=_font(15))
        draw.text((x, y + im.height + 24), f"FI {it.metrics['mean']:.1f}  饱和 {it.metrics['sat_frac']:.1%}",
                  fill=INK2, font=f_small)
        draw.text((x, y + im.height + 42), f"{it.key}  {status}", fill=MUTED, font=f_small)
    sheet.save(path)


def render_slides(deck_path, out_dir, width=1600):
    """Render slides with LibreOffice; returns (png paths, error message)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return [], "没有找到 LibreOffice，跳过幻灯片渲染"
    try:
        import pymupdf
    except ImportError:
        return [], "没有安装 pymupdf，跳过幻灯片渲染"
    out_dir.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        cmd = [soffice, f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "pdf",
               "--outdir", tmp, str(deck_path)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=600)
            pdf = next(Path(tmp).glob("*.pdf"))
        except (subprocess.SubprocessError, StopIteration) as exc:
            return [], f"幻灯片渲染失败（{exc}）"
        paths = []
        with pymupdf.open(pdf) as doc:
            for i, page in enumerate(doc, start=1):
                zoom = width / page.rect.width
                path = out_dir / f"slide-{i:02d}.png"
                page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).save(path)
                paths.append(path)
    return paths, None


def mapping_overlays(res, slide_pngs, out_dir):
    """Draw each picture's assigned concentration and match status on the rendered slide."""
    deck = res.deck
    by_slide = {}
    for it in res.items:
        if it.picture is not None and it.picture.rect is not None:
            by_slide.setdefault(it.picture.slide, []).append(it)
    written = []
    for png in slide_pngs:
        number = int(png.stem.split("-")[1])
        if number not in by_slide:
            continue
        im = Image.open(png).convert("RGB")
        scale = im.width / deck.width
        draw = ImageDraw.Draw(im)
        font = _font(max(14, im.width // 90))
        for it in by_slide[number]:
            r = it.picture.rect
            box = [r.left * scale, r.top * scale, r.right * scale, r.bottom * scale]
            color = STATUS_COLOR.get(it.status, "#d03b3b")
            draw.rectangle(box, outline=color, width=4)
            tag = f"{it.key}  {_conc_text(it)}  [{STATUS_CN.get(it.status, it.status)}]"
            tw, th = draw.textbbox((0, 0), tag, font=font)[2:]
            ty = box[3] - th - 8
            draw.rectangle([box[0], ty, box[0] + tw + 8, box[3]], fill=color)
            draw.text((box[0] + 4, ty + 3), tag, fill=INK, font=font)
        path = out_dir / f"slide-{number:02d}-mapping.png"
        im.save(path)
        written.append(path)
    return written


def dose_response_plot(res, path):
    import logging

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.ticker import FuncFormatter

    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

    names = {f.name for f in font_manager.fontManager.ttflist}
    cjk = [n for n in ("WenQuanYi Zen Hei", "Noto Sans CJK SC", "Microsoft YaHei", "PingFang SC") if n in names]
    plt.rcParams.update({"font.sans-serif": cjk + ["DejaVu Sans"], "axes.unicode_minus": False,
                         "axes.edgecolor": AXIS, "axes.labelcolor": INK2, "xtick.color": MUTED,
                         "ytick.color": MUTED, "axes.facecolor": SURFACE, "figure.facecolor": SURFACE})
    unit = res.unit or "(单位未知)"
    cal = res.calibration if res.calibration is not None and res.calibration.usable else None
    has_ref = bool(res.reference)
    two_panels = has_ref and bool(res.groups) and cal is None
    fig, axes = plt.subplots(1, 2 if two_panels else 1, figsize=(11 if two_panels else 7.5, 4.8), squeeze=False)
    axes = axes[0]

    def style(ax, title):
        ax.set_xscale("log")
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
        ax.grid(True, which="major", color=GRID, linewidth=0.8)
        ax.set_axisbelow(True)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        ax.set_xlabel(f"浓度 ({unit})")
        ax.set_title(title, color=INK, fontsize=11, loc="left")

    positive = [c for c in list(res.groups) + list(res.reference or {}) if c > 0]
    ax = axes[0]
    if has_ref:
        rc = np.array(sorted(res.reference))
        ry = np.array([res.reference[c][0] for c in rc])
        rsd = np.array([res.reference[c][1] for c in rc])
        if res.fit is not None:
            fit = res.fit
            lo_c = min(positive) / 1.5 if positive else fit.x.min()
            inside = np.geomspace(max(fit.x[fit.x > 0].min(), 1e-12), fit.x.max(), 200)
            ax.plot(inside, fit.f(inside), color=INK2, linewidth=2, label=f"拟合（{fit.model}，仅用已有数据）")
            if lo_c < fit.x.min():
                outside = np.geomspace(lo_c, fit.x.min(), 100)
                ax.plot(outside, fit.f(outside), color=INK2, linewidth=2, linestyle="--",
                        label="外推（这里没有已有数据）")
                band = fit.band(outside, prediction=True)
                if band is not None:
                    ax.fill_between(outside, band[0], band[1], color=GRID, alpha=0.7, linewidth=0,
                                    label="外推的 95% 预测区间")
        ax.errorbar(rc, ry, yerr=np.where(np.isfinite(rsd), rsd, 0), fmt="o", color=INK, markersize=7,
                    capsize=3, linewidth=1, label="已有结果（第 2–3 页）")
        style(ax, "浓度–FI（本次测量已换算到已有结果的刻度）" if cal is not None else "已有结果与拟合")
        ax.set_ylabel("FI（已有结果的刻度）")
    if res.groups and (cal is not None or not has_ref or two_panels):
        ax_m = axes[1] if two_panels else ax
        cs = np.array(sorted(res.groups))
        fi = np.array([res.groups[c]["fi"] for c in cs])
        sd = np.array([res.groups[c]["sd"] for c in cs])
        if cal is not None:
            fi, sd = cal.apply(fi), sd * cal.a
        pos = cs > 0
        ax_m.errorbar(cs[pos], fi[pos], yerr=np.where(np.isfinite(sd[pos]), sd[pos], 0), fmt="s",
                      color=SERIES[0], markersize=8, capsize=3, linewidth=1,
                      label="本次同一 ROI 测量（误差棒 = 图内 3×3 分块 SD，不是生物学重复）")
        for c, v in zip(cs[~pos], fi[~pos]):
            ax_m.axhline(v, color=SERIES[0], linestyle=":", linewidth=1.5, label=f"对照 0（{v:.1f}）")
        if ax_m is not ax or not has_ref:
            style(ax_m, "本次测量（图片灰度，未换算）")
            ax_m.set_ylabel("ROI 平均灰度")
    for a in axes:
        if a.get_legend_handles_labels()[0]:
            a.legend(frameon=False, fontsize=8, loc="upper left", labelcolor=INK2)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


# ----------------------------------------------------------------- summary
def _md_table(header, rows):
    out = ["| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    out += ["| " + " | ".join(str(c).replace("|", "/").replace("\n", " ") for c in r) + " |" for r in rows]
    return out


def write_summary(res, out, source, render_note=None):
    L = ["# FI–浓度分析报告", "", f"- 输入：`{source}`",
         f"- ROI（所有图片相同，占图片宽高的比例）：x={res.roi[0]:g}, y={res.roi[1]:g}, "
         f"w={res.roi[2]:g}, h={res.roi[3]:g}",
         f"- 分析通道：{res.channel}（所有图片相同）",
         "- 主指标 = ROI 内平均灰度（mean）。其余列只用于解释原因，不用来挑结果。", ""]
    if res.warnings:
        L += ["**警告**", ""] + [f"- {w}" for w in res.warnings] + [""]

    L += ["## 1. 浓度–图片对应关系", "",
          "规则：浓度标签在上，正下方紧挨着的图片就是对应的共聚焦图。不满足这条规则的图片不分配浓度。", ""]
    L += _md_table(["图片", "位置", "标签", "浓度", "状态", "说明"],
                   [[it.key, it.source, it.label or "—", _conc_text(it), STATUS_CN.get(it.status, it.status),
                     "；".join(x for x in (it.note, it.excluded and f"未计入：{it.excluded}") if x) or ""]
                    for it in res.items])
    unsure = [it for it in res.items if it.conc is None or it.status != "ok"]
    L += [""]
    if unsure:
        L += ["**以下图片的浓度不能确定或需要你确认：** " +
              "、".join(f"{it.key}（{STATUS_CN.get(it.status, it.status)}）" for it in unsure), ""]
    else:
        L += ["所有图片都唯一对应到正上方的浓度标签。", ""]
    if res.unused_labels:
        L += ["没有对应到图片的浓度标签：" + "；".join(f"第 {l.slide} 页“{l.text}”" for l in res.unused_labels), ""]
    if render_note:
        L += [f"（{render_note}）", ""]
    else:
        L += ["每页的对应情况画在 `slides/slide-XX-mapping.png`，请逐页目视确认。", ""]

    L += ["## 2. 图片质控", ""]
    L += _md_table(["图片", "格式", "尺寸", "位深", "问题"],
                   [[it.key, it.plane.fmt if it.plane else "—",
                     f"{it.plane.size[0]}x{it.plane.size[1]}" if it.plane else "—",
                     it.plane.dtype if it.plane else "—",
                     it.error or "；".join(it.flags) or "无"] for it in res.items])
    if res.duplicates:
        L += ["", "**重复/高度相似的图片：**", ""]
        for d in res.duplicates:
            same_file = d["kind"] == "同一个文件"
            L.append(f"- {'、'.join(d['keys'])}：{d['kind']}"
                     + ("，**但标的浓度不同**" if d["different_conc"] else "")
                     + ("" if same_file else "——可能是同一视野用了两次，请核对原始文件（仍计入曲线）"))
    L += [""]

    L += ["## 3. 同一 ROI 的测量结果（按浓度排序）", ""]
    measured = sorted((it for it in res.items if it.metrics is not None and it.conc is not None),
                      key=lambda it: (it.conc.value, it.key))
    L += _md_table(["浓度", "图片", "mean（FI）", "median", "最亮1%(p99)", "空间SD", "饱和%", "0值%", "状态"],
                   [[_conc_text(it), it.key, f"{it.metrics['mean']:.2f}", f"{it.metrics['median']:.1f}",
                     f"{it.metrics['p99']:.1f}", _fmt(it.metrics["tile_sd"]), f"{it.metrics['sat_frac']:.2%}",
                     f"{it.metrics['zero_frac']:.1%}",
                     STATUS_CN.get(it.status, it.status) + ("，未计入" if it.excluded else "")]
                    for it in measured])
    L += ["", "“空间SD”= ROI 切成 3×3 小块后各块平均值的标准差，表示同一张图里换个位置取值会差多少；"
              "它不是生物学重复，不能代替多视野/多样本的误差棒。", ""]

    L += ["## 4. 单调性检查：浓度越高，FI 是否越高", ""]
    if len(res.groups) >= 2:
        L.append("每个浓度的平均 FI：" + "，".join(f"{c:g} → {g['fi']:.2f}" for c, g in res.groups.items())
                 + f"（Spearman ρ = {_fmt(res.spearman)}）")
        L.append("")
        if not res.inversions:
            L += ["没有出现“浓度更高但 FI 更低”的情况。", ""]
        for d in res.inversions:
            L += [f"**{d['high']:g} 的 FI（{d['high_fi']:.2f}）低于 {d['low']:g}（{d['low_fi']:.2f}）**，"
                  f"差 {d['drop']:.2f}，约为取样误差的 {_fmt(d['drop_over_se'], '.1f')} 倍。可能原因：", ""]
            for pair, reasons in d["reasons"].items():
                L.append(f"- {pair}：")
                L += [f"  - {r}" for r in reasons]
            L.append("")
    else:
        L += ["可用于比较的浓度不足两个。", ""]

    L += ["## 5. 已有结果（第 2–3 页）与拟合", ""]
    if res.reference:
        L += [f"来源：{res.reference_source}", ""]
        L += _md_table(["浓度", "FI", "SD"], [[f"{c:g}", f"{v[0]:g}", _fmt(v[1], "g")]
                                             for c, v in sorted(res.reference.items())])
        L.append("")
        if res.fit is not None:
            f = res.fit
            eq = "F = F0 + Fmax·C/(K + C)" if f.model == "langmuir" else "F = F0 + Fmax·Cⁿ/(Kⁿ + Cⁿ)"
            L += [f"模型：{eq}（最小二乘，不加权）", ""]
            L += _md_table(["参数", "估计值", "标准误"],
                           [[n, f"{v:.4g}", _fmt(float(s), ".3g")] for n, v, s in zip(f.names, f.params, f.se)])
            L += ["", f"R² = {f.r2:.4f}，自由度 = {f.dof}"]
            if f.at_bound:
                L.append(f"注意：参数 {', '.join(f.at_bound)} 顶到了边界，对应的标准误不可靠。")
            if f.dof <= 0:
                L.append("注意：点数等于参数个数，曲线必然穿过所有点，无法估计不确定度。")
            L += ["", f"已有数据最低浓度是 {f.x.min():g}。低于它的曲线部分是外推：F0（零浓度时的本底）"
                      "只能靠曲线形状间接推断，已有数据本身并不包含低浓度的信息。", ""]
        else:
            L += [f"没有拟合：{res.fit_error}", ""]
    elif res.reference_candidates:
        L += ["在第 2–3 页找到多组可能的数据，没有自动选择，请指定（`--reference`）：", ""]
        L += [f"- {c['where']}：" + ", ".join(f"({x:g}, {y:g})" for x, y, _ in c["points"])
              for c in res.reference_candidates] + [""]
    else:
        L += ["第 2–3 页没有找到可以直接读取的数据（可能是图片格式的图表）。"
              "请把已有结果的浓度/FI/SD 填到 CSV 里用 `--reference` 提供，"
              "内容见 `deck_content.md` 和 `slides/`。", ""]

    L += ["## 6. 与已有结果的衔接", ""]
    cal = res.calibration
    if cal is None:
        L += ["本次测量和已有结果没有共同的浓度，所以无法把本次的灰度换算到已有结果的刻度上。"
              "两组数据不能直接画在同一条曲线上。可行的做法：在低浓度这批里同时拍/测一个 ≥10 的浓度（例如 10）作为桥接。", ""]
    else:
        pts = "；".join(f"{c:g}：已有 {r:.2f} / 本次 {m:.2f}" for c, r, m in cal.points)
        if cal.method == "ratio":
            L += [f"只有 1 个共同浓度（{pts}），按比例换算：已有 = {cal.a:.4g} × 本次。"
                  "这假设两种测量的本底都为 0；如果有本底，低浓度端会有偏差。", ""]
        else:
            L += [f"共同浓度 {len(cal.points)} 个（{pts}），线性换算：已有 = {cal.a:.4g} × 本次 + {cal.b:.4g}"
                  + (f"，R² = {cal.r2:.3f}" if cal.r2 is not None else "（只有 2 个点，换算本身无法检验）"), ""]
        if not cal.usable:
            L += ["**共同浓度上，本次测量和已有结果对不上（斜率不为正或 R² 低于 0.9）：按同一方法测 PPT 里的图，"
                  "复现不了已有结果。说明已有结果不是用同样的数据/方法得到的（例如原来用的是原始文件，而 PPT 里的图"
                  "调过显示），或者图片本身有第 4 节列出的问题。两组数据不能合并，所以不给出换算值。**", ""]
        if res.comparison:
            L += _md_table(["浓度", "本次测量", "换算后", "曲线预测", "95% 预测区间", "在区间内", "外推"],
                           [[f"{c['conc']:g}", f"{c['measured']:.2f}", f"{c['converted']:.2f}",
                             f"{c['predicted']:.2f}", f"{_fmt(c['pi_low'])} – {_fmt(c['pi_high'])}",
                             {True: "是", False: "**否**", None: "—"}[c["inside"]], "是" if c["extrapolated"] else ""]
                            for c in res.comparison])
            L += ["", "外推区间里，“在区间内”只说明没有和外推曲线矛盾，不说明曲线在那里是对的。", ""]

    L += ["## 7. 需要你确认 / 这次不能确定的事", ""]
    todo = []
    if unsure:
        todo.append("第 1 节列出的图片与浓度的对应关系")
    if res.reference and "自动读取" in res.reference_source:
        todo.append("第 5 节自动读取的已有数据是否就是第 2–3 页的结果")
    if not res.reference:
        todo.append("已有（≥10）结果的具体数值")
    if any(it.plane is not None and it.plane.fmt in ("JPEG", "PNG") for it in res.items):
        todo.append("这些图是从 PPT 里取出的导出图，不是原始共聚焦文件：只有当所有图用同一套采集参数、"
                    "同一套显示设置导出时，灰度才可比；正式结果请用原始文件（.czi/.lif/.oib/.nd2/.tif）重测")
    single = [f"{c:g}" for c, g in res.groups.items() if g["n"] == 1]
    if single:
        todo.append(f"浓度 {', '.join(single)} 只有 1 张图：没有生物学重复，排序差异无法做统计判断")
    L += [f"- {t}" for t in todo] + [""]

    L += ["## 方法（固定规则，没有逐张调整）", "",
          "1. 所有图片用同一个 ROI（按宽高比例定位），同一个通道，不做逐张的亮度/对比度/背景调整。",
          "2. PPT 里的裁剪按显示区域处理；PPT 里的亮度/对比度等调整不会改变被测的嵌入像素，只会被标出来。",
          "3. 比例尺/文字等白色标注自动排除（只对彩色 LUT 图有效）。",
          "4. 饱和像素、0 值像素、合成图通道、重复图片都会被标出。",
          "5. 拟合曲线只用来检查测量值，不用来生成没有测量过的浓度的数值。", ""]
    (out / "summary.md").write_text("\n".join(L), encoding="utf-8")


def write_all(res, out, source, render=True):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    write_tables(res, out)
    contact_sheet(res, out / "roi_sheet.png")
    dose_response_plot(res, out / "dose_response.png")
    render_note = None
    if res.deck is not None:
        write_deck_content(res.deck, out)
        export_images(res, out)
        if render:
            pngs, render_note = render_slides(res.deck.path, out / "slides")
            if pngs:
                mapping_overlays(res, pngs, out / "slides")
        else:
            render_note = "未渲染幻灯片（--no-render）"
    write_summary(res, out, source, render_note)
