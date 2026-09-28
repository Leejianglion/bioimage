"""Same-region fluorescence-intensity (FI) measurement, plus the checks that
decide whether two FI values can be compared at all.

Rules that keep the numbers objective:
  * one analysis channel for every image (chosen once, globally);
  * one ROI, given as fractions of the image, applied unchanged to every image;
  * when two images are compared for diagnosis, one shared threshold for both;
  * no per-image rescaling, background tweaks or ROI moves.
"""
from __future__ import annotations

import io
import os
from dataclasses import dataclass, field

import numpy as np
from PIL import Image

CHANNELS = "RGB"


class UnsupportedImage(Exception):
    pass


@dataclass
class Plane:
    key: str
    data: np.ndarray  # (H, W) gray or (H, W, 3) RGB, original values (after PowerPoint crop)
    kind: str  # "gray" | "rgb"
    fmt: str
    dtype: str
    sat_value: float  # value at which pixels are clipped
    notes: list = field(default_factory=list)

    @property
    def size(self):
        return self.data.shape[1], self.data.shape[0]


def _pil_array(img, notes):
    mode = img.mode
    if mode in ("P", "PA", "CMYK", "YCbCr", "LAB", "HSV"):
        img = img.convert("RGBA" if mode == "PA" or "transparency" in img.info else "RGB")
    elif mode == "1":
        img = img.convert("L")
    elif mode == "LA":
        alpha = np.array(img.getchannel("A"))
        if (alpha < 255).any():
            notes.append("图片含透明像素")
        img = img.getchannel("L")
    arr = np.array(img)
    if arr.ndim == 3 and arr.shape[2] == 4:
        if (arr[..., 3] < 255).any():
            notes.append("图片含透明像素")
        arr = arr[..., :3]
    if arr.dtype.byteorder == ">":
        arr = arr.astype(arr.dtype.newbyteorder("="))
    return arr


def decode(source, key="", crop=(0.0, 0.0, 0.0, 0.0)):
    """Decode image bytes or a file path into a Plane, applying a PowerPoint crop."""
    notes = []
    is_path = isinstance(source, (str, os.PathLike))
    if is_path and str(source).lower().endswith((".tif", ".tiff", ".lsm")):
        import tifffile

        arr = np.squeeze(tifffile.imread(source))
        fmt = "TIFF"
        if arr.ndim == 3 and arr.shape[0] <= 4 < arr.shape[-1]:
            arr = np.moveaxis(arr, 0, -1)  # channel-first -> channel-last
        if arr.ndim > 3 or (arr.ndim == 3 and arr.shape[-1] > 4):
            raise UnsupportedImage(f"{source}: 形状 {arr.shape} 像是 z-stack/时间序列，请先导出单层或投影")
        if arr.ndim == 3 and arr.shape[-1] != 3:
            raise UnsupportedImage(f"{source}: 有 {arr.shape[-1]} 个通道，请单独导出荧光通道")
    else:
        try:
            img = Image.open(source if is_path else io.BytesIO(source))
            img.load()
        except Exception as exc:  # EMF/WMF and other vector formats land here
            raise UnsupportedImage(f"无法读取图片（{exc}）") from exc
        fmt = img.format or "?"
        arr = _pil_array(img, notes)

    h, w = arr.shape[:2]
    l, t, r, b = (max(0.0, c) for c in crop)
    x0, x1 = int(round(l * w)), w - int(round(r * w))
    y0, y1 = int(round(t * h)), h - int(round(b * h))
    if (x0, y0, x1, y1) != (0, 0, w, h):
        arr = arr[y0:y1, x0:x1]
        notes.append(f"已按 PPT 的裁剪取显示区域（左 {l:.1%}，上 {t:.1%}，右 {r:.1%}，下 {b:.1%}）")

    dtype = arr.dtype
    if dtype == np.uint8:
        sat_value = 255.0
    elif np.issubdtype(dtype, np.integer):  # 12/14/16-bit data usually sits in a 16-bit container
        peak = float(arr.max())
        sat_value = float(next(v for v in (4095, 16383, 65535, 2 ** 32 - 1) if peak <= v))
        if sat_value in (4095, 16383):
            notes.append(f"按数据最大值推断为 {int(np.log2(sat_value + 1))}-bit")
    else:
        sat_value = float("inf")

    kind = "gray"
    if arr.ndim == 3:
        spread = arr.max(axis=2).astype(np.int64) - arr.min(axis=2)
        if np.mean(spread <= 3) >= 0.995:  # grey image stored as RGB (JPEG noise tolerated)
            arr = arr.mean(axis=2)
        else:
            kind = "rgb"
    return Plane(key, arr, kind, fmt, str(dtype), sat_value, notes)


def choose_channel(planes, requested="auto"):
    """One channel for all images: the RGB channel with the most total signal."""
    if requested != "auto":
        return requested
    rgb = [p for p in planes if p.kind == "rgb"]
    if not rgb:
        return "gray"
    totals = np.sum([p.data.reshape(-1, 3).astype(float).mean(axis=0) for p in rgb], axis=0)
    return CHANNELS[int(np.argmax(totals))]


def analysis_plane(plane, channel):
    if plane.kind == "gray":
        return plane.data.astype(float)
    if channel == "gray":
        return plane.data.astype(float).mean(axis=2)
    return plane.data[..., CHANNELS.index(channel)].astype(float)


def overlay_mask(plane):
    """White/grey annotations (scale bars, text) burned into a colour-LUT image."""
    if plane.kind != "rgb" or not np.isfinite(plane.sat_value):
        return np.zeros(plane.data.shape[:2], bool)
    return plane.data.min(axis=2) >= 0.4 * plane.sat_value


def channel_report(plane, channel, overlay):
    """Signal in the other colour channels that is not a copy of the analysis channel."""
    if plane.kind != "rgb" or channel == "gray":
        return []
    keep = ~overlay
    y = plane.data[..., CHANNELS.index(channel)][keep].astype(float)
    notes = []
    for ch in CHANNELS:
        if ch == channel:
            continue
        x = plane.data[..., CHANNELS.index(ch)][keep].astype(float)
        if x.mean() < 0.02 * max(y.mean(), 1.0):
            continue
        r = np.corrcoef(x, y)[0, 1] if x.std() > 0 and y.std() > 0 else 0.0
        if r < 0.9:
            notes.append(f"通道 {ch} 有独立信号（r={r:.2f}）：可能是合成图(merge)，或用了会改色的 LUT/饱和提示色")
    return notes


def roi_box(shape, roi):
    h, w = shape[:2]
    fx, fy, fw, fh = roi
    x0, y0 = int(round(fx * w)), int(round(fy * h))
    x1, y1 = int(round((fx + fw) * w)), int(round((fy + fh) * h))
    return max(0, x0), max(0, y0), min(w, x1), min(h, y1)


def roi_values(plane, channel, roi):
    img = analysis_plane(plane, channel)
    excl = overlay_mask(plane)
    x0, y0, x1, y1 = roi_box(img.shape, roi)
    return img[y0:y1, x0:x1], excl[y0:y1, x0:x1]


def measure(plane, channel, roi, tiles=3):
    vals, excl = roi_values(plane, channel, roi)
    v = vals[~excl]
    if v.size == 0:
        raise UnsupportedImage(f"{plane.key}: ROI 内没有像素")
    tile_means = []
    h, w = vals.shape
    for i in range(tiles):
        for j in range(tiles):
            ys = slice(i * h // tiles, (i + 1) * h // tiles)
            xs = slice(j * w // tiles, (j + 1) * w // tiles)
            t = vals[ys, xs][~excl[ys, xs]]
            if t.size:
                tile_means.append(float(t.mean()))
    tile_means = np.array(tile_means)
    return {
        "n_px": int(v.size),
        "mean": float(v.mean()),
        "median": float(np.median(v)),
        "sd": float(v.std()),
        "p99": float(np.percentile(v, 99)),
        "sat_frac": float(np.mean(v >= plane.sat_value)) if np.isfinite(plane.sat_value) else 0.0,
        "zero_frac": float(np.mean(v <= 0)),
        "overlay_frac": float(excl.mean()),
        "tile_sd": float(tile_means.std(ddof=1)) if tile_means.size > 1 else float("nan"),
        # sampling error of the ROI mean, from the tiles (tiles of one field are not replicates)
        "tile_se": float(tile_means.std(ddof=1) / np.sqrt(tile_means.size)) if tile_means.size > 1 else float("nan"),
        "tile_cv": float(tile_means.std(ddof=1) / tile_means.mean())
        if tile_means.size > 1 and tile_means.mean() > 0 else float("nan"),
    }


def pair_coverage(lo, hi, channel, roi):
    """Split two images with one shared threshold, halfway between their common
    background (lower median) and their common bright level (lower 99th pct).

    Returns the threshold and, per image, (fraction of ROI above it, mean above it).
    """
    vals = []
    for plane in (lo, hi):
        v, excl = roi_values(plane, channel, roi)
        vals.append(v[~excl])
    bg = min(float(np.median(v)) for v in vals)
    top = min(float(np.percentile(v, 99)) for v in vals)
    t = (bg + top) / 2
    out = []
    for v in vals:
        above = v > t
        out.append((float(above.mean()), float(v[above].mean()) if above.any() else float("nan")))
    return t, out[0], out[1]


def near_duplicates(planes, channel, roi, threshold=0.9):
    """Pairs of images whose ROI structure matches (same field used twice?).

    Thumbnails are high-pass filtered so that shared illumination fall-off,
    offset or burned-in annotations do not make different fields look alike.
    """
    thumbs = {}
    for p in planes:
        vals, excl = roi_values(p, channel, roi)
        a = vals.copy()
        if excl.any():
            a[excl] = np.median(vals[~excl]) if (~excl).any() else 0.0
        t = np.asarray(Image.fromarray(a.astype(np.float32)).resize((64, 64), Image.BILINEAR), float)
        low = Image.fromarray(t.astype(np.float32)).resize((8, 8), Image.BILINEAR).resize((64, 64), Image.BILINEAR)
        h = t - np.asarray(low, float)
        thumbs[p.key] = (h - h.mean()) / h.std() if h.std() > 0 else None
    keys = list(thumbs)
    pairs = []
    for i, a in enumerate(keys):
        for b in keys[i + 1:]:
            if thumbs[a] is None or thumbs[b] is None:
                continue
            r = float(np.mean(thumbs[a] * thumbs[b]))
            if r >= threshold:
                pairs.append((a, b, r))
    return pairs
