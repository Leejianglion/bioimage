"""Saturation dose-response fit, uncertainty bands, and the checks used to put
new measurements next to an existing curve.

The fitted curve is only ever used to *check* measured values; it is never a
source of values for concentrations that were not measured.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

import numpy as np
from scipy import optimize, stats


def langmuir(c, f0, fmax, k):
    c = np.asarray(c, float)
    return f0 + fmax * c / (k + c)


def hill(c, f0, fmax, k, n):
    c = np.asarray(c, float)
    cn = np.power(c, n)
    return f0 + fmax * cn / (np.power(k, n) + cn)


MODELS = {"langmuir": (langmuir, ["F0", "Fmax", "K"]), "hill": (hill, ["F0", "Fmax", "K", "n"])}


@dataclass
class Fit:
    model: str
    params: np.ndarray
    cov: np.ndarray | None
    x: np.ndarray
    y: np.ndarray
    rss: float
    at_bound: list

    @property
    def names(self):
        return MODELS[self.model][1]

    @property
    def dof(self):
        return len(self.x) - len(self.params)

    def f(self, c):
        return MODELS[self.model][0](c, *self.params)

    @property
    def se(self):
        if self.cov is None:
            return np.full(len(self.params), np.nan)
        return np.sqrt(np.clip(np.diag(self.cov), 0, None))

    @property
    def r2(self):
        sst = float(np.sum((self.y - self.y.mean()) ** 2))
        return 1 - self.rss / sst if sst > 0 else float("nan")

    @property
    def aicc(self):
        n, k = len(self.x), len(self.params)
        if n - k - 1 <= 0:
            return float("inf")
        return n * math.log(max(self.rss, 1e-300) / n) + 2 * k + 2 * k * (k + 1) / (n - k - 1)

    def band(self, c, level=0.95, prediction=False):
        """Delta-method confidence (or prediction) band; None when not estimable."""
        if self.cov is None or self.dof <= 0:
            return None
        c = np.atleast_1d(np.asarray(c, float))
        func = MODELS[self.model][0]
        grad = np.empty((c.size, len(self.params)))
        for i, p in enumerate(self.params):
            h = 1e-6 * max(1.0, abs(p))
            up, dn = self.params.copy(), self.params.copy()
            up[i] += h
            dn[i] -= h
            grad[:, i] = (func(c, *up) - func(c, *dn)) / (2 * h)
        var = np.einsum("ij,jk,ik->i", grad, self.cov, grad)
        if prediction:
            var = var + self.rss / self.dof
        half = stats.t.ppf(0.5 + level / 2, self.dof) * np.sqrt(np.clip(var, 0, None))
        y = self.f(c)
        return y - half, y + half


def fit_saturation(c, y, model="auto"):
    """Unweighted least-squares fit of F = F0 + Fmax*C^n/(K^n + C^n) (n = 1: Langmuir)."""
    c, y = np.asarray(c, float), np.asarray(y, float)
    if model == "auto":
        fits = [fit_saturation(c, y, "langmuir")]
        if len(c) >= 6:  # Hill needs n - 4 - 1 > 0 for AICc
            fits.append(fit_saturation(c, y, "hill"))
        return min(fits, key=lambda f: f.aicc)
    func, names = MODELS[model]
    if len(c) < len(names):
        raise ValueError(f"{model} needs at least {len(names)} points, got {len(c)}")
    ymin, ymax = float(y.min()), float(y.max())
    span = max(ymax - ymin, 1e-9)
    cpos = c[c > 0]
    kmin, kmax = (cpos.min() / 1000, cpos.max() * 1000) if cpos.size else (1e-6, 1e6)
    lo = [min(0.0, ymin - span), 0.0, kmin] + ([0.3] if model == "hill" else [])
    hi = [ymax, 50 * (span + abs(ymax)), kmax] + ([5.0] if model == "hill" else [])
    best = None
    for k0 in np.geomspace(kmin * 10, kmax / 10, 9):
        p0 = [min(max(ymin - 0.1 * span, lo[0]), hi[0]), 1.2 * span, k0] + ([1.0] if model == "hill" else [])
        try:
            p, cov = optimize.curve_fit(func, c, y, p0=p0, bounds=(lo, hi), maxfev=20000)
        except (RuntimeError, ValueError):
            continue
        rss = float(np.sum((y - func(c, *p)) ** 2))
        if best is None or rss < best[2]:
            best = (p, cov, rss)
    if best is None:
        raise RuntimeError(f"{model} fit did not converge")
    p, cov, rss = best
    if len(c) - len(p) <= 0 or not np.all(np.isfinite(cov)):
        cov = None
    at_bound = [n for n, v, a, b in zip(names, p, lo, hi) if abs(v - a) <= 1e-6 * max(1, abs(a))
                or abs(v - b) <= 1e-6 * max(1, abs(b))]
    return Fit(model, np.asarray(p), cov, c, y, rss, at_bound)


@dataclass
class Calibration:
    """Linear map from this analysis' FI scale onto the reference FI scale."""

    a: float
    b: float
    method: str
    points: list  # (conc, reference FI, measured FI)
    r2: float | None

    @property
    def usable(self):
        """False when the shared concentrations do not line up (then the scales cannot be merged)."""
        return self.a > 0 and (self.r2 is None or self.r2 >= 0.9)

    def apply(self, v):
        return self.a * np.asarray(v, float) + self.b


def calibrate(reference, measured):
    """Fit reference = a * measured + b over the concentrations measured in both."""
    points = []
    for c, ref in sorted(reference.items()):
        for m, meas in measured.items():
            if math.isclose(c, m, rel_tol=1e-6, abs_tol=1e-12):
                points.append((c, ref, meas))
    if not points:
        return None
    if len(points) == 1:
        c, ref, meas = points[0]
        return Calibration(ref / meas, 0.0, "ratio", points, None)
    x = np.array([p[2] for p in points])
    y = np.array([p[1] for p in points])
    res = stats.linregress(x, y)
    return Calibration(float(res.slope), float(res.intercept), "linear", points,
                       float(res.rvalue ** 2) if len(points) > 2 else None)


def inversions(groups):
    """Adjacent concentration pairs whose FI goes down as concentration goes up.

    groups: {conc: (mean FI, standard error)} -> (list of dicts, Spearman rho)
    """
    concs = sorted(groups)
    out = []
    for lo_c, hi_c in zip(concs, concs[1:]):
        (lo_v, lo_se), (hi_v, hi_se) = groups[lo_c], groups[hi_c]
        if hi_v < lo_v:
            se = math.hypot(lo_se, hi_se) if np.isfinite(lo_se) and np.isfinite(hi_se) else float("nan")
            out.append({"low": lo_c, "high": hi_c, "low_fi": lo_v, "high_fi": hi_v,
                        "drop": lo_v - hi_v, "drop_over_se": (lo_v - hi_v) / se if se > 0 else float("nan")})
    rho = float(stats.spearmanr(concs, [groups[c][0] for c in concs])[0]) if len(concs) > 2 else float("nan")
    return out, rho
