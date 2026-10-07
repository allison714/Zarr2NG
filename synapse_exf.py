"""
Find synapses in the pan stain and measure the expansion factor from them.

  synapse_exf.py [text]     go through the samples whose names contain the text

Follows M'Saad, Cairns et al., Nat. Biotechnol. 2025 (pan-ExM-t): the linear expansion
factor is the distance between the presynaptic dense projections (DPs) and the postsynaptic
density (PSD), centre to centre along the trans-synaptic axis, divided by its known size
before expansion, 82 nm. The paper drew the line profiles by hand in FIJI; here the
synapses are found and the profiles drawn by the program, and a person accepts or rejects
each one in the log. What the program does, step by step (every number is a setting below):

  1. Each plane of the pan channel is lightly smoothed (Gaussian, SMOOTH_PX; the paper
     smoothed 0.4-1.5 px). Bright, thin, straight-ish bars are picked out with a ridge
     filter (the Hessian's most negative eigenvalue at the scale of a PSD's thickness) and
     kept when they are as long as a PSD.
  2. Across each bar, an intensity profile is taken along the bar's normal and averaged
     along the middle of the bar. A synapse seen side-on shows two peaks, PSD and DPs.
     The second peak is looked for at 82 nm x EXF_SEARCH, a deliberately wide window, so
     that the answer is not built in.
  3. Two Gaussians on a flat baseline are fitted to the profile, to decide whether this is
     a synapse and where its peaks are. The distance is then read as the paper read it:
     peak to peak on a line LINE_PX pixels thick, divided by 81.9 nm. The fitted distance
     is kept beside it. A profile is kept only when the fit is good (FIT_R2), the dip
     between the peaks is real (MIN_DIP), the peaks are of comparable height (PEAK_RATIO)
     and the brighter one is as bright as a PSD (PSD_RANK), which leaves out pairs of membranes.
  4. The antibody channel is measured over the same footprint. A synapse is antibody-
     positive when its brightest smoothed antibody pixel is AB_SIGMAS noise-widths above
     the plane's median. Positive synapses are listed first, when the antibody is synaptic.
  5. The best synapses of each sample get a picture (pan, antibody, overlay, profile with
     the fit) and a row in "Synapse Log.xlsx". Type Y or N in its Accept column and run
     this again: the Summary sheet then also gives the expansion factor from the accepted
     synapses alone. Every synapse that passed, logged or not, is in <sample>_synapses.csv.

Planes are independent: each is read and worked through in tiles, so no image is loaded whole.
"""

import argparse
import configparser
import csv
import json
import math
import re
import sys
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import ab_efficiency as eff
import convert_to_zarr as core
import roi_log as roi

# --------------------------------------------------------------------------- settings

REF_NM = 81.9            # DP-PSD distance before expansion (M'Saad, Cairns et al. 2025, Supplementary Methods)
LINE_PX = 2              # thickness of the line the distance is read from, as in the paper
MIN_PER_FOV = 3          # a field of view needs this many synapses to count towards the mean across fields
EXF_NOMINAL = 20.0       # only sets the size of the filters; the measurement does not use it
EXF_SEARCH = (12.0, 30.0)  # the second peak is looked for at 82 nm times this range
EXF_EXPECTED = (16.0, 24.0)  # the lab's usual range; values outside it are marked, not dropped
SMOOTH_PX = 1.0          # Gaussian smoothing of both channels before anything is measured
PSD_SIGMA_NM = 15.0      # scale of the ridge filter: half the thickness of a PSD
PSD_LENGTH_NM = (120.0, 900.0)   # a bar this long can be a PSD
RIDGE_PERCENTILE = 97.0  # bars are the pixels whose ridge strength is above this share of the plane
BRIGHT_PERCENTILE = 80.0 # ...and a PSD is among the brighter things in the plane
PROFILE_HALF_NM = 260.0  # the profile runs this far either side of the bar
AVERAGE_HALF_NM = 90.0   # ...and is averaged along the bar over at most this far either side of its middle
FIT_R2 = 0.90            # the two-Gaussian fit must explain this share of the profile
MIN_DIP = 0.20           # the trough must fall this share of the smaller peak's height
PEAK_RATIO = (0.25, 4.0) # second peak's height as a share of the first
# The brighter of the two densities must outshine this share of the plane. Two cell membranes
# lying side by side give the same double peak as a synapse but are dim; a PSD is among the
# brightest things in pan-ExM. Set from GluA2-positive synapses (260831_WT4_GluA2ms, 3x3):
# 95% of them ranked above 0.996, against a median of 0.978 for the other double peaks.
PSD_RANK = 0.995
AB_SIGMAS = 6.0          # antibody-positive: this many noise-widths above the plane's median
TOP_N = 60               # synapses pictured and logged per sample
# Antibody labelling, measured at the synapses (see antibody_numbers):
ON_SYNAPSE_X = 2.0       # a punctum this many nominal DP-PSD distances from a synapse's centre is "on" it
SPOT_MIN_PX = 3          # a punctum is at least this many connected pixels above the threshold
# Resolution, from the width of isolated antibody puncta (see spot_width):
SPOT_ALONE_PX = 12       # ...with no other punctum this close
SPOT_FITS = 60           # at most this many puncta are fitted per tile, the brightest first
SPOT_SMALLEST = 10       # the resolution is this percentile of the widths: the smallest puncta, not the typical one
MIN_WIDTHS = 20          # fewer fitted puncta than this give no resolution
MIN_FOR_EXF = 10         # fewer antibody-positive synapses than this cannot carry a sample's ExF
MIN_ACCEPTED = 5         # ...and fewer accepted ones than this cannot either
PLANE_GAP_UM = 4.0       # planes closer than this would show the same synapses twice; they are skipped
TILE_PX = 2048
LOG_NAME = "Synapse Log.xlsx"
HEADER_ROW = 3

# Antibodies recognised in file names: name -> (label, side of the synapse, or None when not synaptic)
ANTIBODIES = {
    "homer1": ("Homer1", "post"), "shank2": ("Shank2", "post"), "psd95": ("PSD-95", "post"),
    "nr1": ("NR1", "post"), "glua1": ("GluA1", "post"), "glua2": ("GluA2", "post"), "glua3": ("GluA3", "post"),
    "glua4": ("GluA4", "post"), "mglur5": ("mGluR5", "post"), "bassoon": ("Bassoon", "pre"),
    "synaptophysin": ("Synaptophysin", "pre"), "cx43": ("Cx43", None), "gfap": ("GFAP", None), "pff1": ("PFF1", None),
}

COLUMNS = [
    ("#", 6), ("Sample", 40), ("Slice", 7), ("X (px)", 9), ("Y (px)", 9), ("Antibody", 12), ("AB Signal (σ)", 13),
    ("AB+", 6), ("DP–PSD (µm, as imaged)", 14), ("ExF (peak to peak)", 12), ("ExF (Gaussian fit)", 12), ("In 16–24", 9), ("Fit R²", 8), ("Dip", 7),
    ("Peak Ratio", 10), ("Accept", 9), ("Notes", 40), ("Picture", 12), ("Neuroglancer", 13),
]
HEADERS = [c[0] for c in COLUMNS]


# --------------------------------------------------------------------------- which samples


def antibodies_in(name):
    """Antibodies named in a file name, as [(label, side)]; 'GluA2ms' is GluA2 raised in mouse."""
    found = []
    for token in re.split(r"[_\s-]+", name.lower()):
        for key, value in ANTIBODIES.items():
            if token.startswith(key) and value not in found:
                found.append(value)
    return found


def ignored(name):
    """Why a sample is left out, or ''. Left out: 'SZ' or 'control' in the name, or two numbers
    joined by an underscore (a repeat such as Sample1_2) unless the first number is the end
    of an antibody's name (Homer1_2 stays) or of a grid size (3x3_1 stays: 3x3 says the image
    is the 2D overview montage, not a repeat)."""
    if re.search(r"(?<![A-Za-z])SZ(?![A-Za-z])", name):
        return "SZ in the name"
    if "control" in name.lower():
        return "control in the name"
    tokens = name.split("_")
    for left, right in zip(tokens[1:], tokens[2:]):       # the leading date is not a repeat number
        if re.fullmatch(r"\d+x\d+", left.lower()):
            continue
        if left[-1:].isdigit() and right[:1].isdigit() and not any(left.lower().startswith(k) for k in ANTIBODIES):
            return f"number_number ({left}_{right})"
    return ""


# --------------------------------------------------------------------------- finding and measuring


def two_peaks(t, b, a1, m1, s1, a2, m2, s2):
    return b + a1 * np.exp(-(t - m1) ** 2 / (2 * s1 * s1)) + a2 * np.exp(-(t - m2) ** 2 / (2 * s2 * s2))


def measure_bar(pan, ab, noise, row, col, theta, length, um):
    """Profile across one bar and its two-Gaussian fit. Returns a dict, or None when the
    profile is not that of a synapse seen side-on.

    pan, ab - smoothed planes (ab may be None); noise - (median, sigma) of the smoothed antibody plane
    theta   - angle of the bar's long axis from the row axis (skimage's 'orientation')
    """
    from scipy import ndimage, optimize, signal

    px_nm = EXF_NOMINAL / (um * 1000)                        # pixels per nanometre, at the nominal expansion
    reach = PROFILE_HALF_NM * px_nm
    along = np.array([math.cos(theta), math.sin(theta)])     # (row, col) along the bar
    across = np.array([-along[1], along[0]])
    half = min(0.3 * length, AVERAGE_HALF_NM * px_nm)
    t = np.arange(-reach, reach + 0.25, 0.5)
    s = np.arange(-half, half + 0.5, 1.0)
    rows = row + t[:, None] * across[0] + s[None, :] * along[0]
    cols = col + t[:, None] * across[1] + s[None, :] * along[1]
    if rows.min() < 1 or cols.min() < 1 or rows.max() > pan.shape[0] - 2 or cols.max() > pan.shape[1] - 2:
        return None
    grid = ndimage.map_coordinates(pan, [rows, cols], order=1)
    if (grid <= 0).any():                                    # stitching padding
        return None
    profile = grid.mean(axis=1)

    spread = profile.max() - profile.min()
    peaks, _ = signal.find_peaks(profile, prominence=0.08 * spread)
    if len(peaks) < 2:
        return None
    first = peaks[np.argmin(np.abs(t[peaks]))]               # the bar itself
    if abs(t[first]) > 3:
        return None
    near, far = (REF_NM * e / (um * 1000) for e in EXF_SEARCH)
    partners = [p for p in peaks if near <= abs(t[p] - t[first]) <= far]
    if not partners:
        return None
    second = max(partners, key=lambda p: profile[p])
    gap = abs(t[second] - t[first])
    lo, hi = min(t[first], t[second]) - 0.9 * gap, max(t[first], t[second]) + 0.9 * gap
    win = (t >= lo) & (t <= hi)
    base = float(profile[win].min())
    start = [base, profile[first] - base, t[first], gap / 4, profile[second] - base, t[second], gap / 4]
    low = [0, 0, t[first] - 2, 0.8, 0, t[second] - 2, 0.8]
    high = [float(profile.max()), 2 * spread + 1, t[first] + 2, gap, 2 * spread + 1, t[second] + 2, gap]
    try:
        fit, _ = optimize.curve_fit(two_peaks, t[win], profile[win], p0=start, bounds=(low, high), maxfev=2000)
    except (RuntimeError, ValueError):
        return None
    model = two_peaks(t[win], *fit)
    resid = float(((profile[win] - model) ** 2).sum())
    total = float(((profile[win] - profile[win].mean()) ** 2).sum())
    r2 = 1 - resid / total if total > 0 else 0.0
    b, a1, m1, s1, a2, m2, s2 = fit
    distance = abs(m2 - m1)
    if not (near <= distance <= far):
        return None
    between = (t >= min(m1, m2)) & (t <= max(m1, m2))
    tops = min(two_peaks(m1, *fit), two_peaks(m2, *fit))
    dip = float((tops - two_peaks(t[between], *fit).min()) / max(tops - b, 1e-6))
    height = float(a2 / a1) if a1 > 0 else 0.0
    if r2 < FIT_R2 or dip < MIN_DIP or not (PEAK_RATIO[0] <= height <= PEAK_RATIO[1]):
        return None

    # The measurement itself, as in the paper: a line LINE_PX pixels thick across the middle
    # of the bar, and the distance from one peak's top to the other's. The fit above only says
    # where the two peaks are; each top is the brightest sample near it, placed between samples
    # by a parabola through it and its two neighbours.
    lane = np.arange(LINE_PX) - (LINE_PX - 1) / 2
    thin = ndimage.map_coordinates(pan, [row + t[:, None] * across[0] + lane[None, :] * along[0],
                                         col + t[:, None] * across[1] + lane[None, :] * along[1]], order=1).mean(axis=1)

    def summit(m):
        near_m = np.flatnonzero(np.abs(t - m) <= distance / 3)
        i = near_m[np.argmax(thin[near_m])]
        if not 0 < i < len(t) - 1:
            return float(t[i])
        bend = thin[i - 1] - 2 * thin[i] + thin[i + 1]
        return float(t[i] + (0.5 * (thin[i - 1] - thin[i + 1]) / bend if bend < 0 else 0.0) * (t[1] - t[0]))

    summits = (summit(m1), summit(m2))
    reading = abs(summits[1] - summits[0])
    if not (near <= reading <= far):
        return None
    # Each density followed along its own length (a line LINE_PX thick through its top), for
    # telling a continuous PSD from a row of dense projections and for finding gaps in a PSD.
    span = np.arange(-0.75 * length, 0.75 * length + 0.5, 0.5)
    runs = []
    for top in summits:
        rr = row + top * across[0] + span[:, None] * along[0] + lane[None, :] * across[0]
        cc = col + top * across[1] + span[:, None] * along[1] + lane[None, :] * across[1]
        inside = (rr.min(axis=1) >= 0) & (cc.min(axis=1) >= 0) & (rr.max(axis=1) <= pan.shape[0] - 1) & (cc.max(axis=1) <= pan.shape[1] - 1)
        runs.append(np.where(inside, ndimage.map_coordinates(pan, [rr, cc], order=1, mode="nearest").mean(axis=1), np.nan))

    out = {
        "row": float(row + (m1 + m2) / 2 * across[0]), "col": float(col + (m1 + m2) / 2 * across[1]),
        "theta": theta, "distance_um": float(reading * um), "exf": float(reading * um * 1000 / REF_NM),
        "fit_um": float(distance * um), "exf_fit": float(distance * um * 1000 / REF_NM), "summits": summits, "thin": thin, "runs": runs, "span": span, "base": float(b), "widths": (float(s1), float(s2)),
        "tops": (float(np.interp(summits[0], t, thin)), float(np.interp(summits[1], t, thin))), "ab_at": None,
        "r2": float(r2), "dip": dip, "ratio": height, "t": t, "profile": profile, "fit": fit, "win": (lo, hi),
        "peak": float(max(two_peaks(m1, *fit), two_peaks(m2, *fit))), "lines": grid,
        "ab_sigma": None, "ab_profile": None,
    }
    if ab is not None:
        # the synapse's footprint: both densities and a quarter of their spacing beyond, along the averaged stretch
        foot = (t >= min(m1, m2) - gap / 4) & (t <= max(m1, m2) + gap / 4)
        grid = ndimage.map_coordinates(ab, [rows, cols], order=1)
        out["ab_sigma"] = float((grid[foot].max() - noise[0]) / noise[1])
        out["ab_profile"] = grid.mean(axis=1)
        out["ab_at"] = float(t[foot][np.argmax(out["ab_profile"][foot])])    # where across the synapse the antibody peaks
    return out


def find_synapses(pan_raw, ab_raw, um):
    """Synapses in one 2D tile. Returns a list of measure_bar() results, positions in tile pixels."""
    from scipy import ndimage
    from skimage import feature, measure

    px_nm = EXF_NOMINAL / (um * 1000)
    valid = pan_raw > 0
    if valid.mean() < 0.05:
        return []
    pan = ndimage.gaussian_filter(pan_raw.astype(np.float32), SMOOTH_PX)
    ab = noise = None
    if ab_raw is not None:
        ab = ndimage.gaussian_filter(ab_raw.astype(np.float32), SMOOTH_PX)
        middle = float(np.median(ab[valid]))
        noise = (middle, 1.4826 * float(np.median(np.abs(ab[valid] - middle))) or 1.0)
    sigma = max(1.0, PSD_SIGMA_NM * px_nm)
    hessian = feature.hessian_matrix(pan, sigma=sigma, order="rc", use_gaussian_derivatives=False)
    upper, lower = feature.hessian_matrix_eigvals(hessian)       # upper >= lower at every pixel
    # a bright line curves down sharply across itself (lower << 0) and hardly along itself
    ridge = np.where((lower < 0) & (np.abs(upper) < 0.5 * np.abs(lower)), -lower, 0)
    bars = (ridge > np.percentile(ridge[valid], RIDGE_PERCENTILE)) & (pan > np.percentile(pan[valid], BRIGHT_PERCENTILE))
    ranks = np.sort(pan[valid][::37])        # a sample of the plane, to rank a brightness against
    found = []
    for region in measure.regionprops(measure.label(bars)):
        length = region.major_axis_length
        if not (PSD_LENGTH_NM[0] * px_nm <= length <= PSD_LENGTH_NM[1] * px_nm):
            continue
        if region.minor_axis_length > 0.6 * length:
            continue
        hit = measure_bar(pan, ab, noise, region.centroid[0], region.centroid[1], region.orientation, length, um)
        if hit:
            # how bright the brighter density is: the share of the plane that is dimmer
            hit["bright"] = float(np.searchsorted(ranks, hit["peak"]) / len(ranks))
            if hit["bright"] >= PSD_RANK:
                found.append(hit)
    return found


def spot_width(raw, row, col):
    """Width (FWHM, pixels) of one punctum: a round 2D Gaussian on a flat background fitted to
    the unsmoothed image, 13 x 13 pixels around it. None when it is not a clean single spot."""
    from scipy import optimize

    r, c, half = int(round(row)), int(round(col)), 6
    if r < half or c < half or r + half >= raw.shape[0] or c + half >= raw.shape[1]:
        return None
    patch = raw[r - half:r + half + 1, c - half:c + half + 1].astype(np.float64)
    yy, xx = np.mgrid[-half:half + 1, -half:half + 1]

    def blob(_, b, a, y0, x0, sd):
        return (b + a * np.exp(-((yy - y0) ** 2 + (xx - x0) ** 2) / (2 * sd * sd))).ravel()

    floor = float(np.median(patch))
    try:
        fit, _ = optimize.curve_fit(blob, None, patch.ravel(), p0=[floor, patch.max() - floor, 0, 0, 1.5],
                                    bounds=([0, 0, -2, -2, 0.4], [patch.max(), 4 * patch.max() + 1, 2, 2, 5]), maxfev=400)
    except (RuntimeError, ValueError):
        return None
    left = float(((patch.ravel() - blob(None, *fit)) ** 2).sum())
    total = float(((patch - patch.mean()) ** 2).sum())
    if total <= 0 or 1 - left / total < 0.6 or not 0.45 < fit[4] < 4.9:
        return None
    return 2.3548 * float(fit[4])


def find_spots(ab_raw, valid):
    """Antibody puncta in one tile, as [(row, col, FWHM in pixels or None)].

    A punctum is SPOT_MIN_PX or more connected pixels of the smoothed antibody channel that
    are AB_SIGMAS noise-widths above the plane's median: the same rule that makes a synapse
    antibody-positive. The width is measured on the brightest puncta that stand alone.
    """
    from scipy import ndimage, spatial
    from skimage import measure

    if valid.mean() < 0.05:
        return []
    ab = ndimage.gaussian_filter(ab_raw.astype(np.float32), SMOOTH_PX)
    middle = float(np.median(ab[valid]))
    sd = 1.4826 * float(np.median(np.abs(ab[valid] - middle))) or 1.0
    regions = [r for r in measure.regionprops(measure.label((ab > middle + AB_SIGMAS * sd) & valid), intensity_image=ab)
               if r.area >= SPOT_MIN_PX]
    if not regions:
        return []
    centres = np.array([r.centroid for r in regions])
    spots = [[float(r), float(c), None] for r, c in centres]
    if len(regions) > 1:
        apart = spatial.cKDTree(centres).query(centres, k=2)[0][:, 1]
    else:
        apart = np.array([np.inf])
    alone = [i for i in np.argsort([-r.intensity_max for r in regions]) if apart[i] > SPOT_ALONE_PX]
    for i in alone[:SPOT_FITS]:
        spots[i][2] = spot_width(ab_raw, *centres[i])
    return [tuple(sp) for sp in spots]


def antibody_numbers(hits, spots, area_px, um):
    """How well the antibody labels the synapses of one sample.

      labelled    share of the synapses found that carry antibody (sensitivity). The synapses
                  are found from the pan stain, without looking at the antibody.
      signal      median antibody signal, in noise-widths, at the synapses that carry it
      on synapse  share of all antibody puncta that sit on a synapse found (specificity)
      enrichment  that share divided by the share expected if the puncta fell at random,
                  which is the part of the field the synapses' neighbourhoods cover.
                  1 = no better than chance.
      resolution  width (FWHM) of the smallest isolated puncta, as imaged: the SPOT_SMALLEST-th
                  percentile of the widths. Most puncta are clusters of receptors and wider
                  than the optics; the smallest ones come closest to a point. Still an
                  upper bound on the lateral resolution.
    """
    from scipy import spatial

    out = {"puncta": len(spots)}
    marked = [h for h in hits if h["ab_sigma"] is not None and h["ab_sigma"] >= AB_SIGMAS]
    if hits:
        out["labelled_pct"] = 100 * len(marked) / len(hits)
    if marked:
        out["signal"] = float(np.median([h["ab_sigma"] for h in marked]))
    reach = ON_SYNAPSE_X * REF_NM * EXF_NOMINAL / (um * 1000)
    if spots and hits and area_px:
        on = 0
        for z in {sp[0] for sp in spots}:
            here = np.array([(h["row"], h["col"]) for h in hits if h["z"] == z])
            there = np.array([(sp[1], sp[2]) for sp in spots if sp[0] == z])
            if len(here):
                on += int((spatial.cKDTree(here).query(there)[0] <= reach).sum())
        chance = min(1.0, len(hits) * math.pi * reach * reach / area_px)
        out["on_pct"] = 100 * on / len(spots)
        out["enrichment"] = (on / len(spots)) / chance if chance > 0 else None
    widths = [sp[3] for sp in spots if sp[3] is not None]
    if len(widths) >= MIN_WIDTHS:
        out["fwhm_nm"], out["fwhm_n"] = float(np.percentile(widths, SPOT_SMALLEST)) * um * 1000, len(widths)
    return out


def unique(hits, radius):
    """Drop repeats: a synapse found from its PSD and again from its DPs, or in two tiles.
    The better-fitted one of any pair closer than `radius` pixels is kept."""
    kept = []
    for h in sorted(hits, key=lambda h: -h["r2"] * h["dip"]):
        if all(h["z"] != k["z"] or (h["row"] - k["row"]) ** 2 + (h["col"] - k["col"]) ** 2 > radius ** 2 for k in kept):
            kept.append(h)
    return kept


def analyse(name, zarr_dir, out_root, synaptic, max_planes):
    """Every synapse in one sample. Returns (hits, info, channels, pan, antibody channel or None,
    um, antibody numbers from antibody_numbers())."""
    series = zarr_dir / "0"
    info = core.read_series(series)
    channels, pan, antibodies, _ = eff.channel_roles(name, zarr_dir, out_root, info)
    if pan is None:
        raise ValueError("no pan stain found")
    raw = eff.open_level(series, info, 0)                         # lazy (c, z, y, x)
    z_um, _, um = eff.voxel_um(info, raw.shape[1:])
    if not um:
        raise ValueError("the image has no pixel size")
    z0, z1 = eff.gel_planes(eff.open_level(series, info, eff.macro_level(info))[pan])
    step = max(1, math.ceil(PLANE_GAP_UM / z_um)) if z_um else 1
    planes = list(range(z0, z1 + 1, step))
    if len(planes) > max_planes:
        planes = [planes[int(round(i))] for i in np.linspace(0, len(planes) - 1, max_planes)]
    ab_c = antibodies[0] if antibodies else None
    print(f"   pan = {channels[pan]['label']}, antibody = {channels[ab_c]['label'] if ab_c is not None else 'none'}; "
          f"{len(planes)} plane(s) between {z0} and {z1}; pixel {um * 1000:.0f} nm as imaged")
    pad = int(math.ceil(2 * PROFILE_HALF_NM * EXF_NOMINAL / (um * 1000)))
    ny, nx = raw.shape[2:]
    hits, spots, area = [], [], 0
    for z in planes:
        before = len(hits)
        for y in range(0, ny, TILE_PX):
            for x in range(0, nx, TILE_PX):
                ya, xa = max(0, y - pad), max(0, x - pad)
                yb, xb = min(ny, y + TILE_PX + pad), min(nx, x + TILE_PX + pad)
                tile = raw[:, z, ya:yb, xa:xb][[pan] + ([ab_c] if ab_c is not None else [])].compute()
                for h in find_synapses(tile[0], tile[1] if ab_c is not None else None, um):
                    h["row"], h["col"], h["z"] = h["row"] + ya, h["col"] + xa, z
                    if y <= h["row"] < y + TILE_PX and x <= h["col"] < x + TILE_PX:   # each tile keeps its own
                        hits.append(h)
                area += int((tile[0][y - ya:y - ya + TILE_PX, x - xa:x - xa + TILE_PX] > 0).sum())
                if ab_c is not None:
                    for r, c, width in find_spots(tile[1], tile[0] > 0):
                        if y <= r + ya < y + TILE_PX and x <= c + xa < x + TILE_PX:
                            spots.append((z, r + ya, c + xa, width))
        print(f"   slice {z}: {len(hits) - before} synapse profile(s)")
    hits = unique(hits, REF_NM * EXF_NOMINAL / (um * 1000))
    # A field of view is one camera frame: a plane of a stack, or one tile of a montage (the
    # tile grid is estimated from the image size, as elsewhere).
    tiled = "montage" in roi.classify(info)
    edges_y, edges_x = eff.tile_edges(ny, ny), eff.tile_edges(nx, nx)
    for h in hits:
        h["ab_pos"] = bool(synaptic and h["ab_sigma"] is not None and h["ab_sigma"] >= AB_SIGMAS)
        h["fov"] = (h["z"], int(np.searchsorted(edges_y, h["row"], side="right")),
                    int(np.searchsorted(edges_x, h["col"], side="right"))) if tiled else (h["z"], 0, 0)
    # Antibody-positive first, cleanest profiles first. The rest by how bright the density is:
    # brightness is what tells a PSD from a membrane, and it does not favour any distance.
    hits.sort(key=lambda h: (not h["ab_pos"], -h["r2"] * h["dip"]) if h["ab_pos"] else (True, -h["peak"]))
    return hits, info, channels, pan, ab_c, um, antibody_numbers(hits, spots, area, um) if ab_c is not None else {}


# --------------------------------------------------------------------------- pictures


def crop(plane, row, col, half):
    r, c = int(round(row)), int(round(col))
    return plane[max(0, r - half):r + half + 1, max(0, c - half):c + half + 1]


def stretch(img, lo, hi):
    return np.clip((img - lo) / max(hi - lo, 1e-6), 0, 1)


def scale_bar(ax, um, width_px):
    """A bar in micrometres as imaged, bottom left, as on the screenshots."""
    target = width_px * um / 4
    length = max((n for n in (0.2, 0.5, 1, 2, 5, 10) if n <= target), default=0.2)
    ax.plot([2, 2 + length / um], [width_px - 3] * 2, color="white", lw=3, solid_capstyle="butt")
    ax.text(2, width_px - 5, f"{length:g} µm", color="white", fontsize=8, va="bottom")


FIT_BLUE = "#4DA3FF"   # the fitted curve: a blue light enough to read on the dark background


def card(png, number, name, h, tile, um, pan_ch, ab_ch):
    """One synapse: pan, antibody, overlay, and the profile with its fit."""
    half = tile["half"]
    pan = stretch(tile["pan"], *np.percentile(tile["pan"], [1, 99.8]))
    fig, axes = eff.figure(1, 4, figsize=(13, 3.4), gridspec_kw={"width_ratios": [1, 1, 1, 1.5]})
    rgb = np.dstack([pan] * 3)
    panels = [(rgb, f"Pan {pan_ch['label']}")]
    if tile["ab"] is not None:
        ab = stretch(tile["ab"], tile["ab_lo"], tile["ab_hi"])
        tint = np.asarray(eff.colour_of(ab_ch), np.float32)
        panels += [(ab[..., None] * tint, f"Antibody {ab_ch['label']}"),
                   (np.clip(rgb * 0.8 + ab[..., None] * tint, 0, 1), "Overlay")]
    for ax, (img, title) in zip(axes, panels):
        ax.imshow(img, interpolation="nearest")
        ax.set_title(title, color=eff.INK, fontsize=9)
        scale_bar(ax, um, img.shape[1])
    for ax in axes[:3]:
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
    # the measured axis, drawn on the overlay only so the other panels stay clean
    across = np.array([-math.sin(h["theta"]), math.cos(h["theta"])])
    reach = h["distance_um"] / um
    axes[2 if tile["ab"] is not None else 0].plot([half - reach * across[1], half + reach * across[1]],
                                                  [half - reach * across[0], half + reach * across[0]],
                                                  color="#FF9900", lw=1, ls=":")
    ax = axes[3]
    t_um = h["t"] * um
    ax.plot(t_um, h["profile"], color=eff.MUTED, lw=1.2, label="Pan, averaged along the bar (fitted)")
    ax.plot(t_um, h["thin"], color="white", lw=2, label=f"Pan, {LINE_PX}-pixel line (measured)")
    inside = (h["t"] >= h["win"][0]) & (h["t"] <= h["win"][1])
    ax.plot(t_um[inside], two_peaks(h["t"][inside], *h["fit"]), color=FIT_BLUE, lw=1.8, label="Two-Gaussian fit")
    for m in h["summits"]:
        ax.axvline(m * um, color="white", lw=1, ls=":")
    if h["ab_profile"] is not None:
        # the antibody along the same line, dashed, rescaled to the pan curve's height so its
        # position can be compared with the two pan peaks (its own counts are not on this axis)
        scaled = h["ab_profile"] - h["ab_profile"].min()
        scaled = scaled / max(scaled.max(), 1e-6) * (h["profile"].max() - h["profile"].min()) + h["profile"].min()
        ax.plot(t_um, scaled, color=eff.colour_of(ab_ch), lw=1.3, ls="--", label="Antibody (rescaled)")
    ax.set_xlabel("Distance across the synapse (µm, as imaged)", color=eff.INK, fontsize=9)
    ax.set_ylabel("Counts", color=eff.INK, fontsize=9)
    ax.grid(color="#2A2A2A", lw=0.6)
    # under the plot, in two columns: inside the axes it sat on top of the peaks
    ax.legend(frameon=False, labelcolor=eff.INK, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=2)
    ax.set_title(f"Peak to peak {h['distance_um']:.2f} µm ÷ {REF_NM:g} nm:  ExF = {h['exf']:.1f}     "
                 f"(Gaussian fit: ExF = {h['exf_fit']:.1f}, R² {h['r2']:.2f})", color=eff.INK, fontsize=9)
    sign = "" if h["ab_sigma"] is None else f"   Antibody {h['ab_sigma']:.1f} σ" + (" (positive)" if h["ab_pos"] else "")
    fig.suptitle(f"#{number}   {name}   Slice {h['z']}   x {h['col']:.0f}, y {h['row']:.0f}{sign}", color=eff.INK, fontsize=10)
    fig.tight_layout()
    eff.save(fig, png)


def contact_sheet(png, name, shown, um, ab_ch):
    """Every logged synapse of a sample at a glance, numbered as in the log."""
    cols = 10
    rows = math.ceil(len(shown) / cols)
    fig, axes = eff.figure(rows, cols, figsize=(cols * 1.5, rows * 1.7 + 0.5), squeeze=False)
    for ax in axes.ravel():
        ax.axis("off")
    for ax, (number, h, tile) in zip(axes.ravel(), shown):
        pan = stretch(tile["pan"], *np.percentile(tile["pan"], [1, 99.8]))
        rgb = np.dstack([pan] * 3) * 0.8
        if tile["ab"] is not None:
            rgb = rgb + stretch(tile["ab"], tile["ab_lo"], tile["ab_hi"])[..., None] * np.asarray(eff.colour_of(ab_ch), np.float32)
        ax.imshow(np.clip(rgb, 0, 1), interpolation="nearest")
        ax.set_title(f"#{number}{'  AB+' if h['ab_pos'] else ''}   ExF = {h['exf']:.1f}", color=eff.INK, fontsize=8)
    scale_bar(axes[0][0], um, shown[0][2]["pan"].shape[1])
    fig.suptitle(f"{name}: Pan (grey) with antibody overlaid. Number as in the Synapse Log, then its expansion factor (ExF)",
                 color=eff.INK, fontsize=10)
    fig.tight_layout()
    eff.save(fig, png)


def synapse_link(out_root, name, number, h, write=True):
    """A Neuroglancer link that opens on one synapse: the sample's own view, moved and zoomed in.
    With write=False nothing is saved on the server; a view saved earlier is still linked."""
    saved = out_root / core.STATE_DIR / "synapses" / name / f"synapse_{number:04d}.json"
    if not write:
        return core.state_url(saved) if saved.is_file() else None
    view = roi.view_file(out_root, name)
    if not view.is_file():
        return None
    state = json.loads(view.read_text(encoding="utf-8"))
    place = {"x": h["col"], "y": h["row"], "z": h["z"]}
    state["position"] = [place.get(n, p) for n, p in zip(state.get("dimensions", {}), state.get("position", []))]
    state["crossSectionScale"] = 0.15
    state["layout"] = "xy"
    return core.state_to_link(state, out_root / core.STATE_DIR / "synapses" / name / f"synapse_{number:04d}.json")


# --------------------------------------------------------------------------- the log


def summarise(values):
    """'mean ± s.d. (median, n)' of expansion factors."""
    if not len(values):
        return "–"
    v = np.asarray(values, float)
    return f"{v.mean():.1f} ± {v.std(ddof=1) if len(v) > 1 else 0:.1f}  (median {np.median(v):.1f}, n = {len(v)})"


def read_decisions(path):
    """What was typed in the log: {(sample, slice): [(x, y, accept, notes)]}."""
    kept = {}
    if not path.is_file():
        return kept
    ws = load_workbook(path)["Synapses"]
    cols = {c.value: i for i, c in enumerate(ws[HEADER_ROW]) if c.value}
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        if row[cols["Sample"]] is None:
            continue
        kept.setdefault((row[cols["Sample"]], row[cols["Slice"]]), []).append(
            (row[cols["X (px)"]], row[cols["Y (px)"]], row[cols["Accept"]], row[cols["Notes"]]))
    return kept


def decision_for(kept, sample, h, radius=8):
    for x, y, accept, notes in kept.get((sample, h["z"]), []):
        if abs(x - h["col"]) <= radius and abs(y - h["row"]) <= radius:
            return accept, notes
    return None, None


METHOD = [
    "How these numbers were made",
    "",
    "Reference: M'Saad, Cairns et al., Nat. Biotechnol. 2025 (pan-ExM-t). Expansion factor = DP–PSD distance ÷ {ref:g} nm,",
    "where DP–PSD is the centre-to-centre distance between the presynaptic dense projections and the postsynaptic density.",
    "The paper (Supplementary Methods) read it peak to peak from 2-pixel-thick line profiles drawn by hand in FIJI, and reported",
    "the mean and s.d. across fields of view. Here the program finds the synapses and places the lines; the reading is the same.",
    "",
    "1. Pan channel, full resolution, one plane at a time (planes closer than {gap:g} µm are skipped so no synapse is counted twice;",
    "   planes outside the gel are left out). Both channels are smoothed with a {smooth:g}-pixel Gaussian.",
    "2. Candidate bars: ridge filter (most negative Hessian eigenvalue, sigma = {sigma:g} nm at a nominal {nominal:g}x), strongest",
    "   {ridge:g}% of pixels that are also brighter than {bright:g}% of the plane; connected bars {len0:g}–{len1:g} nm long (at {nominal:g}x).",
    "3. Profile across each bar, {half:g} nm either side (at {nominal:g}x), averaged along the bar over up to ±{avg:g} nm.",
    "4. Second peak searched at {ref:g} nm x {s0:g}–{s1:g}. This window bounds the answer: no synapse can give an expansion",
    "   factor outside {s0:g}–{s1:g}. It is wider than the expected 16–24 on purpose; check the histogram is not piled at its edges.",
    "5. Two Gaussians + flat baseline fitted to that averaged profile, to decide whether this is a synapse and where its two",
    "   peaks are. Kept when R² ≥ {r2:g}, the dip between peaks ≥ {dip:g} of the smaller peak, and the peak heights are within",
    "   a factor {pr1:g} of each other.",
    "   THE MEASUREMENT: a line {line} pixels thick across the middle of the bar; the top of each peak is its brightest sample,",
    "   placed between samples by a parabola through it and its two neighbours; top-to-top distance ÷ {ref:g} nm = ExF.",
    "   The distance between the fitted Gaussian centres ÷ {ref:g} nm is kept as 'ExF (Gaussian fit)'; it reads about 10% higher.",
    "   The brighter peak must also be brighter than {rank:g}% of the plane. Without this rule most double peaks were pairs of",
    "   membranes, which piled up at 26–30x. The cut-off was set from GluA2-positive synapses (260831_WT4_GluA2ms 3x3 montage):",
    "   95% of them ranked above 99.6%, the other double peaks had a median rank of 97.8%.",
    "6. Antibody-positive: brightest smoothed antibody pixel on the synapse ≥ {ab:g} noise-widths (1.4826 x MAD) above the",
    "   plane's median. Only used when the antibody was confirmed as synaptic at the prompt.",
    "7. The {top} best synapses per sample are logged: antibody-positive first, then by fit quality (R² x dip).",
    "8. 'ExF' on the Summary sheet is the one number quoted for the sample. It uses the accepted synapses if at least {minacc}",
    "   were accepted; otherwise the antibody-positive synapses if the antibody is synaptic and there are at least {minpos};",
    "   otherwise all double peaks. 'ExF Basis' says which. The last kind reads high and should not be quoted.",
    "   As in the paper it is the mean ± s.d. across fields of view: each field's synapses are averaged first. A field of view",
    "   is one plane of a stack or one tile of a montage (tile grid estimated), and needs {minfov} synapses to count. With fewer",
    "   than two such fields the mean ± s.d. is taken across synapses instead, and 'ExF Basis' says so.",
    "",
    "Antibody labelling (provisional)",
    "• Antibody puncta: {spotpx}+ connected pixels of the smoothed antibody channel ≥ {ab:g} noise-widths above the plane's median.",
    "• AB Labelled Synapses (%): share of the synapses found (from the pan stain alone) that are antibody-positive. Membrane",
    "  pairs mistaken for synapses are in the denominator, so this is a lower bound on the true share.",
    "• AB Puncta on Synapses (%): share of all puncta within {onx:g} nominal DP–PSD distances of a synapse found. Only side-on",
    "  synapses are found, so most real synapses are missed and this is far below the true on-target share.",
    "• AB Enrichment (× chance): that share ÷ the share of the field those neighbourhoods cover. 1 = puncta fall at random",
    "  with respect to synapses. This is the number to compare between antibodies; it does not depend on how many synapses were found.",
    "",
    "Resolution",
    "• Resolution FWHM: width of the smallest isolated antibody puncta (the {small}th percentile of the widths), from a round 2D",
    "  Gaussian fitted to the unsmoothed image. The paper (Extended Data Fig. 5) fitted Gaussians to hand-picked sub-diffraction",
    "  antibody features and got 290 ± 60 nm as imaged. Puncta here are mostly receptor clusters and wider than the optics, so",
    "  this is an upper bound and depends on the antibody. Corrected = as imaged ÷ ExF. Lateral only.",
    "• Pixel (corrected) = pixel size as imaged ÷ ExF.",
    "",
    "Caveats",
    "• 81.9 nm is itself a measurement (82 ± 26 nm s.d. between synapses in the paper), so single synapses scatter by ~30%;",
    "  the mean over many synapses is what carries the expansion factor.",
    "• Only synapses lying side-on in the image plane give the true distance. Tilted ones read short. The dip criterion",
    "  favours side-on synapses but cannot guarantee it: reject tilted or doubtful ones in the Accept column.",
    "• Which peak is the PSD and which the DPs is not decided; the distance does not depend on it.",
    "• Synapses without antibody are found from the pan stain alone and still include some membrane pairs (they read high,",
    "  25–30x). The antibody-positive and the accepted values are the ones to trust; compare the median with the mean.",
    "• 'ExF, all found' uses every profile that passed, unreviewed. 'ExF, accepted' uses only rows marked Y.",
]


SUMMARY = [
    "Sample", "Antibody", "Synaptic", "Planes", "Synapses found", "AB+",
    "ExF", "ExF s.d.", "ExF Fields", "ExF n", "ExF Basis", "ExF (Gaussian fit)",
    "ExF, all found", "ExF, AB+ only", "ExF, logged", "ExF, accepted", "Accepted / Rejected / Undecided",
    "Found outside 16–24 (%)",
    "Pixel (nm, as imaged)", "Pixel (nm, corrected)", "Resolution FWHM (nm, as imaged)", "Resolution (nm, corrected)",
    "Resolution n", "AB Labelled Synapses (%)", "AB Signal at Synapses (σ)", "AB Puncta", "AB Puncta on Synapses (%)",
    "AB Enrichment (× chance)", "Contact Sheet",
]


def sample_exf(accepted, positive, every, synaptic):
    """The one expansion factor quoted for a sample, from the best evidence there is, as the
    paper reports it: the mean and s.d. across fields of view of each field's mean.

    Each argument is a list of synapses. Returns a dict for the Summary sheet.
    """
    for group, enough, basis in ((accepted, MIN_ACCEPTED, "accepted by a person"),
                                 (positive if synaptic else [], MIN_FOR_EXF, "antibody-positive synapses, unreviewed"),
                                 (every, 1, "all double peaks, pan only, unreviewed (reads high)")):
        if len(group) < enough:
            continue
        fields = {}
        for h in group:
            fields.setdefault(h["fov"], []).append(h["exf"])
        means = [float(np.mean(v)) for v in fields.values() if len(v) >= MIN_PER_FOV]
        used = sum(len(v) for v in fields.values() if len(v) >= MIN_PER_FOV)
        if len(means) >= 2:
            mean, sd, how, n_fields, n = float(np.mean(means)), float(np.std(means, ddof=1)), "across fields of view", len(means), used
        else:   # not enough fields with MIN_PER_FOV synapses each: fall back on the synapses themselves
            v = np.asarray([h["exf"] for h in group])
            mean, sd, how, n_fields, n = float(v.mean()), float(v.std(ddof=1)) if len(v) > 1 else 0.0, "across synapses", None, len(v)
        return {"ExF": round(mean, 1), "ExF s.d.": round(sd, 1), "ExF Fields": n_fields, "ExF n": n,
                "ExF Basis": f"{basis}; s.d. {how}",
                "ExF (Gaussian fit)": round(float(np.mean([h["exf_fit"] for h in group])), 1)}
    return {}


def write_log(path, rows, summary):
    wb = Workbook()
    ws = wb.active
    ws.title = "Synapses"
    ws["A1"] = "Synapse Log"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Type Y or N under Accept (and anything under Notes), save, close, and run Measure_ExF.bat again. Other columns are rewritten."
    fill, mine = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="E4F1E0")
    for i, (header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(HEADER_ROW, i, header)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    for r, row in enumerate(rows, start=HEADER_ROW + 1):
        for i, header in enumerate(HEADERS, start=1):
            value = row.get(header)
            cell = ws.cell(r, i, value)
            if header in ("Picture", "Neuroglancer") and value:
                cell.value, cell.hyperlink = "open", value
                cell.font = Font(color="0563C1", underline="single")
            elif header in ("Accept", "Notes"):
                cell.fill = mine
    choice = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(choice)
    accept = get_column_letter(HEADERS.index("Accept") + 1)
    choice.add(f"{accept}{HEADER_ROW + 1}:{accept}{HEADER_ROW + max(len(rows), 1)}")
    ws.freeze_panes = ws.cell(HEADER_ROW + 1, 3)

    ws = wb.create_sheet("Summary")
    heads = SUMMARY
    for i, head in enumerate(heads, start=1):
        cell = ws.cell(1, i, head)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), fill
        ws.column_dimensions[get_column_letter(i)].width = 44 if i == 1 else 30 if "ExF" in head else 16
    for r, line in enumerate(summary, start=2):
        for i, head in enumerate(heads, start=1):
            ws.cell(r, i, line.get(head))

    ws = wb.create_sheet("Method")
    ws.column_dimensions["A"].width = 130
    values = dict(ref=REF_NM, gap=PLANE_GAP_UM, smooth=SMOOTH_PX, sigma=PSD_SIGMA_NM, nominal=EXF_NOMINAL,
                  ridge=100 - RIDGE_PERCENTILE, bright=BRIGHT_PERCENTILE, len0=PSD_LENGTH_NM[0], len1=PSD_LENGTH_NM[1],
                  half=PROFILE_HALF_NM, avg=AVERAGE_HALF_NM, s0=EXF_SEARCH[0], s1=EXF_SEARCH[1], r2=FIT_R2, dip=MIN_DIP,
                  pr1=PEAK_RATIO[1], ab=AB_SIGMAS, top=TOP_N, rank=100 * PSD_RANK, minacc=MIN_ACCEPTED,
                  minpos=MIN_FOR_EXF, spotpx=SPOT_MIN_PX, onx=ON_SYNAPSE_X, line=LINE_PX, minfov=MIN_PER_FOV, small=SPOT_SMALLEST)
    for r, line in enumerate(METHOD, start=1):
        ws.cell(r, 1, line.format(**values)).font = Font(
            bold=(r == 1 or line in ("Caveats", "Antibody labelling (provisional)", "Resolution")))
    wb.save(path)


# --------------------------------------------------------------------------- run


def run(args):
    core.load_settings()
    out_root = Path(core.DEFAULT_OUT)
    if not out_root.is_dir():
        sys.exit(f"Cannot reach the zarr folder {out_root}")
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(core.SETTINGS_FILE, encoding="utf-8-sig")
    base = Path(args.out or cfg.get("paths", "synapse_folder", fallback="").strip()
                or Path(core.DEFAULT_SHOTS).parent / "Zarr Synapse ExF")
    base.mkdir(parents=True, exist_ok=True)
    log = base / LOG_NAME
    print(f"Zarr folder : {out_root}\nSynapse log : {log}\n")
    try:
        kept = read_decisions(log)
        if log.is_file():
            open(log, "ab").close()
    except PermissionError:
        sys.exit(f"{LOG_NAME} is open in Excel. Close it and run this again.")

    text = args.text
    if text is None and sys.stdin.isatty():
        text = roi.ask("Which samples? Type text their names must contain (for example GluA2), or Enter for all: ")
    chosen, left_out = [], []
    for name, zarr_dir in roi.samples(out_root):
        if (text or "").lower() in name.lower():
            why = ignored(name)
            (left_out if why else chosen).append((name, why or zarr_dir))
    for name, why in left_out:
        print(f"   left out: {name}   ({why})")
    if not chosen:
        sys.exit("No sample to measure.")

    # Ask about each antibody before any work is done: is it a synaptic antibody?
    synaptic = {}
    for name, _ in chosen:
        for label, side in antibodies_in(name) or [("the antibody in " + name, "?")]:
            if label in synaptic:
                continue
            if args.synaptic == "auto":
                synaptic[label] = side in ("pre", "post")      # as listed in ANTIBODIES; unknown = not synaptic
            elif args.synaptic:
                synaptic[label] = args.synaptic == "y"
            elif side == "?" or not sys.stdin.isatty():
                synaptic[label] = roi.ask_choice(f"Is {label} a synaptic antibody? [Y/N]: ", "YN", "N") == "Y" \
                    if sys.stdin.isatty() else False
            else:
                guess = "Y" if side else "N"
                where = f"{side}synaptic" if side else "not synaptic"
                synaptic[label] = roi.ask_choice(f"{label}: I take it as {where}. Is it a synaptic antibody? "
                                                 f"[Y/N] (Enter = {guess}): ", "YN", guess) == "Y"
    print("\n" + "\n".join(f"   {k}: {'synaptic, positive synapses listed first' if v else 'not synaptic, antibody not used for ranking'}"
                           for k, v in synaptic.items()) + f"\n\n{len(chosen)} sample(s) to go through.\n")

    rows, summary = [], []
    for n, (name, zarr_dir) in enumerate(chosen, start=1):
        print(f"[{n}/{len(chosen)}] {name}")
        named = antibodies_in(name)
        is_synaptic = any(synaptic.get(label) for label, _ in named) if named else synaptic.get("the antibody in " + name, False)
        try:
            hits, info, channels, pan, ab_c, um, ab_n = analyse(name, zarr_dir, out_root, is_synaptic, args.max_planes)
        except Exception as e:
            print(f"   SKIPPED: {e}\n")
            continue
        folder = base / name
        folder.mkdir(exist_ok=True)
        with open(folder / f"{name}_synapses.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Slice", "X (px)", "Y (px)", "DP-PSD peak to peak (µm, as imaged)", "ExF (peak to peak)",
                        "ExF (Gaussian fit)", "Field of View", "Fit R²", "Dip", "Peak Ratio", "Brightness Rank", "AB Signal (σ)", "AB+"])
            for h in hits:
                w.writerow([h["z"], round(h["col"], 1), round(h["row"], 1), round(h["distance_um"], 4), round(h["exf"], 2),
                            round(h["exf_fit"], 2), "z{}-r{}-c{}".format(*h["fov"]),
                            round(h["r2"], 3), round(h["dip"], 3), round(h["ratio"], 3), round(h["bright"], 4),
                            "" if h["ab_sigma"] is None else round(h["ab_sigma"], 2), "Y" if h["ab_pos"] else ""])
        logged = hits[:TOP_N]
        raw = eff.open_level(zarr_dir / "0", info, 0)
        half = int(round(2 * PROFILE_HALF_NM * EXF_NOMINAL / (um * 1000) * 0.75))
        ab_ch = channels[ab_c] if ab_c is not None else None
        label = ", ".join(a for a, _ in named) or (ab_ch["label"] if ab_ch else "")
        shown, decisions = [], []
        from scipy import ndimage
        for number, h in enumerate(logged, start=1):
            r, c = int(round(h["row"])), int(round(h["col"]))
            wide = 6 * half                                     # antibody noise is judged over a wider field
            block = raw[:, h["z"], max(0, r - wide):r + wide + 1, max(0, c - wide):c + wide + 1].compute().astype(np.float32)
            rr, cc = r - max(0, r - wide), c - max(0, c - wide)
            tile = {"half": half, "pan": crop(ndimage.gaussian_filter(block[pan], SMOOTH_PX), rr, cc, half), "ab": None}
            if ab_c is not None:
                ab = ndimage.gaussian_filter(block[ab_c], SMOOTH_PX)
                middle = float(np.median(ab))
                sd = 1.4826 * float(np.median(np.abs(ab - middle))) or 1.0
                tile.update(ab=crop(ab, rr, cc, half), ab_lo=middle + sd, ab_hi=middle + max(12 * sd, 1))
            png = folder / f"{name}_synapse-{number:03d}_z{h['z']:04d}.png"
            card(png, number, name, h, tile, um, channels[pan], ab_ch)
            accept, notes = decision_for(kept, name, h)
            decisions.append(accept)
            shown.append((number, h, tile))
            rows.append({
                "#": number, "Sample": name, "Slice": h["z"], "X (px)": round(h["col"]), "Y (px)": round(h["row"]),
                "Antibody": label, "AB Signal (σ)": None if h["ab_sigma"] is None else round(h["ab_sigma"], 1),
                "AB+": "Y" if h["ab_pos"] else None, "DP–PSD (µm, as imaged)": round(h["distance_um"], 3),
                "ExF (peak to peak)": round(h["exf"], 1), "ExF (Gaussian fit)": round(h["exf_fit"], 1),
                "In 16–24": "Y" if EXF_EXPECTED[0] <= h["exf"] <= EXF_EXPECTED[1] else "no",
                "Fit R²": round(h["r2"], 3), "Dip": round(h["dip"], 2), "Peak Ratio": round(h["ratio"], 2),
                "Accept": accept, "Notes": notes, "Picture": str(png),
                "Neuroglancer": synapse_link(out_root, name, number, h, write=not args.no_links),
            })
        if shown:
            contact_sheet(folder / f"{name}_synapses_contact-sheet.png", name, shown, um, ab_ch)
        every = [h["exf"] for h in hits]
        outside = [e for e in every if not EXF_EXPECTED[0] <= e <= EXF_EXPECTED[1]]
        accepted = [h for h, d in zip(logged, decisions) if str(d or "").strip().upper() == "Y"]
        yes = [h["exf"] for h in accepted]
        no = sum(str(d or "").strip().upper() == "N" for d in decisions)
        quoted = sample_exf(accepted, [h for h in hits if h["ab_pos"]], hits, is_synaptic)
        exf = quoted.get("ExF")
        fwhm = ab_n.get("fwhm_nm")
        rounded = lambda v, d=1: None if v is None else round(v, d)
        summary.append({
            **quoted,
            "Pixel (nm, as imaged)": round(um * 1000, 1), "Pixel (nm, corrected)": rounded(um * 1000 / exf) if exf else None,
            "Resolution FWHM (nm, as imaged)": rounded(fwhm, 0),
            "Resolution (nm, corrected)": rounded(fwhm / exf) if fwhm and exf else None, "Resolution n": ab_n.get("fwhm_n"),
            "AB Labelled Synapses (%)": rounded(ab_n.get("labelled_pct")), "AB Signal at Synapses (σ)": rounded(ab_n.get("signal")),
            "AB Puncta": ab_n.get("puncta"), "AB Puncta on Synapses (%)": rounded(ab_n.get("on_pct")),
            "AB Enrichment (× chance)": rounded(ab_n.get("enrichment")),
            "Contact Sheet": str(folder / f"{name}_synapses_contact-sheet.png") if shown else None,
            "Sample": name, "Antibody": label, "Synaptic": "yes" if is_synaptic else "no",
            "Planes": len({h["z"] for h in hits}), "Synapses found": len(hits), "AB+": sum(h["ab_pos"] for h in hits),
            "ExF, all found": summarise(every), "ExF, AB+ only": summarise([h["exf"] for h in hits if h["ab_pos"]]),
            "ExF, logged": summarise([h["exf"] for h in logged]), "ExF, accepted": summarise(yes),
            "Accepted / Rejected / Undecided": f"{len(yes)} / {no} / {len(logged) - len(yes) - no}",
            "Found outside 16–24 (%)": round(100 * len(outside) / len(every), 1) if every else None,
        })
        s = summary[-1]
        print(f"   {s['Synapses found']} synapse profile(s), {s['AB+']} antibody-positive\n"
              f"   ExF, all found : {s['ExF, all found']}\n   ExF, AB+ only  : {s['ExF, AB+ only']}\n"
              f"   ExF, accepted  : {s['ExF, accepted']}\n"
              f"   ExF quoted     : {s.get('ExF')} ± {s.get('ExF s.d.')}  ({s.get('ExF Basis')}; fields {s.get('ExF Fields')}, "
              f"n = {s.get('ExF n')}); Gaussian fit {s.get('ExF (Gaussian fit)')}\n"
              f"   antibody       : {s['AB Labelled Synapses (%)']}% of synapses labelled; {s['AB Puncta on Synapses (%)']}% of "
              f"{s['AB Puncta']} puncta on synapses ({s['AB Enrichment (× chance)']}x chance)\n"
              f"   resolution     : {s['Resolution FWHM (nm, as imaged)']} nm as imaged, {s['Resolution (nm, corrected)']} nm corrected "
              f"(n = {s['Resolution n']}); pixel {s['Pixel (nm, corrected)']} nm corrected\n")
    # rows of samples not gone through this time stay in the log
    done = {s["Sample"] for s in summary}
    if log.is_file():
        old = load_workbook(log)
        heads = [c.value for c in old["Synapses"][HEADER_ROW]]
        for sheet, key, target in (("Synapses", HEADER_ROW, rows), ("Summary", 1, summary)):
            names = [c.value for c in old[sheet][key]]
            for values in old[sheet].iter_rows(min_row=key + 1):
                row = {h: (c.hyperlink.target if c.hyperlink else c.value) for h, c in zip(names, values)}
                if row.get("Sample") and row["Sample"] not in done:
                    target.append(row)
    write_log(log, rows, summary)
    print(f"Done. Open {log}, mark Accept Y or N, and run this again for the accepted-only expansion factor.")
    return 0


def main():
    p = argparse.ArgumentParser(description="Find synapses in the pan stain and measure the expansion factor (DP-PSD / 82 nm).")
    p.add_argument("text", nargs="?", help="go through the samples whose names contain this (omit to be asked)")
    p.add_argument("--synaptic", choices=("y", "n", "auto"),
                   help="answer the 'is it a synaptic antibody' question in advance; auto = as listed in ANTIBODIES")
    p.add_argument("--max-planes", type=int, default=12, help="at most this many planes per sample, evenly spread (default 12)")
    p.add_argument("--no-links", action="store_true", help="do not save Neuroglancer views of the synapses on the server")
    p.add_argument("--out", help="folder for the log and pictures, instead of the one in settings.ini")
    return run(p.parse_args())


if __name__ == "__main__":
    sys.exit(main())
