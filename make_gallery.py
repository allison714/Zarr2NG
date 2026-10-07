"""
Build the gallery page ('2026 AC Gallery.html') in the screenshot folder: one page, opened in a browser, for
looking through every sample quickly.

  - a grid of thumbnails, one card per sample
  - search, filter by type, sort by name / date / antibody density / rating
  - click a card for all of that sample's pictures, its metadata, and its Neuroglancer link

It is built from what is already on disk - the pictures in each sample's subfolder, the
screenshot log (ratings, notes), the image list (metadata, links) and the synapse log
(expansion factor, resolution, antibody labelling; see synapse_exf.py) - so it needs no new
exports and can be rebuilt at any time. Samples or pictures that do not exist yet are simply
left out. Small JPEG copies of the pictures are kept in '_gallery_thumbs' so the page loads
fast; the full-size PNG opens when a picture is clicked.
"""

import datetime as dt
import html
import json
import re
import sys
import urllib.parse
from pathlib import Path

from openpyxl import load_workbook

import convert_to_zarr as core
import roi_log

GALLERY_TITLE = "2026 AC Gallery"   # shown at the top of the page; also the file name
GALLERY_NAME = GALLERY_TITLE + ".html"
OLD_NAMES = ["Zarr Gallery.html"]    # earlier names of the page, removed when it is rebuilt
THUMB_DIR = "_gallery_thumbs"
THUMB_PX = 420
SYNAPSE_THUMB_PX = 1100   # synapse panels are wide strips, so their small copies are wider
SYNAPSE_FOLDER = "Zarr Synapse ExF"   # beside the screenshot folder; written by Measure_ExF.bat
# the order pictures appear in a sample's page; the first one found is the card's picture
VIEW_ORDER = ["Volume", "MIP-XY", "MIP-XZ", "MIP-YZ", "Overview", "Full-image"] + \
             [f"{kind}-{q:02d}pct" for q in roi_log.PLANE_PERCENTS for kind in ("Plane", "Zoom-200pct")] + \
             ["Zoom-200pct", "Seam-crop", "Crop-centre", "Crop-top-left", "Crop-top-right",
              "Crop-bottom-left", "Crop-bottom-right"]
CARD_ORDER = ["Volume", "Plane-50pct", "Overview", "Full-image", "MIP-XY"]


def nice_view(tag):
    """'Zoom-200pct-25pct' -> '200% close-up at 25% depth'."""
    m = re.fullmatch(r"Plane-(\d+)pct", tag)
    if m:
        return f"Plane at {int(m.group(1))}% depth"
    m = re.fullmatch(r"Zoom-(\d+)pct-(\d+)pct", tag)
    if m:
        return f"{m.group(1)}% close-up at {int(m.group(2))}% depth"
    m = re.fullmatch(r"Zoom-(\d+)pct", tag)
    if m:
        return f"{m.group(1)}% close-up"
    return {"MIP-XY": "Projection, top view (XY)", "MIP-XZ": "Projection, side view (XZ)",
            "MIP-YZ": "Projection, side view (YZ)", "Volume": "Tilted volume",
            "Full-image": "Full image"}.get(tag, tag.replace("-", " "))


def thumbnail(png, dest, size=THUMB_PX):
    """Make (or reuse) the small JPEG copy of a picture. Returns False if it cannot be read."""
    from PIL import Image

    try:
        if dest.is_file() and dest.stat().st_mtime >= png.stat().st_mtime:
            return True
        img = Image.open(png).convert("RGB")
        img.thumbnail((size, size))
        dest.parent.mkdir(parents=True, exist_ok=True)
        img.save(dest, "JPEG", quality=82)
        return True
    except Exception:
        return False   # still being written, or not a picture


def read_metadata():
    """{image name: metadata} from the image list, and {name: [log rows]} from the screenshot log."""
    meta = {}
    try:
        ws = load_workbook(core.DEFAULT_LIST)["Images"]
        header_row = next(r for r in range(1, 12) if "Image Name" in [c.value for c in ws[r]])
        heads = [c.value for c in ws[header_row]]
        for r in range(header_row + 1, ws.max_row + 1):
            row = {h: ws.cell(r, i + 1).value for i, h in enumerate(heads) if h}
            link = ws.cell(r, heads.index("Image Name") + 1).hyperlink
            if row.get("Image Name"):
                row["url"] = None if link is None else (link.target or "") + ("#" + link.location if link.location else "")
                meta[row["Image Name"]] = row
    except Exception as e:
        print(f"(image list not read: {e})")
    logged = {}
    try:
        for row in roi_log.read_log(Path(core.DEFAULT_SHOTS) / roi_log.LOG_NAME):
            logged.setdefault(row.get(roi_log.HEADERS[0]), []).append(row)
    except Exception as e:
        print(f"(screenshot log not read: {e})")
    return meta, logged


def synapse_pictures(base, name):
    """The synapse panels of one sample, in the order they are shown: the contact sheet
    first, then each synapse by its number in the Synapse Log. Empty if none were made."""
    folder = base.parent / SYNAPSE_FOLDER / name
    if not folder.is_dir():
        return []
    found = []
    for png in folder.glob("*.png"):
        sheet = png.name == f"{name}_synapses_contact-sheet.png"
        m = re.fullmatch(re.escape(name) + r"_synapse-(\d+)_z(\d+)\.png", png.name)
        if not sheet and not m:
            continue
        thumb = base / THUMB_DIR / "_synapses" / name / (png.stem + ".jpg")
        if not thumbnail(png, thumb, SYNAPSE_THUMB_PX):
            continue
        found.append({
            "tag": "synapse", "group": "synapse", "number": 0 if sheet else int(m.group(1)),
            "view": "All synapses (contact sheet)" if sheet else f"Synapse #{int(m.group(1))}",
            "slice": None if sheet else int(m.group(2)),
            "file": urllib.parse.quote(f"../{SYNAPSE_FOLDER}/{name}/{png.name}"),
            "thumb": urllib.parse.quote(f"{THUMB_DIR}/_synapses/{name}/{thumb.name}"),
            "rating": None, "notes": None,
        })
    return sorted(found, key=lambda p: p["number"])


def read_synapse_summary(base):
    """{sample: its row of the Synapse Log's Summary sheet} (synapse_exf.py), and the folder the
    synapse pictures are in. Empty when nothing has been measured yet."""
    import configparser

    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(core.SETTINGS_FILE, encoding="utf-8-sig")
    folder = Path(cfg.get("paths", "synapse_folder", fallback="").strip() or base.parent / "Zarr Synapse ExF")
    rows = {}
    try:
        ws = load_workbook(folder / "Synapse Log.xlsx", read_only=True)["Summary"]
        lines = ws.iter_rows(values_only=True)
        heads = next(lines)
        for line in lines:
            row = dict(zip(heads, line))
            if row.get("Sample"):
                rows[row["Sample"]] = row
    except Exception as e:
        print(f"(synapse log not read: {e})")
    return rows, folder


def collect(base):
    """One record per sample subfolder that holds at least one picture."""
    meta, logged = read_metadata()
    measured, synapse_folder = read_synapse_summary(base)
    samples = []
    folders = sorted((d for d in base.iterdir() if d.is_dir() and not d.name.startswith("_")),
                     key=lambda d: d.name.lower())
    for n, folder in enumerate(folders, start=1):
        name = folder.name
        notes = {r.get("Screenshot"): r for r in logged.get(name, [])}
        pictures = []
        for png in folder.glob("*.png"):
            m = re.fullmatch(re.escape(name) + r"_(.+)_z(all|\d+)(?:_\d+)?\.png", png.name)
            if not m:
                continue
            tag = m.group(1)
            thumb = base / THUMB_DIR / name / (png.stem + ".jpg")
            if not thumbnail(png, thumb):
                continue
            row = notes.get(png.name, {})
            pictures.append({
                "tag": tag, "view": nice_view(tag), "slice": None if m.group(2) == "all" else int(m.group(2)),
                "file": urllib.parse.quote(f"{name}/{png.name}"),
                "thumb": urllib.parse.quote(f"{THUMB_DIR}/{name}/{thumb.name}"),
                "rating": row.get("Rating"), "notes": row.get("Notes"),
            })
        if not pictures:
            continue
        pictures.sort(key=lambda p: (VIEW_ORDER.index(p["tag"]) if p["tag"] in VIEW_ORDER else len(VIEW_ORDER),
                                     p["tag"], p["slice"] or 0))
        # the synapses the expansion factor was measured on: their contact sheet, last
        x = measured.get(name, {})
        sheet = synapse_folder / name / f"{name}_synapses_contact-sheet.png"
        thumb = base / THUMB_DIR / name / (sheet.stem + ".jpg")
        if x and sheet.is_file() and thumbnail(sheet, thumb):
            pictures.append({
                "tag": "Synapses", "view": "Synapses measured for the expansion factor", "slice": None,
                "file": urllib.parse.quote(f"../{synapse_folder.name}/{name}/{sheet.name}"),
                "thumb": urllib.parse.quote(f"{THUMB_DIR}/{name}/{thumb.name}"), "rating": None, "notes": None,
            })
        card = next((p for tag in CARD_ORDER for p in pictures if p["tag"] == tag), pictures[0])
        synapses = synapse_pictures(base, name)
        pictures = pictures + synapses   # always after every other picture of the sample
        m = meta.get(name, {})
        ratings = [p["rating"] for p in pictures if isinstance(p["rating"], (int, float))]
        date = m.get("Date Acquired")
        samples.append({
            "name": name, "card": card["thumb"], "pictures": pictures, "synapses": len([p for p in synapses if p["number"]]),
            "type": next((r.get("Type") for r in logged.get(name, []) if r.get("Type")), ""),
            "date": date.strftime("%Y-%m-%d") if hasattr(date, "strftime") else (str(date) if date else ""),
            "channels": m.get("Channels (laser)") or "", "microscope": m.get("Microscope") or "",
            "objective": m.get("Objective") or "", "size": m.get("Size X × Y × Z") or "",
            "extent": m.get("Extent X × Y × Z (µm)") or "", "density": m.get("AB Spots / 1000 µm²"),
            "density_text": m.get("AB Density by Channel") or "", "link_status": m.get("Link Status") or "",
            "url": m.get("url"), "rating": max(ratings) if ratings else None,
            "notes": " | ".join(str(p["notes"]) for p in pictures if p["notes"]),
            "folder": urllib.parse.quote(name) + "/",
            # from the synapse log; absent until Measure_ExF.bat has been run on the sample
            "antibody": x.get("Antibody") or "", "exf": x.get("ExF"), "exf_sd": x.get("ExF s.d."),
            "exf_fields": x.get("ExF Fields"), "exf_n": x.get("ExF n"), "exf_basis": x.get("ExF Basis") or "",
            "exf_fit": x.get("ExF (Gaussian fit)"),
            # an expansion factor from pan-only double peaks reads high and is shown as doubtful
            "exf_sure": bool(x.get("ExF")) and not str(x.get("ExF Basis") or "").startswith("all double peaks"),
            "pixel": x.get("Pixel (nm, corrected)"), "pixel_raw": x.get("Pixel (nm, as imaged)"),
            "res": x.get("Resolution (nm, corrected)"), "res_raw": x.get("Resolution FWHM (nm, as imaged)"),
            "synapses": x.get("Synapses found"), "ab_pos": x.get("AB+"), "ab_labelled": x.get("AB Labelled Synapses (%)"),
            "ab_signal": x.get("AB Signal at Synapses (σ)"), "ab_puncta": x.get("AB Puncta"),
            "ab_on": x.get("AB Puncta on Synapses (%)"), "ab_enrich": x.get("AB Enrichment (× chance)"),
            "synapse_folder": urllib.parse.quote(f"../{synapse_folder.name}/{name}") + "/" if x else "",
        })
        if n % 20 == 0:
            print(f"   {n}/{len(folders)} samples")
    return samples


PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root { --bg:#111418; --panel:#1b2027; --line:#2c333d; --text:#e6e9ee; --dim:#9aa4b2; --accent:#4fc3b0; --warn:#ff8a80; }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--text); font:14px/1.45 "Segoe UI",system-ui,sans-serif; }
header { position:sticky; top:0; z-index:5; background:rgba(17,20,24,.96); border-bottom:1px solid var(--line);
         padding:12px 18px; display:flex; flex-wrap:wrap; gap:10px; align-items:center; }
h1 { font-size:18px; margin:0 14px 0 0; font-weight:600; }
input,select { background:var(--panel); color:var(--text); border:1px solid var(--line); border-radius:6px; padding:7px 9px; font:inherit; }
input[type=search] { width:260px; }
#count { color:var(--dim); margin-left:auto; }
#grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(230px,1fr)); gap:14px; padding:18px; }
.card { background:var(--panel); border:1px solid var(--line); border-radius:8px; overflow:hidden; cursor:pointer; }
.card:hover,.card:focus { border-color:var(--accent); outline:none; }
.card img { width:100%; aspect-ratio:1/1; object-fit:contain; background:#000; display:block; }
.card .t { padding:8px 10px 10px; }
.name { font-weight:600; font-size:12.5px; word-break:break-word; }
.sub { color:var(--dim); font-size:12px; margin-top:3px; }
.badge { display:inline-block; font-size:11px; padding:1px 7px; border-radius:10px; border:1px solid var(--line); color:var(--dim); margin:4px 4px 0 0; }
.badge.r { color:#ffd54f; border-color:#6b5a1e; } .badge.w { color:var(--warn); border-color:#6b2a25; }
.badge.x { color:var(--accent); border-color:#27695f; } .badge.q { border-style:dashed; }
#detail { position:fixed; inset:0; background:rgba(0,0,0,.82); display:none; z-index:10; overflow:auto; }
#detail.on { display:block; }
#box { background:var(--bg); max-width:1500px; margin:22px auto; border:1px solid var(--line); border-radius:10px; padding:18px 22px 26px; }
#box h2 { margin:0 0 4px; font-size:17px; word-break:break-word; }
.bar { display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin:10px 0 14px; }
.btn { background:var(--panel); border:1px solid var(--line); color:var(--text); border-radius:6px; padding:6px 12px; text-decoration:none; cursor:pointer; font:inherit; }
.btn:hover { border-color:var(--accent); } .btn.go { background:#17443d; border-color:#27695f; }
table { border-collapse:collapse; margin:4px 0 16px; } td { padding:2px 14px 2px 0; vertical-align:top; } td:first-child { color:var(--dim); white-space:nowrap; }
#pics { display:grid; grid-template-columns:repeat(auto-fill,minmax(250px,1fr)); gap:14px; }
figure { margin:0; background:var(--panel); border:1px solid var(--line); border-radius:8px; overflow:hidden; }
figure img { width:100%; aspect-ratio:1/1; object-fit:contain; background:#000; display:block; }
figcaption { padding:7px 10px 9px; font-size:12.5px; } figcaption span { color:var(--dim); }
.empty { padding:60px; text-align:center; color:var(--dim); }
h3 { margin:26px 0 10px; font-size:14px; font-weight:600; color:var(--dim); border-top:1px solid var(--line); padding-top:16px; }
#syn { display:grid; grid-template-columns:repeat(auto-fill,minmax(620px,1fr)); gap:14px; }
#syn figure img { aspect-ratio:auto; }
</style></head><body>
<header>
  <h1>__TITLE__</h1>
  <input id="q" type="search" placeholder="Search name, channels, notes…" autofocus>
  <select id="type"><option value="">All types</option></select>
  <select id="sort">
    <option value="name">Sort: name</option><option value="date">Sort: date, newest first</option>
    <option value="density">Sort: antibody density</option><option value="rating">Sort: rating</option>
    <option value="exf">Sort: expansion factor</option><option value="res">Sort: resolution, finest first</option>
    <option value="enrich">Sort: antibody on synapses (enrichment)</option>
    <option value="labelled">Sort: synapses labelled (%)</option>
  </select>
  <label><input id="rated" type="checkbox"> rated only</label>
  <label><input id="measured" type="checkbox"> with expansion factor only</label>
  <span id="count"></span>
</header>
<div id="grid"></div>
<div id="detail"><div id="box"></div></div>
<script>
const DATA = __DATA__;
const BUILT = "__BUILT__";
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
[...new Set(DATA.map(s => s.type).filter(Boolean))].sort().forEach(t => $("type").insertAdjacentHTML("beforeend", `<option>${esc(t)}</option>`));
let shown = [];
function draw() {
  const q = $("q").value.trim().toLowerCase(), type = $("type").value, sort = $("sort").value, rated = $("rated").checked;
  const measured = $("measured").checked;
  shown = DATA.filter(s => (!type || s.type === type) && (!rated || s.rating != null) && (!measured || s.exf != null) &&
    (!q || q.split(/\s+/).every(w => (s.name + " " + s.channels + " " + s.notes + " " + s.type + " " + s.date + " " + s.antibody).toLowerCase().includes(w))));
  const by = { name:(a,b)=>a.name.localeCompare(b.name), date:(a,b)=>b.date.localeCompare(a.date),
               density:(a,b)=>(b.density ?? -1)-(a.density ?? -1), rating:(a,b)=>(b.rating ?? -1)-(a.rating ?? -1),
               exf:(a,b)=>(b.exf ?? -1)-(a.exf ?? -1), res:(a,b)=>(a.res ?? 1e9)-(b.res ?? 1e9),
               enrich:(a,b)=>(b.ab_enrich ?? -1)-(a.ab_enrich ?? -1), labelled:(a,b)=>(b.ab_labelled ?? -1)-(a.ab_labelled ?? -1) }[sort];
  shown.sort(by);
  $("count").textContent = `${shown.length} of ${DATA.length} samples · built ${BUILT}`;
  $("grid").innerHTML = shown.length ? shown.map((s,i) => `
    <div class="card" tabindex="0" data-i="${i}">
      <img loading="lazy" src="${s.card}" alt="">
      <div class="t"><div class="name">${esc(s.name)}</div>
        <div class="sub">${esc(s.date)}${s.type ? " · " + esc(s.type) : ""} · ${s.pictures.length - (s.synapses ? s.synapses + 1 : 0)} pictures</div>
        ${s.rating != null ? `<span class="badge r">★ ${s.rating}</span>` : ""}
        ${s.density != null ? `<span class="badge">${Number(s.density).toFixed(1)} spots / 1000 µm²</span>` : ""}
        ${s.exf != null ? `<span class="badge x${s.exf_sure ? "" : " q"}" title="${esc(s.exf_basis)}">ExF ${s.exf_sure ? "" : "~"}${Number(s.exf).toFixed(1)}×</span>` : ""}
        ${s.res != null ? `<span class="badge" title="width of the smallest antibody puncta ÷ ExF: an upper bound">≤ ${Number(s.res).toFixed(0)} nm</span>` : ""}
        ${s.ab_enrich != null ? `<span class="badge" title="antibody puncta on synapses, × chance">AB ${Number(s.ab_enrich).toFixed(1)}× on synapses</span>` : ""}
        ${s.synapses ? `<span class="badge">${s.synapses} synapses</span>` : ""}
        ${s.link_status && s.link_status !== "OK" ? `<span class="badge w">link problem</span>` : ""}
      </div></div>`).join("") : `<div class="empty">No samples match.</div>`;
}
const fig = p => `<figure><a href="${p.file}" target="_blank"><img loading="lazy" src="${p.thumb}" alt=""></a>
  <figcaption>${esc(p.view)}${p.slice != null ? ` <span>· slice ${p.slice}</span>` : ""}
  ${p.rating != null ? ` <span>· ★ ${p.rating}</span>` : ""}${p.notes ? `<br><span>${esc(p.notes)}</span>` : ""}</figcaption></figure>`;
function open_(i) {
  const s = shown[i]; if (!s) return;
  const rows = [["Date", s.date],["Type", s.type],["Channels", s.channels],["Microscope", s.microscope],["Objective", s.objective],
    ["Size (voxels)", s.size],["Extent (µm)", s.extent],["Antibody density", s.density_text],["Link status", s.link_status],["Notes", s.notes],
    ["Expansion factor", s.exf == null ? "" : `${s.exf_sure ? "" : "about "}${s.exf} ± ${s.exf_sd}×  (DP–PSD peak to peak ÷ 81.9 nm; ${s.exf_basis}; ` +
      `${s.exf_fields ? s.exf_fields + " fields, " : ""}${s.exf_n} synapses; Gaussian fit ${s.exf_fit}×)` +
      (s.exf_sure ? "" : "  — not reliable: no synaptic antibody to confirm the synapses, reads high")],
    ["Pixel size", s.pixel == null ? "" : `${s.pixel} nm corrected  (${s.pixel_raw} nm as imaged)`],
    ["Resolution", s.res == null ? "" : `≤ ${s.res} nm corrected  (${s.res_raw} nm as imaged; smallest antibody puncta, lateral, an upper bound)`],
    ["Synapses found", s.synapses == null ? "" : `${s.synapses}${s.antibody ? `, ${s.ab_pos} with ${s.antibody}` : ""}`],
    ["Antibody labelling", s.ab_labelled == null ? "" : `${s.ab_labelled}% of synapses found carry antibody` +
      (s.ab_signal != null ? ` (median ${s.ab_signal} σ)` : "") +
      (s.ab_on != null ? `; ${s.ab_on}% of ${s.ab_puncta} puncta sit on a synapse found = ${s.ab_enrich}× chance` : "") + "  (provisional)"]]
    .filter(r => r[1]).map(r => `<tr><td>${r[0]}</td><td>${esc(r[1])}</td></tr>`).join("");
  $("box").innerHTML = `<h2>${esc(s.name)}</h2>
    <div class="bar">
      ${s.url ? `<a class="btn go" href="${esc(s.url)}" target="_blank" rel="noopener">Open in Neuroglancer</a>` : ""}
      <a class="btn" href="${s.folder}" target="_blank">Open folder</a>
      ${s.synapse_folder ? `<a class="btn" href="${s.synapse_folder}" target="_blank">Synapse pictures</a>` : ""}
      <button class="btn" data-go="${i-1}">◀ Previous</button><button class="btn" data-go="${i+1}">Next ▶</button>
      <button class="btn" data-go="x">Close (Esc)</button>
    </div>
    <table>${rows}</table>
    <div id="pics">${s.pictures.filter(p => p.group !== "synapse").map(fig).join("")}</div>
    ${s.pictures.some(p => p.group === "synapse") ? `<h3>Synapses · ${s.synapses} measured for expansion factor</h3>
      <div id="syn">${s.pictures.filter(p => p.group === "synapse").map(fig).join("")}</div>` : ""}`;
  $("detail").classList.add("on"); $("detail").scrollTop = 0; $("detail").dataset.i = i;
  history.replaceState(null, "", "#" + encodeURIComponent(s.name));   // the address now points at this sample
}
const close_ = () => { $("detail").classList.remove("on"); history.replaceState(null, "", location.pathname); };
$("grid").addEventListener("click", e => { const c = e.target.closest(".card"); if (c) open_(+c.dataset.i); });
$("grid").addEventListener("keydown", e => { if (e.key === "Enter") { const c = e.target.closest(".card"); if (c) open_(+c.dataset.i); } });
$("detail").addEventListener("click", e => {
  if (e.target.id === "detail") return close_();
  const b = e.target.closest("[data-go]"); if (!b) return;
  b.dataset.go === "x" ? close_() : open_(Math.max(0, Math.min(shown.length - 1, +b.dataset.go)));
});
document.addEventListener("keydown", e => {
  if (!$("detail").classList.contains("on")) return;
  const i = +$("detail").dataset.i;
  if (e.key === "Escape") close_(); else if (e.key === "ArrowRight") open_(Math.min(shown.length - 1, i + 1)); else if (e.key === "ArrowLeft") open_(Math.max(0, i - 1));
});
["q","type","sort","rated","measured"].forEach(id => $(id).addEventListener("input", draw));
draw();
if (location.hash) {   // opened with a sample's name after '#': go straight to it
  const want = decodeURIComponent(location.hash.slice(1)), at = shown.findIndex(s => s.name === want);
  if (at >= 0) open_(at);
}
</script></body></html>
"""


def main():
    core.load_settings()
    base = Path(core.DEFAULT_SHOTS)
    if not base.is_dir():
        sys.exit(f"The screenshot folder does not exist yet: {base}\nRun Log_Screenshots.bat first.")
    print(f"Screenshot folder: {base}")
    print("Reading pictures and making thumbnails (only new or changed pictures take time)...")
    samples = collect(base)
    if not samples:
        sys.exit("No pictures found. Run Log_Screenshots.bat first.")
    data = json.dumps(samples, ensure_ascii=False, default=str).replace("</", "<\\/")
    page = PAGE.replace("__DATA__", data).replace("__BUILT__", html.escape(f"{dt.datetime.now():%Y-%m-%d %H:%M}"))
    page = page.replace("__TITLE__", html.escape(GALLERY_TITLE))
    out = base / GALLERY_NAME
    out.write_text(page, encoding="utf-8")
    for old in OLD_NAMES:   # the page under a previous name would otherwise linger, out of date
        if old != GALLERY_NAME and (base / old).is_file():
            (base / old).unlink()
    print(f"\n{len(samples)} samples, {sum(len(s['pictures']) for s in samples)} pictures")
    print(f"Gallery: {out}")
    return out


if __name__ == "__main__":
    path = main()
    if "--open" in sys.argv:
        import os
        os.startfile(path)
