"""Self-contained HTML X-ray report.

No CDN, no external fonts, no network. The report embeds structure only -- cell
coordinates and types, never cell values -- so it can be circulated without
carrying confidential figures out of the environment.
"""
from __future__ import annotations

import html
import json
from dataclasses import asdict

from .scan import WorkbookXray

KIND_COLOR = {
    "data_table": "#2D6CA2",
    "calculation": "#B4531F",
    "key_value": "#6B4E9E",
    "matrix": "#167B75",
    "notes": "#7A8290",
    "title": "#4A5568",
    "unknown": "#C0392B",
}

CSS = """
*{box-sizing:border-box}
:root{
  --ink:#10192B; --paper:#F7F8FA; --panel:#FFFFFF; --rule:#D5DAE3;
  --dim:#5B6577; --plate:#182437; --grid:#26344B;
}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);
  font:15px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.mono{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
.wrap{max-width:1180px;margin:0 auto;padding:32px 24px 80px}
header.top{border-bottom:2px solid var(--ink);padding-bottom:20px;margin-bottom:28px}
.eyebrow{font:600 11px/1 ui-monospace,monospace;letter-spacing:.18em;
  text-transform:uppercase;color:var(--dim);margin-bottom:10px}
h1{margin:0;font-size:30px;letter-spacing:-.02em;font-weight:650;word-break:break-all}
.meta{margin-top:8px;color:var(--dim);font-size:13px}
.vitals{display:flex;flex-wrap:wrap;gap:0;margin-top:22px;
  border:1px solid var(--rule);border-radius:3px;overflow:hidden;background:var(--panel)}
.vital{flex:1 1 130px;padding:12px 14px;border-right:1px solid var(--rule)}
.vital:last-child{border-right:0}
.vital b{display:block;font:600 22px/1.1 ui-monospace,monospace;letter-spacing:-.02em}
.vital span{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;
  color:var(--dim);margin-top:5px}
.warn{border-left:3px solid #B4531F;background:#FFF6F0;padding:11px 14px;
  margin:8px 0;font-size:13.5px;border-radius:0 3px 3px 0}
.sheet{background:var(--panel);border:1px solid var(--rule);border-radius:4px;
  margin:20px 0;overflow:hidden}
.sheet>h2{margin:0;padding:13px 16px;font-size:16px;font-weight:600;
  border-bottom:1px solid var(--rule);display:flex;align-items:center;gap:10px;
  flex-wrap:wrap}
.tag{font:600 10px/1 ui-monospace,monospace;letter-spacing:.1em;text-transform:uppercase;
  padding:4px 7px;border-radius:2px;background:#EDF0F5;color:var(--dim)}
.tag.hidden{background:#F5E9E9;color:#8B3A3A}
.body{display:grid;grid-template-columns:minmax(240px,340px) 1fr;gap:0}
@media(max-width:820px){.body{grid-template-columns:1fr}}
.plateWrap{padding:16px;background:var(--plate);border-right:1px solid var(--rule)}
.plate{position:relative;width:100%}
.plate canvas{display:block;width:100%;image-rendering:pixelated;border:1px solid var(--grid)}
.plateCap{color:#8A9AB5;font:11px/1.5 ui-monospace,monospace;margin-top:9px}
.regions{padding:6px 16px 16px}
.reg{border-bottom:1px solid var(--rule);padding:13px 0}
.reg:last-child{border-bottom:0}
.regHead{display:flex;align-items:baseline;gap:9px;flex-wrap:wrap}
.dot{width:9px;height:9px;border-radius:2px;flex:none;transform:translateY(1px)}
.kind{font:600 12px/1 ui-monospace,monospace;letter-spacing:.06em;text-transform:uppercase}
.ref{font-family:ui-monospace,monospace;font-size:12.5px;color:var(--dim)}
.conf{margin-left:auto;font:600 12px/1 ui-monospace,monospace;padding:3px 6px;
  border-radius:2px;background:#EDF0F5}
.conf.low{background:#FBE9E7;color:#A33A2A}
.hdrs{margin:8px 0 0;font-family:ui-monospace,monospace;font-size:12px;
  color:var(--ink);display:flex;flex-wrap:wrap;gap:4px}
.hdrs i{font-style:normal;background:#EDF0F5;padding:2px 6px;border-radius:2px}
.hdrs i.blank{color:#A6AEBC}
.ev{margin:8px 0 0;padding:0;list-style:none;font-size:12.5px;color:var(--dim)}
.ev li{padding-left:13px;position:relative}
.ev li:before{content:"";position:absolute;left:0;top:9px;width:5px;height:1px;
  background:var(--dim)}
table.sk{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}
table.sk th{text-align:left;font:600 10px/1 ui-monospace,monospace;letter-spacing:.1em;
  text-transform:uppercase;color:var(--dim);padding:7px 8px;border-bottom:1px solid var(--rule)}
table.sk td{padding:6px 8px;border-bottom:1px solid #EEF1F5;vertical-align:top}
table.sk td.n{text-align:right;font-family:ui-monospace,monospace;white-space:nowrap;
  color:var(--dim)}
table.sk td.f{font-family:ui-monospace,monospace;word-break:break-all}
.section{padding:14px 16px;border-top:1px solid var(--rule)}
.section>h3{margin:0 0 4px;font:600 11px/1 ui-monospace,monospace;letter-spacing:.14em;
  text-transform:uppercase;color:var(--dim)}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:14px;font-size:12px;color:var(--dim)}
.legend span{display:flex;align-items:center;gap:6px}
footer{margin-top:34px;padding-top:16px;border-top:1px solid var(--rule);
  font-size:12.5px;color:var(--dim)}
"""

JS = """
function drawPlate(cv, data){
  const {rows, cols, cells, regions} = data;
  const R = Math.max(rows, 8), C = Math.max(cols, 6);
  const cs = Math.max(3, Math.min(9, Math.floor(560 / C)));
  const w = C * cs, h = R * cs;
  const dpr = window.devicePixelRatio || 1;
  cv.width = w * dpr; cv.height = h * dpr;
  cv.style.height = (h * (cv.clientWidth / w)) + 'px';
  const g = cv.getContext('2d');
  g.scale(dpr, dpr);
  g.fillStyle = '#182437'; g.fillRect(0, 0, w, h);
  g.strokeStyle = '#26344B'; g.lineWidth = 0.5;
  for(let c=0;c<=C;c++){g.beginPath();g.moveTo(c*cs+.25,0);g.lineTo(c*cs+.25,h);g.stroke();}
  for(let r=0;r<=R;r++){g.beginPath();g.moveTo(0,r*cs+.25);g.lineTo(w,r*cs+.25);g.stroke();}
  const FLAG = {1:'#7E93B8', 2:'#E0913F', 3:'#E05252', 4:'#A9BBD6'};
  for(const [r,c,f] of cells){
    g.fillStyle = FLAG[f] || '#7E93B8';
    g.fillRect((c-1)*cs+1, (r-1)*cs+1, cs-1.5, cs-1.5);
  }
  for(const rg of regions){
    g.strokeStyle = rg.color; g.lineWidth = 1.5;
    if(rg.conf < 0.7) g.setLineDash([3,2]); else g.setLineDash([]);
    g.strokeRect((rg.left-1)*cs+.5, (rg.top-1)*cs+.5,
                 (rg.right-rg.left+1)*cs-1, (rg.bottom-rg.top+1)*cs-1);
  }
  g.setLineDash([]);
}
document.querySelectorAll('canvas[data-plate]').forEach(cv=>{
  const d = JSON.parse(cv.getAttribute('data-plate'));
  drawPlate(cv, d);
  window.addEventListener('resize', ()=>drawPlate(cv, d));
});
"""


def _esc(s) -> str:
    return html.escape(str(s if s is not None else ""))


def build_report(wx: WorkbookXray) -> str:
    total_regions = sum(len(s.regions) for s in wx.sheets)
    total_formulas = sum(s.formula_profile.get("total", 0) for s in wx.sheets)
    distinct = sum(s.formula_profile.get("distinct_skeletons", 0) for s in wx.sheets)
    low_conf = [(s.name, r) for s in wx.sheets for r in s.regions
                if r.detect_confidence < 0.70]
    errors = [e for s in wx.sheets for e in s.error_cells]
    hidden = [s.name for s in wx.sheets if s.state != "visible"]

    parts: list[str] = []
    A = parts.append

    A(f"<!doctype html><meta charset='utf-8'>"
      f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
      f"<title>X-ray: {_esc(wx.filename)}</title><style>{CSS}</style>"
      f"<div class='wrap'>")

    A("<header class='top'><div class='eyebrow'>Workbook X-ray &middot; structural scan</div>"
      f"<h1>{_esc(wx.filename)}</h1>"
      f"<div class='meta mono'>{_esc(wx.sha256[:16])}&hellip; &middot; "
      f"{wx.size_bytes:,} bytes &middot; modified {_esc(wx.fs_modified)} &middot; "
      f"last saved by {_esc((wx.app_props or {}).get('application') or 'unknown')}</div>")

    A("<div class='vitals'>")
    for val, label in [
        (len(wx.sheets), "sheets"),
        (total_regions, "regions found"),
        (f"{total_formulas:,}", "formula cells"),
        (distinct, "distinct formulas"),
        (f"{(total_formulas/distinct):.0f}x" if distinct else "&ndash;", "compression"),
        (len(low_conf), "need review"),
    ]:
        A(f"<div class='vital'><b>{val}</b><span>{label}</span></div>")
    A("</div></header>")

    for w in wx.warnings:
        A(f"<div class='warn'>{_esc(w)}</div>")
    if errors:
        A(f"<div class='warn'><b>{len(errors)} cached error cell(s):</b> "
          f"<span class='mono'>{_esc(', '.join(errors[:12]))}</span></div>")
    if hidden:
        A(f"<div class='warn'>Hidden sheet(s): <span class='mono'>"
          f"{_esc(', '.join(hidden))}</span>. Hidden sheets often hold the working "
          f"logic a report depends on &mdash; check before consolidating.</div>")

    for s in wx.sheets:
        fp = s.formula_profile
        plate = {
            "rows": s.max_row, "cols": s.max_col, "cells": s.occupancy,
            "regions": [{"top": r.top, "left": r.left, "bottom": r.bottom,
                         "right": r.right, "conf": r.detect_confidence,
                         "color": KIND_COLOR.get(r.kind, "#888")} for r in s.regions],
        }
        A(f"<section class='sheet'><h2>{_esc(s.name)}"
          f"<span class='tag'>#{s.position + 1}</span>")
        if s.state != "visible":
            A(f"<span class='tag hidden'>{_esc(s.state)}</span>")
        A(f"<span class='tag'>{s.populated_cells:,} cells</span>"
          f"<span class='tag'>{len(s.regions)} regions</span></h2>")

        A("<div class='body'><div class='plateWrap'><div class='plate'>"
          f"<canvas data-plate='{_esc(json.dumps(plate))}'></canvas></div>"
          f"<div class='plateCap'>{_esc(s.dimension or '')} &middot; density "
          f"{s.density:.0%}"
          + (" &middot; plate truncated" if s.occupancy_truncated else "")
          + "</div></div>")

        A("<div class='regions'>")
        if not s.regions:
            A("<div class='reg'><span class='ref'>No populated regions.</span></div>")
        for r in s.regions:
            color = KIND_COLOR.get(r.kind, "#888")
            cls = "conf low" if r.detect_confidence < 0.70 else "conf"
            A(f"<div class='reg'><div class='regHead'>"
              f"<span class='dot' style='background:{color}'></span>"
              f"<span class='kind' style='color:{color}'>{_esc(r.kind)}</span>"
              f"<span class='ref'>{_esc(r.ref)}</span>"
              f"<span class='ref'>{r.n_data_rows} data rows &middot; "
              f"{r.formula_cells} formulas"
              + (f" &middot; totals row {r.totals_row}" if r.totals_row else "")
              + f"</span><span class='{cls}'>{r.detect_confidence:.2f}</span></div>")
            if r.headers:
                A("<div class='hdrs'>" + "".join(
                    f"<i class='blank'>&empty;</i>" if not h
                    else f"<i>{_esc(h)}</i>" for h in r.headers) + "</div>")
            if r.evidence:
                A("<ul class='ev'>" + "".join(
                    f"<li>{_esc(e)}</li>" for e in r.evidence) + "</ul>")
            A("</div>")
        A("</div></div>")

        if fp.get("total"):
            A("<div class='section'><h3>Calculation profile</h3>"
              f"<p style='margin:6px 0 10px;font-size:13px;color:var(--dim)'>"
              f"{fp['total']:,} formula cells reduce to {fp['distinct_skeletons']} "
              f"distinct calculations ({fp['compression']}x). "
              f"{fp['cross_sheet_count']} reach another sheet, "
              f"{fp['external_count']} reach another file, "
              f"{fp['hardcoded_literal_count']} contain a hardcoded number."
              + (f" Volatile functions: {fp['volatile_count']}."
                 if fp['volatile_count'] else "") + "</p>")
            A("<table class='sk'><tr><th>Normalised formula (R1C1)</th>"
              "<th style='text-align:right'>Cells</th></tr>")
            for sk, n in fp["top_skeletons"][:10]:
                A(f"<tr><td class='f'>{_esc(sk)}</td><td class='n'>{n}</td></tr>")
            A("</table></div>")

    A("<div class='legend'>")
    for k, c in KIND_COLOR.items():
        A(f"<span><i class='dot' style='display:inline-block;width:9px;height:9px;"
          f"border-radius:2px;background:{c}'></i>{_esc(k)}</span>")
    A("<span>Dashed outline = confidence below 0.70, needs a human look.</span></div>")

    A("<footer>Structure only &mdash; this report contains cell coordinates, types "
      "and normalised formula shapes. No cell values are embedded. "
      "Region boundaries on hand-built sheets are inferred, not declared: treat "
      "anything below 0.70 as a prompt to open the file.</footer>")
    A(f"</div><script>{JS}</script>")
    return "".join(parts)


def write_report(wx: WorkbookXray, path: str) -> str:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(build_report(wx))
    return path
