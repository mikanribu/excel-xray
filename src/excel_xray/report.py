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
from .tabular import FILE_FIELDS, TAB_FIELDS, file_rows, tab_rows
from .util import safe_scan_message

BASIS_COLOR = {
    "extracted": "#167B75",
    "derived": "#2D6CA2",
    "drafted": "#B4531F",
    "inferred": "#6B4E9E",
    "needs_llm": "#7A8290",
    "needs_human": "#7A8290",
    "needs_corpus": "#7A8290",
}

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
.plate canvas{display:block;margin:0 auto;max-width:100%;image-rendering:pixelated;border:1px solid var(--grid)}
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
.assess{background:var(--panel);border:1px solid var(--rule);border-radius:4px;
  margin:20px 0;padding:6px 16px 16px}
.assess>h2{font-size:16px;font-weight:600;margin:12px 2px 6px}
.assess>h3{font:600 11px/1 ui-monospace,monospace;letter-spacing:.14em;
  text-transform:uppercase;color:var(--dim);margin:18px 2px 8px}
.scroll{overflow-x:auto}
table.euc{width:100%;border-collapse:collapse;font-size:12.5px}
table.euc th{text-align:left;font:600 10px/1.3 ui-monospace,monospace;
  letter-spacing:.08em;text-transform:uppercase;color:var(--dim);
  padding:7px 9px;border-bottom:1px solid var(--rule);vertical-align:bottom}
table.euc td{padding:6px 9px;border-bottom:1px solid #EEF1F5;vertical-align:top}
td.grp{font:600 10px/1.3 ui-monospace,monospace;letter-spacing:.06em;
  text-transform:uppercase;color:var(--dim);white-space:nowrap}
td.fld{font-weight:600;white-space:nowrap;color:var(--ink)}
table.matrix td:first-child{font-weight:600;white-space:nowrap}
table.matrix td{min-width:120px;max-width:320px;white-space:normal;vertical-align:top}
.cellv{display:block;overflow-wrap:anywhere;word-break:break-word}
.cellv.trunc{cursor:pointer}
.cellv.trunc .moreTag{color:#2D6CA2;font-weight:600;white-space:nowrap}
.cellv.trunc:hover .moreTag{text-decoration:underline}
.tblTools{display:flex;align-items:center;gap:10px;margin:10px 2px 2px}
.tblTools input[type=search]{flex:1;max-width:380px;font:13px ui-monospace,monospace;
  padding:7px 10px;border:1px solid var(--rule);border-radius:3px;
  background:var(--panel);color:var(--ink)}
.tblTools input[type=search]:focus{outline:2px solid #2D6CA2;outline-offset:-1px}
.tblTools .rowCount{font-size:11.5px;color:var(--dim);white-space:nowrap}
.badge{display:inline-block;font:600 9px/1.4 ui-monospace,monospace;
  letter-spacing:.04em;padding:1px 5px;border-radius:2px;color:#fff;
  white-space:nowrap;vertical-align:middle}
.legend2{margin:10px 2px 0;font-size:11.5px;color:var(--dim);
  display:flex;flex-wrap:wrap;gap:12px}
.tabbar{position:sticky;top:0;z-index:10;display:flex;flex-wrap:nowrap;
  overflow-x:auto;gap:3px;background:var(--paper);padding:10px 0 0;margin:0 0 20px;
  border-bottom:2px solid var(--ink)}
.tabbtn{flex:none;font:600 12.5px/1 ui-monospace,monospace;padding:10px 15px;
  background:var(--panel);color:var(--dim);border:1px solid var(--rule);
  border-bottom:none;border-radius:4px 4px 0 0;cursor:pointer;white-space:nowrap}
.tabbtn:hover{color:var(--ink)}
.tabbtn.active{background:var(--ink);color:#fff;border-color:var(--ink)}
.tabbtn .tct{margin-left:7px;opacity:.6;font-weight:400}
.tabpanel{display:none}
.tabpanel.active{display:block}
"""

JS = """
const PLATE_MAX_ROWS = 120, PLATE_MAX_COLS = 60, BOX_W = 300, BOX_H = 340;
function drawPlate(cv, data){
  const {rows, cols, cells, regions} = data;
  // Bound the grid so a 10k-row sheet is a thumbnail, not a giant strip. Cells
  // beyond the cap are not drawn (occupancy is already clipped server-side).
  const R = Math.max(Math.min(rows, PLATE_MAX_ROWS), 8);
  const C = Math.max(Math.min(cols, PLATE_MAX_COLS), 6);
  let cs = Math.floor(Math.min(BOX_W / C, BOX_H / R));
  cs = Math.max(2, Math.min(10, cs));
  const w = C * cs, h = R * cs;
  const dpr = window.devicePixelRatio || 1;
  cv.width = w * dpr; cv.height = h * dpr;
  cv.style.width = w + 'px'; cv.style.height = h + 'px';  // natural size, no stretch
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

(function initTabs(){
  const buttons = Array.from(document.querySelectorAll('.tabbtn'));
  const panels = Array.from(document.querySelectorAll('.tabpanel'));
  if(!buttons.length) return;
  function activate(id, scroll){
    buttons.forEach(b => b.classList.toggle('active', b.dataset.target === id));
    panels.forEach(p => p.classList.toggle('active', p.id === id));
    if(scroll){
      const btn = buttons.find(b => b.dataset.target === id);
      if(btn) btn.scrollIntoView({block:'nearest', inline:'nearest'});
    }
  }
  buttons.forEach(b => b.addEventListener('click', () => {
    activate(b.dataset.target, false);
    history.replaceState(null, '', '#' + b.dataset.target);
  }));
  const ids = panels.map(p => p.id);
  const wanted = (location.hash || '').slice(1);
  const fallback = ids.includes('tab-assessment') ? 'tab-assessment' : ids[0];
  activate(ids.includes(wanted) ? wanted : fallback, true);
})();

document.querySelectorAll('.cellv.trunc').forEach(el=>{
  const tag = el.querySelector('.moreTag');
  el.addEventListener('click', ()=>{
    const open = el.classList.toggle('open');
    el.childNodes[0].nodeValue = open ? el.dataset.full : el.dataset.preview;
    if(tag) tag.textContent = open ? ' show less' : ' \\u2026 show more';
  });
});

document.querySelectorAll('.rowFilter').forEach(inp=>{
  const table = document.getElementById(inp.dataset.table);
  if(!table) return;
  const countEl = inp.parentElement.querySelector('.rowCount');
  const rows = Array.from(table.querySelectorAll('tr')).filter(tr => !tr.querySelector('th'));
  function apply(){
    const q = inp.value.trim().toLowerCase();
    let shown = 0;
    rows.forEach(tr=>{
      const hit = !q || tr.textContent.toLowerCase().includes(q);
      tr.style.display = hit ? '' : 'none';
      if(hit) shown++;
    });
    if(countEl) countEl.textContent = shown + ' / ' + rows.length + ' rows';
  }
  inp.addEventListener('input', apply);
  apply();
});
"""


def _esc(s) -> str:
    return html.escape(str(s if s is not None else ""))


def _badge(basis: str) -> str:
    return (f"<span class='badge' style='background:{BASIS_COLOR.get(basis, '#888')}'>"
            f"{_esc(basis)}</span>")


CELL_TRUNCATE_AT = 220


def _cellval(value, evidence: str = "") -> str:
    """Render a table cell value, clamping long text behind a click-to-expand.

    Some fields (external link lists, ambiguous-region dumps, narrative
    summaries) run to thousands of characters -- inline in full they blow a
    single table row to a huge height and make the whole table unreadable.
    Short values render unchanged; long ones show a truncated preview with
    a "show more" toggle (JS in `JS`), and the full text is always still
    reachable via the hover title.
    """
    s = str(value if value is not None else "")
    title_attr = f" title='{_esc(evidence)}'" if evidence else ""
    if len(s) <= CELL_TRUNCATE_AT:
        return f"<span class='cellv'{title_attr}>{_esc(s)}</span>"
    cut = s.rfind(";", 0, CELL_TRUNCATE_AT)
    if cut < CELL_TRUNCATE_AT * 0.5:
        cut = s.rfind(" ", 0, CELL_TRUNCATE_AT)
    if cut < CELL_TRUNCATE_AT * 0.5:
        cut = CELL_TRUNCATE_AT
    preview = s[:cut].rstrip(" ;,")
    return (f"<span class='cellv trunc'{title_attr} "
            f"data-full='{_esc(s)}' data-preview='{_esc(preview)}'>{_esc(preview)}"
            f"<span class='moreTag'> &hellip; show more</span></span>")


def _table_tools(table_id: str, placeholder: str) -> str:
    return (f"<div class='tblTools'><input type='search' class='rowFilter' "
            f"data-table='{table_id}' placeholder='{_esc(placeholder)}'>"
            f"<span class='rowCount'></span></div>")


def _error_header(review) -> str:
    """One-line, high-level warning for the always-visible header — the
    grouped table and per-cell detail live in the Diagnostics tab instead."""
    groups = review.error_groups
    if not groups:
        return ""
    total = sum(g.count for g in groups)
    return (f"<div class='warn'><b>{total} cached error cell(s)</b> across "
           f"{len(groups)} group(s) — see the Diagnostics tab for the grouped "
           f"summary, downstream impact and reviewer actions.</div>")


def _error_detail(review) -> str:
    """Grouped-by-type/sheet/region error summary plus the per-cell detail
    table — the "too detailed to keep in the header" material, Diagnostics tab."""
    groups = review.error_groups
    if not groups:
        return ""
    rows = []
    for g in groups:
        recalc = "Yes" if g.recalc_required else "No — formula repair needed"
        downstream = ", ".join(g.downstream_outputs) or "none identified"
        rows.append(
            f"<tr><td class='mono'>{_esc(g.error_type)}</td><td class='n'>{g.count}</td>"
            f"<td>{_esc(g.sheet)}</td><td class='mono'>{_esc(g.region_ref)}</td>"
            f"<td>{_esc(g.area_type)}</td><td>{_esc(downstream)}</td>"
            f"<td>{_esc(recalc)}</td><td>{_esc(g.reviewer_action)}</td></tr>"
        )
    total = sum(g.count for g in groups)
    table = (
        f"<h3>Error summary</h3>"
        f"<p style='margin:6px 2px 10px;font-size:13px;color:var(--dim)'>"
        f"{total} cached error cell(s) across {len(groups)} group(s).</p>"
        "<div class='scroll'><table class='euc' style='margin:6px 0 10px'>"
        "<tr><th>Type</th><th>Count</th><th>Sheet</th><th>Region</th><th>Area</th>"
        "<th>Downstream outputs</th><th>Recalc required</th><th>Reviewer action</th></tr>"
        + "".join(rows) + "</table></div>"
    )
    details = review.error_details
    if details:
        trs = "".join(
            f"<tr><td>{_esc(d['sheet'])}</td><td class='mono'>{_esc(d['cell'])}</td>"
            f"<td class='mono'>{_esc(d['error_type'])}</td>"
            f"<td class='mono'>{_esc(d['region'])}</td><td>{_esc(d['area_type'])}</td></tr>"
            for d in details
        )
        table += (
            "<details style='margin:0 0 16px'><summary style='cursor:pointer;"
            f"color:var(--dim);font-size:12.5px'>Detailed error cell references "
            f"({len(details)})</summary><div class='scroll'>"
            "<table class='euc' style='margin:6px 0'><tr><th>Sheet</th><th>Cell</th>"
            "<th>Type</th><th>Region</th><th>Area</th></tr>" + trs
            + "</table></div></details>"
        )
    return table


def _hidden_header(review) -> str:
    """One-line, high-level warning for the always-visible header — the
    purpose grouping lives in the Diagnostics tab instead."""
    hs = review.hidden_summary
    if not hs:
        return ""
    if not hs.get("hidden_count"):
        return f"<div class='warn'>{_esc(hs.get('header', '0 worksheets are hidden.'))}</div>"
    return (f"<div class='warn'>{_esc(hs['header'])} Hidden status is never treated "
           f"as evidence a worksheet or workbook is obsolete — see the Diagnostics "
           f"tab for the purpose grouping.</div>")


def _hidden_detail(review) -> str:
    """'X of Y worksheets are hidden', grouped by likely purpose — never a
    bare list, and never framed as evidence of obsolescence."""
    hs = review.hidden_summary
    if not hs or not hs.get("hidden_count"):
        return ""
    rows = []
    for g in hs["groups"]:
        rows.append(
            f"<tr><td>{_esc(g.purpose)}</td><td class='mono'>{_esc(', '.join(g.sheets))}</td>"
            f"<td class='n'>{g.count}</td>"
            f"<td>{_esc(', '.join(g.depended_on_by))}</td>"
            f"<td>{_esc(g.explanation)}</td></tr>"
        )
    return (
        f"<h3>Hidden sheet groups</h3>"
        f"<p style='margin:6px 2px 10px;font-size:13px;color:var(--dim)'>"
        f"{_esc(hs['header'])} Hidden sheets often hold working logic a visible "
        f"report depends on — hidden status is never treated as evidence a "
        f"worksheet or workbook is obsolete.</p>"
        "<div class='scroll'><table class='euc' style='margin:6px 0 16px'>"
        "<tr><th>Purpose</th><th>Worksheets</th><th>Count</th>"
        "<th>Depended on by</th><th>Likely purpose</th></tr>"
        + "".join(rows) + "</table></div>"
    )


def _render_diagnostics(review) -> str:
    """Diagnostics tab: everything too detailed for the always-visible
    header — grouped and per-cell error detail, hidden-sheet purpose
    grouping. Only rendered when there is something to show."""
    error_html = _error_detail(review)
    hidden_html = _hidden_detail(review)
    if not error_html and not hidden_html:
        return ""
    parts = ["<section class='assess'><h2>Diagnostics</h2>"]
    parts.append(error_html)
    parts.append(hidden_html)
    parts.append("</section>")
    return "".join(parts)


def _render_assessment(assessment) -> str:
    """Render the EUC assessment as the File level / Tab level review tables."""
    parts: list[str] = []
    A = parts.append
    A("<section class='assess'><h2>EUC assessment</h2>")

    # ---- File level summary (Type | Field | Value | Basis) --------------
    A("<h3>File level summary</h3>")
    A(_table_tools("euc-file", "Filter rows… (type, field, value, basis)"))
    A("<table class='euc' id='euc-file'>"
      "<tr><th>Type</th><th>Field</th><th>Value</th><th>Basis</th></tr>")
    prev = None
    for r in file_rows(assessment):
        grp = "" if r["type"] == prev else r["type"]
        prev = r["type"]
        A(f"<tr><td class='grp'>{_esc(grp)}</td>"
          f"<td class='fld'>{_esc(r['field'])}</td>"
          f"<td>{_cellval(r['value'], r['evidence'])}</td>"
          f"<td>{_badge(r['basis'])}</td></tr>")
    A("</table>")

    # ---- Tab level details (matrix: rows = tabs, cols = fields) ----------
    if assessment.tabs:
        labels = [label for _sec, _attr, label in TAB_FIELDS]
        rows = tab_rows(assessment)
        by_tab: dict[str, dict] = {}
        for r in rows:
            by_tab.setdefault(r["tab"], {})[r["field"]] = r
        A("<h3>Tab level details</h3>")
        A(_table_tools("euc-tabs", "Filter tabs… (name, category, logic…)"))
        A("<div class='scroll'><table class='euc matrix' id='euc-tabs'><tr>"
          + "".join(f"<th>{_esc(l)}</th>" for l in labels) + "</tr>")
        for tab, fields in by_tab.items():
            A("<tr>")
            for label in labels:
                cell = fields.get(label, {"value": "—", "basis": "", "evidence": ""})
                badge = _badge(cell["basis"]) if cell["basis"] in (
                    "drafted", "inferred") else ""
                A(f"<td>{_cellval(cell['value'], cell['evidence'])} {badge}</td>")
            A("</tr>")
        A("</table></div>")

    A("<div class='legend2'>"
      "<span>extracted = read from the file</span>"
      "<span>derived = heuristic over evidence</span>"
      "<span>drafted = offline narrative template</span>"
      "<span>inferred = model-written</span>"
      "<span>needs_* = pending model / human / corpus</span></div>")
    A("</section>")
    return "".join(parts)


def build_report(wx: WorkbookXray, assessment=None) -> str:
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
      f"<div class='meta'>Scan status: {_esc(wx.parse_status)} &middot; "
      f"last saved by {_esc((wx.app_props or {}).get('application') or 'unknown')}</div>")

    need_review = (sum(1 for t in assessment.tabs if t.human_validation_required.value == "Y")
                  if assessment is not None else len(low_conf))
    A("<div class='vitals'>")
    for val, label in [
        (len(wx.sheets), "sheets"),
        (total_regions, "regions found"),
        (f"{total_formulas:,}", "formula cells"),
        (distinct, "distinct formulas"),
        (f"{(total_formulas/distinct):.0f}x" if distinct else "&ndash;", "compression"),
        (need_review, "need review"),
    ]:
        A(f"<div class='vital'><b>{val}</b><span>{label}</span></div>")
    A("</div></header>")

    for w in wx.warnings:
        A(f"<div class='warn'>{_esc(safe_scan_message(w, wx.path))}</div>")
    if assessment is not None:
        A(_error_header(assessment.review))
        A(_hidden_header(assessment.review))
    else:
        if errors:
            A(f"<div class='warn'><b>{len(errors)} cached error cell(s)</b> — run with "
              f"an assessment for a grouped, business-facing summary.</div>")
        if hidden:
            A(f"<div class='warn'>{len(hidden)} of {len(wx.sheets)} worksheet(s) are "
              f"hidden. Hidden status is never evidence a sheet is obsolete.</div>")

    # Visible worksheets are presented (and analysed) first; original position
    # is preserved on each tab as evidence, not lost by the reorder.
    if assessment is not None:
        order = [t.tab_name.value for t in assessment.tabs]
        sheets_ordered = sorted(
            wx.sheets, key=lambda s: order.index(s.name) if s.name in order else 10**6)
    else:
        sheets_ordered = wx.sheets

    # ---- Tab bar: Assessment first (selected by default — see initTabs in
    # JS, which opens ids[0] absent a URL hash), then Diagnostics (error/
    # hidden detail — too much for the always-visible header), then one tab
    # per sheet, so a many-sheet workbook navigates by click instead of by
    # scrolling past every sheet's plate and region list. --------------------
    diagnostics_html = _render_diagnostics(assessment.review) if assessment is not None else ""
    A("<div class='tabbar'>")
    if assessment is not None:
        A("<button class='tabbtn' data-target='tab-assessment'>Assessment</button>")
    if diagnostics_html:
        A("<button class='tabbtn' data-target='tab-diagnostics'>Diagnostics</button>")
    for i, s in enumerate(sheets_ordered):
        hidden_tag = " &middot; hidden" if s.state != "visible" else ""
        A(f"<button class='tabbtn' data-target='tab-sheet-{i}'>{_esc(s.name)}"
          f"<span class='tct'>{len(s.regions)}{hidden_tag}</span></button>")
    A("</div>")

    if assessment is not None:
        A("<div class='tabpanel' id='tab-assessment'>")
        A(_render_assessment(assessment))
        A("</div>")
    if diagnostics_html:
        A("<div class='tabpanel' id='tab-diagnostics'>")
        A(diagnostics_html)
        A("</div>")

    for i, s in enumerate(sheets_ordered):
        fp = s.formula_profile
        plate = {
            "rows": s.max_row, "cols": s.max_col, "cells": s.occupancy,
            "regions": [{"top": r.top, "left": r.left, "bottom": r.bottom,
                         "right": r.right, "conf": r.detect_confidence,
                         "color": KIND_COLOR.get(r.kind, "#888")} for r in s.regions],
        }
        A(f"<div class='tabpanel' id='tab-sheet-{i}'>")
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
          + (f" &middot; overview shows first {min(s.max_row, 120):,}&times;"
             f"{min(s.max_col, 60)} of {s.max_row:,}&times;{s.max_col}"
             if s.occupancy_truncated else "")
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

        A("</section></div>")

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


def write_report(wx: WorkbookXray, path: str, assessment=None) -> str:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(build_report(wx, assessment))
    return path
