"""Local-only portfolio dashboard with paged drill-down and review exports."""

from __future__ import annotations

import json
import os
import re
import tempfile
from contextlib import closing
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit

from .corpus import SIMILAR_THRESHOLD, SKELETON_DUP, similarity
from .portfolio import (DB_NAME, DIAGNOSTIC_COLUMNS, EXCEL_NAME, FINDING_COLUMNS,
                        SUMMARY_COLUMNS, TAB_COLUMNS, assessment_from_json, bundle,
                        connect, diagnostic_rows, finding_rows, selected_files,
                        summary_rows, tab_detail_rows, write_csv, write_excel,
                        source_is_current, _restore_fp, _display_scan_error)
from .tabular import fmt_value

DASHBOARD = r"""<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel X-ray · EUC portfolio</title>
<style>
:root{font:14px/1.45 system-ui,-apple-system,Segoe UI,sans-serif;color:#17243b;background:#f3f6fa}
*{box-sizing:border-box}body{margin:0}button,input,select{font:inherit}button{cursor:pointer;border:0;border-radius:7px;padding:9px 12px;background:#26558b;color:#fff}button.secondary{background:#e9eef5;color:#18334f}button:disabled{opacity:.45;cursor:default}input,select{border:1px solid #c7d2e2;border-radius:7px;padding:9px;background:white}header{background:#173b65;color:white;padding:22px max(24px,4vw)}header h1{margin:0;font-size:23px}header p{margin:5px 0 0;opacity:.83}main{max-width:1600px;margin:0 auto;padding:20px max(20px,3vw)}.cards{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:12px;margin-bottom:16px}.card,.panel{background:white;border:1px solid #d8e1ec;border-radius:10px;box-shadow:0 2px 10px #152b4810}.card{padding:17px}.card b{display:block;font-size:26px}.card span{color:#58677a}.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:12px 0}.toolbar input{min-width:250px;flex:1}.panel{padding:16px;margin:14px 0}.panel h2{margin:0 0 12px;font-size:18px}.tablewrap{overflow:auto;max-height:520px}table{border-collapse:collapse;width:100%;min-width:1300px}th,td{text-align:left;padding:9px 11px;border-bottom:1px solid #e6ebf1;vertical-align:top}th{background:#f0f4f9;position:sticky;top:0;z-index:1;white-space:nowrap}tr:hover td{background:#f8fbff}.muted{color:#607083}.pill{display:inline-block;border-radius:12px;padding:2px 8px;background:#e9f1fa;color:#214c76}.pill.high{background:#fee7e7;color:#a32626}.pill.partial{background:#fff1c5;color:#805d00}.split{display:grid;grid-template-columns:minmax(0,1fr) minmax(380px,.7fr);gap:14px}.detail{max-height:720px;overflow:auto}.detail h3{margin:18px 0 7px}.detail dl{display:grid;grid-template-columns:180px 1fr;gap:4px 9px}.detail dt{font-weight:650}.detail dd{margin:0;white-space:pre-wrap}.tabline{border-top:1px solid #e4eaf1;padding:10px 0}.tabline summary{cursor:pointer;font-weight:650}.actions{display:flex;gap:6px;flex-wrap:wrap}.notice{background:#fff8e0;border:1px solid #efdfaa;padding:10px;border-radius:7px;margin-top:8px}.pager{display:flex;gap:8px;align-items:center;margin-top:10px}.hidden{display:none}.comparison{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}.comparison .card b{font-size:15px}a{color:#1f5a97}
@media(max-width:1000px){.split{display:block}.cards{grid-template-columns:repeat(2,1fr)}}
</style>
<header><h1>Workbook X-ray · EUC portfolio</h1><p>Consolidated review of submitted workbooks. Structural evidence and assessment only; source workbooks are downloaded on request.</p></header>
<main>
<section class="cards" id="stats"></section>
<section class="panel"><h2>Portfolio findings</h2><div id="findings"></div></section>
<section class="panel"><h2>Consolidated file-level summary</h2>
 <div class="toolbar"><input id="search" type="search" placeholder="Search file name, process or output" aria-label="Search files"><select id="status" aria-label="Scan status"><option value="">All statuses</option><option value="full">Full</option><option value="partial">Partial</option><option value="failed">Failed</option></select><button class="secondary" id="refresh">Search</button></div>
 <div class="toolbar"><span id="chosen">0 selected</span><button id="compare" class="secondary">Compare selected</button><button id="csv">Summary CSV</button><button id="excel">Review Excel</button><button id="tabs" class="secondary">Tab detail CSV</button><button id="diagnostics" class="secondary">Diagnostics CSV</button><button id="findings-csv" class="secondary">Findings CSV</button><button id="bundle" class="secondary">Originals + analysis ZIP</button></div>
 <p class="muted">Exports use selected EUCs. With none selected, they include every EUC. The ZIP contains original files, so handle it according to your team's data rules.</p>
 <div class="tablewrap"><table><thead><tr><th><input type="checkbox" id="select-page" aria-label="Select visible page"></th><th>File</th><th>Scan status</th><th>Business Area Purpose</th><th>Process</th><th>Sub-Process</th><th>Key output</th><th>No. of Sheets - Total</th><th>No. of Sheets - Hidden</th><th>Errors</th><th>Review</th><th>Assessment status</th></tr></thead><tbody id="files"></tbody></table></div>
 <div class="pager"><button id="prev" class="secondary">Previous</button><span id="page"></span><button id="next" class="secondary">Next</button></div>
</section>
<section class="split"><section class="panel detail"><h2>Individual EUC analysis</h2><div id="detail" class="muted">Choose a file above to see its file assessment, worksheet detail and diagnostics.</div></section><section class="panel"><h2>Comparative review</h2><div id="comparison" class="muted">Select two or more EUCs, then choose Compare selected.</div></section></section>
</main>
<script>
const $=id=>document.getElementById(id), selected=new Set();let offset=0,total=0,pageRows=[];const limit=50;
function esc(x){return String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function val(f){if(!f)return '—';if(f.display!==undefined)return f.display;const v=f.value;if(v===null||v===undefined)return '—';if(Array.isArray(v))return v.join('; ');if(typeof v==='object')return JSON.stringify(v);return String(v)}
function short(x,n=150){x=String(x??'—');return x.length>n?x.slice(0,n-1)+'…':x}
function label(s){if(s==='business_area_process')return 'Business Area Purpose';return s.replaceAll('_',' ').replace(/\b\w/g,x=>x.toUpperCase())}
async function get(url){const r=await fetch(url);if(!r.ok)throw Error(await r.text());return r.json()}
function syncSelected(){$('chosen').textContent=`${selected.size} selected`;}
async function overview(){const d=await get('/api/overview');$('stats').innerHTML=[['Submitted EUCs',d.total],['Scanned',d.scanned],['Cached errors',d.errors],['Hidden sheets',d.hidden]].map(([a,b])=>`<div class="card"><b>${b}</b><span>${a}</span></div>`).join('');$('findings').innerHTML=d.findings.length?`<div class="tablewrap"><table><thead><tr><th>Type</th><th>Finding</th><th>EUCs</th><th>Severity</th><th>Evidence</th><th>Reviewer action</th></tr></thead><tbody>${d.findings.map(f=>`<tr><td>${esc(f.type)}</td><td>${esc(f.title)}</td><td>${f.affected_eucs}</td><td><span class="pill ${f.severity==='High'?'high':''}">${esc(f.severity)}</span></td><td>${esc(f.evidence)}</td><td>${esc(f.action)}</td></tr>`).join('')}</tbody></table></div>`:'No portfolio findings.'}
async function files(){const q=new URLSearchParams({offset,limit,q:$('search').value,status:$('status').value});const d=await get('/api/files?'+q);total=d.total;pageRows=d.rows;$('files').innerHTML=d.rows.map(r=>`<tr><td><input type="checkbox" data-id="${r.id}" ${selected.has(r.id)?'checked':''} aria-label="Select ${esc(r.file_name)}"></td><td><a href="#" data-open="${r.id}">${esc(r.file_name)}</a><br><small class="muted">${esc(r.file_id||'')}</small></td><td><span class="pill ${r.scan_status!=='full'?'partial':''}">${esc(r.scan_status)}</span></td><td>${esc(r.business_area||'Not assessed')}</td><td>${esc(r.process||'Not assessed')}</td><td>${esc(r.sub_process||'Not assessed')}</td><td title="${esc(r.output||'Not assessed')}">${esc(short(r.output||'Not assessed'))}</td><td>${esc(r.sheet_count)}</td><td>${esc(r.hidden_count)}</td><td>${esc(r.cached_error_count)}</td><td>${esc(r.review_count)}</td><td title="${esc(r.assessment_error||'')}">${esc(r.assessment_status||'Not assessed')}</td></tr>`).join('');$('page').textContent=`${total?offset+1:0}–${Math.min(total,offset+limit)} of ${total}`;$('prev').disabled=offset===0;$('next').disabled=offset+limit>=total;$('select-page').checked=d.rows.length>0&&d.rows.every(r=>selected.has(r.id));}
async function detail(id){const d=await get(`/api/files/${id}`);const a=d.assessment;let out=`<div class="actions">${a?`<button class="secondary" onclick="window.open('/report/${id}.html','_blank')">Open full HTML report</button><button class="secondary" onclick="location.href='/report/${id}.xlsx'">Download individual Excel</button>`:''}<button onclick="location.href='/original/${id}'">Download original EUC</button></div><h3>${esc(d.file.file_name)}</h3><p class="muted">Scan status: ${esc(d.file.scan_status)} · Scan error: ${esc(d.file.scan_error||'N/A')} · Assessment: ${esc(d.file.assessment_status||'Not assessed')}</p>`;
if(a){const core=['file_id','business_area_process','process','sub_process','purpose_of_file','key_output_outcome','complexity','key_inputs','key_outputs'];const show=keys=>'<dl>'+keys.map(k=>{const v=a.file[k];return `<dt>${esc(label(k))}</dt><dd>${esc(val(v))}<br><small class="muted">${esc(v.basis)}${v.confidence!==null?' · '+v.confidence:''}</small></dd>`}).join('')+'</dl>';out+='<h3>File assessment</h3>'+show(core)+'<details class="tabline"><summary>All file assessment fields</summary>'+show(Object.keys(a.file).filter(k=>!core.includes(k)))+'</details><h3>Worksheet detail</h3>'+a.tabs.map(t=>`<details class="tabline"><summary>${esc(val(t.tab_name))} · ${esc(val(t.tab_visibility))} · ${esc(val(t.tab_category))}</summary><dl>${Object.entries(t).filter(([k])=>k!=='tab_name').map(([k,v])=>`<dt>${esc(label(k))}</dt><dd>${esc(val(v))}<br><small class="muted">${esc(v.basis)}${v.evidence?.length?' · '+esc(v.evidence.join('; ')):''}</small></dd>`).join('')}</dl></details>`).join('');out+='<h3>Diagnostics</h3><p>'+esc(a.review.hidden_summary.header||'')+'</p>'+(a.review.hidden_summary.groups||[]).map(g=>`<div class="notice"><b>${esc(g.purpose)}</b> · ${g.count} hidden tab(s)<br>${esc(g.sheets.join(', '))}<br><small>Used by: ${esc(g.depended_on_by.join(', ')||'none detected')}</small></div>`).join('')+a.review.error_groups.map(g=>`<div class="notice"><b>${esc(g.error_type)}</b> · ${g.count} cells · ${esc(g.sheet)} · ${esc(g.area_type)}<br>${esc(g.reviewer_action)}</div>`).join('');}
if(!a){out+=`<div class="notice"><b>Assessment not produced</b><br>${esc(d.file.scan_error||'The workbook could not be scanned.')}. No worksheet-level findings or cached-error diagnostics were produced. Resolve the scan error and rescan this workbook.</div>`;}else if(d.file.assessment_error&&d.file.assessment_error!=='N/A'){out+=`<div class="notice"><b>AI narrative unavailable; offline assessment retained</b><br>${esc(d.file.assessment_error)}</div>`;}
$('detail').innerHTML=out;document.querySelector('.detail').scrollTop=0;}
async function compare(){const ids=[...selected].slice(0,20);if(ids.length<2){$('comparison').textContent='Select at least two EUCs.';return}const d=await get('/api/compare?ids='+ids.join(','));$('comparison').innerHTML='<p class="muted">Showing up to 20 selected EUCs; exports include the full selection.</p><div class="comparison">'+d.files.map(r=>`<div class="card"><b>${esc(r.file_name)}</b><p><strong>Business Area Purpose:</strong> ${esc(r.business_area||'—')}</p><p><strong>Process:</strong> ${esc(r.process||'—')}</p><p><strong>Sub-process:</strong> ${esc(r.sub_process||'—')}</p><p title="${esc(r.output||'')}"><strong>Output:</strong> ${esc(short(r.output))}</p><p><strong>Logic:</strong> ${esc(r.logic||'—')}</p><p><strong>No. of Sheets:</strong> ${r.sheet_count}; <strong>errors:</strong> ${r.cached_error_count}</p></div>`).join('')+'</div><h3>Structural matches</h3>'+(d.pairs.length?d.pairs.map(p=>`<div class="notice"><b>${esc(p.a)} ↔ ${esc(p.b)}</b> · ${esc(p.relationship)} · score ${p.score}</div>`).join(''):'No selected pair met the structural similarity threshold.');}
async function download(kind){const r=await fetch('/api/export',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({kind,ids:[...selected]})});if(!r.ok){alert(await r.text());return}const blob=await r.blob(),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=({'csv':'file_summary.csv','excel':'portfolio_review.xlsx','tabs':'worksheet_details.csv','diagnostics':'diagnostics.csv','findings':'portfolio_findings.csv','bundle':'euc_review_bundle.zip'})[kind];a.click();setTimeout(()=>URL.revokeObjectURL(a.href),30000)}
$('files').addEventListener('change',e=>{if(e.target.dataset.id){let id=Number(e.target.dataset.id);e.target.checked?selected.add(id):selected.delete(id);syncSelected()}});$('files').addEventListener('click',e=>{const id=e.target.dataset.open;if(id){e.preventDefault();detail(id)}});$('select-page').onchange=e=>{for(const r of pageRows)e.target.checked?selected.add(r.id):selected.delete(r.id);syncSelected();files()};$('refresh').onclick=()=>{offset=0;files()};$('search').onkeydown=e=>{if(e.key==='Enter'){offset=0;files()}};$('status').onchange=()=>{offset=0;files()};$('prev').onclick=()=>{offset-=limit;files()};$('next').onclick=()=>{offset+=limit;files()};$('compare').onclick=compare;for(const kind of ['csv','excel','tabs','diagnostics'])$(kind).onclick=()=>download(kind);$('findings-csv').onclick=()=>download('findings');$('bundle').onclick=()=>download('bundle');overview();files();
</script></html>"""


def make_handler(run_dir):
    run_dir = Path(run_dir).resolve()

    class Handler(BaseHTTPRequestHandler):
        def _db(self):
            return connect(run_dir)

        def _json(self, value, status=200):
            body = json.dumps(value, ensure_ascii=False, default=str).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _error(self, code, message):
            self._json({"error": message}, code)

        def _send_file(self, path, filename, mime):
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Disposition", f"attachment; filename=download; filename*=UTF-8''{quote(filename)}")
            self.send_header("Content-Length", str(os.path.getsize(path)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            with open(path, "rb") as fh:
                while block := fh.read(1024 * 1024):
                    self.wfile.write(block)

        def do_GET(self):
            u = urlsplit(self.path)
            query = parse_qs(u.query)
            if u.path == "/":
                body = DASHBOARD.encode()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            try:
                with closing(self._db()) as db:
                    if u.path == "/api/overview":
                        r = db.execute("SELECT COUNT(*) total,SUM(scan_status!='failed') scanned,SUM(cached_error_count) errors,SUM(hidden_count) hidden FROM files").fetchone()
                        findings = [dict(x) for x in db.execute("SELECT * FROM findings ORDER BY CASE severity WHEN 'High' THEN 0 ELSE 1 END,affected_eucs DESC")]
                        return self._json({"total": r["total"], "scanned": r["scanned"] or 0,
                                           "errors": r["errors"] or 0, "hidden": r["hidden"] or 0,
                                           "findings": findings})
                    if u.path == "/api/files":
                        q = query.get("q", [""])[0].strip()
                        status = query.get("status", [""])[0]
                        limit = min(100, max(1, int(query.get("limit", [50])[0])))
                        offset = max(0, int(query.get("offset", [0])[0]))
                        where, args = [], []
                        if q:
                            where.append("(file_name LIKE ? OR summary_json LIKE ?)")
                            args.extend([f"%{q}%", f"%{q}%"])
                        if status in ("full", "partial", "failed"):
                            where.append("scan_status=?")
                            args.append(status)
                        clause = " WHERE " + " AND ".join(where) if where else ""
                        total = db.execute("SELECT COUNT(*) FROM files" + clause, args).fetchone()[0]
                        rows = []
                        for r in db.execute("SELECT * FROM files" + clause + " ORDER BY file_name,id LIMIT ? OFFSET ?", (*args, limit, offset)):
                            s = json.loads(r["summary_json"]) if r["summary_json"] else {}
                            failed = r["scan_status"] == "failed"
                            rows.append({"id": r["id"], "file_id": r["file_id"] or ("Not available — scan failed" if failed else "Not recorded"), "file_name": r["file_name"],
                                         "scan_status": r["scan_status"], "business_area": s.get("Business Area Purpose"),
                                         "process": s.get("Process"), "sub_process": s.get("Sub-Process"),
                                         "output": s.get("Key Output / Outcome"),
                                         "sheet_count": "Not assessed" if failed else r["sheet_count"],
                                         "hidden_count": "Not assessed" if failed else r["hidden_count"],
                                         "cached_error_count": "Not assessed" if failed else r["cached_error_count"],
                                         "review_count": "Not assessed" if failed else r["review_count"],
                                         "assessment_status": ("Not assessed" if failed else s.get("Assessment Status", "Not recorded — rescan to populate")),
                                         "assessment_error": ("Workbook could not be scanned; see Scan Error." if failed else s.get("Assessment Error", "N/A"))})
                        return self._json({"total": total, "rows": rows})
                    m = re.fullmatch(r"/api/files/(\d+)", u.path)
                    if m:
                        r = db.execute("SELECT * FROM files WHERE id=?", (int(m.group(1)),)).fetchone()
                        if r is None:
                            return self._error(404, "File not found")
                        a = json.loads(r["assessment_json"]) if r["assessment_json"] else None
                        if a:
                            for field in a["file"].values():
                                field["display"] = fmt_value(field["value"])
                            for tab in a["tabs"]:
                                for field in tab.values():
                                    field["display"] = fmt_value(field["value"])
                        public_file = {key: r[key] for key in (
                            "id", "file_id", "file_name", "scan_status",
                            "sheet_count", "hidden_count", "cached_error_count", "review_count",
                        )}
                        public_file["scan_error"] = _display_scan_error(r["scan_error"])
                        public_file["assessment_status"] = (
                            "Not assessed" if r["scan_status"] == "failed" else
                            (a.get("scan", {}).get("assessment_status") if a else "Not recorded — rescan to populate"))
                        public_file["assessment_error"] = (
                            "Workbook could not be scanned; see Scan Error." if r["scan_status"] == "failed" else
                            (a.get("scan", {}).get("assessment_error") or "N/A" if a else "N/A"))
                        return self._json({"file": public_file, "assessment": a})
                    if u.path == "/api/compare":
                        ids = [int(x) for x in query.get("ids", [""])[0].split(",") if x.isdigit()][:20]
                        rows = []
                        fps = []
                        for r in selected_files(db, ids):
                            s = json.loads(r["summary_json"]) if r["summary_json"] else {}
                            failed = r["scan_status"] == "failed"
                            fallback = "Not assessed" if failed else "Not recorded — rescan to populate"
                            rows.append({"file_name": r["file_name"], "scan_status": r["scan_status"],
                                         "business_area": s.get("Business Area Purpose") or fallback,
                                         "process": s.get("Process") or fallback,
                                         "sub_process": s.get("Sub-Process") or fallback,
                                         "output": s.get("Key Output / Outcome") or fallback,
                                         "logic": s.get("Logic Types") or fallback,
                                         "sheet_count": "Not assessed" if failed else r["sheet_count"],
                                         "cached_error_count": "Not assessed" if failed else r["cached_error_count"]})
                            if r["fingerprint_json"]:
                                fps.append((r["file_name"], _restore_fp(r["fingerprint_json"])))
                        pairs = []
                        for i, (a_name, a_fp) in enumerate(fps):
                            for b_name, b_fp in fps[i + 1:]:
                                score = similarity(a_fp, b_fp)
                                if score["overall"] >= SIMILAR_THRESHOLD or score["skeleton"] >= SKELETON_DUP:
                                    pairs.append({"a": a_name, "b": b_name, "score": score["overall"],
                                                  "relationship": "Structural similarity review"})
                        pairs.sort(key=lambda p: -p["score"])
                        return self._json({"files": rows, "pairs": pairs[:30]})
                    m = re.fullmatch(r"/original/(\d+)", u.path)
                    if m:
                        r = db.execute("SELECT * FROM files WHERE id=?", (int(m.group(1)),)).fetchone()
                        if r is None:
                            return self._error(404, "Original workbook not found")
                        if not source_is_current(r):
                            return self._error(409, "Original workbook missing or changed since scan")
                        return self._send_file(r["source_path"], f"{r['id']}_{Path(r['source_path']).name}",
                                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    m = re.fullmatch(r"/report/(\d+)\.(html|xlsx)", u.path)
                    if m:
                        r = db.execute("SELECT * FROM files WHERE id=?", (int(m.group(1)),)).fetchone()
                        if r is None or not r["assessment_json"]:
                            return self._error(404, "Individual report unavailable")
                        if not source_is_current(r):
                            return self._error(409, "Original workbook missing or changed since scan")
                        from .scan import xray_workbook
                        wx = xray_workbook(r["source_path"])
                        a = assessment_from_json(r["assessment_json"])
                        if m.group(2) == "html":
                            from .report import build_report
                            body = build_report(wx, a).encode()
                            self.send_response(200)
                            self.send_header("Content-Type", "text/html; charset=utf-8")
                            self.send_header("Content-Length", str(len(body)))
                            self.end_headers()
                            self.wfile.write(body)
                            return
                        from .xlsx_report import write_xlsx_report
                        with tempfile.TemporaryDirectory() as tmp:
                            out = Path(tmp) / f"{r['id']}_analysis.xlsx"
                            write_xlsx_report(wx, a, str(out))
                            return self._send_file(out, out.name,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except (ValueError, KeyError) as exc:
                return self._error(400, str(exc))
            except Exception as exc:
                return self._error(500, str(exc))
            self._error(404, "Route not found")

        def do_POST(self):
            if self.path != "/api/export":
                return self._error(404, "Route not found")
            length = int(self.headers.get("Content-Length", "0"))
            if length > 100_000:
                return self._error(413, "Selection too large")
            try:
                request = json.loads(self.rfile.read(length))
                kind = request["kind"]
                ids = request.get("ids") or None  # no selection means all
                if ids is not None:
                    if not isinstance(ids, list) or len(ids) > 10_000 or any(type(x) is not int for x in ids):
                        return self._error(400, "Invalid file selection")
                choices = {
                    "csv": ("file_summary.csv", "text/csv", SUMMARY_COLUMNS, summary_rows),
                    "tabs": ("worksheet_details.csv", "text/csv", TAB_COLUMNS, tab_detail_rows),
                    "diagnostics": ("diagnostics.csv", "text/csv", DIAGNOSTIC_COLUMNS, diagnostic_rows),
                    "findings": ("portfolio_findings.csv", "text/csv", FINDING_COLUMNS, finding_rows),
                    "excel": ("portfolio_review.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", None, None),
                    "bundle": ("euc_review_bundle.zip", "application/zip", None, None),
                }
                if kind not in choices:
                    return self._error(400, "Unknown export type")
                name, mime, headers, rows = choices[kind]
                with closing(self._db()) as db, tempfile.TemporaryDirectory() as tmp:
                    out = Path(tmp) / name
                    if kind == "excel":
                        write_excel(out, db, ids)
                    elif kind == "bundle":
                        bundle(out, db, ids)
                    else:
                        write_csv(out, headers, rows(db, ids))
                    self._send_file(out, name, mime)
            except Exception as exc:
                self._error(500, str(exc))

    return Handler


def serve(run_dir, host="127.0.0.1", port=8765):
    if not (Path(run_dir) / DB_NAME).is_file():
        raise FileNotFoundError(f"No {DB_NAME} in {run_dir}")
    server = ThreadingHTTPServer((host, port), make_handler(run_dir))
    print(f"Portfolio dashboard: http://{host}:{server.server_port}/", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
