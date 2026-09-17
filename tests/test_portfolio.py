"""Portfolio output is consolidated, selectable and faithful to source files."""

from __future__ import annotations

import csv
import hashlib
import json
import zipfile
from pathlib import Path

import openpyxl

from excel_xray.portfolio import (DB_NAME, EXCEL_NAME, SUMMARY_NAME, _compare,
                                  assessment_from_json, bundle, connect,
                                  scan_portfolio, source_is_current, write_excel)
from excel_xray.report import build_report
from excel_xray.scan import xray_workbook


FIXTURES = Path(__file__).parent / "fixtures"


def test_portfolio_one_file_row_per_euc_and_tab_drilldown(tmp_path):
    paths = [str(FIXTURES / "complex_euc_model.xlsx"),
             str(FIXTURES / "messy_reserving_model.xlsx")]
    result = scan_portfolio(paths, tmp_path)
    assert result["stored"] == 2
    assert not list(tmp_path.glob("*.json"))
    assert not list(tmp_path.glob("xray_*.xlsx"))
    with (tmp_path / SUMMARY_NAME).open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert len(rows) == 2
    assert {r["File Name"] for r in rows} == {Path(p).name for p in paths}
    assert all(r["File ID"] and r["Scan Status"] == "full" for r in rows)
    wb = openpyxl.load_workbook(tmp_path / EXCEL_NAME, read_only=True)
    assert wb.sheetnames == ["File summary", "Worksheet details", "Diagnostics", "Portfolio findings"]
    assert sum(1 for _ in wb["File summary"].values) == 3
    assert sum(1 for _ in wb["Worksheet details"].values) == 16
    with connect(tmp_path) as db:
        assert db.execute("SELECT COUNT(*) FROM tabs").fetchone()[0] == 15
        row = db.execute("SELECT * FROM files ORDER BY id LIMIT 1").fetchone()
        assert source_is_current(row)
        a = assessment_from_json(row["assessment_json"])
        html = build_report(xray_workbook(row["source_path"]), a)
        assert "EUC assessment" in html


def test_selected_bundle_contains_only_selected_original_and_analysis(tmp_path):
    paths = [str(FIXTURES / "complex_euc_model.xlsx"),
             str(FIXTURES / "messy_reserving_model.xlsx")]
    scan_portfolio(paths, tmp_path)
    with connect(tmp_path) as db:
        row = db.execute("SELECT * FROM files ORDER BY id LIMIT 1").fetchone()
        dest = tmp_path / "selected.zip"
        bundle(dest, db, [row["id"]])
        with zipfile.ZipFile(dest) as zf:
            originals = [n for n in zf.namelist() if n.startswith("originals/")]
            assert len(originals) == 1
            assert hashlib.sha256(zf.read(originals[0])).hexdigest() == row["sha256"]
            with zf.open(SUMMARY_NAME) as fh:
                assert len(list(csv.reader(line.decode("utf-8-sig") for line in fh))) == 2
            with zf.open(EXCEL_NAME) as fh:
                assert fh.read(2) == b"PK"
        selected_excel = tmp_path / "selected.xlsx"
        write_excel(selected_excel, db, [row["id"]])
        wb = openpyxl.load_workbook(selected_excel, read_only=True)
        assert sum(1 for _ in wb["File summary"].values) == 2


def test_resume_reuses_unchanged_workbooks_and_failed_rows_are_retained(tmp_path):
    good = str(FIXTURES / "messy_reserving_model.xlsx")
    bad = tmp_path / "corrupt.xlsx"
    bad.write_bytes(b"not an OOXML workbook")
    first = scan_portfolio([good, str(bad)], tmp_path / "run")
    assert first["stored"] == 2 and first["failed"] == 1
    second = scan_portfolio([good, str(bad)], tmp_path / "run", resume=True)
    assert second["reused"] == 1 and second["failed"] == 1
    with connect(tmp_path / "run") as db:
        assert db.execute("SELECT COUNT(*) FROM files").fetchone()[0] == 2
        assert db.execute("SELECT scan_error FROM files WHERE scan_status='failed'").fetchone()[0]


def test_source_change_blocks_original_export(tmp_path):
    source = tmp_path / "source.xlsx"
    source.write_bytes((FIXTURES / "messy_reserving_model.xlsx").read_bytes())
    scan_portfolio([str(source)], tmp_path / "run")
    with connect(tmp_path / "run") as db:
        row = db.execute("SELECT * FROM files").fetchone()
        assert source_is_current(row)
        source.write_bytes(source.read_bytes() + b"changed")
        assert not source_is_current(row)
        try:
            bundle(tmp_path / "changed.zip", db, [row["id"]])
        except FileNotFoundError:
            pass
        else:
            assert False, "bundle should reject changed originals"


def test_two_thousand_records_compare_without_individual_reports(tmp_path):
    with connect(tmp_path) as db:
        for i in range(2000):
            fp = {"file_id": str(i), "file_name": f"f{i}.xlsx", "logic_type": "Calculation",
                  "skeletons": [f"shape_{i}"], "headers": [f"header_{i}"],
                  "functions": [], "sheet_names": [f"sheet_{i}"]}
            db.execute("""INSERT INTO files(source_path,file_id,file_name,scan_status,
                       summary_json,assessment_json,fingerprint_json) VALUES(?,?,?,?,?,?,?)""",
                       (f"/synthetic/f{i}.xlsx", str(i), fp["file_name"], "full", "{}",
                        json.dumps({"file": {}, "tabs": [], "review": {}}), json.dumps(fp)))
        _compare(db)
        assert db.execute("SELECT COUNT(*) FROM files").fetchone()[0] == 2000
        assert db.execute("SELECT COUNT(*) FROM pairs").fetchone()[0] == 0
        assert db.execute("SELECT COUNT(*) FROM files WHERE assessment_json IS NOT NULL").fetchone()[0] == 2000


def test_identical_templates_keep_a_bounded_match_shortlist(tmp_path):
    with connect(tmp_path) as db:
        for i in range(100):
            fp = {"file_id": str(i), "file_name": f"f{i}.xlsx", "logic_type": "Calculation",
                  "skeletons": ["same_shape"], "headers": ["same_header"],
                  "functions": ["SUM"], "sheet_names": ["same_sheet"]}
            db.execute("""INSERT INTO files(source_path,file_id,file_name,scan_status,
                       summary_json,assessment_json,fingerprint_json) VALUES(?,?,?,?,?,?,?)""",
                       (f"/synthetic/f{i}.xlsx", str(i), fp["file_name"], "full", "{}",
                        json.dumps({"file": {}, "tabs": [], "review": {}}), json.dumps(fp)))
        _compare(db)
        stored = db.execute("SELECT COUNT(*) FROM pairs").fetchone()[0]
        assert 0 < stored <= 100 * 20
        assert stored < 100 * 99 // 2
