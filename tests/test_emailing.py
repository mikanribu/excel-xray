from __future__ import annotations

from pathlib import Path

import pytest

from excel_xray import emailing
from excel_xray.emailing import EmailDeliveryError, send_excel_report


class FakeSMTP:
    instances = []

    def __init__(self, host, port, timeout):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.login_args = None
        self.message = None
        self.to_addrs = None
        self.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def login(self, sender, password):
        self.login_args = (sender, password)

    def send_message(self, message, to_addrs=None):
        self.message = message
        self.to_addrs = to_addrs


def test_send_excel_report_attaches_file_and_uses_configured_gmail(monkeypatch, tmp_path):
    FakeSMTP.instances.clear()
    monkeypatch.setattr(emailing.smtplib, "SMTP_SSL", FakeSMTP)
    report = tmp_path / "portfolio_review.xlsx"
    report.write_bytes(b"excel workbook bytes")

    send_excel_report(report, sender="sender@gmail.com", app_password="app-password")

    smtp = FakeSMTP.instances[0]
    assert (smtp.host, smtp.port, smtp.timeout) == ("smtp.gmail.com", 465, 30)
    assert smtp.login_args == ("sender@gmail.com", "app-password")
    assert smtp.message["To"] == "jacobweglarz@gmail.com, clementine.pages@gmail.com"
    assert smtp.to_addrs == ["jacobweglarz@gmail.com", "clementine.pages@gmail.com"]
    attachment = list(smtp.message.iter_attachments())[0]
    assert attachment.get_filename() == "portfolio_review.xlsx"
    assert attachment.get_content() == b"excel workbook bytes"


def test_send_excel_report_reads_credentials_from_environment(monkeypatch, tmp_path):
    FakeSMTP.instances.clear()
    monkeypatch.setattr(emailing.smtplib, "SMTP_SSL", FakeSMTP)
    monkeypatch.setenv("GMAIL_ADDRESS", "sender@gmail.com")
    monkeypatch.setenv("GMAIL_APP_PASSWORD", "app-password")
    report = tmp_path / "report.xlsx"
    report.write_bytes(b"xlsx")

    send_excel_report(report, recipient="reviewer@example.com, clementine@example.com")

    smtp = FakeSMTP.instances[0]
    assert smtp.login_args == ("sender@gmail.com", "app-password")
    assert smtp.message["To"] == "reviewer@example.com, clementine@example.com"
    assert smtp.to_addrs == ["reviewer@example.com", "clementine@example.com"]


def test_send_excel_report_requires_credentials_without_attempting_smtp(monkeypatch, tmp_path):
    monkeypatch.delenv("GMAIL_ADDRESS", raising=False)
    monkeypatch.delenv("GMAIL_APP_PASSWORD", raising=False)
    report = tmp_path / "report.xlsx"
    report.write_bytes(b"xlsx")

    with pytest.raises(EmailDeliveryError, match="GMAIL_ADDRESS and GMAIL_APP_PASSWORD"):
        send_excel_report(report)


def test_send_excel_report_rejects_oversize_and_non_excel_attachments(tmp_path):
    report = tmp_path / "large.xlsx"
    with report.open("wb") as stream:
        stream.truncate(emailing.MAX_ATTACHMENT_BYTES + 1)
    with pytest.raises(EmailDeliveryError, match="limited"):
        send_excel_report(report, sender="sender@gmail.com", app_password="secret")

    csv_report = tmp_path / "report.csv"
    csv_report.write_text("a,b\n", encoding="utf-8")
    with pytest.raises(EmailDeliveryError, match=".xlsx reports only"):
        send_excel_report(csv_report, sender="sender@gmail.com", app_password="secret")


def test_cli_folder_sends_only_the_consolidated_workbook(monkeypatch, tmp_path, capsys):
    from excel_xray import cli

    source_dir = tmp_path / "inputs"
    source_dir.mkdir()
    (source_dir / "euc.xlsx").write_bytes(b"source")
    calls = []

    def fake_scan(paths, run_dir, **kwargs):
        assert len(paths) == 1
        run_dir = Path(run_dir)
        run_dir.mkdir(parents=True)
        (run_dir / "portfolio_review.xlsx").write_bytes(b"consolidated")
        return {"stored": 1, "failed": 0, "reused": 0}

    def fake_send(path, *, recipient):
        calls.append((Path(path).name, recipient))

    monkeypatch.setattr("excel_xray.portfolio.scan_portfolio", fake_scan)
    monkeypatch.setattr(cli, "send_excel_report", fake_send)
    monkeypatch.setattr(
        "sys.argv",
        ["excel-xray", str(source_dir)],
    )

    assert cli.main() == 0
    assert calls == [(
        "portfolio_review.xlsx",
        "jacobweglarz@gmail.com, clementine.pages@gmail.com",
    )]
    assert "Emailed portfolio_review.xlsx" in capsys.readouterr().err


def test_cli_folder_can_keep_report_local(monkeypatch, tmp_path):
    from excel_xray import cli

    source_dir = tmp_path / "inputs"
    source_dir.mkdir()
    (source_dir / "euc.xlsx").write_bytes(b"source")

    def fake_scan(paths, run_dir, **kwargs):
        run_dir = Path(run_dir)
        run_dir.mkdir(parents=True)
        (run_dir / "portfolio_review.xlsx").write_bytes(b"consolidated")
        return {"stored": 1, "failed": 0, "reused": 0}

    def unexpected_send(*args, **kwargs):
        pytest.fail("email should be disabled with --no-email-report")

    monkeypatch.setattr("excel_xray.portfolio.scan_portfolio", fake_scan)
    monkeypatch.setattr(cli, "send_excel_report", unexpected_send)
    monkeypatch.setattr("sys.argv", ["excel-xray", str(source_dir), "--no-email-report"])

    assert cli.main() == 0


def test_cli_single_workbook_emails_generated_report_by_default(
    monkeypatch, tmp_path, fixture_path,
):
    from excel_xray import cli

    calls = []

    def fake_send(path, *, recipient):
        calls.append((Path(path).name, recipient))

    monkeypatch.setattr(cli, "send_excel_report", fake_send)
    monkeypatch.setattr(
        "sys.argv",
        ["excel-xray", fixture_path, "-o", str(tmp_path)],
    )

    assert cli.main() == 0
    assert len(calls) == 1
    assert calls[0][0].endswith(".xlsx")
    assert calls[0][1] == "jacobweglarz@gmail.com, clementine.pages@gmail.com"
