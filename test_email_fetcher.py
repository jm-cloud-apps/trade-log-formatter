"""Fetcher regression tests — the 08/07/2026 stranding.

That report was in Gmail with a valid PDF attached, and the fetcher skipped it
on every run for eight days because its UID sat in `processed_uids` while the
PDF had never been written to disk. One desync between the side-car state and
the folder cost a whole trading day, and a missing day doesn't just drop rows —
it re-points the FIFO matcher at the wrong inventory.

The fix makes the *disk* the authority: state may say "seen", but if the PDF
isn't there, the report is downloaded anyway, and nothing is marked processed
until a file actually lands.

Run: python3 -m unittest test_email_fetcher -v
"""

import email.message
import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import mock

import email_fetcher


def _report_message(stamp="20260807", subject_date="08/07/2026", attach=True,
                    msgid="<API.1@ny5ibop3>"):
    msg = email.message.EmailMessage()
    msg["Subject"] = f"Daily Trade Report for {subject_date}"
    msg["From"] = "donotreply@interactivebrokers.com"
    msg["Message-ID"] = msgid
    msg.set_content("Your daily trade report is attached.")
    if attach:
        msg.add_attachment(b"%PDF-1.4 fake", maintype="application",
                           subtype="octet-stream",
                           filename=f"DailyTradeReport.{stamp}.pdf")
    return msg


class FakeIMAP:
    """Just enough IMAP for fetch_new_reports: one message, uid 3286."""

    def __init__(self, msg, uid=b"3286"):
        self._msg = msg
        self._uid = uid
        self.stored_labels = []

    # context-manager protocol — fetch_new_reports uses `with imaplib...`
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def login(self, *a):
        return "OK", []

    def select(self, *a, **kw):
        return "OK", [b"1"]

    def noop(self):
        return "OK", []

    def logout(self):
        return "OK", []

    def search(self, *a):
        return "OK", [self._uid]

    def store(self, uid, cmd, value):
        self.stored_labels.append((uid, value))
        return "OK", []

    def fetch(self, uid_bytes, spec):
        raw = self._msg.as_bytes()
        if "HEADER" in spec:
            head = b"\r\n".join(
                line for line in raw.split(b"\r\n\r\n", 1)[0].split(b"\r\n")
                if line.lower().startswith((b"subject:", b"message-id:"))
            ) + b"\r\n\r\n"
            return "OK", [(b"1 (BODY[HEADER]", head)]
        return "OK", [(b"1 (RFC822", raw)]


class ReportOnDiskTests(unittest.TestCase):
    def test_finds_the_pdf_by_its_date_stamp(self):
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            folder = base / "08.2026"
            folder.mkdir()
            pdf = folder / "DailyTradeReport.20260807.pdf"
            pdf.write_bytes(b"x")
            found = email_fetcher.report_pdf_on_disk(base, datetime(2026, 8, 7))
            self.assertEqual(found, pdf)

    def test_a_renamed_file_still_counts_as_present(self):
        # Matching on the stamp, not an exact name, keeps a manually recovered
        # or safe_filename-mangled report from being downloaded twice.
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            folder = base / "08.2026"
            folder.mkdir()
            (folder / "recovered_DailyTradeReport.20260807 (1).pdf").write_bytes(b"x")
            self.assertIsNotNone(email_fetcher.report_pdf_on_disk(base, datetime(2026, 8, 7)))

    def test_absent_report_and_absent_folder_are_both_none(self):
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            (base / "08.2026").mkdir()
            self.assertIsNone(email_fetcher.report_pdf_on_disk(base, datetime(2026, 8, 7)))
            self.assertIsNone(email_fetcher.report_pdf_on_disk(base, datetime(2026, 9, 9)))

    def test_another_days_pdf_does_not_satisfy_the_check(self):
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            folder = base / "08.2026"
            folder.mkdir()
            (folder / "DailyTradeReport.20260806.pdf").write_bytes(b"x")
            self.assertIsNone(email_fetcher.report_pdf_on_disk(base, datetime(2026, 8, 7)))


class FetchTests(unittest.TestCase):
    def _run(self, msg, state, tmp):
        config = {
            "imap_host": "h", "imap_port": 993, "email": "e", "app_password": "p",
            "ib_sender": "donotreply@interactivebrokers.com",
            "subject_prefix": "Daily Trade Report for",
            "download_base": tmp, "lookback_days": 30,
            "processed_label": "",          # don't exercise labelling here
        }
        fake = FakeIMAP(msg)
        with mock.patch.object(email_fetcher.imaplib, "IMAP4_SSL", return_value=fake), \
             mock.patch.object(email_fetcher, "save_state"):
            saved = email_fetcher.fetch_new_reports(config, state)
        return saved, state

    def test_a_uid_marked_processed_is_redownloaded_when_the_pdf_is_missing(self):
        """The exact 08/07/2026 failure: state says seen, disk says otherwise."""
        with TemporaryDirectory() as tmp:
            state = {"processed_uids": ["3286"], "processed_message_ids": []}
            saved, state = self._run(_report_message(), state, tmp)

            self.assertEqual(len(saved), 1, "the stranded report should be re-downloaded")
            self.assertTrue(saved[0].exists())
            self.assertEqual(saved[0].name, "DailyTradeReport.20260807.pdf")
            self.assertEqual(saved[0].parent.name, "08.2026")

    def test_a_report_already_on_disk_is_not_downloaded_again(self):
        with TemporaryDirectory() as tmp:
            folder = Path(tmp) / "08.2026"
            folder.mkdir()
            (folder / "DailyTradeReport.20260807.pdf").write_bytes(b"original")

            state = {"processed_uids": [], "processed_message_ids": []}
            saved, state = self._run(_report_message(), state, tmp)

            self.assertEqual(saved, [], "should not re-download what's already here")
            self.assertEqual((folder / "DailyTradeReport.20260807.pdf").read_bytes(), b"original")
            # and the state self-repairs so the next run skips it cheaply
            self.assertIn("3286", state["processed_uids"])

    def test_state_records_the_uid_and_message_id_after_a_real_download(self):
        with TemporaryDirectory() as tmp:
            state = {"processed_uids": [], "processed_message_ids": []}
            saved, state = self._run(_report_message(), state, tmp)

            self.assertEqual(len(saved), 1)
            self.assertIn("3286", state["processed_uids"])
            self.assertIn("<API.1@ny5ibop3>", state["processed_message_ids"])

    def test_a_message_with_no_pdf_is_left_unmarked_for_the_next_run(self):
        # Marking unconditionally is what turned one bad run into a permanently
        # missing day — an attachment-less message must stay retryable.
        with TemporaryDirectory() as tmp:
            state = {"processed_uids": [], "processed_message_ids": []}
            saved, state = self._run(_report_message(attach=False), state, tmp)

            self.assertEqual(saved, [])
            self.assertNotIn("3286", state["processed_uids"])
            self.assertNotIn("<API.1@ny5ibop3>", state["processed_message_ids"])


if __name__ == "__main__":
    unittest.main()
