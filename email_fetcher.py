"""Fetch new Interactive Brokers Daily Trade Report PDFs from Gmail via IMAP.

Connects to Gmail using an app password, searches for messages from
donotreply@interactivebrokers.com with subject "Daily Trade Report for MM/DD/YYYY",
extracts PDF attachments, and saves them into the monthly trades folder
(MM.YYYY) matching the existing trade-log-formatter convention.

Tracks processed messages in email_state.json so reruns skip duplicates.
"""

import email
import email.message
import imaplib
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from pathlib import Path

logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
STATE_PATH = SCRIPT_DIR / "email_state.json"

SUBJECT_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing {CONFIG_PATH.name}. Copy config.example.json to config.json "
            f"and fill in your Gmail app password."
        )
    with open(CONFIG_PATH) as f:
        return json.load(f)


def load_state() -> dict:
    if not STATE_PATH.exists():
        return {"last_fetched_utc": None, "processed_uids": [], "processed_message_ids": []}
    with open(STATE_PATH) as f:
        return json.load(f)


def save_state(state: dict) -> None:
    tmp = STATE_PATH.with_suffix(".json.tmp")
    with open(tmp, "w") as f:
        json.dump(state, f, indent=2)
    os.replace(tmp, STATE_PATH)


def decode_subject(raw_subject: str) -> str:
    parts = decode_header(raw_subject or "")
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def parse_report_date(subject: str) -> datetime | None:
    """Extract the MM/DD/YYYY date from an IB Daily Trade Report subject."""
    m = SUBJECT_DATE_RE.search(subject)
    if not m:
        return None
    month, day, year = (int(x) for x in m.groups())
    try:
        return datetime(year, month, day)
    except ValueError:
        return None


def imap_since_string(days_back: int) -> str:
    """IMAP SINCE wants a date like '01-Jan-2026'.

    SINCE filters server-side on INTERNALDATE by *date only*. We derive the
    window from UTC now, which can already be the next calendar day when a run
    fires late in the local evening — that used to clip a report sitting right
    on the boundary (the 06/23/2026 report was skipped this way by a delayed
    run). Subtract one extra day of margin so the boundary is always covered;
    dedup (processed_uids / message-ids) makes the overlap free.
    """
    since = (datetime.now(timezone.utc) - timedelta(days=days_back + 1)).date()
    return since.strftime("%d-%b-%Y")


def extract_pdf_attachments(msg: email.message.Message) -> list[tuple[str, bytes]]:
    out = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if not filename:
            continue
        filename = decode_subject(filename)  # filenames can be encoded the same way
        if not filename.lower().endswith(".pdf"):
            continue
        payload = part.get_payload(decode=True)
        if payload:
            out.append((filename, payload))
    return out


def ensure_month_folder(base: Path, report_date: datetime) -> Path:
    folder = base / report_date.strftime("%m.%Y")
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def safe_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._\-]", "_", name)


def report_pdf_on_disk(base: Path, report_date: datetime) -> Path | None:
    """The already-downloaded PDF for this report date, if there is one.

    Matches on the YYYYMMDD stamp anywhere in the filename rather than on an
    exact name, so a renamed or safe_filename-mangled attachment still counts
    as present instead of being downloaded a second time.
    """
    folder = base / report_date.strftime("%m.%Y")
    if not folder.is_dir():
        return None
    stamp = report_date.strftime("%Y%m%d")
    for p in sorted(folder.glob("*.pdf")):
        if stamp in p.name:
            return p
    return None


def fetch_new_reports(config: dict, state: dict) -> list[Path]:
    """Connect to IMAP, download new IB report PDFs, return list of saved paths."""
    host = config["imap_host"]
    port = config["imap_port"]
    user = config["email"]
    password = config["app_password"]
    sender = config["ib_sender"]
    subject_prefix = config["subject_prefix"]
    download_base = Path(config["download_base"])
    lookback = int(config.get("lookback_days", 7))

    start_after_str = config.get("start_after_date")
    start_after = (
        datetime.strptime(start_after_str, "%Y-%m-%d") if start_after_str else None
    )

    processed_uids = set(state.get("processed_uids", []))
    processed_msgids = set(state.get("processed_message_ids", []))

    saved_paths: list[Path] = []

    apply_label = config.get("processed_label", "IB/Processed")

    logger.info("Connecting to %s:%s as %s", host, port, user)
    with imaplib.IMAP4_SSL(host, port) as M:
        M.login(user, password)
        # Need read-write so we can apply a Gmail label after processing.
        M.select("INBOX", readonly=False)

        # Gmail IMAP returns stale SEARCH results on a freshly-selected
        # mailbox — the first run after a new email arrives often shows zero
        # matches, while a second run a few seconds later sees them. NOOP
        # forces the server to flush pending EXISTS / EXPUNGE updates into
        # this session, and a single retry covers the rare case where the
        # server is still warming up.
        try:
            M.noop()
        except Exception as e:
            logger.debug("NOOP raised %s; continuing anyway", e)

        since = imap_since_string(lookback)
        uids: list[bytes] = []
        retry_attempts = max(1, int(config.get("search_retries", 2)))
        retry_sleep = float(config.get("search_retry_sleep", 2.5))

        for attempt in range(1, retry_attempts + 1):
            # Gmail IMAP supports standard SEARCH; FROM + SINCE narrows server-side.
            typ, data = M.search(None, "FROM", f'"{sender}"', "SINCE", since)
            if typ != "OK":
                logger.error("IMAP search failed: %s", typ)
                return saved_paths

            uids = data[0].split()
            logger.info(
                "Found %d candidate message(s) since %s (attempt %d/%d)",
                len(uids), since, attempt, retry_attempts,
            )
            if uids or attempt == retry_attempts:
                break
            # No matches yet — re-NOOP and wait before retrying once more.
            logger.info("No matches; waiting %.1fs and retrying SEARCH…", retry_sleep)
            time.sleep(retry_sleep)
            try:
                M.noop()
            except Exception:
                pass

        for uid_bytes in uids:
            uid = uid_bytes.decode()

            # Headers first. They're cheap, and they carry the report date —
            # which is what lets the *disk* decide whether this one is already
            # in hand. The processed-uid set is only an optimisation here; it
            # is never allowed to veto a download, because a single desync
            # between that list and the folder used to strand a report
            # permanently (the 08/07/2026 report was lost exactly this way:
            # uid marked, message-id absent, PDF never written).
            typ, hdr = M.fetch(uid_bytes, "(BODY.PEEK[HEADER.FIELDS (SUBJECT MESSAGE-ID)])")
            if typ != "OK" or not hdr or not hdr[0]:
                logger.warning("Failed to fetch headers for UID %s", uid)
                continue

            head = email.message_from_bytes(hdr[0][1])
            subject = decode_subject(head.get("Subject", ""))
            message_id = (head.get("Message-ID") or "").strip()

            if not subject.startswith(subject_prefix):
                logger.debug("UID %s subject %r doesn't match prefix, skipping", uid, subject)
                continue

            report_date = parse_report_date(subject)
            if not report_date:
                logger.warning("UID %s subject %r has no parseable date", uid, subject)
                continue

            if start_after and report_date <= start_after:
                logger.info(
                    "Skipping %s (report date %s <= start_after %s)",
                    subject, report_date.date(), start_after.date(),
                )
                processed_uids.add(uid)
                if message_id:
                    processed_msgids.add(message_id)
                continue

            existing = report_pdf_on_disk(download_base, report_date)
            if existing:
                logger.debug("Report %s already on disk (%s)", report_date.date(), existing.name)
                # Also repairs a state file that lost track of a report it does
                # in fact have.
                processed_uids.add(uid)
                if message_id:
                    processed_msgids.add(message_id)
                continue

            if uid in processed_uids or (message_id and message_id in processed_msgids):
                logger.warning(
                    "UID %s (%s) is marked processed but no PDF is on disk — re-downloading",
                    uid, subject,
                )

            typ, msg_data = M.fetch(uid_bytes, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                logger.warning("Failed to fetch UID %s", uid)
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            attachments = extract_pdf_attachments(msg)
            if not attachments:
                logger.warning("UID %s (%s) has no PDF attachments", uid, subject)
                continue

            dest_folder = ensure_month_folder(download_base, report_date)
            landed = False
            for fname, blob in attachments:
                fname = safe_filename(fname)
                dest = dest_folder / fname
                if dest.exists():
                    logger.info("Already on disk: %s", dest)
                    landed = True
                    continue
                with open(dest, "wb") as f:
                    f.write(blob)
                logger.info("Saved %s (%d bytes)", dest, len(blob))
                saved_paths.append(dest)
                landed = True

            # Mark processed only once something is actually on disk. Marking
            # unconditionally is what turned one bad run into a permanently
            # missing trading day.
            if not landed:
                logger.warning(
                    "UID %s (%s) produced no file — leaving unmarked so the next run retries",
                    uid, subject,
                )
                continue

            processed_uids.add(uid)
            if message_id:
                processed_msgids.add(message_id)

            if apply_label:
                try:
                    # Gmail-specific: X-GM-LABELS adds the label without moving the message.
                    typ, _ = M.store(uid_bytes, "+X-GM-LABELS", f'"{apply_label}"')
                    if typ == "OK":
                        logger.info("Applied label %r to UID %s", apply_label, uid)
                    else:
                        logger.warning("Label STORE returned %s for UID %s", typ, uid)
                except Exception as e:
                    logger.warning("Failed to apply label to UID %s: %s", uid, e)

        M.logout()

    state["processed_uids"] = sorted(processed_uids)
    state["processed_message_ids"] = sorted(processed_msgids)
    state["last_fetched_utc"] = datetime.now(timezone.utc).isoformat()
    save_state(state)

    return saved_paths


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    config = load_config()
    state = load_state()
    saved = fetch_new_reports(config, state)
    print()
    print("=" * 60)
    if saved:
        print(f"Downloaded {len(saved)} new report(s):")
        for p in saved:
            print(f"  ✓ {p.name}  →  {p.parent}")
    else:
        print("No new reports. Everything up to date.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
