"""Failure notifier: writes errors to last_error.log and optionally emails them.

Uses the same Gmail account / app password configured in config.json to send
self-alerts via SMTP when the fetcher or formatter raises.
"""

import logging
import smtplib
import traceback
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent
ERROR_LOG = SCRIPT_DIR / "last_error.log"


def write_error_log(stage: str, exc: BaseException) -> None:
    timestamp = datetime.now().isoformat(timespec="seconds")
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    body = f"[{timestamp}] FAILURE during {stage}\n{type(exc).__name__}: {exc}\n\n{tb}\n"
    with open(ERROR_LOG, "a") as f:
        f.write(body)
        f.write("-" * 70 + "\n")
    logger.error("Wrote failure to %s", ERROR_LOG)


def send_error_email(config: dict, stage: str, exc: BaseException) -> None:
    """Best-effort: email the user with the failure. Swallows its own errors."""
    if not config.get("alert_email_enabled", True):
        return
    user = config["email"]
    password = config["app_password"]
    to_addr = config.get("alert_email_to", user)

    timestamp = datetime.now().isoformat(timespec="seconds")
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))

    msg = EmailMessage()
    msg["Subject"] = f"[trade-log-formatter] FAILED in {stage}: {type(exc).__name__}"
    msg["From"] = user
    msg["To"] = to_addr
    msg.set_content(
        f"trade-log-formatter failed during {stage} at {timestamp}.\n\n"
        f"{type(exc).__name__}: {exc}\n\n"
        f"Traceback:\n{tb}\n"
    )

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as s:
            s.login(user, password)
            s.send_message(msg)
        logger.info("Sent failure alert email to %s", to_addr)
    except Exception as e:
        logger.error("Failed to send alert email: %s", e)


def report_failure(config: dict, stage: str, exc: BaseException) -> None:
    write_error_log(stage, exc)
    try:
        send_error_email(config, stage, exc)
    except Exception:
        pass
