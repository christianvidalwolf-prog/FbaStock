#!/usr/bin/env python3
"""
fetch_minerales_email.py
Downloads the daily Minerales (Vives Cortada) stock CSV from Gmail.

Looks for the newest email from info@vivescortadaimport.com with subject
"[Vivescortadaimport] Stocks Vidal", saves its CSV attachment as
minerales_stock_YYYYMMDD.csv in WORK_DIR and refreshes minerales_feed.xml.
auto_update_stock_fast.py then picks it up as the latest manual Minerales file.

Auth: IMAP with a Gmail app password stored in the macOS Keychain:
  security add-generic-password -s stock-gmail-imap -a <gmail address> -w
(or GMAIL_APP_PASSWORD / GMAIL_USER environment variables).

Usage:
  fetch_minerales_email.py                 # check once, today's email only
  fetch_minerales_email.py --wait-minutes 60   # poll until today's email arrives
  fetch_minerales_email.py --any-date      # accept the newest email from the last days
"""
import argparse
import csv
import email
import imaplib
import io
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

WORK_DIR = "/Users/christianvidalwolf/Stock"
STATE_FILE = os.path.join(WORK_DIR, "logs", "minerales_email_state.json")
KEYCHAIN_SERVICE = "stock-gmail-imap"
DEFAULT_USER = "vidalregals@gmail.com"
SENDER = "info@vivescortadaimport.com"
SUBJECT = "[Vivescortadaimport] Stocks Vidal"
POLL_SECONDS = 300
MIN_ROWS = 1000


def get_credentials():
    user = os.environ.get("GMAIL_USER", DEFAULT_USER)
    password = os.environ.get("GMAIL_APP_PASSWORD")
    if not password:
        result = subprocess.run(
            ["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", user, "-w"],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            password = result.stdout.strip()
    if not password:
        raise RuntimeError(
            f"No Gmail app password found. Store it with: "
            f"security add-generic-password -s {KEYCHAIN_SERVICE} -a {user} -w"
        )
    return user, password.replace(" ", "")


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def all_mail_folder(imap):
    """Gmail's All Mail folder name is localized ([Gmail]/Todos, [Gmail]/All Mail...)."""
    typ, folders = imap.list()
    if typ == "OK":
        for raw in folders or []:
            line = raw.decode("utf-8", errors="ignore")
            if "\\All" in line:
                match = re.search(r'"([^"]+)"\s*$', line) or re.search(r"(\S+)\s*$", line)
                if match:
                    return match.group(1)
    return "INBOX"


def decoded_subject(msg):
    try:
        return str(make_header(decode_header(msg.get("Subject", ""))))
    except Exception:
        return msg.get("Subject", "")


def extract_csv_attachment(msg):
    """Returns (filename, bytes) of the first CSV attachment, or (None, None)."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        if filename:
            try:
                filename = str(make_header(decode_header(filename)))
            except Exception:
                pass
        is_csv = (filename or "").lower().endswith(".csv") or part.get_content_type() in (
            "text/csv",
            "application/csv",
        )
        if is_csv:
            payload = part.get_payload(decode=True)
            if payload:
                return filename or "attachment.csv", payload
    return None, None


def validate_csv(data):
    """Checks the attachment looks like the Minerales stock export. Returns row count."""
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("latin-1")
    reader = csv.reader(io.StringIO(text), delimiter=";")
    header = [h.strip().lower() for h in next(reader, [])]
    if not any("referencia" in h for h in header) or not any("cantidad" in h for h in header):
        raise ValueError(f"Unexpected CSV header: {header}")
    rows = sum(1 for row in reader if row)
    if rows < MIN_ROWS:
        raise ValueError(f"CSV has only {rows} rows (expected >= {MIN_ROWS})")
    return rows


def find_latest_email(imap):
    """Returns (message, internal datetime) of the newest matching email, or (None, None)."""
    imap.select(f'"{all_mail_folder(imap)}"', readonly=True)
    query = f'from:{SENDER} subject:"{SUBJECT}" has:attachment newer_than:3d'
    typ, data = imap.search(None, "X-GM-RAW", f'"{query.replace(chr(34), chr(92) + chr(34))}"')
    if typ != "OK" or not data or not data[0]:
        return None, None
    best_msg, best_date = None, None
    for num in data[0].split():
        typ, fetched = imap.fetch(num, "(RFC822)")
        if typ != "OK" or not fetched or not isinstance(fetched[0], tuple):
            continue
        msg = email.message_from_bytes(fetched[0][1])
        if SUBJECT.lower() not in decoded_subject(msg).lower():
            continue
        if SENDER.lower() not in msg.get("From", "").lower():
            continue
        try:
            msg_date = parsedate_to_datetime(msg.get("Date")).astimezone()
        except Exception:
            continue
        if best_date is None or msg_date > best_date:
            best_msg, best_date = msg, msg_date
    return best_msg, best_date


def import_attachment(msg, msg_date, state):
    filename, data = extract_csv_attachment(msg)
    if not data:
        print("Email found but it has no CSV attachment.")
        return False
    rows = validate_csv(data)

    target = os.path.join(WORK_DIR, f"minerales_stock_{msg_date.strftime('%Y%m%d')}.csv")
    tmp_target = target + ".tmp"
    with open(tmp_target, "wb") as f:
        f.write(data)
    os.replace(tmp_target, target)
    print(f"Saved '{filename}' ({rows} rows, email {msg_date:%Y-%m-%d %H:%M}) -> {target}")

    from update_minerales_feed import convert_manual_csv_to_feed

    if not convert_manual_csv_to_feed(target):
        print("Warning: minerales_feed.xml could not be regenerated from the new file.")
        return False

    state.update(
        {
            "message_id": msg.get("Message-ID", ""),
            "email_date": msg_date.isoformat(),
            "file": target,
            "imported_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    save_state(state)
    return True


def check_once(user, password, state, any_date):
    """Returns True when there is nothing left to wait for (imported or already up to date)."""
    with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
        imap.login(user, password)
        msg, msg_date = find_latest_email(imap)
    if msg is None:
        print("No Minerales stock email found in the last 3 days.")
        return False
    if msg.get("Message-ID", "") == state.get("message_id"):
        print(f"Latest Minerales email ({msg_date:%Y-%m-%d %H:%M}) already imported.")
        return msg_date.date() == datetime.now().date() or any_date
    if not any_date and msg_date.date() != datetime.now().date():
        print(f"Latest Minerales email is from {msg_date:%Y-%m-%d}; waiting for today's.")
        return False
    return import_attachment(msg, msg_date, state)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--wait-minutes", type=int, default=0, help="poll Gmail until today's email arrives")
    parser.add_argument("--any-date", action="store_true", help="accept the newest email even if not from today")
    args = parser.parse_args()

    print(f"--- Minerales email check @ {datetime.now():%Y-%m-%d %H:%M:%S} ---")
    try:
        user, password = get_credentials()
    except RuntimeError as e:
        print(f"Error: {e}")
        return 1

    state = load_state()
    deadline = time.time() + args.wait_minutes * 60
    while True:
        try:
            if check_once(user, password, state, args.any_date):
                return 0
        except (imaplib.IMAP4.error, OSError, ValueError) as e:
            print(f"Error checking Gmail: {e}")
        if time.time() + POLL_SECONDS > deadline:
            print("Giving up for now; the pipeline will use the last imported Minerales file.")
            return 0
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    sys.exit(main())
