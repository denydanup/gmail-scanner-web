#!/usr/bin/env python3
"""
grab.py - Subject-only mailbox scanner (Gmail/IMAP) — ALL mailboxes → Excel output

Scans ALL mailboxes (Inbox, Spam, Trash, All Mail, labels) and searches SUBJECT only
for:  "domain <something>.my.id sudah aktif"  (case-insensitive)

Extracts:
- domain
- user_name (nama penerima)
- user_email (email penerima)
- to, cc, bcc

Outputs:
- hasil_subject.xlsx (Excel file)

Usage:
    python grab.py

NOTE:
This version prints immediate progress and flushes stdout so logs appear in real-time.
"""

import imaplib
import email
from email.message import Message
from email.header import decode_header
from email.utils import getaddresses, parseaddr
import re
import os
import sys
import time
from typing import List, Optional

# Excel
try:
    from openpyxl import Workbook
except:
    Workbook = None

# ===============================
#  STATIC CONFIG (EDIT DI SINI)
# ===============================
IMAP_SERVER   = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_PORT     = int(os.getenv("IMAP_PORT", "993"))
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
APP_PASSWORD  = os.getenv("APP_PASSWORD")

OUTPUT_FILE   = "hasil_subject.xlsx"
# Subject pattern: "domain xxx.my.id sudah aktif"
SUBJECT_PATTERN = re.compile(
    r'domain\s+[A-Za-z0-9\-\._]+\.my\.id\s+sudah\s+aktif',
    flags=re.IGNORECASE
)

PRINT_EVERY = 20   # progress frequency


# ===============================
#  Helper Functions
# ===============================
def print_flush(*args, **kwargs):
    """Print and flush immediately to stdout (helpful for real-time logs)."""
    kwargs.setdefault("flush", True)
    print(*args, **kwargs)


def decode_header_value(value):
    if not value:
        return ""
    parts = decode_header(value)
    out = []
    for v, enc in parts:
        if isinstance(v, bytes):
            try:
                out.append(v.decode(enc or "utf-8", errors="replace"))
            except:
                out.append(v.decode("utf-8", errors="replace"))
        else:
            out.append(v)
    return "".join(out).strip()


def list_mailboxes(mail: imaplib.IMAP4_SSL) -> List[str]:
    """List all mailboxes (Inbox, Spam, Trash, All Mail, Labels)."""
    try:
        status, data = mail.list()
    except Exception as e:
        print_flush("ERROR list mailboxes:", e)
        return []
    if status != "OK" or not data:
        return []

    boxes = []
    for item in data:
        try:
            s = item.decode() if isinstance(item, bytes) else str(item)
            m = re.search(r'"([^"]+)"\s*$', s)
            if m:
                boxes.append(m.group(1))
            else:
                parts = s.split()
                if parts:
                    boxes.append(parts[-1].strip('"'))
        except Exception:
            pass

    # remove duplicates while preserving order
    seen = set()
    out = []
    for b in boxes:
        if b not in seen:
            seen.add(b)
            out.append(b)
    return out


def safe_select(mail, mailbox: str) -> bool:
    try:
        status, _ = mail.select(mailbox)
        if status == "OK":
            return True
        # fallback quoted
        status, _ = mail.select(f'"{mailbox}"')
        return status == "OK"
    except Exception as e:
        print_flush(f"  select({mailbox}) failed:", e)
        return False


def fetch_header_subject(mail, msg_id: bytes) -> str:
    """Only fetch Subject header — fast."""
    try:
        status, data = mail.fetch(msg_id, '(BODY.PEEK[HEADER.FIELDS (SUBJECT)])')
        if status != "OK" or not data:
            return ""
        raw = None
        for part in data:
            if isinstance(part, tuple) and isinstance(part[1], (bytes, bytearray)):
                raw = part[1]
        if not raw:
            return ""
        try:
            h = raw.decode("utf-8", errors="replace")
        except:
            h = raw.decode("latin-1", errors="replace")

        m = re.search(r"^Subject:\s*(.*)$", h, flags=re.IGNORECASE | re.MULTILINE)
        return decode_header_value(m.group(1)) if m else ""

    except Exception as e:
        # do not crash: return empty subject
        print_flush("    (warning) fetch_header_subject failed for id", getattr(msg_id, "decode", lambda: msg_id)(), ":", e)
        return ""


def safe_fetch_rfc822(mail, msg_id: bytes):
    try:
        status, data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK" or not data:
            return None
        for part in data:
            if isinstance(part, tuple):
                return part[1]
        return None
    except Exception as e:
        print_flush("    (warning) safe_fetch_rfc822 failed for id", getattr(msg_id, "decode", lambda: msg_id)(), ":", e)
        return None


def extract_recipients(msg: Message):
    tos = msg.get_all("to", []) or []
    ccs = msg.get_all("cc", []) or []
    bccs = msg.get_all("bcc", []) or []

    all_to = getaddresses(tos)
    all_cc = getaddresses(ccs)
    all_bcc = getaddresses(bccs)

    # primary receiver
    if all_to:
        primary_name = all_to[0][0] or ""
        primary_email = all_to[0][1] or ""
    else:
        primary_name = ""
        primary_email = ""

    def join(pairs):
        seen = set()
        out = []
        for name, addr in pairs:
            if addr and addr not in seen:
                seen.add(addr)
                out.append(addr)
        return ";".join(out)

    return (
        primary_name,
        primary_email,
        join(all_to),
        join(all_cc),
        join(all_bcc),
    )


# ===============================
#  MAIN SCANNER
# ===============================
def scan_all(mail, out_xlsx):

    mailboxes = list_mailboxes(mail)
    if not mailboxes:
        mailboxes = ["INBOX"]

    print_flush(f"[{time.strftime('%H:%M:%S')}] Mailboxes to scan ({len(mailboxes)}): {mailboxes}")

    rows = []
    total_matches = 0
    start = time.time()

    for idx_box, mailbox in enumerate(mailboxes, start=1):
        print_flush(f"\n[{time.strftime('%H:%M:%S')}] [{idx_box}/{len(mailboxes)}] Selecting mailbox: {mailbox!r}...")

        if not safe_select(mail, mailbox):
            print_flush(f"[{time.strftime('%H:%M:%S')}] SKIP {mailbox}")
            continue
        print_flush(f"[{time.strftime('%H:%M:%S')}] Selected {mailbox}")

        # get message count if possible
        try:
            status_count, count_data = mail.select(mailbox)
            if status_count == "OK" and count_data and isinstance(count_data[0], bytes):
                try:
                    total_msgs_in_box = int(count_data[0].decode().split()[0])
                except:
                    total_msgs_in_box = None
            else:
                total_msgs_in_box = None
        except Exception:
            total_msgs_in_box = None

        status, data = mail.search(None, "ALL")
        if status != "OK" or not data or not data[0]:
            print_flush(f"[{time.strftime('%H:%M:%S')}]  No messages or search failed in {mailbox!r}")
            continue

        msg_ids = data[0].split()
        print_flush(f"[{time.strftime('%H:%M:%S')}]  Total messages: {len(msg_ids)} (reported: {total_msgs_in_box if total_msgs_in_box is not None else 'n/a'})")

        matched = []

        # SUBJECT FILTER (fast pass)
        for mid in msg_ids:
            subj = fetch_header_subject(mail, mid)
            if subj and SUBJECT_PATTERN.search(subj):
                matched.append(mid)

        print_flush(f"[{time.strftime('%H:%M:%S')}]  Matched subject count in {mailbox}: {len(matched)}")

        # FETCH DATA
        for i, mid in enumerate(matched, start=1):

            raw = safe_fetch_rfc822(mail, mid)
            if not raw:
                print_flush(f"[{time.strftime('%H:%M:%S')}]    Skip {getattr(mid, 'decode', lambda: mid)()} (fetch failed)")
                continue

            try:
                msg = email.message_from_bytes(raw)
            except Exception as e:
                print_flush(f"[{time.strftime('%H:%M:%S')}]    Skip {getattr(mid, 'decode', lambda: mid)()} (parse failed: {e})")
                continue

            subject = decode_header_value(msg.get("Subject"))

            # domain extract
            m = re.search(r'([A-Za-z0-9\-\._]+\.my\.id)', subject or "", flags=re.IGNORECASE)
            domain = m.group(1).lower() if m else ""

            user_name, user_email, to_s, cc_s, bcc_s = extract_recipients(msg)

            rows.append({
                "mailbox": mailbox,
                "msg_id": mid.decode() if isinstance(mid, (bytes, bytearray)) else str(mid),
                "subject": subject,
                "domain": domain,
                "user_name": user_name,
                "user_email": user_email,
                "to": to_s,
                "cc": cc_s,
                "bcc": bcc_s
            })

            total_matches += 1

            # PROGRESS
            if i % PRINT_EVERY == 0 or i == len(matched):
                elapsed = int(time.time() - start)
                print_flush(f"[{time.strftime('%H:%M:%S')}]    Progress: {i}/{len(matched)} in {mailbox} | total {total_matches} | {elapsed}s elapsed")

    # ===============================
    # WRITE EXCEL
    # ===============================
    if Workbook is None:
        print_flush("\nERROR: openpyxl tidak terinstall. Install dengan: pip install openpyxl")
        print_flush("Tidak bisa membuat Excel.")
        sys.exit(1)

    print_flush(f"\n[{time.strftime('%H:%M:%S')}] Writing Excel: {out_xlsx}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    headers = ["mailbox", "msg_id", "subject", "domain", "user_name", "user_email", "to", "cc", "bcc"]
    ws.append(headers)

    for r in rows:
        ws.append([
            r["mailbox"], r["msg_id"], r["subject"], r["domain"],
            r["user_name"], r["user_email"], r["to"], r["cc"], r["bcc"]
        ])

    wb.save(out_xlsx)

    print_flush(f"\n[{time.strftime('%H:%M:%S')}] DONE! Total matches: {total_matches}")
    print_flush(f"[{time.strftime('%H:%M:%S')}] Saved to: {out_xlsx}")


# ===============================
#  MAIN
# ===============================
def main():
    print_flush(f"[{time.strftime('%H:%M:%S')}] Script started.")
    if not EMAIL_ADDRESS or not APP_PASSWORD:
        print_flush("ERROR: EMAIL_ADDRESS or APP_PASSWORD not set in environment variables.")
        print_flush("Set IMAP_SERVER, IMAP_PORT, EMAIL_ADDRESS, APP_PASSWORD as environment variables.")
        sys.exit(1)

    print_flush(f"[{time.strftime('%H:%M:%S')}] Connecting to Gmail IMAP ({IMAP_SERVER}:{IMAP_PORT})...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        print_flush(f"[{time.strftime('%H:%M:%S')}] Attempting login for {EMAIL_ADDRESS} ...")
        mail.login(EMAIL_ADDRESS, APP_PASSWORD)
        print_flush(f"[{time.strftime('%H:%M:%S')}] LOGIN SUCCESS")
    except Exception as e:
        print_flush(f"[{time.strftime('%H:%M:%S')}] LOGIN FAILED: {e}")
        sys.exit(1)

    try:
        scan_all(mail, OUTPUT_FILE)
    finally:
        try:
            mail.logout()
        except:
            pass


if __name__ == "__main__":
    main()
