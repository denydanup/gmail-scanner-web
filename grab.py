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
Password tersimpan di dalam file sesuai permintaan Anda.
JANGAN UPLOAD ke GitHub atau share ke pihak lain.
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
    except:
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
        except:
            pass

    # remove duplicates
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
    except:
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

    except:
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
    except:
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

    print(f"Mailboxes to scan ({len(mailboxes)}): {mailboxes}")

    rows = []
    total_matches = 0
    start = time.time()

    for idx_box, mailbox in enumerate(mailboxes, start=1):
        print(f"\n[{idx_box}/{len(mailboxes)}] Selecting mailbox: {mailbox!r}...", end=" ")

        if not safe_select(mail, mailbox):
            print("SKIP")
            continue
        print("OK")

        status, data = mail.search(None, "ALL")
        if status != "OK":
            print(f"  Cannot read mailbox {mailbox}")
            continue

        msg_ids = data[0].split()
        print(f"  Total messages: {len(msg_ids)}")

        matched = []

        # SUBJECT FILTER
        for mid in msg_ids:
            subj = fetch_header_subject(mail, mid)
            if subj and SUBJECT_PATTERN.search(subj):
                matched.append(mid)

        print(f"  Matched subject count: {len(matched)}")

        # FETCH DATA
        for i, mid in enumerate(matched, start=1):

            raw = safe_fetch_rfc822(mail, mid)
            if not raw:
                continue

            msg = email.message_from_bytes(raw)
            subject = decode_header_value(msg.get("Subject"))

            # domain extract
            m = re.search(r'([A-Za-z0-9\-\._]+\.my\.id)', subject, flags=re.IGNORECASE)
            domain = m.group(1).lower() if m else ""

            user_name, user_email, to_s, cc_s, bcc_s = extract_recipients(msg)

            rows.append({
                "mailbox": mailbox,
                "msg_id": mid.decode(),
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
            if i % PRINT_EVERY == 0:
                elapsed = int(time.time() - start)
                print(f"    Progress: {i}/{len(matched)} in {mailbox} | total {total_matches} | {elapsed}s elapsed")

    # ===============================
    # WRITE EXCEL
    # ===============================
    if Workbook is None:
        print("\nERROR: openpyxl tidak terinstall. Install dengan: pip install openpyxl")
        print("Tidak bisa membuat Excel.")
        sys.exit(1)

    print(f"\nWriting Excel: {out_xlsx}")

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

    print(f"\nDONE! Total matches: {total_matches}")
    print(f"Saved to: {out_xlsx}")


# ===============================
#  MAIN
# ===============================
def main():
    print("Connecting to Gmail IMAP...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, APP_PASSWORD)
    except Exception as e:
        print("LOGIN FAILED:", e)
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
