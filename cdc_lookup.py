#!/usr/bin/env python3
"""
CDC (SAP Customer Data Cloud / Gigya) login-issue lookup.

Reads the "Users" sheet of the login-issues workbook, and for every user:
  1. Looks up CDC accounts by the user's phone  -> which email is connected to it.
  2. If the user also gave an email: notes whether it matches the CDC email
     of the phone account, and checks whether a separate CDC account exists
     on that email (and which phone that account has).
  3. Writes the results into NEW columns of a COPY of the workbook, next to
     the existing User ID / Ticket numbers columns, so every finding stays
     attached to the tickets that raised it.

The original workbook is never modified. This script only READS from CDC
(accounts.search) - it never changes accounts.

Configuration (environment variables - never put credentials in this file):
  CDC_API_KEY      site API key                                   (required)
  CDC_USER_KEY     application/user key                           (required)
  CDC_SECRET       application/user secret                        (required)
  CDC_DATA_CENTER  e.g. eu1, us1, au1                             (default eu1)
  CDC_PHONE_FIELDS comma-separated account fields holding phones
                   (default: phoneNumber,profile.phones.number)
  CDC_EMAIL_FIELDS comma-separated account fields holding emails
                   (default: profile.email,loginIDs.emails)

Usage:
  python3 cdc_lookup.py INPUT.xlsx OUTPUT.xlsx [--limit 5] [--delay 0.2]
"""
import argparse
import re
import datetime as dt
import json
import os
import sys
import time
import urllib.parse
import urllib.request

import openpyxl

NEW_COLUMNS = [
    "CDC lookup date",
    "CDC result",
    "CDC accounts on phone (UID)",
    "CDC email on phone account",
    "User email matches CDC email",
    "CDC accounts on user email (UID)",
    "Phone on email account",
    "CDC accounts on other emails",
    "CDC lookup error",
]


# ---------------------------------------------------------------- CDC client
class CDCClient:
    def __init__(self, api_key, user_key, secret, data_center="eu1",
                 phone_fields=None, email_fields=None, delay=0.2):
        self.url = f"https://accounts.{data_center}.gigya.com/accounts.search"
        self.auth = {"apiKey": api_key, "userKey": user_key, "secret": secret}
        self.phone_fields = phone_fields or ["phoneNumber", "profile.phones.number"]
        self.email_fields = email_fields or ["profile.email", "loginIDs.emails"]
        self.delay = delay

    @classmethod
    def from_env(cls, delay=0.2):
        missing = [k for k in ("CDC_API_KEY", "CDC_USER_KEY", "CDC_SECRET")
                   if not os.environ.get(k)]
        if missing:
            sys.exit("Missing environment variables: " + ", ".join(missing))
        split = lambda v: [s.strip() for s in v.split(",") if s.strip()] if v else None
        return cls(os.environ["CDC_API_KEY"], os.environ["CDC_USER_KEY"],
                   os.environ["CDC_SECRET"], os.environ.get("CDC_DATA_CENTER", "eu1"),
                   split(os.environ.get("CDC_PHONE_FIELDS")),
                   split(os.environ.get("CDC_EMAIL_FIELDS")), delay)

    def _search(self, fields, value):
        value = value.replace("\\", "\\\\").replace('"', '\\"')
        where = " OR ".join(f'{f} = "{value}"' for f in fields)
        query = f"SELECT UID, profile.email, loginIDs.emails, phoneNumber, profile.phones FROM accounts WHERE {where} LIMIT 20"
        data = urllib.parse.urlencode({**self.auth, "query": query}).encode()
        for attempt in range(4):
            try:
                with urllib.request.urlopen(self.url, data=data, timeout=30) as r:
                    res = json.load(r)
                break
            except OSError:
                if attempt == 3:
                    raise
                time.sleep(2 ** (attempt + 1))
        time.sleep(self.delay)
        if res.get("errorCode", 0) != 0:
            raise RuntimeError(f"CDC error {res.get('errorCode')}: {res.get('errorMessage')} {res.get('errorDetails', '')}")
        return res.get("results", [])

    def by_phone(self, phone):
        """Search every common spelling of the number (+972..., 972..., 05...)."""
        seen, results = set(), []
        for variant in phone_variants(phone):
            for acc in self._search(self.phone_fields, variant):
                if acc["UID"] not in seen:
                    seen.add(acc["UID"])
                    results.append(acc)
        return results

    def by_email(self, email):
        return self._search(self.email_fields, email.lower())


# ------------------------------------------------------------- account utils
def account_emails(acc):
    out = []
    pe = (acc.get("profile") or {}).get("email")
    if pe:
        out.append(pe)
    out += (acc.get("loginIDs") or {}).get("emails") or []
    seen, res = set(), []
    for e in out:
        if e and e.lower() not in seen:
            seen.add(e.lower())
            res.append(e)
    return res


def account_phones(acc):
    out = []
    if acc.get("phoneNumber"):
        out.append(acc["phoneNumber"])
    phones = (acc.get("profile") or {}).get("phones") or []
    if isinstance(phones, dict):
        phones = [phones]
    out += [p.get("number") for p in phones if isinstance(p, dict) and p.get("number")]
    return list(dict.fromkeys(out))


def phone_variants(phone):
    """'+972504668080' -> ['+972504668080', '972504668080', '0504668080'].
    Local numbers ('0504668080') are also expanded to the +972 forms."""
    digits = re.sub(r"\D", "", str(phone or ""))
    if not digits:
        return []
    if digits.startswith("972"):
        national = digits[3:]
    elif digits.startswith("0"):
        national = digits[1:]
    else:  # foreign number: only try with / without the plus
        return [f"+{digits}", digits]
    return [f"+972{national}", f"972{national}", f"0{national}"]


def norm_email(e):
    return (e or "").strip().lower()


def split_list(v):
    return [s.strip() for s in str(v or "").replace(";", ",").split(",") if s.strip()]


# --------------------------------------------------------------- core logic
def check_user(client, phone, email, other_emails, other_phones=None):
    """Return dict of NEW_COLUMNS values for one user."""
    r = {c: None for c in NEW_COLUMNS}
    r["CDC lookup date"] = dt.date.today().isoformat()
    email = norm_email(email)
    phone_accs, email_accs = None, None
    try:
        if phone:
            phone_accs = client.by_phone(str(phone))
            for op in split_list(other_phones):  # a second number the customer gave
                for acc in client.by_phone(op):
                    if acc["UID"] not in {a["UID"] for a in phone_accs}:
                        phone_accs.append(acc)
            r["CDC accounts on phone (UID)"] = ", ".join(a["UID"] for a in phone_accs) or "none"
            r["CDC email on phone account"] = "; ".join(
                ", ".join(account_emails(a)) or "(no email)" for a in phone_accs) or None
        if email:
            email_accs = client.by_email(email)
            r["CDC accounts on user email (UID)"] = ", ".join(a["UID"] for a in email_accs) or "none"
            r["Phone on email account"] = "; ".join(
                ", ".join(account_phones(a)) or "(no phone)" for a in email_accs) or None
        others = []
        for oe in split_list(other_emails):
            if norm_email(oe) == email:
                continue
            accs = client.by_email(oe)
            others.append(f"{oe}: {', '.join(a['UID'] for a in accs) or 'none'}")
        r["CDC accounts on other emails"] = " | ".join(others) or None
    except Exception as ex:  # keep going; record the error on the row
        r["CDC lookup error"] = str(ex)[:500]
        r["CDC result"] = "Lookup error"
        return r

    if phone_accs is not None and email:
        cdc_emails = {norm_email(e) for a in phone_accs for e in account_emails(a)}
        r["User email matches CDC email"] = "Yes" if email in cdc_emails else ("No" if phone_accs else "n/a")

    r["CDC result"] = classify(phone, email, phone_accs, email_accs)
    return r


def classify(phone, email, phone_accs, email_accs):
    phone_uids = {a["UID"] for a in phone_accs or []}
    email_uids = {a["UID"] for a in email_accs or []}
    if phone and email:
        if not phone_accs and not email_accs:
            return "Phone and email not in CDC (likely erased)"
        if not phone_accs:
            return "Phone not in CDC; email has account"
        if len(phone_accs) > 1:
            return "Phone on multiple accounts"
        if phone_uids & email_uids:
            return "Match: phone and email on same account"
        if email_accs:
            return "Mismatch: email on a different account"
        return "Mismatch: phone account has other email; no account on user email"
    if phone:
        if not phone_accs:
            return "Phone not in CDC (likely erased)"
        return "Phone on multiple accounts" if len(phone_accs) > 1 else "Phone found"
    if email:
        if not email_accs:
            return "Email not in CDC (likely erased)"
        return "Email on multiple accounts" if len(email_accs) > 1 else "Email found"
    return "No phone or email to check"


# -------------------------------------------------------------- workbook io
def run(inp, out, client, limit=None, save_every=25):
    wb = openpyxl.load_workbook(inp)
    results = {}
    ws = wb["Users"]
    header = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(header)}
    need = ["User ID", "Phone (intl, for CDC)", "Email", "Ticket numbers"]
    missing = [h for h in need if h not in col]
    if missing:
        sys.exit(f"Users sheet missing columns: {missing}")

    start = ws.max_column + 1
    for i, name in enumerate(NEW_COLUMNS):
        ws.cell(row=1, column=start + i, value=name).font = openpyxl.styles.Font(bold=True)

    done = 0
    for row in range(2, ws.max_row + 1):
        uid = ws.cell(row=row, column=col["User ID"] + 1).value
        if not uid:
            continue
        if limit and done >= limit:
            break
        phone = ws.cell(row=row, column=col["Phone (intl, for CDC)"] + 1).value
        email = ws.cell(row=row, column=col["Email"] + 1).value
        others = ws.cell(row=row, column=col["Other emails"] + 1).value if "Other emails" in col else None
        other_phones = ws.cell(row=row, column=col["Other phones"] + 1).value if "Other phones" in col else None
        res = check_user(client, phone, email, others, other_phones)
        for i, name in enumerate(NEW_COLUMNS):
            ws.cell(row=row, column=start + i, value=res[name])
        if "CDC finding" in col:
            c = ws.cell(row=row, column=col["CDC finding"] + 1)
            if not c.value:
                c.value = res["CDC result"]
        results[uid] = res
        done += 1
        print(f"{uid}: {res['CDC result']}", flush=True)
        if done % save_every == 0:  # checkpoint so a network drop loses little
            wb.save(out)

    annotate_tickets(wb, results)
    wb.save(out)
    print(f"Checked {done} users -> {out}")


TICKET_COLUMNS = ["CDC result", "CDC email on phone account",
                  "CDC accounts on phone (UID)", "CDC accounts on user email (UID)",
                  "CDC lookup date"]


def annotate_tickets(wb, results):
    """Copy each user's finding onto every one of their rows in the Tickets
    sheet, so a ticket number alone is enough to find the CDC finding later."""
    if "Tickets" not in wb.sheetnames:
        return
    ts = wb["Tickets"]
    header = [c.value for c in ts[1]]
    if "User ID" not in header:
        return
    uid_col = header.index("User ID") + 1
    start = ts.max_column + 1
    for i, name in enumerate(TICKET_COLUMNS):
        ts.cell(row=1, column=start + i, value=name).font = openpyxl.styles.Font(bold=True)
    for row in range(2, ts.max_row + 1):
        res = results.get(ts.cell(row=row, column=uid_col).value)
        if res:
            for i, name in enumerate(TICKET_COLUMNS):
                ts.cell(row=row, column=start + i, value=res[name])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--limit", type=int, help="only check the first N users (test run)")
    ap.add_argument("--delay", type=float, default=0.2, help="seconds between CDC calls")
    a = ap.parse_args()
    if os.path.abspath(a.input) == os.path.abspath(a.output):
        sys.exit("Output must be a different file - the original is never overwritten.")
    run(a.input, a.output, CDCClient.from_env(a.delay), a.limit)


if __name__ == "__main__":
    main()
