#!/usr/bin/env python3
"""
Build a copy-paste-ready "Replies" sheet from the CDC-checked workbook.

One row per user:  ticket numbers | complaint | reply (Hebrew, placeholders
filled) | what the customer gave us (colour 1) | what CDC says (colour 2).

The reply text comes from the "Reply templates" sheet (created on first run
from REPLY_TEMPLATES below). Edit the Hebrew there, rerun, and every reply is
regenerated. Placeholders:  {cdc_email}  {phone}  {email}  {ticket}  {name}  {last_login}

Usage:  python3 make_replies.py CHECKED.xlsx OUT.xlsx [--issues "A;B;C"]
"""
import argparse, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DEFAULT_ISSUES = ["Password reset not working", "Can't log in", "Phone / email already in use"]

# (main issue or "*", CDC result) -> (English note from support lead, Hebrew reply)
REPLY_TEMPLATES = [
    ("*", "Phone found",
     "Please login with the correct email and we provide them the email they are signed up with. "
     "Here we explain also we only allow the use of one phone per account.",
     "שלום {name},\nבדקנו את החשבון המשויך למספר הטלפון {phone}. החשבון רשום עם כתובת המייל: {cdc_email}\n"
     "אנא נסו להתחבר עם כתובת מייל זו. שימו לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד."),
    ("*", "Match: phone and email on same account",
     "Give them the email they should login with attached to this phone number and ask them to try login again with this email.",
     "שלום {name},\nמספר הטלפון {phone} והמייל {cdc_email} משויכים לאותו חשבון. אנא נסו להתחבר שוב עם המייל {cdc_email}. "
     "אם ההתחברות עדיין נכשלת, לחצו על 'שכחתי סיסמה' ועדכנו אותנו."),
    ("*", "Phone on multiple accounts",
     "This one we will check manually - its probably a lifestyle user - we need to provide them with the email of the login account so they can login.",
     "[בדיקה ידנית] מספר הטלפון {phone} משויך ליותר מחשבון אחד. יש לבחור את חשבון ההתחברות הנכון (ראו פרטי חשבונות) ולשלוח ללקוח את המייל שלו."),
    ("*", "Phone not in CDC (likely erased)",
     "We can not see the phone attached to an account please create a new account. Or give us the email you are trying to login with to check.",
     "שלום {name},\nלא מצאנו חשבון המשויך למספר הטלפון {phone}. אנא הירשמו מחדש באתר/באפליקציה, "
     "או שלחו לנו את כתובת המייל שאיתה ניסיתם להתחבר כדי שנבדוק."),
    ("*", "Phone and email not in CDC (likely erased)",
     "We can not see the phone attached to an account please create a new account.",
     "שלום {name},\nלא מצאנו חשבון המשויך למספר הטלפון {phone} או למייל {email}. אנא הירשמו מחדש באתר/באפליקציה."),
    ("*", "Mismatch: phone account has other email; no account on user email",
     "Please login with the correct email and we provide them the email they are signed up with.",
     "שלום {name},\nהמייל {email} אינו משויך לחשבון. החשבון המשויך למספר {phone} רשום עם המייל: {cdc_email}\nאנא נסו להתחבר עם מייל זה."),
    ("*", "Mismatch: email on a different account",
     "Please login with the correct email and we provide them the email they are signed up with. Here we explain also we only allow the use of one phone per account.",
     "שלום {name},\nהמייל {email} משויך לחשבון אחר מזה של מספר הטלפון {phone}. החשבון של מספר הטלפון רשום עם המייל: {cdc_email}\n"
     "אנא נסו להתחבר עם מייל זה. שימו לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד."),
    ("*", "LOGGED IN AFTER CONTACT",
     "CDC shows a login on the phone's account after the customer first contacted us: tell them we saw they managed to log in.",
     "שלום {name},\nראינו שהצלחתם להתחבר לחשבון עם המייל {cdc_email} בתאריך {last_login}. "
     "אם עדיין יש בעיה, אנא כתבו לנו מה קורה כשמנסים להתחבר."),
    ("*", "Phone not in CDC; email has account",
     "Email exists but without this phone.",
     "שלום {name},\nהמייל {email} משויך לחשבון, אך מספר הטלפון {phone} אינו מופיע בו. אנא התחברו עם המייל והוסיפו את מספר הטלפון בפרופיל."),
]

CUSTOMER_COLS = ["Customer name", "Phone", "Email", "Other phones", "Other emails"]
CDC_COLS = ["CDC result", "CDC accounts on phone (UID)", "CDC email on phone account",
            "User email matches CDC email", "CDC accounts on user email (UID)",
            "Phone on email account", "CDC accounts on other emails", "CDC account details (per UID)"]
import re, datetime as dt

def last_login_after_contact(details, first_contact):
    """Latest 'last login YYYY-MM-DD' in the CDC details that is on/after the
    customer's first contact date; None otherwise."""
    if not details or not first_contact: return None
    fc = first_contact.date() if isinstance(first_contact, dt.datetime) else first_contact
    dates = [dt.date.fromisoformat(d) for d in re.findall(r"last login (\d{4}-\d{2}-\d{2})", str(details))]
    later = [d for d in dates if d >= fc]
    return max(later).isoformat() if later else None

HEAD_COLS = ["User ID", "Priority", "Ticket numbers", "Latest ticket #", "Main issue", "Issue (most specific)",
             "Customer's words (all login tickets)", "First contact"]

FILL_REPLY = PatternFill("solid", fgColor="FFF2CC")     # yellow
FILL_CUSTOMER = PatternFill("solid", fgColor="DDEBF7")  # blue
FILL_CDC = PatternFill("solid", fgColor="E2EFDA")       # green


def ensure_templates(wb):
    if "Reply templates" in wb.sheetnames:
        return
    ts = wb.create_sheet("Reply templates")
    ts.append(["Main issue (* = any)", "CDC result", "English note", "Hebrew reply (edit me; placeholders {name} {phone} {email} {cdc_email} {ticket})"])
    for c in ts[1]: c.font = Font(bold=True)
    for row in REPLY_TEMPLATES: ts.append(list(row))
    ts.column_dimensions["B"].width = 55; ts.column_dimensions["C"].width = 60; ts.column_dimensions["D"].width = 90
    for r in ts.iter_rows(min_row=2):
        for c in r: c.alignment = Alignment(wrap_text=True, vertical="top")


def load_templates(wb):
    t = {}
    for issue, result, _en, he in wb["Reply templates"].iter_rows(min_row=2, values_only=True):
        if result: t[(issue or "*", result)] = he or ""
    return t


def first_cdc_email(v):
    v = str(v or "")
    for part in v.replace(";", ",").split(","):
        p = part.strip()
        if p and "@" in p: return p
    return "(לא נמצא מייל בחשבון)"


def build(inp, out, issues):
    wb = openpyxl.load_workbook(inp)
    ensure_templates(wb)
    templates = load_templates(wb)
    us = wb["Users"]; h = [c.value for c in us[1]]; col = {n: i for i, n in enumerate(h)}
    missing = [c for c in HEAD_COLS + CUSTOMER_COLS + CDC_COLS if c not in col]
    if missing: sys.exit(f"Users sheet missing columns: {missing}")
    if "Replies" in wb.sheetnames: del wb["Replies"]
    rs = wb.create_sheet("Replies", 0)
    header = HEAD_COLS + ["Reply (copy-paste)", "Logged in after contact"] + ["Customer gave: " + c for c in CUSTOMER_COLS] + CDC_COLS
    rs.append(header)
    for i, c in enumerate(rs[1], 1):
        c.font = Font(bold=True)
        if i in (len(HEAD_COLS) + 1, len(HEAD_COLS) + 2): c.fill = FILL_REPLY
        elif len(HEAD_COLS) + 2 < i <= len(HEAD_COLS) + 2 + len(CUSTOMER_COLS): c.fill = FILL_CUSTOMER
        elif i > len(HEAD_COLS) + 2 + len(CUSTOMER_COLS): c.fill = FILL_CDC
    n = 0
    for row in us.iter_rows(min_row=2, values_only=True):
        g = lambda name: row[col[name]]
        if g("Main issue") not in issues or not g("User ID"): continue
        result = g("CDC result") or ""
        ll = last_login_after_contact(g("CDC account details (per UID)"), g("First contact")) if g("CDC accounts on phone (UID)") not in (None, "none") else None
        key = "LOGGED IN AFTER CONTACT" if ll and result in ("Phone found", "Match: phone and email on same account") else result
        tpl = templates.get((g("Main issue"), key)) or templates.get(("*", key)) or f"[אין תבנית לתוצאה: {result}]"
        reply = tpl.format(name=g("Customer name") or "", phone=g("Phone") or "", email=g("Email") or "",
                           cdc_email=first_cdc_email(g("CDC email on phone account")), ticket=g("Latest ticket #") or "",
                           last_login=ll or "")
        vals = [g(c) for c in HEAD_COLS] + [reply, "Yes" if ll else "No"] + [g(c) for c in CUSTOMER_COLS] + [g(c) for c in CDC_COLS]
        rs.append(vals); n += 1
        r = rs.max_row
        for i in range(1, len(header) + 1):
            c = rs.cell(r, i); c.alignment = Alignment(wrap_text=True, vertical="top")
            if i in (len(HEAD_COLS) + 1, len(HEAD_COLS) + 2): c.fill = FILL_REPLY
            elif len(HEAD_COLS) + 2 < i <= len(HEAD_COLS) + 2 + len(CUSTOMER_COLS): c.fill = FILL_CUSTOMER
            elif i > len(HEAD_COLS) + 2 + len(CUSTOMER_COLS): c.fill = FILL_CDC
    widths = {"Ticket numbers": 22, "Customer's words (all login tickets)": 50, "Reply (copy-paste)": 60,
              "CDC result": 30, "CDC account details (per UID)": 60, "Main issue": 24, "Issue (most specific)": 30}
    for i, name in enumerate(header, 1):
        rs.column_dimensions[get_column_letter(i)].width = widths.get(name.replace("Customer gave: ", ""), 18)
    rs.freeze_panes = "D2"; rs.auto_filter.ref = rs.dimensions
    wb.save(out); print(f"{n} replies -> {out}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input"); ap.add_argument("output")
    ap.add_argument("--issues", default=";".join(DEFAULT_ISSUES), help="main issues to include, separated by ';'")
    a = ap.parse_args()
    build(a.input, a.output, [s.strip() for s in a.issues.split(";") if s.strip()])
