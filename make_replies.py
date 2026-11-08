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
import argparse, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from hebrew_names import hebrew_name_and_gender, email_matches_name, mask_email
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DEFAULT_ISSUES = ["Password reset not working", "Can't log in", "Phone / email already in use"]

# (main issue or "*", CDC result, English note, Hebrew female, Hebrew male, Hebrew neutral)
# Neutral is used when the gender could not be guessed from the first name.
# {greet} = "שלום <first name>," or "שלום," when the name is unknown.
REPLY_TEMPLATES = [
    ("*", "Phone found",
     "Please login with the correct email and we provide them the email they are signed up with. "
     "Here we explain also we only allow the use of one phone per account.",
     "{greet}\nבדקנו את החשבון המשויך למספר הטלפון {phone}. החשבון רשום עם כתובת המייל: {cdc_email}\nאנא נסי להתחבר עם כתובת מייל זו. שימי לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד.",
     "{greet}\nבדקנו את החשבון המשויך למספר הטלפון {phone}. החשבון רשום עם כתובת המייל: {cdc_email}\nאנא נסה להתחבר עם כתובת מייל זו. שים לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד.",
     "{greet}\nבדקנו את החשבון המשויך למספר הטלפון {phone}. החשבון רשום עם כתובת המייל: {cdc_email}\nיש להתחבר עם כתובת מייל זו. לתשומת לבך: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד."),
    ("*", "Match: phone and email on same account",
     "Give them the email they should login with attached to this phone number and ask them to try login again with this email.",
     "{greet}\nמספר הטלפון {phone} והמייל {cdc_email} משויכים לאותו חשבון. אנא נסי להתחבר שוב עם המייל {cdc_email}. אם ההתחברות עדיין נכשלת, לחצי על 'שכחתי סיסמה' ועדכני אותנו.",
     "{greet}\nמספר הטלפון {phone} והמייל {cdc_email} משויכים לאותו חשבון. אנא נסה להתחבר שוב עם המייל {cdc_email}. אם ההתחברות עדיין נכשלת, לחץ על 'שכחתי סיסמה' ועדכן אותנו.",
     "{greet}\nמספר הטלפון {phone} והמייל {cdc_email} משויכים לאותו חשבון. יש להתחבר שוב עם המייל {cdc_email}. אם ההתחברות עדיין נכשלת, יש ללחוץ על 'שכחתי סיסמה' ולעדכן אותנו."),
    ("*", "Phone on multiple accounts",
     "Phone sits on more than one account. The login account is the one with the old-style UID: tell them to log in with its email.",
     "{greet}\nמספר הטלפון {phone} משויך ליותר מחשבון אחד. חשבון ההתחברות שלך רשום עם כתובת המייל: {cdc_email}\nאנא נסי להתחבר עם כתובת מייל זו.",
     "{greet}\nמספר הטלפון {phone} משויך ליותר מחשבון אחד. חשבון ההתחברות שלך רשום עם כתובת המייל: {cdc_email}\nאנא נסה להתחבר עם כתובת מייל זו.",
     "{greet}\nמספר הטלפון {phone} משויך ליותר מחשבון אחד. חשבון ההתחברות רשום עם כתובת המייל: {cdc_email}\nיש להתחבר עם כתובת מייל זו."),
    ("*", "LOGGED IN AFTER CONTACT",
     "CDC shows a login on the account since the go-live cutoff (default 2026-10-07): tell them we saw they managed to log in.",
     "{greet}\nראינו שהצלחת להתחבר לחשבון עם המייל {cdc_email} בתאריך {last_login}. אם עדיין יש בעיה, אנא כתבי לנו מה קורה כשאת מנסה להתחבר.",
     "{greet}\nראינו שהצלחת להתחבר לחשבון עם המייל {cdc_email} בתאריך {last_login}. אם עדיין יש בעיה, אנא כתוב לנו מה קורה כשאתה מנסה להתחבר.",
     "{greet}\nראינו שההתחברות לחשבון עם המייל {cdc_email} הצליחה בתאריך {last_login}. אם עדיין יש בעיה, נשמח לדעת מה קורה בניסיון ההתחברות."),
    ("*", "LOGGED IN AFTER CONTACT (mismatch)",
     "Customer wrote with one email but their phone's account has another email, and that account was logged into after they contacted us.",
     "{greet}\nראינו שהצלחת להתחבר לחשבון עם המייל {cdc_email} בתאריך {last_login}. שימי לב: המייל {email} שציינת אינו המייל של חשבון זה, ולכן יש להתחבר עם {cdc_email}. אם עדיין יש בעיה, אנא כתבי לנו מה קורה כשאת מנסה להתחבר.",
     "{greet}\nראינו שהצלחת להתחבר לחשבון עם המייל {cdc_email} בתאריך {last_login}. שים לב: המייל {email} שציינת אינו המייל של חשבון זה, ולכן יש להתחבר עם {cdc_email}. אם עדיין יש בעיה, אנא כתוב לנו מה קורה כשאתה מנסה להתחבר.",
     "{greet}\nראינו שההתחברות לחשבון עם המייל {cdc_email} הצליחה בתאריך {last_login}. לתשומת לבך: המייל {email} שצוין אינו המייל של חשבון זה, ולכן יש להתחבר עם {cdc_email}. אם עדיין יש בעיה, נשמח לדעת מה קורה בניסיון ההתחברות."),
    ("*", "Phone not in CDC (likely erased)",
     "We can not see the phone attached to an account please create a new account. Or give us the email you are trying to login with to check.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone}. אנא הירשמי מחדש באתר/באפליקציה, או שלחי לנו את כתובת המייל שאיתה ניסית להתחבר כדי שנבדוק.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone}. אנא הירשם מחדש באתר/באפליקציה, או שלח לנו את כתובת המייל שאיתה ניסית להתחבר כדי שנבדוק.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone}. יש להירשם מחדש באתר/באפליקציה, או לשלוח לנו את כתובת המייל ששימשה לניסיון ההתחברות כדי שנבדוק."),
    ("*", "Phone and email not in CDC (likely erased)",
     "We can not see the phone attached to an account please create a new account.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone} או למייל {email}. אנא הירשמי מחדש באתר/באפליקציה.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone} או למייל {email}. אנא הירשם מחדש באתר/באפליקציה.",
     "{greet}\nלא מצאנו חשבון המשויך למספר הטלפון {phone} או למייל {email}. יש להירשם מחדש באתר/באפליקציה."),
    ("*", "Mismatch: phone account has other email; no account on user email",
     "Please login with the correct email and we provide them the email they are signed up with.",
     "{greet}\nהמייל {email} אינו משויך לחשבון. החשבון המשויך למספר {phone} רשום עם המייל: {cdc_email}\nאנא נסי להתחבר עם מייל זה.",
     "{greet}\nהמייל {email} אינו משויך לחשבון. החשבון המשויך למספר {phone} רשום עם המייל: {cdc_email}\nאנא נסה להתחבר עם מייל זה.",
     "{greet}\nהמייל {email} אינו משויך לחשבון. החשבון המשויך למספר {phone} רשום עם המייל: {cdc_email}\nיש להתחבר עם מייל זה."),
    ("*", "Mismatch: email on a different account",
     "Please login with the correct email and we provide them the email they are signed up with. Here we explain also we only allow the use of one phone per account.",
     "{greet}\nהמייל {email} משויך לחשבון אחר מזה של מספר הטלפון {phone}. החשבון של מספר הטלפון רשום עם המייל: {cdc_email}\nאנא נסי להתחבר עם מייל זה. שימי לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד.",
     "{greet}\nהמייל {email} משויך לחשבון אחר מזה של מספר הטלפון {phone}. החשבון של מספר הטלפון רשום עם המייל: {cdc_email}\nאנא נסה להתחבר עם מייל זה. שים לב: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד.",
     "{greet}\nהמייל {email} משויך לחשבון אחר מזה של מספר הטלפון {phone}. החשבון של מספר הטלפון רשום עם המייל: {cdc_email}\nיש להתחבר עם מייל זה. לתשומת לבך: ניתן לשייך מספר טלפון אחד לחשבון אחד בלבד."),
    ("*", "Email found",
     "Customer gave only an email and it exists in CDC: tell them to log in with it (and which phone is on it).",
     "{greet}\nמצאנו את החשבון הרשום עם המייל {email}. אנא נסי להתחבר עם מייל זה. אם ההתחברות נכשלת, לחצי על 'שכחתי סיסמה' ועדכני אותנו.",
     "{greet}\nמצאנו את החשבון הרשום עם המייל {email}. אנא נסה להתחבר עם מייל זה. אם ההתחברות נכשלת, לחץ על 'שכחתי סיסמה' ועדכן אותנו.",
     "{greet}\nמצאנו את החשבון הרשום עם המייל {email}. יש להתחבר עם מייל זה. אם ההתחברות נכשלת, יש ללחוץ על 'שכחתי סיסמה' ולעדכן אותנו."),
    ("*", "Email not in CDC (likely erased)",
     "Customer gave only an email and it is not in CDC: register again, or send us the phone number to check.",
     "{greet}\nלא מצאנו חשבון הרשום עם המייל {email}. אנא הירשמי מחדש באתר/באפליקציה, או שלחי לנו את מספר הטלפון שלך כדי שנבדוק.",
     "{greet}\nלא מצאנו חשבון הרשום עם המייל {email}. אנא הירשם מחדש באתר/באפליקציה, או שלח לנו את מספר הטלפון שלך כדי שנבדוק.",
     "{greet}\nלא מצאנו חשבון הרשום עם המייל {email}. יש להירשם מחדש באתר/באפליקציה, או לשלוח לנו את מספר הטלפון כדי שנבדוק."),
    ("*", "Email on multiple accounts",
     "Email sits on more than one account - check manually which is the login one.",
     "[בדיקה ידנית] המייל {email} משויך ליותר מחשבון אחד. יש לבחור את חשבון ההתחברות (ראו פרטי חשבונות) ולעדכן את הלקוחה.",
     "[בדיקה ידנית] המייל {email} משויך ליותר מחשבון אחד. יש לבחור את חשבון ההתחברות (ראו פרטי חשבונות) ולעדכן את הלקוח.",
     "[בדיקה ידנית] המייל {email} משויך ליותר מחשבון אחד. יש לבחור את חשבון ההתחברות (ראו פרטי חשבונות) ולעדכן את הלקוח/ה."),
    ("*", "No phone or email to check",
     "Nothing to look up: ask the customer for the phone number and email they registered with.",
     "{greet}\nכדי שנוכל לבדוק את החשבון, אנא שלחי לנו את מספר הטלפון ואת כתובת המייל שאיתם נרשמת.",
     "{greet}\nכדי שנוכל לבדוק את החשבון, אנא שלח לנו את מספר הטלפון ואת כתובת המייל שאיתם נרשמת.",
     "{greet}\nכדי שנוכל לבדוק את החשבון, יש לשלוח לנו את מספר הטלפון ואת כתובת המייל ששימשו להרשמה."),
    ("*", "Phone not in CDC; email has account",
     "Email exists but without this phone.",
     "{greet}\nהמייל {email} משויך לחשבון, אך מספר הטלפון {phone} אינו מופיע בו. אנא התחברי עם המייל והוסיפי את מספר הטלפון בפרופיל.",
     "{greet}\nהמייל {email} משויך לחשבון, אך מספר הטלפון {phone} אינו מופיע בו. אנא התחבר עם המייל והוסף את מספר הטלפון בפרופיל.",
     "{greet}\nהמייל {email} משויך לחשבון, אך מספר הטלפון {phone} אינו מופיע בו. יש להתחבר עם המייל ולהוסיף את מספר הטלפון בפרופיל."),
]

CUSTOMER_COLS = ["Customer name", "Phone", "Email", "Other phones", "Other emails"]
CDC_COLS = ["CDC result", "CDC accounts on phone (UID)", "CDC email on phone account",
            "User email matches CDC email", "CDC accounts on user email (UID)",
            "Phone on email account", "CDC accounts on other emails", "CDC account details (per UID)"]
import re, datetime as dt

LOGIN_CUTOFF = dt.date(2026, 10, 7)   # the CDC go-live: a login on/after this date counts as "managed to log in"


def last_login_after_contact(details, cutoff=None):
    """Latest 'last login YYYY-MM-DD' in the CDC details that is on/after the
    cutoff date (default LOGIN_CUTOFF); None otherwise."""
    if not details: return None
    fc = cutoff or LOGIN_CUTOFF
    dates = [dt.date.fromisoformat(d) for d in re.findall(r"last login (\d{4}-\d{2}-\d{2})", str(details))]
    later = [d for d in dates if d >= fc]
    return max(later).isoformat() if later else None

HEAD_COLS = ["User ID", "Priority", "Ticket numbers", "Latest ticket #", "Main issue", "Issue (most specific)",
             "Customer's words (all login tickets)", "First contact"]

FILL_REPLY = PatternFill("solid", fgColor="FFF2CC")     # yellow
FILL_CUSTOMER = PatternFill("solid", fgColor="DDEBF7")  # blue
FILL_CDC = PatternFill("solid", fgColor="E2EFDA")       # light green
FILL_LOGGED_IN = PatternFill("solid", fgColor="92D050")  # green: customer logged in after contacting us


def ensure_templates(wb):
    if "Reply templates" in wb.sheetnames:
        return
    ts = wb.create_sheet("Reply templates")
    ts.append(["Main issue (* = any)", "CDC result", "English note", "Hebrew reply - female", "Hebrew reply - male",
               "Hebrew reply - gender unknown", "Placeholders: {greet} {phone} {email} {cdc_email} {last_login} {ticket}"])
    for c in ts[1]: c.font = Font(bold=True)
    for row in REPLY_TEMPLATES: ts.append(list(row))
    ts.column_dimensions["B"].width = 45; ts.column_dimensions["C"].width = 50
    for L in "DEF": ts.column_dimensions[L].width = 70
    for r in ts.iter_rows(min_row=2):
        for c in r: c.alignment = Alignment(wrap_text=True, vertical="top")


def load_templates(wb):
    t = {}
    for row in wb["Reply templates"].iter_rows(min_row=2, values_only=True):
        issue, result, _en, f, m, n = (list(row) + [None] * 6)[:6]
        if result: t[(issue or "*", result)] = {"F": f or n or "", "M": m or n or "", "": n or f or ""}
    return t


def login_email(uids, emails, details):
    """Email of the login account. One account: its email. Several: the account
    with the old-style UID (no dashes); if several of those, the one logged
    into most recently."""
    uids = [u.strip() for u in str(uids or "").split(",") if u.strip() and u.strip() != "none"]
    per_acc = [e.strip() for e in str(emails or "").split(";")]
    last = {m.group(1): m.group(2) for m in re.finditer(r"^([0-9a-f-]{32,36}) \|.*?last login (\S+)", str(details or ""), re.M)}
    if not uids: return "(לא נמצא מייל בחשבון)"
    cands = [u for u in uids if "-" not in u] or uids
    cands.sort(key=lambda u: last.get(u, "never") if last.get(u, "never") != "never" else "", reverse=True)
    i = uids.index(cands[0])
    em = per_acc[i] if i < len(per_acc) else ""
    for part in em.split(","):
        if "@" in part: return part.strip()
    return "(לא נמצא מייל בחשבון)"


def build(inp, out, issues, cutoff=None):
    wb = openpyxl.load_workbook(inp)
    ensure_templates(wb)
    templates = load_templates(wb)
    us = wb["Users"]; h = [c.value for c in us[1]]; col = {n: i for i, n in enumerate(h)}
    missing = [c for c in HEAD_COLS + CUSTOMER_COLS + CDC_COLS if c not in col]
    if missing: sys.exit(f"Users sheet missing columns: {missing}")
    if "Replies" in wb.sheetnames: del wb["Replies"]
    rs = wb.create_sheet("Replies", 0)
    header = HEAD_COLS + ["Reply (copy-paste)", "Logged in since cutoff", "Email check", "First name (Hebrew)", "Gender guess", "Name needs check"] + ["Customer gave: " + c for c in CUSTOMER_COLS] + CDC_COLS
    rs.append(header)
    for i, c in enumerate(rs[1], 1):
        c.font = Font(bold=True)
        if len(HEAD_COLS) + 1 <= i <= len(HEAD_COLS) + 6: c.fill = FILL_REPLY
        elif len(HEAD_COLS) + 6 < i <= len(HEAD_COLS) + 6 + len(CUSTOMER_COLS): c.fill = FILL_CUSTOMER
        elif i > len(HEAD_COLS) + 6 + len(CUSTOMER_COLS): c.fill = FILL_CDC
    n = 0
    for row in us.iter_rows(min_row=2, values_only=True):
        g = lambda name: row[col[name]]
        if not g("User ID") or (issues != ["all"] and g("Main issue") not in issues): continue
        result = g("CDC result") or ""
        has_acc = g("CDC accounts on phone (UID)") not in (None, "none") or (result == "Email found")
        ll = last_login_after_contact(g("CDC account details (per UID)"), cutoff) if has_acc else None
        key = result
        if ll and result in ("Phone found", "Match: phone and email on same account", "Phone on multiple accounts", "Email found"):
            key = "LOGGED IN AFTER CONTACT"
        elif ll and result in ("Mismatch: phone account has other email; no account on user email", "Mismatch: email on a different account"):
            key = "LOGGED IN AFTER CONTACT (mismatch)"
        heb_name, gender, check = hebrew_name_and_gender(g("Customer name"))
        tpl = templates.get((g("Main issue"), key)) or templates.get(("*", key)) or {"F": f"[אין תבנית לתוצאה: {result}]"}
        text = tpl.get(gender) or tpl.get("") or tpl.get("F", "")
        cdc_email = login_email(g("CDC accounts on phone (UID)"), g("CDC email on phone account"), g("CDC account details (per UID)"))
        if "@" not in cdc_email and result in ("Email found", "Email on multiple accounts") and g("Email"):
            cdc_email = str(g("Email")).strip()
        gave = {e.strip().lower() for e in f"{g('Email') or ''},{g('Other emails') or ''}".replace(";", ",").split(",") if "@" in e}
        email_ok = "@" not in cdc_email or cdc_email.lower() in gave or email_matches_name(cdc_email, g("Customer name"))
        email_check = "" if email_ok else "masked - name does not match email"
        reply = text.format(greet=f"שלום {heb_name}," if heb_name else "שלום,", phone=g("Phone") or "", email=g("Email") or "",
                            cdc_email=cdc_email if email_ok else mask_email(cdc_email),
                            ticket=g("Latest ticket #") or "", last_login=ll or "")
        vals = [g(c) for c in HEAD_COLS] + [reply, "Yes" if ll else "No", email_check, heb_name, gender or "?", "check" if check else ""] + [g(c) for c in CUSTOMER_COLS] + [g(c) for c in CDC_COLS]
        rs.append(vals); n += 1
        r = rs.max_row
        for i in range(1, len(header) + 1):
            c = rs.cell(r, i); c.alignment = Alignment(wrap_text=True, vertical="top")
            if i in (len(HEAD_COLS) + 1, len(HEAD_COLS) + 2) and ll: c.fill = FILL_LOGGED_IN
            elif len(HEAD_COLS) + 1 <= i <= len(HEAD_COLS) + 6: c.fill = FILL_REPLY
            elif len(HEAD_COLS) + 6 < i <= len(HEAD_COLS) + 6 + len(CUSTOMER_COLS): c.fill = FILL_CUSTOMER
            elif i > len(HEAD_COLS) + 6 + len(CUSTOMER_COLS): c.fill = FILL_CDC
    widths = {"Ticket numbers": 22, "Customer's words (all login tickets)": 50, "Reply (copy-paste)": 60,
              "CDC result": 30, "CDC account details (per UID)": 60, "Main issue": 24, "Issue (most specific)": 30}
    for i, name in enumerate(header, 1):
        rs.column_dimensions[get_column_letter(i)].width = widths.get(name.replace("Customer gave: ", ""), 18)
    rs.freeze_panes = "D2"; rs.auto_filter.ref = rs.dimensions
    wb.save(out); print(f"{n} replies -> {out}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input"); ap.add_argument("output")
    ap.add_argument("--issues", default=";".join(DEFAULT_ISSUES), help="main issues to include, separated by ';' (or 'all')")
    ap.add_argument("--login-since", default=LOGIN_CUTOFF.isoformat(), help="a CDC login on/after this date counts as 'managed to log in'")
    a = ap.parse_args()
    build(a.input, a.output, [s.strip() for s in a.issues.split(";") if s.strip()], dt.date.fromisoformat(a.login_since))
