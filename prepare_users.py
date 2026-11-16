#!/usr/bin/env python3
"""
Turn a raw ticket export into the Users/Tickets workbook that cdc_lookup.py
and make_replies.py expect.

Expected export columns (header row optional, first non-empty column):
  Ticket #, Status, Created, Updated, Channel, Contact (email or phone),
  Customer text, Phone, Email
Tickets are joined into one user when they share a phone or an email
(phones/emails found inside the customer text count too).

Usage: python3 prepare_users.py EXPORT.xlsx OUT.xlsx
"""
import re, sys, datetime as dt
import openpyxl
from openpyxl.styles import Font

PHONE_RE = re.compile(r"(?<!\d)(?:\+?972[-\s]?|0)(5\d)[-\s]?(\d{3})[-\s]?(\d{4})(?!\d)")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def norm_phone(v):
    d = re.sub(r"\D", "", str(v or ""))
    if d.startswith("972"): d = "0" + d[3:]
    return d if re.fullmatch(r"05\d{8}", d) else None


def intl(p): return "+972" + p[1:] if p else None


def classify(text):
    t = text or ""
    if re.search(r"בשימוש|תפוס|קיים במערכת|already in use", t): return "Phone / email already in use", "הטלפון/המייל כבר בשימוש"
    if re.search(r"איפוס|אפס|סיסמ|password|reset", t): return "Password reset not working", "איפוס סיסמה לא עובד"
    if re.search(r"קוד|sms|otp", t, re.I): return "Verification code (SMS/OTP)", "קוד אימות"
    return "Can't log in", "לא מצליח/ה להתחבר"


def main(inp, out):
    ws = openpyxl.load_workbook(inp, data_only=True).active
    tickets = []
    for r in ws.iter_rows(values_only=True):
        r = [c for c in r]
        while r and r[0] in (None, ""): r = r[1:]          # drop leading empty column(s)
        if not r or not isinstance(r[0], (int, float)) : continue
        r += [None] * (9 - len(r))
        tid, status, created, updated, channel, contact, text, phone, email = r[:9]
        phones, emails = set(), set()
        for v in (contact, phone):
            p = norm_phone(v)
            if p: phones.add(p)
        for v in (contact, email):
            if v and "@" in str(v): emails.add(str(v).strip().lower())
        for m in PHONE_RE.finditer(str(text or "")): phones.add("0" + "".join(m.groups()))
        for m in EMAIL_RE.finditer(str(text or "")): emails.add(m.group(0).lower())
        tickets.append(dict(tid=int(tid), status=status, created=created, updated=updated, channel=channel,
                            text=(text or "").strip(), phones=phones, emails=emails,
                            pri_phone=norm_phone(phone) or norm_phone(contact), pri_email=(email or (contact if "@" in str(contact) else "") or "").strip().lower()))
    # union-find on shared identifiers
    parent = list(range(len(tickets)))
    def find(i):
        while parent[i] != i: parent[i] = parent[parent[i]]; i = parent[i]
        return i
    seen = {}
    for i, t in enumerate(tickets):
        for k in [("p", p) for p in t["phones"]] + [("e", e) for e in t["emails"]]:
            if k in seen: parent[find(i)] = find(seen[k])
            else: seen[k] = i
    groups = {}
    for i in range(len(tickets)): groups.setdefault(find(i), []).append(tickets[i])

    wb = openpyxl.Workbook(); us = wb.active; us.title = "Users"
    uh = ["User ID", "Priority", "Main issue", "Issue (most specific)", "Issue (Hebrew)", "Customer name", "Phone", "Phone (intl, for CDC)",
          "Email", "Has phone + email", "Other phones", "Other emails", "# login tickets", "Ticket numbers", "Latest ticket #",
          "First contact", "Last contact", "Channels", "Status", "Customer's words (all login tickets)", "Suggested action", "CDC finding", "Notes"]
    us.append(uh)
    ts = wb.create_sheet("Tickets")
    th = ["User ID", "Ticket #", "Date", "Channel", "Status", "Phone(s) in ticket", "Email(s) in ticket", "Customer text"]
    ts.append(th)
    for c in list(us[1]) + list(ts[1]): c.font = Font(bold=True)
    ordered = sorted(groups.values(), key=lambda g: (-len(g), min(t["created"] for t in g)))
    for n, g in enumerate(ordered, 1):
        uid = f"N{n:04d}"
        g.sort(key=lambda t: t["created"])
        phones = [t["pri_phone"] for t in g if t["pri_phone"]] + sorted({p for t in g for p in t["phones"]})
        emails = [t["pri_email"] for t in g if t["pri_email"]] + sorted({e for t in g for e in t["emails"]})
        phones = list(dict.fromkeys(phones)); emails = list(dict.fromkeys(emails))
        phone = phones[0] if phones else None; email = emails[0] if emails else None
        words = "\n".join(f"[{t['tid']}] {t['text']}" for t in g if t["text"])
        mi, he = classify(" ".join(t["text"] for t in g))
        pri = "P1" if len(g) >= 3 else ("P2" if len(g) == 2 else "P3")
        us.append([uid, pri, mi, mi, he, None, phone, intl(phone), email, "Yes" if phone and email else "No",
                   ", ".join(phones[1:]) or None, ", ".join(emails[1:]) or None, len(g), ", ".join(str(t["tid"]) for t in g), g[-1]["tid"],
                   g[0]["created"], g[-1]["created"], ", ".join(sorted({t["channel"] for t in g if t["channel"]})),
                   ", ".join(sorted({str(t["status"]) for t in g})), words or None, None, None, None])
        for t in g:
            ts.append([uid, t["tid"], t["created"], t["channel"], t["status"], ", ".join(sorted(t["phones"])), ", ".join(sorted(t["emails"])), t["text"]])
    wb.save(out)
    print(f"{len(tickets)} tickets -> {len(groups)} users -> {out}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
