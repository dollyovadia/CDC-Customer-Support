#!/usr/bin/env python3
"""Small file for the tech team: users whose phone sits on a CDC account whose
email does not look like theirs (masked in the replies) AND the CDC result is a
mismatch. Lists the account(s) to erase / release so the customer can register
again. Usage: python3 erase_list.py REPLIES.xlsx OUT.xlsx"""
import sys, openpyxl
from openpyxl.styles import Font, Alignment
inp, out = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(inp, read_only=True)["Replies"]
rows = list(ws.iter_rows(values_only=True)); h = rows[0]
sel = [dict(zip(h, r)) for r in rows[1:]]
sel = [d for d in sel if d["Email check"] and str(d["CDC result"]).startswith("Mismatch")]
wb = openpyxl.Workbook(); o = wb.active; o.title = "Accounts to erase"
cols = ["User ID", "Ticket numbers", "Customer name", "Customer phone", "Customer email (keep)", "CDC result",
        "UID to ERASE (phone account)", "Email on that account", "Account details", "UID of customer's own email account (KEEP)",
        "Logged in since cutoff", "Action"]
o.append(cols); [setattr(c, "font", Font(bold=True)) for c in o[1]]
for d in sel:
    keep = d["CDC accounts on user email (UID)"] if d["CDC accounts on user email (UID)"] not in (None, "none") else ""
    action = ("Erase the phone account (or remove the phone from it) so the customer can register again with phone + own email."
              if not keep else
              "Remove the phone from the listed account (or erase it) and attach it to the customer's own email account.")
    o.append([d["User ID"], d["Ticket numbers"], d["Customer gave: Customer name"], d["Customer gave: Phone"], d["Customer gave: Email"],
              d["CDC result"], d["CDC accounts on phone (UID)"], d["CDC email on phone account"], d["CDC account details (per UID)"],
              keep, d["Logged in since cutoff"], action])
for i, w in enumerate([8, 22, 18, 14, 28, 40, 38, 30, 60, 38, 10, 60], 1):
    o.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for r in o.iter_rows(min_row=2):
    for c in r: c.alignment = Alignment(wrap_text=True, vertical="top")
o.freeze_panes = "B2"; wb.save(out); print(f"{len(sel)} users -> {out}")
