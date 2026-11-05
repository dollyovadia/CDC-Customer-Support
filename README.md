# CDC Customer Support

## `cdc_lookup.py` – login-issue lookup in CDC

For each user in the **Users** sheet of the login-issues workbook it:

1. Looks up CDC accounts by phone (`Phone (intl, for CDC)`, plus `Other phones`), trying every spelling (`+972…`, `972…`, `05…`), and records the email connected to it.
2. If the user also has an email: records whether it matches the CDC email, and whether a separate CDC account exists on that email (and that account's phone). Emails in `Other emails` are checked too.
3. Saves a **copy** of the workbook with new columns (`CDC lookup date`, `CDC result`, UIDs, emails, phones, errors). `User ID` and `Ticket numbers` stay on every row, so each finding stays linked to its tickets. The empty `CDC finding` column is filled with the result.
4. Copies the finding onto every row of the **Tickets** sheet (`CDC result`, email on phone account, UIDs, date), so a ticket number alone is enough to find the finding later.

Progress is saved to the output file every 25 users, so a dropped connection loses little; rerun to redo.

Read-only: it uses `accounts.search` only and never changes accounts. The original workbook is never overwritten.

### Setup
Store credentials as environment variables / secrets in the environment settings (never in code or chat):

| Variable | Meaning |
|---|---|
| `CDC_API_KEY` | site API key |
| `CDC_USER_KEY` | application key (read-only permissions recommended) |
| `CDC_SECRET` | application secret |
| `CDC_DATA_CENTER` | `eu1` (default), `us1`, `au1`, … |
| `CDC_PHONE_FIELDS` | optional, default `phoneNumber,profile.phones.number` |
| `CDC_EMAIL_FIELDS` | optional, default `profile.email,loginIDs.emails` |

The host `accounts.<dc>.gigya.com` must be allowed in the network settings.

### Run
```bash
python3 cdc_lookup.py input.xlsx output.xlsx --limit 5   # test on 5 users
python3 cdc_lookup.py input.xlsx output.xlsx             # all users
```

### `CDC result` values
- `Match: phone and email on same account`
- `Mismatch: email on a different account`
- `Mismatch: phone account has other email; no account on user email`
- `Phone not in CDC; email has account`
- `Phone and email not in CDC (likely erased)` / `Phone not in CDC (likely erased)` / `Email not in CDC (likely erased)`
- `Phone found` / `Email found` (only one identifier available)
- `Phone on multiple accounts` / `Email on multiple accounts`
- `Lookup error` (details in `CDC lookup error`)
