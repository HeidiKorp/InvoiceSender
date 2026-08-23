# Arvete Saatja (InvoiceSender)

Eesti keeles: [README.et.md](README.et.md)

Desktop app for sending apartment invoices by email. It splits a combined invoice file, matches apartments to clients, saves one PDF per apartment, and creates Outlook drafts.

## What it does

1. Choose invoice type: **Kommunaalarved** or **Küttearved**. The type mainly sets the email subject and body templates.
2. Choose the invoice file and the clients file.
   - Kommunaal invoices: PDF
   - Küte invoices: Excel (`.xls`, `.xlsx`)
   - Clients: Excel (`.xls`, `.xlsx`, `.xlsm`)
3. Invoices are split and matched to clients **by apartment number**. A progress bar shows work in progress. Apartments with no invoice, or invoices with no client, are dropped and an error is shown.
4. Matched invoices are saved under `arved/<address>/<period>/` next to the invoice file (one `{apartment}.pdf` per apartment).
5. An email template editor opens. Default subject is `{address} arve {period} {year}`, for example `Õismäe tee 48 arve august 2026`. `{period}` must be an Estonian month name and `{year}` a year from 2001 to 2999; if one invoice page parses those poorly, other invoices in the same file are used. You can edit subject and body, and load or save a `.cfg` template.
6. Drafts are created in Outlook (must be installed and logged in) with category `ArveteSaatja`, the matching client email(s), and the apartment PDF attached. Several emails on the same client row all get a draft. Outlook should show the **Drafts** folder.
7. **Saada mustandid** sends only drafts in that category.

On error, the UI returns to the previous usable state (inputs stay selected; cancel restores the idle bar). Unexpected failures are written to `utils/error.log` (next to the exe when frozen), with timestamp, operation, and traceback.

## Requirements

- Windows
- Classic Outlook, signed in
- Tesseract OCR with Estonian language data (`est`) — used for kommunaal PDFs
- Excel — used for küte invoices
- Python 3, plus project dependencies (`pandas`, `openpyxl` for `.xlsx`/`.xlsm` clients, `xlrd` for `.xls`, `pytesseract`, `pypdf`, PyMuPDF, `pywin32`, `ttkbootstrap`)

## Usage

```
python -m run_app
```

or:

```
python run_app.py
```

Build an exe:

```
rm -rf build dist *.spec
pyinstaller --onedir --noconsole --name ArveteSaatja --paths . --add-data "tesseract:tesseract" --add-data "config.cfg:." run_app.py
```

## Clients file

Required columns: `klient_mail`, `korter`, `yhistu`, `maj_nr`.

- `korter` must be digits only
- `klient_mail` is required; several addresses on one row may be separated by `,` or `;`

Matching is by apartment number only, not by email or street address.

## Email templates

Defaults live in `config.cfg`. Placeholders: `{address}`, `{period}`, `{year}`, `{apartment}`, `{ky_name}`. The subject is filled from invoices that have a valid address, Estonian month name, and year (2001–2999); a bad first page does not keep `periood` in the title if another invoice has `august`.

Both invoice types use:

```
SUBJECT={address} arve {period} {year}
```

## Setting up the Gmail account in Outlook

- Make sure you have the classic Outlook installed
    - https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5
- Go to **Control Panel** and search for "Mail (Microsoft Outlook)"
- This opens up the wizard where you can manage and add new accounts.
- When you can choose between which account to add, choose to add an account manually (opposed to the Microsoft 365 account)
- If your Gmail has a 2-step verification, you need to generate an app password specifically for login into Outlook
    - Go to your Gmail
    - Go to Google Account Security (from your profile)
    - Under "Signing in to Google" -> App Passwords (or search for it)
    - Generate one for *Mail / Outlook*
    - Save the generated password, remove spaces and paste it into the Outlook password field
    - Account type: IMAP
    - Incoming mail server: `imap.gmail.com`
    - Outgoing mail server: `smtp.gmail.com`
- As the username, set it to your full gmail account like korpheidi@gmail.com
- In the bottom right corner click "More settings..."
- On the **Outgoing Server** tab:
    - Check "My outgoing server (SMTP) requires authentication"
    - Select "Use same settings as my incoming mail server"
- On **Advanced** tab:
    - Incoming server (IMAP): **993**, encryption **SSL/TLS**
    - Outgoing server (SMTP): **587**, encryption **STARTTLS**
- If you skip the *More settings...* step, Outlook will try "no encryption" on port 25 -> Gmail rejects with `530 5.7.0 Authentication Required`
