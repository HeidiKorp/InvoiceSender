# InvoiceSender
Gets a list of people's email addressed and a PDF with a collection of their respective invoices and sends the bills accordingly.

Program plan:
* 2 inputs:
    * **xls** file with house name, person name, email address
    * **PDF** file with multiple invoices grouped together, each corresponding to a certain person
* Solution:
    * Read the xls file
    * Get the person's email address
        * Create a local database and later update it
        * There might be 2 email addresses, send to both
    * Read the PDF file
    * Extract each file in the PDF file
    * Access my email tool from this program
    * Send the correct file to the correct email address (local database)
        * Local database can be a dict at first, later a better database


 ## Usage
`python invoice_sender.py --clients data/kliendid.xls --invoices data/palman_aug_25.pdf`

Use with GUI:
In InvoiceSender:
`python -m run_app`

Create an exe:
`rm -rf build dist *.spec`
`pyinstaller --onedir --noconsole --name ArveteSaatja   --paths .   --add-data "tesseract:tesseract" --add-data "config.cfg:." run_app.py`

Notes:
* The address in the PDF must match the one provided in the client's table

## Setting up the Gmail account in Outlook
* Make sure you have the classic Outlook installed 
    * https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5
* Go to **Control Panel** and search for "Mail (Microsoft Outlook)"
* This opens up the wizard where you can manage and add new accounts.
* When you can choose between which account to add, choose to add an account manually (opposed to the Microsoft 365 account)
* If your Gmail has a 2-step verification, you need to generate an app password specifically for login into Outlook
    * Go to your Gmail
    * Go to Google Account Security (from your profile)\
    * Under "Signing in to Google" -> App Passwords (or search for it)
    * Generate one for *Mail / Outlook*
    * Save the generated password, remove spaces and paste it into the Outlook password field\
    * Account type: IMAP
    * Incoming mail server: `imap.gmail.com`
    * Outgoing mail server: `smtp.gmail.com`
* As the username, set it to your full gmail account like korpheidi@gmail.com\
* In the bottom right corner click "More settings..."
* On the **Outgoing Server** tab:
    * Check "My outgoing server (SMTP) requires authentication"
    * Select "Use same settings as my incoming mail server"
* On **Advanced** tab:
    * Incoming server (IMAP): **993**, encryption **SSL/TLS**
    * Outgoing server (SMTP): **587**, encryption **STARTTLS**
* If you skip the *More settings...* step, Outlook will try "no encryption" on port 25 -> Gmail rejects with `530 5.7.0 Authentication Required`