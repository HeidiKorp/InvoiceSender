import argparse
from pathlib import Path

from xls_extractor import extract_person_data
from pdf_extractor import separate_invoices, save_each_invoice_as_file
from email_sender import send_emails_with_invoices, ensure_outlook_ready

def main():
    print("Invoice sender started.")
    parser = argparse.ArgumentParser(description="Send invoices to clients.")
    parser.add_argument('--clients', type=str, required=True, help='XLS file with client data')
    parser.add_argument('--invoices', type=str, required=True, help='PDF file with invoices')

    # Parse the command-line arguments
    args = parser.parse_args()

    print(f"Clients file: {args.clients}")
    print(f"Invoices file: {args.invoices}")

    persons = extract_person_data(args.clients)
    print(f"Extracted {len(persons)} persons from the clients file.")

    print(f'\n\n')
    invoices = separate_invoices(args.invoices)
    print(f"Extracted {len(invoices)} invoices from the PDF file.")

    current_path = Path(__file__).resolve().parent
    dest = current_path / "arved"
    invoices_dir = save_each_invoice_as_file(invoices, dest) # returns the full folder path to all individual invoices
    print(f"Invoices saved to directory: {dest}")

    # Compose emails and send them
    ensure_outlook_ready()
    send_emails_with_invoices(persons, invoices_dir)
    


if __name__ == "__main__":
    main()