import pandas as pd
import re

class Person:
    def __init__(self, first_name, last_name, apartment, address, email=None):
        self.first_name = first_name
        self.last_name = last_name
        self.apartment = apartment
        self.address = address
        self.emails = split_emails(email, first_name, last_name)

    def __repr__(self):
        return f"Person({self.first_name} {self.last_name}, \nemails={self.emails}, \naddress={self.address})"


def split_emails(email_string, first_name, last_name):
    emails = []
    if email_string:
        emails = [email.strip() for email in re.split(r'[;,]', email_string) if email.strip()]
        emails = [email for email in emails if validate_email(email)]
    else:
        print(f"Warning: No email provided for {first_name} {last_name}")
        raise ValueError("Email is required")
    return emails


# TODO: Improve email validation
def validate_email(email):
    if not "@" in email and not "." in email and len(email) < 5:
        raise ValueError(f"Invalid email address: {email}")
    return True


def extract_person_data(input_file):
    df = pd.read_excel(input_file)
    # print(df.head())  # Print the first few rows for debugging

    persons = []
    for _, row in df.iterrows():
        person = Person(
            first_name = row['ees_nimi'],
            last_name = row['pere_nimi'],
            email = row['klient_mail'],
            apartment = row['korter'],
            address = row['yhistu'].lower() + " " + str(row['maj_nr'])
        )
        persons.append(person)
    return persons