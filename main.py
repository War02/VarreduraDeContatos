import imaplib
import csv
from email import message_from_bytes
import re

def save_emails_to_csv(emails, filename):
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['From']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for email in emails:
            writer.writerow(email)

def fetch_emails_from_outlook(username, password, server='email-ssl.com.br', ssl_port=993, limit=200):
    print("Conectando ao servidor IMAP...")
    imap = imaplib.IMAP4_SSL(server, ssl_port)
    print("Autenticando...")
    imap.login(username, password)
    imap.select('INBOX')

    print("Buscando e-mails...")
    result, data = imap.search(None, 'ALL')
    emails = []

    for num in data[0].split():
        if len(emails) >= limit:
            break
        result, raw_email = imap.fetch(num, '(RFC822)')
        email_msg = message_from_bytes(raw_email[0][1])

        match = re.search(r'<([^>]+)>', email_msg['From'])
        if match:
            from_email = match.group(1)
        else:
            from_email = email_msg['From']

        emails.append(from_email)

    imap.close()
    imap.logout()

    unique_emails = list(set(emails))
    unique_emails.sort()

    return [{"From": email} for email in unique_emails]

def main():
    username = 'm.barreto@mecpar.com'
    password = 'Dw1#8T1O'
    emails = fetch_emails_from_outlook(username, password, limit=200)
    save_emails_to_csv(emails, 'outlook_emails.csv')
    print("Endereços de e-mail salvos em outlook_emails.csv")

if __name__ == "__main__":
    main()
