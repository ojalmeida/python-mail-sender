import csv
import smtplib
import os
import time
from getpass import getpass


def load_recipients(path):
    recipients = []

    with open(path, newline='') as csv_file:
        file = csv.reader(csv_file)
        for row in file:
            recipient = {}
            recipient['attributes'] = row[0]
            recipients.append(recipient)
        return recipients


def prepare_recipient(recipient):
    attributes = recipient['attributes']

    _ = attributes.split(';')
    for i in range(0, len(_)):
        if _[i] != '' and i != 0:
            recipient[i] = _[i]

    recipient['email'] = _[0]
    recipient.pop('attributes')

    return recipient


def get_message():
    def add_message():
        found_file = False
        while not found_file:
            clear()
            print('# Extern message file')
            path = input('\n File path: ')
            try:
                file = open(path, newline='')
                found_file = True
                message = file.readlines()
                return message
            except OSError:
                print('\nFile not found')
                time.sleep(3)


    print(
        'Type the email, using the patterns. Press Enter for insert new line, type \'-p\' for new paragraph% or '
        '\'-end\' to finish input ')
    print('\n(Or type -insert to add extern message file)')
    message = []
    line = input()
    ended = False
    while not ended:
        if line == '-end':
            ended = True
        elif line == '-p':
            line = ''
        elif line == '-insert':
            message = add_message()
            ended = True
        else:
            message.append(line)
            line = input()

    return message


def prepare_message(subject, message, recipient, patterns):
    prepared_message = 'Subject: {}\n\n'.format(subject).encode('iso-8859-1')
    for line in message:
        for pattern in patterns:
            to_be_replaced = pattern[0]
            index_of_attribute = int(pattern[1])
            line = line.replace(to_be_replaced, recipient[index_of_attribute])

        prepared_message += (line + '\n').encode('iso-8859-1')

    recipient['message'] = prepared_message

    return recipient


def load_patterns():
    # load patterns
    print('e.g. \'%name, 0\'\n')
    print('Before: Dear $name. After: Dear John (\'John\' was index 0 in .csv file)\n')
    print('---------- Pattern declaration ----------\n')
    print('Type the patterns to be used in message formatting: \n')
    patterns = []
    ready = False
    while not ready:
        pattern = input().split(', ')
        if pattern != ['']:
            patterns.append(pattern)
        else:
            ready = True

    return patterns


def get_credentials():
    email = input('Email: ')
    password = input('Password: ')

    return email, password


def clear():
    os.system('cls')
    # for i in range(50):
    #     print('\n')


accepted = False
while not accepted:
    clear()
    print('---------- / Mail Sender \\ ----------')
    print('# E-mail service\n')
    option = input('1. Gmail\n2. Outlook\n\nAnswer: ')
    if option == '1':
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        accepted = True
    elif option == '2':
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.ehlo()
        accepted = True

authentication_success = False
while not authentication_success:
    try:
        clear()
        print('# Login \n')
        email, password = get_credentials()
        server.login(email, password)
        authentication_success = True
    except smtplib.SMTPAuthenticationError:
        print('\nServer error, try turn off two-factor security of your e-mail service')
        time.sleep(3)

clear()
print('# Subject\n')
subject = input('Subject: ')

clear()
print('# Patterns \n')
patterns = load_patterns()

clear()
print('# Message\n')
message = get_message()

# sends e-mail for each recipient
found_file = False
while not found_file:
    clear()
    print('# CSV')
    try:
        path = input('Type recipients .csv file : ')
        prepared_recipients = []
        for recipient in load_recipients(path):
            found_file = True
            prepared_recipient = prepare_recipient(recipient)
            prepared_recipient = prepare_message(subject, message, prepared_recipient, patterns)

            prepared_recipients.append(prepared_recipient)
    except OSError:
        print('\nFile not found')
        time.sleep(3)

clear()
print('{} e-mail(s) queued. \n'.format(len(prepared_recipients)))
for i in range(len(prepared_recipients)):
    print('{}. to: {}\n'.format(i + 1, prepared_recipients[i]['email']))

print('Message: \n\n{}'.format(message))

print('\n')
print('Are you sure you want to send? (yes/no)')
resp = input().lower()

if resp == 'yes' or resp == 'y':
    clear()

    sent_emails = 0
    errors = 0
    for prepared_recipient in prepared_recipients:
        print('# Sending...\n')
        print('{}/{} emails sent\n'.format(sent_emails, len(prepared_recipients)))

        try:
            server.sendmail(email, prepared_recipient['email'], prepared_recipient['message'])
            sent_emails += 1
        except:
            print('An error occurred on mail sending to {}'.format(prepared_recipient['email']))
            errors += 1

        clear()

    print('\n')
    if sent_emails == len(prepared_recipients):
        print('All emails sent. No errors')
    else:
        print('{} could be sent. {} errors'.format(sent_emails, errors))

server.quit()
