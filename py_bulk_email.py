#!/usr/bin/env python

from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
import os
import smtplib

from pyexcel_xls import get_data


def rows_to_dicts(lst):
    headers = lst[0]
    out_lst = []
    for el in lst[1:]:
        out_lst.append({k: v for k, v in zip(headers, el)})
    return out_lst


def cols_to_dicts(lst):
    headers = [el[0] for el in lst]
    out_lst = []
    for el_idx in range(1, len(lst[0])):
        out_lst.append({k: v for k, v in zip(headers,
                                             [el[el_idx] for el in lst])})
    return out_lst


def batch_send_email(xls='py_bulk_email.xlsx'):
    # Get data from xls
    try:
        try:
            data = get_data(xls)
        except IOError:
            data = get_data(xls[:-1])
    except IOError:
        raise IOError('No such files: send_email.xlsx nor send_email.xls. '
                      'You need to provide an xls file with the email '
                      'information')

    # Get the data out
    account_info = rows_to_dicts(data['Account info'])[0]
    email_content = cols_to_dicts(data['Email content'])[0]
    contacts = rows_to_dicts(data['Contacts'])

    subject = email_content['subject'].encode('utf-8').strip()
    prim_email_f = email_content['primary email field'].encode('utf-8').strip()
    sec_email_f = email_content['secondary email field'].encode(
        'utf-8').strip()

    html = email_content['html'].encode('utf-8')

    from_email = account_info['email'].encode('utf-8').strip()
    smtp = account_info['smtp'].encode('utf-8').strip()
    port = account_info['port']
    username = account_info['username'].encode('utf-8').strip()
    password = account_info['password'].encode('utf-8').strip()

    # Connect and authenticate
    try:
        mail = smtplib.SMTP(smtp, port, timeout=5)
        mail.ehlo_or_helo_if_needed()
        mail.starttls()
    except smtplib.SMTPServerDisconnected:
        mail = smtplib.SMTP_SSL(smtp, port, timeout=5)
        mail.ehlo_or_helo_if_needed()
    mail.login(username, password)

    for ct in contacts:
        # Create message container
        msg = MIMEMultipart('related', 'utf-8')
        msg['Message-ID'] = make_msgid()  # or you'll look like spam!
        msg['Content-Type'] = 'text/html; charset=utf-8'
        msg['Subject'] = subject
        msg['From'] = from_email
        msg['Date'] = formatdate(localtime=True)
        msg.preamble = 'This is a multi-part message in MIME format.'

        # Create the body of the message
        # Record the MIME type - text/html
        text = MIMEText(html.format(**ct), 'html')
        msg.attach(text)

        # Load attachments
        for fname in os.listdir('attachments'):
            with open(os.path.join('attachments', fname), 'rb') as f:
                part = MIMEApplication(f.read(), Name=fname)
                part['Content-Disposition'] = ('attachment; filename="%s"'
                                               '' % fname)
                msg.attach(part)

        # Load images
        for i, fname in enumerate(os.listdir('inline_images')):
            with open(os.path.join('inline_images', fname), 'rb') as f:
                img = MIMEImage(f.read(), Name=fname)
                img.add_header('Content-ID', '<image{}>'.format(i+1))
                msg.attach(img)

        # Batch send emails
        to_email = ct[prim_email_f].encode('utf-8')
        # If there is no primary email and secondary email field is set
        if to_email in (None, '') and sec_email_f not in (None, ''):
            to_email = ct[sec_email_f].encode('utf-8')
            if to_email is None or to_email == '':
                continue
        msg['To'] = to_email
        mail.sendmail(from_email, to_email, msg.as_string())

    mail.quit()


if __name__ == '__main__':
    batch_send_email()
