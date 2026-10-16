# -*- coding: utf-8 -*-
"""
lbv_email.py - Mehrarbeit-PDFs per SMTP versenden
Liest Parameter aus einer JSON-Datei und versendet E-Mails.
Aufruf: lbv_email.exe <json_pfad>
"""
import sys
import json
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def sende_emails(json_pfad):
    with open(json_pfad, 'r', encoding='utf-8') as f:
        data = json.load(f)

    smtp_server     = data['smtp_server']
    smtp_port       = int(data['smtp_port'])
    verschluesselung = data.get('verschluesselung', 'TLS').upper()
    benutzername    = data['benutzername']
    passwort        = data['passwort']
    absender        = data.get('absender', benutzername)
    mails           = data['mails']

    gesendet  = 0
    fehler    = 0

    # SMTP-Verbindung aufbauen
    try:
        if smtp_port == 465 or verschluesselung == 'SSL':
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            if verschluesselung in ('TLS', 'STARTTLS'):
                server.starttls()

        server.login(benutzername, passwort)
        print(f"Verbunden mit {smtp_server}:{smtp_port}")

    except Exception as e:
        print(f"FEHLER Verbindung: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)

    # Mails versenden
    for mail_data in mails:
        an        = mail_data['an']
        pdf_pfad  = mail_data['pdf_pfad']
        dateiname = mail_data.get('dateiname', os.path.basename(pdf_pfad))

        try:
            msg = MIMEMultipart()
            msg['From']    = absender
            msg['To']      = an
            msg['Subject'] = f"Mehrarbeitsverg\u00fctung - {dateiname}"

            body = (
                "Sehr geehrte Damen und Herren,\n\n"
                "anbei erhalten Sie Ihre Mehrarbeitsverg\u00fctungsabrechnung.\n\n"
                "Mit freundlichen Gr\u00fc\u00dfen"
            )
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # PDF-Anhang
            with open(pdf_pfad, 'rb') as f:
                anhang = MIMEBase('application', 'octet-stream')
                anhang.set_payload(f.read())
            encoders.encode_base64(anhang)
            anhang.add_header(
                'Content-Disposition',
                'attachment',
                filename=dateiname
            )
            msg.attach(anhang)

            server.sendmail(absender, an, msg.as_string())
            print(f"OK: {an} <- {dateiname}")
            gesendet += 1

        except Exception as e:
            print(f"FEHLER {an}: {type(e).__name__}: {e}", file=sys.stderr)
            fehler += 1

    server.quit()
    print(f"Fertig: {gesendet} gesendet, {fehler} Fehler")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Aufruf: lbv_email.exe <json_pfad>")
        sys.exit(1)
    sende_emails(sys.argv[1])
