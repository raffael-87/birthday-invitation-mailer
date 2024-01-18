"""
Seit dem 01.01.2024 funktioniert der Versand per Python-Skript über Gmail nur noch mit einem sogenannten "App-Password".
Hierfür muss folgendes aktiviert sein bzw. werden: "Google Konto Verwalten" => "Sicherheit/Security" => "Bestätigen in zwei Schritten"

Anschließend direkt über den Link myaccount.google.com/apppasswords ein App-Password generieren lassen. Dieses muss dann in die .env-Datei statt dem Passwort, welches man sonst für den ganz normalen Login in seinen Google-Account benötigt.
"""

import os
from dotenv import load_dotenv
import pandas as pd
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()

# Email-Server Konfiguration
EMAIL_SERVER = "smtp.gmail.com"
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASSWORD = os.environ.get("APP_PASSWORD")

# Pfad zur Excel-Datei
file_path_sheet = "test.xlsx"  # Pfad zur Excel-Datei mit den E-Mail-Adressen.
file_path_attachments = [ # Optional. Eckige Klammern leer lassen, wenn es keine Anhänge gibt.
    "attachments/party_1.jpg",
    "attachments/party_2.jpg",
    "attachments/party_3.jpg"
]  

# Textbausteine
subject = "Einladung zu meiner Geburtstagsfeier ..."
message = """
Hallo RUFNAME_PLATZHALTER,
hiermit lade ich dich herzlichst zu meiner Geburtstagsfeier am ...
"""

######### Excel-File einlesen  #####################################################
Tabelle1 = pd.read_excel(file_path_sheet, sheet_name="Tabelle1")

for index, row in Tabelle1.iterrows():
    message_body = message.replace("RUFNAME_PLATZHALTER", row["Rufname"])

    # E-Mail-Objekt erstellen
    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = row["E-Mail Adresse"]
    msg["Subject"] = subject
    msg.attach(MIMEText(message_body, "plain"))

    # Anhänge hinzufügen, falls vorhanden
    for file_path in file_path_attachments:
        if file_path:  # Überprüft, ob der Pfad nicht leer ist
            with open(file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(file_path)}"',
            )
            msg.attach(part)

    with smtplib.SMTP(EMAIL_SERVER) as connection:
        connection.starttls()
        connection.login(user=EMAIL_USER, password=EMAIL_PASSWORD)
        connection.send_message(msg)

print('EMAILS ERFOLGREICH VERSANDT!')

