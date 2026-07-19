import streamlit as st
import pandas as pd
import os
import smtplib
import shutil
import zipfile
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- Konfiguration & UI ---
st.set_page_config(page_title="Lehrevaluations-Tool", layout="wide")

st.title("📧 E-Mail Versand Tool für Lehrevaluationen")

# Anleitung
with st.expander("ℹ️ Anleitung: So funktioniert es"):
    st.write("""
    1. **Excel-Datei hochladen**: Die Excel-Datei muss die Spalten `Emailempfänger`, `LVEName`, `Kennung`, `Anrede` und `Nachname` enthalten.
    2. **ZIP-Verzeichnis hochladen**: Laden Sie die ZIP-Datei hoch, die alle Unterordner (benannt nach den Nachnamen) enthält. Das Tool entpackt diese intern.
    3. **Konfiguration**: Geben Sie die SMTP-Server-Daten und Absender-E-Mail ein.
    4. **Textbearbeitung**: Passen Sie den E-Mail-Text im Editor an.
    5. **Versand**: Starten Sie den Versand.
    """)

# Sidebar - Einstellungen
st.sidebar.header("SMTP Konfiguration")
smtp_server = st.sidebar.text_input("SMTP Server", "mail.uni-ulm.de")
smtp_port = st.sidebar.number_input("Port", 587)
sender_email = st.sidebar.text_input("Absender E-Mail")
smtp_password = st.sidebar.text_input("Passwort", type="password")
bcc_email = st.sidebar.text_input("BCC E-Mail", sender_email)

# Hauptbereich - Inputs
col1, col2 = st.columns(2)
uploaded_excel = col1.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])
uploaded_zip = col2.file_uploader("ZIP-Datei mit Dokumenten hochladen (.zip)", type=["zip"])

st.subheader("E-Mail Text bearbeiten")
email_body_template = st.text_area("E-Mail Body", height=400, value="""Sehr geehrte(r) {anrede},

Wir möchten Sie heute freundlich daran erinnern, dass die Lehrevaluation {veranstaltungen_de} noch aussteht. 

Die Unterlagen sollten Sie bereits erhalten haben, aber wir haben diese nochmal angehängt. 

Vielen Dank für Ihre Mühe!

---

Dear {anrede_en},

We would like to kindly remind you that the course evaluation for {veranstaltungen_en} is still pending.

Thank you very much for your effort!

Tatjana Shevchik
Universität Ulm""")

if st.button("🚀 E-Mails versenden"):
    if not (uploaded_excel and uploaded_zip and sender_email and smtp_password):
        st.error("Bitte laden Sie alle Dateien hoch und füllen Sie die SMTP-Einstellungen aus.")
    else:
        # Prozess starten
        with st.spinner("Verarbeite Daten und versende E-Mails..."):
            # Temporäres Verzeichnis für entpackte Dateien
            with tempfile.TemporaryDirectory() as tmpdir:
                # Zip entpacken
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(tmpdir)
                
                df = pd.read_excel(uploaded_excel)
                df.columns = df.columns.str.strip()
                
                # Gruppieren
                grouped = df.groupby(['Emailempfänger', 'Anrede', 'Nachname'])
                
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, smtp_password)
                
                for (email, anrede, nachname), group in grouped:
                    # Logik für Textbausteine
                    veranstaltungen = (group['LVEName'].astype(str) + " " + group['Kennung'].astype(str)).unique().tolist()
                    
                    # Pfad zum Ordner
                    folder_path = os.path.join(tmpdir, str(nachname).strip())
                    
                    # E-Mail erstellen
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = email
                    msg['Subject'] = Header(f"Erinnerung Lehrevaluation: {nachname}", 'utf-8').encode()
                    
                    # Platzhalter ersetzen
                    body = email_body_template.format(
                        anrede=anrede,
                        anrede_en=anrede.replace("Sehr geehrte Frau", "Ms.").replace("Sehr geehrter Herr", "Mr."),
                        veranstaltungen_de=", ".join(veranstaltungen),
                        veranstaltungen_en=", ".join(veranstaltungen)
                    )
                    msg.attach(MIMEText(body, 'plain', 'utf-8'))
                    
                    # Datei anhängen falls vorhanden
                    if os.path.exists(folder_path):
                        # Zippen des Ordners
                        zip_out = shutil.make_archive(os.path.join(tmpdir, nachname), 'zip', folder_path)
                        with open(zip_out, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header("Content-Disposition", f"attachment; filename={nachname}.zip")
                            msg.attach(part)
                            
                    server.sendmail(sender_email, [email, bcc_email], msg.as_string())
                
                server.quit()
                st.success("Alle E-Mails wurden erfolgreich versandt!")
