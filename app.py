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

# Seitenkonfiguration
st.set_page_config(page_title="Lehrevaluations-Tool", layout="wide")
st.title("📧 E-Mail Versand Tool für Lehrevaluationen")

# ℹ️ Anleitung & Variablen-Übersicht
with st.expander("ℹ️ Anleitung & Verfügbare Variablen"):
    st.markdown("""
    **Anleitung:**
    1. Laden Sie Ihre Excel-Datei (`.xlsx`) hoch. Die Daten müssen in **'Tabelle1'** stehen.
    2. Optional: Laden Sie eine ZIP-Datei mit Unterordnern (Nachnamen) hoch, falls Anhänge benötigt werden.
    3. Konfigurieren Sie in der Seitenleiste Ihren SMTP-Server.
    4. Der E-Mail-Text kann live bearbeitet werden.
    """)

# Sidebar - SMTP Konfiguration
st.sidebar.header("SMTP Konfiguration")
sender_email = st.sidebar.text_input("Absender E-Mail")
bcc_email = st.sidebar.text_input("BCC E-Mail", sender_email)
smtp_server = st.sidebar.text_input("SMTP Server", "mail.uni-ulm.de")
smtp_port = st.sidebar.number_input("Port", value=587)
smtp_password = st.sidebar.text_input("Passwort", type="password")

# Upload Bereich
col1, col2 = st.columns(2)
uploaded_excel = col1.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])
uploaded_zip = col2.file_uploader("ZIP-Datei mit Unterordnern (optional)", type=["zip"])

# Editor
email_body_template = st.text_area("E-Mail Text bearbeiten", height=400, value="""{anrede_de},

Wir möchten Sie heute freundlich daran erinnern, dass die Lehrevaluation {veranstaltungen_de} noch aussteht. 

Die Unterlagen zur Evaluation sollten Sie bereits erhalten haben, aber wir haben diese nochmal hier angehängt. Bitte senden Sie uns dazu auch noch den ausgefüllten Rücklaufbogen zu.

Laden Sie dazu die ZIP-Datei vollständig herunter, entpacken diese und öffnen Sie die enthaltene HTML-Datei lokal (nicht aus einer Vorschau).

Vielen Dank für Ihre Mühe!

Mit den besten Grüßen

---

{anrede_en},

We would like to kindly remind you today that the course evaluation for {veranstaltungen_en} is still pending.

You should have already received the evaluation documents, but we have attached them here again. Please also send us the completed response form.

To do this, please download the ZIP file completely, unpack it, and open the contained HTML file locally (do not open it from a preview).

Thank you very much for your effort!

Best regards,

Tatjana Shevchik""")

# Versand-Logik
if st.button("🚀 E-Mails versenden"):
    if not (uploaded_excel and sender_email and smtp_password):
        st.error("Bitte laden Sie die Excel-Datei hoch und füllen Sie die SMTP-Einstellungen aus.")
    else:
        with st.spinner("Verarbeite Daten und versende E-Mails..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Zip entpacken, falls vorhanden
                if uploaded_zip:
                    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                        zip_ref.extractall(tmpdir)
                
                # Excel laden (Immer Tabelle1)
                df = pd.read_excel(uploaded_excel, sheet_name="Tabelle1", engine="openpyxl")
                df.columns = df.columns.str.strip()
                df['LVEName_Kennung_Kombiniert'] = df['LVEName'].astype(str) + " " + df['Kennung'].astype(str)
                grouped = df.groupby(['Emailempfänger', 'Anrede', 'Nachname'])
                
                # Server Login
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, smtp_password)
                
                for (email, anrede_excel, nachname), group in grouped:
                    veranstaltungen = group['LVEName_Kennung_Kombiniert'].dropna().unique().tolist()
                    
                    # Anrede-Übersetzung
                    anrede_en_basis = anrede_excel.strip()
                    replacements = {
                        "Sehr geehrte Frau Professorin": "Dear Prof.",
                        "Sehr geehrter Herr Professor": "Dear Prof.",
                        "Guten Tag Frau Professorin": "Dear Prof.",
                        "Guten Tag Herr Professor": "Dear Prof.",
                        "Guten Tag Professorin": "Dear Prof.",
                        "Guten Tag Professor": "Dear Prof.",
                        "Sehr geehrte Frau": "Dear Ms.",
                        "Sehr geehrter Herr": "Dear Mr.",
                        "Guten Tag Frau": "Dear Ms.",
                        "Guten Tag Herr": "Dear Mr.",
                        "Guten Tag": "Dear"
                    }
                    for deutsch, englisch in replacements.items():
                        anrede_en_basis = anrede_en_basis.replace(deutsch, englisch)
                    
                    # Kurs-Text Generierung
                    liste_v = "\n".join([f'- "{v}"' for v in veranstaltungen])
                    v_en = f"your following courses:\n\n{liste_v}"
                    
                    # E-Mail Bau
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = email
                    msg['Subject'] = Header(f"Erinnerung: Lehrevaluation ({nachname})", 'utf-8').encode()
                    
                    body = email_body_template.format(
                        anrede_de=f"{anrede_excel.strip()}",
                        anrede_en=f"{anrede_en_basis}",
                        veranstaltungen_en=v_en
                    )
                    msg.attach(MIMEText(body, 'plain', 'utf-8'))
                    
                    # Anhang finden und zippen, falls ZIP hochgeladen wurde
                    if uploaded_zip:
                        target_folder_path = os.path.join(tmpdir, str(nachname).strip())
                        if os.path.isdir(target_folder_path):
                            zip_file = shutil.make_archive(target_folder_path, 'zip', target_folder_path)
                            with open(zip_file, "rb") as f:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(f.read())
                                encoders.encode_base64(part)
                                part.add_header("Content-Disposition", f"attachment; filename={nachname}.zip")
                                msg.attach(part)
                    
                    server.sendmail(sender_email, [email, bcc_email], msg.as_string())
                
                server.quit()
                st.success("✅ Alle E-Mails wurden erfolgreich versandt!")
