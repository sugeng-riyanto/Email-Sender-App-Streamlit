import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import subprocess
import sys

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# Install openpyxl if not already installed
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# Define allowed files
ALLOWED_EXTENSIONS = {'xlsx'}

# Utility function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def main():
    st.title('Email Sender App')

    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        email_list = df.to_dict(orient='records')

        for idx, entry in enumerate(email_list):
            subject = entry['Subject']
            grade = entry['Grade']
            va = entry['virtual_account']
            name = entry['customer_name']
            email = entry['customer_email']
            nominal = "{:,.2f}".format(entry['trx_amount'])
            expired_date = entry['expired_date']
            expired_time = entry['expired_time']
            description = entry['description']
            link = entry['link']

            message = f"""
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Email Template</title>
                    <style>
                        /* Paste the CSS styles here */
                        
                        /* ... */
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="header">
                            <h2>{subject}</h2>
                        </div>
                        <div class="content">
                            <p>Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)</p>
                            <p>Salam Hormat,</p>
                            <p>Kami hendak menyampaikan info mengenai:</p>
                            <ul>
                                <li><strong>Subject:</strong> {subject}</li>
                                <li><strong>Batas Tanggal Pembayaran:</strong> {expired_date}</li>
                                <li><strong>Sebesar:</strong> Rp. {nominal}</li>
                                <li><strong>Pembayaran via nomor virtual account (VA) BNI/Bank:</strong> {va}</li>
                            </ul>
                            <p>Terima kasih atas kerjasamanya.</p>
                            <p>Admin Sekolah</p>
                            <p><strong>Catatan:</strong><br>Mohon diabaikan jika sudah melakukan pembayaran.</p>
                            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:<br>
                                <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
                                <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
                            </p>
                        </div>
                        <div class="footer">
                            <p>Best Regards,<br>{your_name}</p>
                        </div>
                    </div>
                </body>
                </html>
            """

            msg = MIMEMultipart()
            msg['From'] = your_email
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'html'))

            try:
                server.sendmail(your_email, email, msg.as_string())
                st.success(f'Email {idx + 1} to {email} successfully sent!')
            except Exception as e:
                st.error(f'Failed to send email {idx + 1} to {email}: {e}')

        st.dataframe(df)

if __name__ == '__main__':
    main()
