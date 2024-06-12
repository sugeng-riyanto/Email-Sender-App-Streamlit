import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
from prettytable import PrettyTable
import os

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# Utility function to check allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def main():
    st.title('Email Sender App')

    menu = ["Home", "Send Reminder", "Announcement", "Invoice", "Send Proof of Payment"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Home":
        st.subheader("Home")
        st.write("Welcome to the Email Sender App!")

    elif choice == "Send Reminder":
        st.subheader("Send Reminder")
        handle_file_upload()

    elif choice == "Announcement":
        st.subheader("Announcement")
        handle_file_upload(announcement=True)

    elif choice == "Invoice":
        st.subheader("Invoice")
        handle_file_upload(invoice=True)

    elif choice == "Send Proof of Payment":
        st.subheader("Send Proof of Payment")
        handle_file_upload(proof_payment=True)

def handle_file_upload(announcement=False, invoice=False, proof_payment=False):
    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        email_list = df.to_dict(orient='records')
        st.dataframe(df)

        if st.button("Send Emails"):
            for idx, entry in enumerate(email_list):
                if announcement:
                    subject = entry['Subject']
                    name = entry['Nama_Siswa']
                    email = entry['Email']
                    description = entry['Description']
                    link = entry['Link']
                    message = f"""
                    Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span><br>
                    <p>Salam Hormat,</p>
                    <p>Kami hendak menyampaikan info mengenai:</p>
                    <ul>
                        <li><strong>Subject:</strong> {subject}</li>
                        <li><strong>Description:</strong> {description}</li>
                        <li><strong>Link:</strong> {link}</li>
                    </ul>
                    <p>Terima kasih atas kerjasamanya.</p>
                    <p>Admin Sekolah</p>
                    <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
                    <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
                    <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
                    """
                elif invoice:
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
                    Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
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
                    <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
                    <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
                    <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
                    """
                elif proof_payment:
                    subject = entry['Subject']
                    va = entry['virtual_account']
                    name = entry['customer_name']
                    email = entry['customer_email']
                    grade = entry['Grade']
                    sppbuljal = "{:,.2f}".format(entry['bulan_berjalan'])
                    ket1 = entry['Ket_1']
                    spplebih = "{:,.2f}".format(entry['SPP_30hari'])
                    ket2 = entry['Ket_2']
                    denda = "{:,.2f}".format(entry['Denda'])
                    ket3 = entry['Ket_3']
                    ket4 = entry['Ket_4']
                    total = "{:,.2f}".format(entry['Total'])
                    message = f"""
                    Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
                    <p>Salam Hormat,</p>
                    <p>Kami hendak menyampaikan info mengenai SPP:</p>
                    <ul>
                        <li><strong>SPP yang sedang berjalan:</strong> {sppbuljal} ({ket1})</li>
                        <li><strong>Denda:</strong> {denda} ({ket3})</li>
                        <li><strong>SPP bulan-bulan sebelumnya:</strong> {spplebih} ({ket2})</li>
                        <li><strong>Keterangan:</strong> {ket4}</li>
                        <li><strong>Total tagihan:</strong> {total}</li>
                    </ul>
                    <p>Terima kasih atas kerjasamanya.</p>
                    <p>Admin Sekolah</p>
                    <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
                    <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
                    <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
                    """
                else:
                    continue

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

if __name__ == '__main__':
    main()
