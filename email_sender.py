import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import subprocess

def verify_openpyxl_installation():
    """
    Verify if openpyxl is installed in the Streamlit environment.
    """
    result = subprocess.run(["pip", "show", "openpyxl"], capture_output=True, text=True)
    if result.returncode == 0:
        print("openpyxl is installed in the Streamlit environment.")
        print(result.stdout)
    else:
        print("openpyxl is not installed in the Streamlit environment.")

def check_streamlit_permissions():
    """
    Check if the Streamlit environment has necessary permissions for package installations.
    """
    # You can add specific checks for permissions here if needed
    print("Streamlit environment permissions checked.")

def run_script_outside_streamlit():
    """
    Run the script directly using Python outside Streamlit environment.
    """
    # Replace 'your_script.py' with the name of your script
    result = subprocess.run(["python", "your_script.py"], capture_output=True, text=True)
    if result.returncode == 0:
        print("Script ran successfully outside Streamlit environment.")
        print(result.stdout)
    else:
        print("Error running script outside Streamlit environment.")
        print(result.stderr)

def update_streamlit_environment():
    """
    Update Streamlit to the latest version.
    """
    result = subprocess.run(["pip", "install", "--upgrade", "streamlit"], capture_output=True, text=True)
    if result.returncode == 0:
        print("Streamlit environment updated successfully.")
        print(result.stdout)
    else:
        print("Error updating Streamlit environment.")
        print(result.stderr)

# Call functions as needed
verify_openpyxl_installation()
check_streamlit_permissions()
run_script_outside_streamlit()
update_streamlit_environment()
# Install openpyxl if not already installed
subprocess.run(["pip", "install", "openpyxl"])

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

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
