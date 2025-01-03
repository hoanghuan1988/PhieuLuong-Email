import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_log.log"),
        logging.StreamHandler()  # This keeps logs visible in the console as well
    ]
)

def send_email(receiver_email, subject, body, attachment_path):
    sender_email = "maruichi.sunsco.hn@gmail.com"
    sender_password = "osgo smmj osxp vhyk"

    if not sender_email or not sender_password:
        logging.error("Sender email credentials are missing. Please set the environment variables.")
        return False

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    

    if attachment_path and os.path.exists(attachment_path):
        try:
            filename = os.path.basename(attachment_path)
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                # Encode filename properly
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{filename}"'
            )
            msg.attach(part)
            logging.info(f"Attachment added: {filename}")
        except Exception as e:
            logging.error(f"Error reading attachment: {e}")
            return False
    else:
        logging.warning(f"Attachment not found: {attachment_path}")


    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        logging.info(f"Email sent to {receiver_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {receiver_email}. Error: {e}")
        return False


def process_and_send_emails(file_path):

    try:
        xls = pd.ExcelFile(file_path)
        data = xls.parse(sheet_name='Salary Sheet')
        data = data.iloc[10:]  # Skip irrelevant rows
    except Exception as e:
        print(f"Failed to read Excel file. Error: {e}")
        return


    for index, row in data.iterrows():
        name = row.get('Unnamed: 2')
        email = row.get('Unnamed: 48')
        staff_id = row.get('Unnamed: 1')

        if pd.isna(email) or not email:
            logging.warning(f"Skipping row {index}: Invalid or missing email.")
            continue

        subject = "Noreply_PHIẾU LƯƠNG/THƯỞNG_THÁNG 12 NĂM 2024 _ MARUICHI SUN STEEL (HÀ NỘI)"
        body = f"Xin Chào {name}!\n\nVui lòng tải file đính kèm để xem phiếu lương chi tiếp.\nTrân Trọng!\n\n 1-Yêu cầu bảo mật phiếu lương.\n 2-Mọi thắc mắc xin liên hệ lại phòng hành chính Ms.Giang hoặc Ms.Thùy!\n\n Chú ý:Đây là email tự động. Vui lòng không trả lời!"
        attachment_path = os.path.join("salary_slips", f"{staff_id}_Salary_Slip.pdf")

        if not send_email(email, subject, body, attachment_path):
            logging.error(f"Email not sent for row {index}.")

if __name__ == "__main__":
    salary_file_path = "Salary_data.xls"
    process_and_send_emails(salary_file_path)
