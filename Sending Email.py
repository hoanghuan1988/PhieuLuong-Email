import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd

def send_email(receiver_email, subject, body, attachment_path):
    #sender_email = "notice.hn@maruichisunsteel.com"
    sender_email = "maruichi.sunsco.hn@gmail.com"
    #sender_password = "&65FINtp"  # app password
    sender_password = "osgo smmj osxp vhyk"  # app password

    if not sender_email or not sender_password:
        print("Sender email credentials are missing. Please set the environment variables.")
        return False

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Add body
    msg.attach(MIMEText(body, 'plain'))

    # Attach file
    if attachment_path:
        print(f"Attempting to attach file: {attachment_path}")
        if os.path.exists(attachment_path):
            filename = os.path.basename(attachment_path)
            try:
                with open(attachment_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)

                    # Ensure filename is properly set
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{filename}"'
                    )
                    msg.attach(part)
                    print(f"Attachment added successfully: {filename}")
            except Exception as e:
                print(f"Error reading attachment: {e}")
                return False
        else:
            print(f"Attachment file not found: {attachment_path}")
            return False

    # Send email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print(f"Email sent to {receiver_email}")
        return True
    except Exception as e:
        print(f"Failed to send email to {receiver_email}. Error: {e}")
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

        if pd.isna(email) or email == 0:
            print(f"Skipping row {index}: Invalid or missing email.")
            continue

        subject = "PHIẾU LƯƠNG/THƯỞNG"
        body = f"Dear {name},\n\nPlease find attached your salary slip for this month."
        attachment_path = f"salary_slips/{name}_Salary_Slip.pdf"

        if not send_email(email, subject, body, attachment_path):
            print(f"Email not sent for row {index}.")

# Run the script
if __name__ == "__main__":
    salary_file_path = "Salary_data.xls"
    process_and_send_emails(salary_file_path)
