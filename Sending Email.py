import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(to_email, subject, body):
    #sender_email = "notice.hn@maruichisunsteel.com"
    sender_email = "maruichi.sunsco.hn@gmail.com"
    #sender_password = "&65FINtp"  # app password
    sender_password = "osgo smmj osxp vhyk"  # app password

    # Cấu hình email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Kết nối và gửi email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
        #with smtplib.SMTP('mail.maruichisunsteel.com', 995) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {e}")

# Danh sách nhân viên
employees = [
    {"name": "Nguyen Thi Huong Giang", "email": "huonggiang@maruichisunsteel.com"},
    {"name": "Nguyen Thi Thuy", "email": "nguyenthuy@maruichisunsteel.com"}
]

for emp in employees:
    send_email(emp["email"], "PHIẾU LƯƠNG/THƯỞNG_THÁNG NĂM ", f"Xin Chào {emp['name']}.Đây là phiếu lương của bạn! Xin Vui lòng kiểm tra")