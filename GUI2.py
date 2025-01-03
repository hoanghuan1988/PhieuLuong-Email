import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import smtplib
from email.mime.text import MIMEText
import threading

CONFIG_FILE = "email_config.json"

class SalarySlipMailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Salary Slip Mailer")
        self.root.geometry("500x400")

        # Menu Bar
        self.menu_bar = tk.Menu(root)
        self.root.config(menu=self.menu_bar)

        # Settings Menu
        self.settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.settings_menu.add_command(label="Email Settings", command=self.open_email_settings)
        self.menu_bar.add_cascade(label="Settings", menu=self.settings_menu)

        # Excel File Selection
        self.excel_label = tk.Label(root, text="Select Salary Data Excel File:")
        self.excel_label.pack(pady=5)
        self.excel_frame = tk.Frame(root)
        self.excel_frame.pack(pady=5)
        self.excel_entry = tk.Entry(self.excel_frame, width=40)
        self.excel_entry.pack(side=tk.LEFT, padx=5)
        self.excel_button = tk.Button(self.excel_frame, text="Browse", command=self.browse_excel)
        self.excel_button.pack(side=tk.LEFT)

        # Process Button
        self.process_button = tk.Button(root, text="Process Salary Slips", command=self.process_salary_slips)
        self.process_button.pack(pady=10)

        # Salary Slip Folder Selection
        self.folder_label = tk.Label(root, text="Select Salary Slip Folder:")
        self.folder_label.pack(pady=5)
        self.folder_frame = tk.Frame(root)
        self.folder_frame.pack(pady=5)
        self.folder_entry = tk.Entry(self.folder_frame, width=40)
        self.folder_entry.pack(side=tk.LEFT, padx=5)
        self.folder_button = tk.Button(self.folder_frame, text="Browse", command=self.browse_folder)
        self.folder_button.pack(side=tk.LEFT)

        # Start Button
        self.start_button = tk.Button(root, text="Send Emails", command=self.start_sending_emails)
        self.start_button.pack(pady=20)

        # Status Label
        self.status_label = tk.Label(root, text="Status: Waiting to start...", fg="blue")
        self.status_label.pack(pady=10)

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])
        self.excel_entry.delete(0, tk.END)
        self.excel_entry.insert(0, file_path)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, folder_path)

    def process_salary_slips(self):
        excel_file = self.excel_entry.get()
        folder_path = self.folder_entry.get()

        if not os.path.exists(excel_file):
            messagebox.showerror("Error", "Excel file does not exist.")
            return
        if not os.path.exists(folder_path):
            messagebox.showerror("Error", "Folder path does not exist.")
            return

        self.status_label.config(text="Status: Processing salary slips...", fg="orange")
        threading.Thread(target=self.run_process_slips, args=(excel_file, folder_path)).start()

    def run_process_slips(self, excel_file, folder_path):
        try:
            # Add your salary slip processing logic here
            # Replace with the actual function that generates salary slips
            process_salary_slips(excel_file, folder_path)  # Placeholder for actual logic
            self.status_label.config(text="Status: Salary slips processed successfully!", fg="green")
        except Exception as e:
            self.status_label.config(text="Status: Error occurred during processing.", fg="red")
            messagebox.showerror("Error", str(e))

    def start_sending_emails(self):
        excel_file = self.excel_entry.get()
        folder_path = self.folder_entry.get()

        if not os.path.exists(excel_file):
            messagebox.showerror("Error", "Excel file does not exist.")
            return
        if not os.path.exists(folder_path):
            messagebox.showerror("Error", "Folder path does not exist.")
            return

        self.status_label.config(text="Status: Sending emails...", fg="green")
        threading.Thread(target=self.run_email_process, args=(excel_file, folder_path)).start()

    def run_email_process(self, excel_file, folder_path):
        try:
            # Call your existing function here
            process_and_send_emails(excel_file)  # Replace with your email-sending logic
            self.status_label.config(text="Status: Emails sent successfully!", fg="green")
        except Exception as e:
            self.status_label.config(text="Status: Error occurred while sending emails.", fg="red")
            messagebox.showerror("Error", str(e))

    def open_email_settings(self):
        EmailSettingsWindow(self.root)


class EmailSettingsWindow:
    def __init__(self, master):
        self.top = tk.Toplevel(master)
        self.top.title("Email Settings")
        self.top.geometry("400x300")

        # Load existing settings
        self.settings = self.load_settings()

        # Input fields for email configuration
        self.sender_label = tk.Label(self.top, text="Sender Email:")
        self.sender_label.pack(pady=5)
        self.sender_entry = tk.Entry(self.top, width=30)
        self.sender_entry.pack(pady=5)
        self.sender_entry.insert(0, self.settings.get("sender_email", ""))

        self.password_label = tk.Label(self.top, text="Password:")
        self.password_label.pack(pady=5)
        self.password_entry = tk.Entry(self.top, width=30, show="*")
        self.password_entry.pack(pady=5)
        self.password_entry.insert(0, self.settings.get("password", ""))

        self.smtp_label = tk.Label(self.top, text="SMTP Server:")
        self.smtp_label.pack(pady=5)
        self.smtp_entry = tk.Entry(self.top, width=30)
        self.smtp_entry.pack(pady=5)
        self.smtp_entry.insert(0, self.settings.get("smtp_server", "smtp.gmail.com"))

        self.port_label = tk.Label(self.top, text="Port:")
        self.port_label.pack(pady=5)
        self.port_entry = tk.Entry(self.top, width=30)
        self.port_entry.pack(pady=5)
        self.port_entry.insert(0, self.settings.get("port", "587"))

        # Buttons
        self.save_button = tk.Button(self.top, text="Save", command=self.save_settings)
        self.save_button.pack(pady=10)

        self.test_button = tk.Button(self.top, text="Send Test Email", command=self.send_test_email)
        self.test_button.pack(pady=10)

    def load_settings(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        return {}

    def save_settings(self):
        self.settings = {
            "sender_email": self.sender_entry.get(),
            "password": self.password_entry.get(),
            "smtp_server": self.smtp_entry.get(),
            "port": self.port_entry.get(),
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.settings, f)
        messagebox.showinfo("Success", "Email settings saved successfully!")

    def send_test_email(self):
        try:
            sender_email = self.sender_entry.get()
            password = self.password_entry.get()
            smtp_server = self.smtp_entry.get()
            port = int(self.port_entry.get())

            with smtplib.SMTP(smtp_server, port) as server:
                server.starttls()
                server.login(sender_email, password)
                test_email = "test@example.com"  # Replace with a valid email for testing
                message = MIMEText("This is a test email.")
                message["From"] = sender_email
                message["To"] = test_email
                message["Subject"] = "Test Email"
                server.sendmail(sender_email, test_email, message.as_string())

            messagebox.showinfo("Success", "Test email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send test email. Error: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SalarySlipMailerApp(root)
    root.mainloop()
