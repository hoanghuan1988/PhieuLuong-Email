import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
#from email_sending_script import process_and_send_emails  # Import your existing function

class SalarySlipMailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Salary Slip Mailer")
        self.root.geometry("500x350")

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
            process_and_send_emails(excel_file)
            self.status_label.config(text="Status: Emails sent successfully!", fg="green")
        except Exception as e:
            self.status_label.config(text="Status: Error occurred while sending emails.", fg="red")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = SalarySlipMailerApp(root)
    root.mainloop()
