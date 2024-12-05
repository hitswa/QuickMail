import smtplib
import pandas as pd
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, ttk

# Send emails and update status
def send_bulk_emails():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    smtp_server = smtp_server_entry.get()
    smtp_port = smtp_port_entry.get()

    if not sender_email or not sender_password or not smtp_server or not smtp_port:
        messagebox.showerror("Error", "SMTP configuration is required!")
        return

    try:
        with smtplib.SMTP(smtp_server, int(smtp_port)) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            for index, row in email_data.iterrows():
                try:
                    message = f"Subject: {row['Subject']}\n\n{row['Body']}"
                    server.sendmail(sender_email, row['Email'], message)
                    email_data.at[index, 'Status'] = 'Done'
                except Exception as e:
                    email_data.at[index, 'Status'] = f"Failed: {e}"

                update_table_status(index, email_data.at[index, 'Status'])

            messagebox.showinfo("Success", "All emails processed!")
    except Exception as e:
        messagebox.showerror("Error", f"SMTP Error: {e}")

# Update table status
def update_table_status(index, status):
    tree.item(tree.get_children()[index], values=(
        email_data.at[index, 'Email'],
        email_data.at[index, 'Subject'],
        email_data.at[index, 'Body'],
        status
    ))

# Load Excel and display data
def load_excel():
    global email_data
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    try:
        email_data = pd.read_excel(file_path)
        if not {'Email', 'Subject', 'Body'}.issubset(email_data.columns):
            messagebox.showerror("Error", "Excel must contain Email, Subject, and Body columns!")
            return

        email_data['Status'] = 'Pending'
        populate_table()
    except Exception as e:
        messagebox.showerror("Error", f"Error loading Excel file: {e}")

# Populate table with email data
def populate_table():
    tree.delete(*tree.get_children())
    for index, row in email_data.iterrows():
        tree.insert("", "end", values=(row['Email'], row['Subject'], row['Body'], row['Status']))

# GUI setup
root = Tk()
root.title("Bulk Email Sender")
root.geometry("800x600")

# SMTP configuration
Label(root, text="Sender Email:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
sender_email_entry = Entry(root, width=30)
sender_email_entry.grid(row=0, column=1, padx=10, pady=5)

Label(root, text="Sender Password:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
sender_password_entry = Entry(root, show="*", width=30)
sender_password_entry.grid(row=1, column=1, padx=10, pady=5)

Label(root, text="SMTP Server:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
smtp_server_entry = Entry(root, width=30)
smtp_server_entry.grid(row=2, column=1, padx=10, pady=5)
smtp_server_entry.insert(0, "smtp.gmail.com")

Label(root, text="SMTP Port:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
smtp_port_entry = Entry(root, width=30)
smtp_port_entry.grid(row=3, column=1, padx=10, pady=5)
smtp_port_entry.insert(0, "587")

# Load Excel button
Button(root, text="Load Excel", command=load_excel).grid(row=4, column=0, columnspan=2, pady=10)

# Email table
columns = ("Email", "Subject", "Body", "Status")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=200, anchor="w")
tree.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Send Emails button
Button(root, text="Send Emails", command=send_bulk_emails).grid(row=6, column=0, columnspan=2, pady=10)

root.mainloop()
