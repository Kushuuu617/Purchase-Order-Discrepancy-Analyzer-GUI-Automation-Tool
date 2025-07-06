from automation import start_automation
import tkinter as tk
from tkinter import messagebox
import threading

def run_gui():
    root = tk.Tk()
    root.title("PO Discrepancy Checker")
    root.geometry("500x450")

    tk.Label(root, text="Username:").pack(pady=5)
    username_entry = tk.Entry(root, width=50)
    username_entry.pack()

    tk.Label(root, text="Password:").pack(pady=5)
    password_entry = tk.Entry(root, width=50, show="*")
    password_entry.pack()

    tk.Label(root, text="Status:").pack(pady=5)
    status_text = tk.Text(root, height=15, width=65, wrap=tk.WORD)
    status_text.pack()

    def on_start():
        username = username_entry.get()
        password = password_entry.get()
        if not username or not password:
            messagebox.showerror("Error", "Username and Password are required.")
            return
        threading.Thread(target=start_automation, args=(username, password, status_text), daemon=True).start()

    tk.Button(root, text="Start Automation", command=on_start).pack(pady=10)
    root.mainloop()
