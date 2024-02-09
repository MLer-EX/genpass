# import
import tkinter as tk
from tkinter import messagebox
import random
from openpyxl import Workbook
import string


# Function to generate a secure password
def generate_password(length=12):
    special_chars = "()_-@"
    password_chars = (
            [random.choice(string.ascii_uppercase)] +
            [random.choice(string.ascii_lowercase)] +
            [random.choice(string.digits)] +
            [random.choice(special_chars)] +
            [random.choice(string.ascii_letters + string.digits + special_chars) for _ in range(length - 4)]
    )
    random.shuffle(password_chars)
    return ''.join(password_chars)


# Function to save username and password to an Excel file
def save_to_excel(username, password, filename="user_passwords.xlsx"):
    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Username", "Password"])
        ws.append([username, password])
        wb.save(filename)
        messagebox.showinfo("Success", "Username and password saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", "Failed to save to Excel.\n" + str(e))


# GUI Application
class PasswordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Password Generator")

        # Username
        self.username_label = tk.Label(root, text="Username:")
        self.username_label.grid(row=0, column=0, padx=10, pady=10)
        self.username_entry = tk.Entry(root)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        # Password
        self.password_label = tk.Label(root, text="Generated Password:")
        self.password_label.grid(row=1, column=0, padx=10, pady=10)
        self.password_entry = tk.Entry(root, state="readonly")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        # Buttons
        self.generate_button = tk.Button(root, text="Generate Password", command=self.generate_password)
        self.generate_button.grid(row=2, column=0, padx=10, pady=10)
        self.save_button = tk.Button(root, text="Save to Excel", command=self.save_password)
        self.save_button.grid(row=2, column=1, padx=10, pady=10)

    def generate_password(self):
        password = generate_password()
        self.password_entry.config(state="normal")
        self.password_entry.delete(0, tk.END)
        self.password_entry.insert(0, password)
        self.password_entry.config(state="readonly")

    def save_password(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        if username and password:
            save_to_excel(username, password)
        else:
            messagebox.showwarning("Warning", "Please generate a password first.")


if __name__ == "__main__":
    root = tk.Tk()
    app = PasswordApp(root)
    root.mainloop()
