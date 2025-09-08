import tkinter as tk
from tkinter import filedialog, messagebox
from process import process_file

def run_app():
    file_path = file_entry.get()
    try:
        top_cards = int(cards_entry.get())
        top_cashiers = int(cashiers_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Please enter valid numbers for Top Cards and Top Cashiers.")
        return

    password = password_entry.get().strip()
    if not password:
        messagebox.showerror("Error", "Please enter a password.")
        return

    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        process_file(file_path, top_n_cards=top_cards, top_n_cashiers=top_cashiers, password=password)
        messagebox.showinfo("Success", "Processing complete! Check the TopTransactionsPerMonth folder.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def toggle_password():
    if show_password_var.get():
        password_entry.config(show="")  # show actual text
    else:
        password_entry.config(show="*")  # mask with *

# Tkinter window
root = tk.Tk()
root.title("VScan Report Generator")

# File input
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

# Top N Cards
tk.Label(root, text="Top N Cards:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
cards_entry = tk.Entry(root, width=10)
cards_entry.insert(0, "20")
cards_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

# Top N Cashiers
tk.Label(root, text="Top N Cashiers:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
cashiers_entry = tk.Entry(root, width=10)
cashiers_entry.insert(0, "20")
cashiers_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

# Password
tk.Label(root, text="Password:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
password_entry = tk.Entry(root, width=20, show="*")
password_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

# Show Password checkbox
show_password_var = tk.BooleanVar()
show_password_checkbox = tk.Checkbutton(root, text="Show Password", variable=show_password_var, command=toggle_password)
show_password_checkbox.grid(row=3, column=2, padx=5, pady=5, sticky="w")

# Run button
run_button = tk.Button(root, text="Run", command=run_app, bg="green", fg="white")
run_button.grid(row=4, column=0, columnspan=3, pady=10)

root.mainloop()
