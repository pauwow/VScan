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

    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        process_file(file_path, top_n_cards=top_cards, top_n_cashiers=top_cashiers)
        messagebox.showinfo("Success", "Processing complete! Check the TopTransactionsPerMonth folder.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

# Tkinter window
root = tk.Tk()
root.title("VScan Report Generator")

tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Top N Cards:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
cards_entry = tk.Entry(root, width=10)
cards_entry.insert(0, "20")
cards_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

tk.Label(root, text="Top N Cashiers:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
cashiers_entry = tk.Entry(root, width=10)
cashiers_entry.insert(0, "20")
cashiers_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

run_button = tk.Button(root, text="Run", command=run_app, bg="green", fg="white")
run_button.grid(row=3, column=0, columnspan=3, pady=10)

root.mainloop()
