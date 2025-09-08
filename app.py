import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import sys
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
        # Process file (auto-generates password inside process.py)
        output_folder, last_output_file, _ = process_file(
            file_path, top_n_cards=top_cards, top_n_cashiers=top_cashiers
        )

        # Show report summary (NO PASSWORD DISPLAYED)
        preview_text.config(state="normal")
        preview_text.delete(1.0, tk.END)
        preview_text.insert(tk.END, "=== Report Summary Preview ===\n\n")
        preview_text.insert(tk.END, f"Source File: {os.path.basename(file_path)}\n")
        preview_text.insert(tk.END, f"Top N Cards: {top_cards}\n")
        preview_text.insert(tk.END, f"Top N Cashiers: {top_cashiers}\n")
        preview_text.insert(tk.END, f"Output Folder: {output_folder}\n")
        preview_text.insert(tk.END, f"Last Generated File: {os.path.basename(last_output_file)}\n")
        preview_text.config(state="disabled")

        # Enable open button
        open_button.config(
            state="normal",
            text=f"Open {os.path.basename(last_output_file)}",
            command=lambda: open_output_file(last_output_file)
        )

        messagebox.showinfo("Success", "Processing complete! Check the Report Summary Preview below.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def open_output_file(file_path):
    """Open the output Excel file directly."""
    try:
        if os.name == "nt":  # Windows
            os.startfile(file_path)
        elif os.name == "posix":  # macOS/Linux
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", file_path])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")

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

# Run button
run_button = tk.Button(root, text="Run", command=run_app, bg="green", fg="white")
run_button.grid(row=3, column=0, columnspan=3, pady=10)

# Preview box
tk.Label(root, text="Report Summary Preview:").grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 0))
preview_text = tk.Text(root, width=100, height=10, wrap="word", state="disabled", bg="#f9f9f9")
preview_text.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

# Open file button (disabled by default)
open_button = tk.Button(root, text="Open Output File", state="disabled")
open_button.grid(row=6, column=0, columnspan=3, pady=5)

root.mainloop()
