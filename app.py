import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import sys
import pandas as pd
from process import process_file, process_entity_details

# Keep global storage of values for search filtering
card_values_full = []
cashier_values_full = []

# ---------- Helpers ----------
def detect_available_fields(file_path):
    """Check available columns in the uploaded file and decide which inputs to show."""
    try:
        df = pd.read_excel(file_path, nrows=0)
        headers = set(col.lower() for col in df.columns)

        available = {
            "has_cards": "card_no" in headers,
            "has_cashiers": "cashier" in headers,
        }
        return available
    except Exception:
        return {"has_cards": False, "has_cashiers": False}

# ---------- Tab 1 ----------
def run_app():
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    available = detect_available_fields(file_path)

    try:
        top_cards = None
        top_cashiers = None

        if available["has_cards"]:
            try:
                top_cards = int(cards_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number for Top Cards.")
                return

        if available["has_cashiers"]:
            try:
                top_cashiers = int(cashiers_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number for Top Cashiers.")
                return

        # Process file
        output_folder, last_output_file, password = process_file(
            file_path,
            top_n_cards=top_cards,
            top_n_cashiers=top_cashiers,
            encrypt=encrypt_var.get(),
            separate_cards=separate_var.get(),
            include_intervals=interval_var.get()  # NEW: pass transaction interval choice
        )

        # Show report summary
        preview_text.config(state="normal")
        preview_text.delete(1.0, tk.END)
        preview_text.insert(tk.END, "=== Report Summary Preview ===\n\n")
        preview_text.insert(tk.END, f"Source File: {os.path.basename(file_path)}\n")

        if available["has_cards"]:
            preview_text.insert(tk.END, f"Top N Cards: {top_cards}\n")

        if available["has_cashiers"]:
            preview_text.insert(tk.END, f"Top N Cashiers: {top_cashiers}\n")

        preview_text.insert(tk.END, f"Output Folder: {output_folder}\n")
        preview_text.insert(tk.END, f"Last Generated File: {os.path.basename(last_output_file)}\n")

        if encrypt_var.get():
            preview_text.insert(tk.END, "Encryption: ENABLED\n")
            preview_text.insert(tk.END, "(Password saved in password_log.txt)\n")
        else:
            preview_text.insert(tk.END, "Encryption: DISABLED\n")

        if separate_var.get():
            preview_text.insert(tk.END, "Card Separation: ENABLED (8880 = Blue, 8881 = Yellow)\n")
        else:
            preview_text.insert(tk.END, "Card Separation: DISABLED\n")

        if interval_var.get():
            preview_text.insert(tk.END, "Transaction Intervals: INCLUDED\n")
        else:
            preview_text.insert(tk.END, "Transaction Intervals: EXCLUDED\n")

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
    if not file_path:
        return
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

    # Dynamically update input fields based on available columns
    available = detect_available_fields(file_path)

    if available["has_cards"]:
        cards_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        cards_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    else:
        cards_label.grid_remove()
        cards_entry.grid_remove()

    if available["has_cashiers"]:
        cashiers_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        cashiers_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
    else:
        cashiers_label.grid_remove()
        cashiers_entry.grid_remove()

def open_output_file(file_path):
    """Open the output Excel file directly."""
    try:
        if os.name == "nt":  # Windows
            os.startfile(file_path)
        elif os.name == "posix":  # macOS/Linux
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", file_path])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")

# ---------- Tab 2 ----------
def browse_file_tab2():
    global card_values_full, cashier_values_full
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    file_entry_tab2.delete(0, tk.END)
    file_entry_tab2.insert(0, file_path)

    try:
        df = pd.read_excel(file_path)
        if "card_no" in df.columns:
            df["card_no"] = df["card_no"].astype(str)

        card_values_full = sorted(df["card_no"].dropna().unique().tolist()) if "card_no" in df.columns else []
        cashier_values_full = sorted(df["cashier"].dropna().unique().tolist()) if "cashier" in df.columns else []

        card_dropdown["values"] = card_values_full
        cashier_dropdown["values"] = cashier_values_full

        card_dropdown.config(state="normal")
        cashier_dropdown.config(state="normal")
        card_var.set("")
        cashier_var.set("")
    except Exception as e:
        messagebox.showerror("Error", f"Could not load file: {e}")

def on_card_selected(event):
    if card_var.get():
        cashier_var.set("")
        cashier_dropdown.config(state="disabled")
    else:
        cashier_dropdown.config(state="normal")

def on_cashier_selected(event):
    if cashier_var.get():
        card_var.set("")
        card_dropdown.config(state="disabled")
    else:
        card_dropdown.config(state="normal")

def filter_card_list(event):
    value = card_var.get().lower()
    filtered = [v for v in card_values_full if value in str(v).lower()]
    card_dropdown["values"] = filtered
    if filtered:
        card_dropdown.after(50, lambda: card_dropdown.event_generate("<Down>"))

def filter_cashier_list(event):
    value = cashier_var.get().lower()
    filtered = [v for v in cashier_values_full if value in str(v).lower()]
    cashier_dropdown["values"] = filtered
    if filtered:
        cashier_dropdown.after(50, lambda: cashier_dropdown.event_generate("<Down>"))

def run_tab2():
    file_path = file_entry_tab2.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    chosen_card = card_var.get()
    chosen_cashier = cashier_var.get()

    if chosen_card and chosen_cashier:
        messagebox.showerror("Error", "Please choose only one: either a Card OR a Cashier.")
        return

    if not chosen_card and not chosen_cashier:
        messagebox.showerror("Error", "Please select a Card or Cashier.")
        return

    try:
        output_file = process_entity_details(file_path, card_no=chosen_card if chosen_card else None,
                                             cashier=chosen_cashier if chosen_cashier else None)
        messagebox.showinfo("Success", f"Details exported to:\n{output_file}")
        open_output_file(output_file)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------- Build UI ----------
root = tk.Tk()
root.title("VScan Report Generator")

# Notebook (tabs)
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# --- Tab 1: Main ---
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Generate Top Cards/Cashiers")

# File input
tk.Label(tab1, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
file_frame = tk.Frame(tab1)
file_frame.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="we")
tab1.grid_columnconfigure(1, weight=1)

file_entry = tk.Entry(file_frame)
file_entry.pack(side="left", fill="x", expand=True)
browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side="left", padx=(5, 0))

# Top N Cards
cards_label = tk.Label(tab1, text="Top N Cards:")
cards_entry = tk.Entry(tab1, width=10)
cards_entry.insert(0, "20")

# Top N Cashiers
cashiers_label = tk.Label(tab1, text="Top N Cashiers:")
cashiers_entry = tk.Entry(tab1, width=10)
cashiers_entry.insert(0, "20")

cards_label.grid_remove()
cards_entry.grid_remove()
cashiers_label.grid_remove()
cashiers_entry.grid_remove()

# Configuration label
tk.Label(tab1, text="Configuration:", font=("Arial", 10, "bold")).grid(
    row=3, column=0, padx=5, pady=(10, 0), sticky="w"
)

options_frame = tk.Frame(tab1)
options_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="w")

encrypt_var = tk.BooleanVar()
separate_var = tk.BooleanVar()
interval_var = tk.BooleanVar()  # NEW: transaction intervals checkbox

encrypt_checkbox = tk.Checkbutton(options_frame, text="Encrypt Output File", variable=encrypt_var)
encrypt_checkbox.pack(side="left", padx=(0, 15))

separate_checkbox = tk.Checkbutton(options_frame, text="Separate Card/Cashier", variable=separate_var)
separate_checkbox.pack(side="left", padx=(0, 15))

interval_checkbox = tk.Checkbutton(options_frame, text="Include Transaction Intervals", variable=interval_var)
interval_checkbox.pack(side="left")

run_button = tk.Button(tab1, text="Generate Report", command=run_app, bg="green", fg="white")
run_button.grid(row=5, column=0, columnspan=3, pady=10)

tk.Label(tab1, text="Report Summary Preview:", font=("Arial", 10, "bold")).grid(
    row=6, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 0)
)
preview_text = tk.Text(tab1, width=100, height=10, wrap="word", state="disabled", bg="#f9f9f9")
preview_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

open_button = tk.Button(tab1, text="Open Output File", state="disabled")
open_button.grid(row=8, column=0, columnspan=3, pady=5)

# --- Tab 2: Card/Cashier Details ---
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Extract Card/Cashier Details")

tk.Label(tab2, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
file_frame_tab2 = tk.Frame(tab2)
file_frame_tab2.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="we")
tab2.grid_columnconfigure(1, weight=1)

file_entry_tab2 = tk.Entry(file_frame_tab2)
file_entry_tab2.pack(side="left", fill="x", expand=True)
browse_button_tab2 = tk.Button(file_frame_tab2, text="Browse", command=browse_file_tab2)
browse_button_tab2.pack(side="left", padx=(5, 0))

# Dropdowns
tk.Label(tab2, text="Select Card No:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
card_var = tk.StringVar()
card_dropdown = ttk.Combobox(tab2, textvariable=card_var, state="normal")
card_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="we")
card_dropdown.bind("<KeyRelease>", filter_card_list)

tk.Label(tab2, text="Select Cashier:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
cashier_var = tk.StringVar()
cashier_dropdown = ttk.Combobox(tab2, textvariable=cashier_var, state="normal")
cashier_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky="we")
cashier_dropdown.bind("<KeyRelease>", filter_cashier_list)

card_dropdown.bind("<<ComboboxSelected>>", on_card_selected)
cashier_dropdown.bind("<<ComboboxSelected>>", on_cashier_selected)

run_button_tab2 = tk.Button(tab2, text="Export Details", command=run_tab2, bg="green", fg="white")
run_button_tab2.grid(row=3, column=0, columnspan=3, pady=10)

# --- Tab 3 (empty for now) ---
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="Tab 3")

root.mainloop()
