import os
import pandas as pd
from datetime import datetime
import secrets
import string
import io

# Encryption-related imports are done inside encrypt_excel to avoid platform/library issues
# (pywin32, msoffcrypto, pyAesCrypt may not be available on all systems)

def generate_password(length=14):
    """Return a cryptographically secure random alphanumeric password."""
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def encrypt_excel(input_path, desired_output_path, password):
    """
    Try to encrypt an Excel file. Returns absolute path to encrypted file.
    Tries (in order):
      1) Windows COM Excel (requires pywin32 & Excel on Windows)
      2) msoffcrypto OOXML encryption (cross-platform; uses OOXMLFile.encrypt)
      3) pyAesCrypt AES encrypt fallback (creates .aes file)
    """
    # 1) Try Windows COM (Excel built-in encryption) - Windows only
    try:
        import pythoncom
        import win32com.client as win32  # pywin32
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        out_abs = os.path.abspath(desired_output_path)
        # FileFormat=51 -> xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(out_abs, FileFormat=51, Password=password)
        wb.Close(SaveChanges=False)
        excel.Application.Quit()
        return out_abs
    except Exception:
        # ignore and try next method
        pass

    # 2) Try msoffcrypto OOXML encryption (recommended & cross-platform for .xlsx)
    try:
        # Use the OOXMLFile class specifically for encryption (see msoffcrypto docs)
        from msoffcrypto.format.ooxml import OOXMLFile
        out_abs = os.path.abspath(desired_output_path)
        with open(input_path, "rb") as f_in:
            ooxml = OOXMLFile(f_in)
            with open(out_abs, "wb") as f_out:
                # correct signature: OOXMLFile.encrypt(password, outfile)
                ooxml.encrypt(password, f_out)
        return out_abs
    except Exception:
        # ignore and try fallback
        pass

    # 3) Fallback: AES encrypt raw bytes using pyAesCrypt (creates .aes file)
    try:
        import pyAesCrypt
        bufferSize = 64 * 1024
        if not desired_output_path.lower().endswith(".aes"):
            out_aes = desired_output_path + ".aes"
        else:
            out_aes = desired_output_path
        pyAesCrypt.encryptFile(input_path, out_aes, password, bufferSize)
        return os.path.abspath(out_aes)
    except Exception as e_aes:
        # If all methods fail, propagate an explanatory error
        raise RuntimeError(f"Encryption failed with all methods: {e_aes}")

def process_file(input_file, top_n_cards=20, top_n_cashiers=20,
                 output_folder="TopTransactionsPerMonth", encrypt=True):
    """
    Process the input Excel and create monthly reports.
    If encrypt=True: encrypt each monthly .xlsx and remove the unencrypted copy.
    Returns: (output_folder, last_output_file, password_or_None)
    """

    # Generate password only if encryption is requested
    password = generate_password() if encrypt else None

    # Prepare folders
    os.makedirs(output_folder, exist_ok=True)
    password_log_folder = os.path.join(output_folder, "passwordlogs")
    os.makedirs(password_log_folder, exist_ok=True)
    log_file = os.path.join(password_log_folder, "password_log.txt")

    df = pd.read_excel(input_file, dtype={"card_no": str} if "card_no" in pd.read_excel.__code__.co_varnames else None)

    # Ensure TransactionDateTime exists and is parsed
    df["TransactionDateTime"] = pd.to_datetime(df["TransactionDateTime"])
    df["YearMonth"] = df["TransactionDateTime"].dt.to_period("M")

    has_trans_total = "trans_total" in df.columns

    last_encrypted_file = None

    # Process per month
    for month, month_data in df.groupby("YearMonth"):
        month_data = month_data.sort_values("TransactionDateTime").copy()
        # Ensure card_no remains as string
        if "card_no" in month_data.columns:
            month_data["card_no"] = month_data["card_no"].astype(str)

        # Prepare output filename for this month
        # month is a Period like 2025-08
        output_file = os.path.join(output_folder, f"top_transaction_{month}.xlsx")

        # Write RawData + TopCards + TopCashiers into Excel workbook
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # RawData (drop YearMonth for raw)
            raw_data = month_data.drop(columns=["YearMonth"])
            raw_data.to_excel(writer, sheet_name="RawData", index=False)

            # Top Cards
            if "card_no" in month_data.columns:
                top_cards = month_data["card_no"].value_counts().head(top_n_cards).index
            else:
                top_cards = []

            card_summaries = []
            for card in top_cards:
                card_data = month_data[month_data["card_no"] == card].sort_values("TransactionDateTime")
                day_counts = card_data["TransactionDateTime"].dt.date.value_counts()
                peak_day = day_counts.idxmax() if not day_counts.empty else None
                peak_count = int(day_counts.max()) if not day_counts.empty else 0
                low_day = day_counts.idxmin() if not day_counts.empty else None
                low_count = int(day_counts.min()) if not day_counts.empty else 0

                summary = {
                    "Card Number": str(card),
                    "Month": str(month),
                    "Total Transactions": len(card_data),
                    "First Transaction": card_data["TransactionDateTime"].min(),
                    "Last Transaction": card_data["TransactionDateTime"].max(),
                    "Distinct Branches": int(card_data["branch_code"].nunique()) if "branch_code" in card_data.columns else 0,
                    "Branch List": ", ".join(card_data["branch_code"].unique().astype(str)) if "branch_code" in card_data.columns else "",
                    "Distinct Cashiers": int(card_data["cashier"].nunique()) if "cashier" in card_data.columns else 0,
                    "Cashier List": ", ".join(card_data["cashier"].unique().astype(str)) if "cashier" in card_data.columns else "",
                    "Distinct Registers": int(card_data["register_no"].nunique()) if "register_no" in card_data.columns else 0,
                    "Register List": ", ".join(card_data["register_no"].unique().astype(str)) if "register_no" in card_data.columns else "",
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = float(card_data["trans_total"].sum())
                card_summaries.append(summary)
            pd.DataFrame(card_summaries).to_excel(writer, sheet_name="TopCards", index=False)

            # Top Cashiers
            if "cashier" in month_data.columns:
                top_cashiers = month_data["cashier"].value_counts().head(top_n_cashiers).index
            else:
                top_cashiers = []

            cashier_summaries = []
            for cashier in top_cashiers:
                cashier_data = month_data[month_data["cashier"] == cashier].sort_values("TransactionDateTime")
                day_counts = cashier_data["TransactionDateTime"].dt.date.value_counts()
                peak_day = day_counts.idxmax() if not day_counts.empty else None
                peak_count = int(day_counts.max()) if not day_counts.empty else 0
                low_day = day_counts.idxmin() if not day_counts.empty else None
                low_count = int(day_counts.min()) if not day_counts.empty else 0

                summary = {
                    "Cashier": cashier,
                    "Month": str(month),
                    "Total Transactions": len(cashier_data),
                    "First Transaction": cashier_data["TransactionDateTime"].min(),
                    "Last Transaction": cashier_data["TransactionDateTime"].max(),
                    "Distinct Branches": int(cashier_data["branch_code"].nunique()) if "branch_code" in cashier_data.columns else 0,
                    "Branch List": ", ".join(cashier_data["branch_code"].unique().astype(str)) if "branch_code" in cashier_data.columns else "",
                    "Distinct Cards Handled": int(cashier_data["card_no"].nunique()) if "card_no" in cashier_data.columns else 0,
                    "Card List": ", ".join(cashier_data["card_no"].unique().astype(str)) if "card_no" in cashier_data.columns else "",
                    "Distinct Registers": int(cashier_data["register_no"].nunique()) if "register_no" in cashier_data.columns else 0,
                    "Register List": ", ".join(cashier_data["register_no"].unique().astype(str)) if "register_no" in cashier_data.columns else "",
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = float(cashier_data["trans_total"].sum())
                cashier_summaries.append(summary)
            pd.DataFrame(cashier_summaries).to_excel(writer, sheet_name="TopCashiers", index=False)

        # If encryption requested, encrypt the output file and remove the unencrypted copy
        if encrypt:
            encrypted_target = output_file.replace(".xlsx", "_encrypted.xlsx")
            try:
                enc_path = encrypt_excel(output_file, encrypted_target, password)
                last_encrypted_file = enc_path
            except Exception as e:
                # If encryption fails, make the failure explicit
                raise RuntimeError(f"Failed to encrypt '{output_file}': {e}")

            # Remove the original unencrypted file if it still exists
            if os.path.exists(output_file):
                try:
                    os.remove(output_file)
                except Exception:
                    # ignore removal errors
                    pass
            output_to_log = last_encrypted_file
        else:
            last_encrypted_file = output_file
            output_to_log = output_file

        # Write log entry (append)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file, "a", encoding="utf-8") as log:
            if encrypt:
                log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | Output: {os.path.basename(output_to_log)} | Encryption: ENABLED | Password: {password}\n")
            else:
                log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | Output: {os.path.basename(output_to_log)} | Encryption: DISABLED\n")

        print(f"Saved {'and encrypted ' if encrypt else ''}{output_to_log}")

    print(f"\nAll monthly reports created in '{output_folder}' folder.")
    return output_folder, last_encrypted_file, (password if encrypt else None)
