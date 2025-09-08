import os
import pandas as pd
from datetime import datetime
import secrets
import string

# optional imports for encryption - imported inside functions to avoid platform errors
import io

def generate_password(length=14):
    alphabet = string.ascii_letters + string.digits  # alphanumeric only
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def encrypt_excel(input_path, desired_output_path, password):
    # 1) Try Windows COM (requires pywin32 and real Excel installed; Windows only)
    try:
        import pythoncom
        import win32com.client as win32  # requires pywin32
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        # FileFormat=51 -> xlOpenXMLWorkbook (.xlsx)
        out_abs = os.path.abspath(desired_output_path)
        wb.SaveAs(out_abs, FileFormat=51, Password=password)
        wb.Close(SaveChanges=False)
        excel.Application.Quit()
        return out_abs
    except Exception as e_com:
        pass

    # 2) Try msoffcrypto (best-effort)
    try:
        import msoffcrypto
        with open(input_path, "rb") as f_in:
            office_file = msoffcrypto.OfficeFile(f_in)
            out_abs = os.path.abspath(desired_output_path)
            with open(out_abs, "wb") as f_out:
                try:
                    office_file.encrypt(f_out, password)
                except TypeError:
                    office_file.encrypt(f_out, password=password)
        return out_abs
    except Exception:
        pass

    # 3) Fallback: AES encrypt the file bytes with pyAesCrypt (creates .aes)
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
        raise RuntimeError(f"Encryption failed with all methods: {e_aes}")

def process_file(input_file, top_n_cards=20, top_n_cashiers=20,
                 output_folder="TopTransactionsPerMonth"):
    # Auto-generate password
    password = generate_password()

    # Prepare folders
    os.makedirs(output_folder, exist_ok=True)
    password_log_folder = os.path.join(output_folder, "passwordlogs")
    os.makedirs(password_log_folder, exist_ok=True)
    log_file = os.path.join(password_log_folder, "password_log.txt")

    # Read input (force card_no as string)
    df = pd.read_excel(input_file, dtype={"card_no": str})
    df["TransactionDateTime"] = pd.to_datetime(df["TransactionDateTime"])
    df["YearMonth"] = df["TransactionDateTime"].dt.to_period("M")

    has_trans_total = "trans_total" in df.columns

    for month, month_data in df.groupby("YearMonth"):
        month_data = month_data.sort_values("TransactionDateTime").copy()
        month_data["card_no"] = month_data["card_no"].astype(str)

        output_file = os.path.join(output_folder, f"top_transaction_{month}.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            raw_data = month_data.drop(columns=["YearMonth"])
            raw_data.to_excel(writer, sheet_name="RawData", index=False)

            # Top Cards
            top_cards = month_data["card_no"].value_counts().head(top_n_cards).index
            card_summaries = []
            for card in top_cards:
                card_data = month_data[month_data["card_no"] == card].sort_values("TransactionDateTime")
                day_counts = card_data["TransactionDateTime"].dt.date.value_counts()
                peak_day = day_counts.idxmax() if not day_counts.empty else None
                peak_count = day_counts.max() if not day_counts.empty else 0
                low_day = day_counts.idxmin() if not day_counts.empty else None
                low_count = day_counts.min() if not day_counts.empty else 0

                summary = {
                    "Card Number": str(card),
                    "Month": str(month),
                    "Total Transactions": len(card_data),
                    "First Transaction": card_data["TransactionDateTime"].min(),
                    "Last Transaction": card_data["TransactionDateTime"].max(),
                    "Distinct Branches": card_data["branch_code"].nunique(),
                    "Branch List": ", ".join(card_data["branch_code"].unique().astype(str)),
                    "Distinct Cashiers": card_data["cashier"].nunique(),
                    "Cashier List": ", ".join(card_data["cashier"].unique().astype(str)),
                    "Distinct Registers": card_data["register_no"].nunique(),
                    "Register List": ", ".join(card_data["register_no"].unique().astype(str)),
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = card_data["trans_total"].sum()
                card_summaries.append(summary)
            pd.DataFrame(card_summaries).to_excel(writer, sheet_name="TopCards", index=False)

            # Top Cashiers
            top_cashiers = month_data["cashier"].value_counts().head(top_n_cashiers).index
            cashier_summaries = []
            for cashier in top_cashiers:
                cashier_data = month_data[month_data["cashier"] == cashier].sort_values("TransactionDateTime")
                day_counts = cashier_data["TransactionDateTime"].dt.date.value_counts()
                peak_day = day_counts.idxmax() if not day_counts.empty else None
                peak_count = day_counts.max() if not day_counts.empty else 0
                low_day = day_counts.idxmin() if not day_counts.empty else None
                low_count = day_counts.min() if not day_counts.empty else 0

                summary = {
                    "Cashier": cashier,
                    "Month": str(month),
                    "Total Transactions": len(cashier_data),
                    "First Transaction": cashier_data["TransactionDateTime"].min(),
                    "Last Transaction": cashier_data["TransactionDateTime"].max(),
                    "Distinct Branches": cashier_data["branch_code"].nunique(),
                    "Branch List": ", ".join(cashier_data["branch_code"].unique().astype(str)),
                    "Distinct Cards Handled": cashier_data["card_no"].nunique(),
                    "Card List": ", ".join(cashier_data["card_no"].unique().astype(str)),
                    "Distinct Registers": cashier_data["register_no"].nunique(),
                    "Register List": ", ".join(cashier_data["register_no"].unique().astype(str)),
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = cashier_data["trans_total"].sum()
                cashier_summaries.append(summary)
            pd.DataFrame(cashier_summaries).to_excel(writer, sheet_name="TopCashiers", index=False)

        encrypted_file_target = output_file.replace(".xlsx", "_encrypted.xlsx")
        try:
            enc_path = encrypt_excel(output_file, encrypted_file_target, password)
            last_encrypted_file = enc_path
        except Exception as e:
            raise RuntimeError(f"Failed to encrypt '{output_file}': {e}")

        if os.path.exists(output_file):
            try:
                os.remove(output_file)
            except Exception:
                pass

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file, "a", encoding="utf-8") as log:
            log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | Output: {os.path.basename(enc_path)} | Password: {password}\n")

        print(f"Saved and encrypted {enc_path}")

    print(f"\nAll monthly reports created in '{output_folder}' folder.")

    return output_folder, last_encrypted_file, password
