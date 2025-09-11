import os
import pandas as pd
from datetime import datetime
import secrets
import string

def generate_password(length=14):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def encrypt_excel(input_path, desired_output_path, password):
    # 1) Windows COM method
    try:
        import pythoncom
        import win32com.client as win32
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        out_abs = os.path.abspath(desired_output_path)
        wb.SaveAs(out_abs, FileFormat=51, Password=password)
        wb.Close(SaveChanges=False)
        excel.Application.Quit()
        return out_abs
    except Exception:
        pass

    # 2) msoffcrypto method
    try:
        from msoffcrypto.format.ooxml import OOXMLFile
        out_abs = os.path.abspath(desired_output_path)
        with open(input_path, "rb") as f_in:
            ooxml = OOXMLFile(f_in)
            with open(out_abs, "wb") as f_out:
                ooxml.encrypt(password, f_out)
        return out_abs
    except Exception:
        pass

    # 3) AES fallback
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

def detect_schema(df):
    new_schema_cols = {
        "csn", "card_no", "transaction_type", "branch_name",
        "transaction_code", "scheme_id", "transaction_datetime",
        "post_date", "point_earned", "transaction_amount"
    }
    if new_schema_cols.issubset(set(df.columns)):
        return "new"
    return "old"

#SHELL
def process_new_schema(df, output_file, top_n_cards=20):
    df["transaction_datetime"] = pd.to_datetime(df["transaction_datetime"], errors="coerce")
    df["Month"] = df["transaction_datetime"].dt.to_period("M")

    summaries = []
    for (card, month), group in df.groupby(["card_no", "Month"]):
        day_counts = group["transaction_datetime"].dt.date.value_counts()
        peak_day = day_counts.idxmax() if not day_counts.empty else None
        peak_count = int(day_counts.max()) if not day_counts.empty else 0
        low_day = day_counts.idxmin() if not day_counts.empty else None
        low_count = int(day_counts.min()) if not day_counts.empty else 0

        summaries.append({
            "Card Number": str(card),
            "Month": str(month),
            "Total Transactions": len(group),
            "First Transaction": group["transaction_datetime"].min(),
            "Last Transaction": group["transaction_datetime"].max(),
            "Distinct Branches": group["branch_name"].nunique(),
            "Branch List": ", ".join(sorted(group["branch_name"].dropna().unique().astype(str))),
            "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
            "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
            "transaction_amount_total": group["transaction_amount"].sum(),
            "total_points": group["point_earned"].sum(),
        })

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RawData", index=False)
        pd.DataFrame(summaries).to_excel(writer, sheet_name="TopCards", index=False)

#MARKETS/UNIQLO/SACI
def process_old_schema(df, output_file, top_n_cards=20, top_n_cashiers=20):
    df["TransactionDateTime"] = pd.to_datetime(df["TransactionDateTime"], errors="coerce")
    df["YearMonth"] = df["TransactionDateTime"].dt.to_period("M")

    has_trans_total = "trans_total" in df.columns

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RawData", index=False)

        # Top Cards
        top_cards = df["card_no"].value_counts().head(top_n_cards).index if "card_no" in df.columns else []
        card_summaries = []
        for card in top_cards:
            card_data = df[df["card_no"] == card]
            day_counts = card_data["TransactionDateTime"].dt.date.value_counts()
            peak_day = day_counts.idxmax() if not day_counts.empty else None
            peak_count = int(day_counts.max()) if not day_counts.empty else 0
            low_day = day_counts.idxmin() if not day_counts.empty else None
            low_count = int(day_counts.min()) if not day_counts.empty else 0

            summary = {
                "Card Number": str(card),
                "Month": str(card_data["YearMonth"].iloc[0]),
                "Total Transactions": len(card_data),
                "First Transaction": card_data["TransactionDateTime"].min(),
                "Last Transaction": card_data["TransactionDateTime"].max(),
                "Distinct Branches": card_data["branch_code"].nunique() if "branch_code" in card_data.columns else 0,
                "Branch List": ", ".join(card_data["branch_code"].unique().astype(str)) if "branch_code" in card_data.columns else "",
                "Distinct Cashiers": card_data["cashier"].nunique() if "cashier" in card_data.columns else 0,
                "Cashier List": ", ".join(card_data["cashier"].unique().astype(str)) if "cashier" in card_data.columns else "",
                "Distinct Registers": card_data["register_no"].nunique() if "register_no" in card_data.columns else 0,
                "Register List": ", ".join(card_data["register_no"].unique().astype(str)) if "register_no" in card_data.columns else "",
                "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
            }
            if has_trans_total:
                summary["Sum of Transaction Total"] = float(card_data["trans_total"].sum())

            card_summaries.append(summary)
        pd.DataFrame(card_summaries).to_excel(writer, sheet_name="TopCards", index=False)

        # Top Cashiers
        top_cashiers = df["cashier"].value_counts().head(top_n_cashiers).index if "cashier" in df.columns else []
        cashier_summaries = []
        for cashier in top_cashiers:
            cashier_data = df[df["cashier"] == cashier]
            day_counts = cashier_data["TransactionDateTime"].dt.date.value_counts()
            peak_day = day_counts.idxmax() if not day_counts.empty else None
            peak_count = int(day_counts.max()) if not day_counts.empty else 0
            low_day = day_counts.idxmin() if not day_counts.empty else None
            low_count = int(day_counts.min()) if not day_counts.empty else 0

            summary = {
                "Cashier": cashier,
                "Month": str(cashier_data["YearMonth"].iloc[0]),
                "Total Transactions": len(cashier_data),
                "First Transaction": cashier_data["TransactionDateTime"].min(),
                "Last Transaction": cashier_data["TransactionDateTime"].max(),
                "Distinct Branches": cashier_data["branch_code"].nunique() if "branch_code" in cashier_data.columns else 0,
                "Branch List": ", ".join(cashier_data["branch_code"].unique().astype(str)) if "branch_code" in cashier_data.columns else "",
                "Distinct Cards Handled": cashier_data["card_no"].nunique() if "card_no" in cashier_data.columns else 0,
                "Card List": ", ".join(cashier_data["card_no"].unique().astype(str)) if "card_no" in cashier_data.columns else "",
                "Distinct Registers": cashier_data["register_no"].nunique() if "register_no" in cashier_data.columns else 0,
                "Register List": ", ".join(cashier_data["register_no"].unique().astype(str)) if "register_no" in cashier_data.columns else "",
                "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
            }
            if has_trans_total:
                summary["Sum of Transaction Total"] = float(cashier_data["trans_total"].sum())

            cashier_summaries.append(summary)
        pd.DataFrame(cashier_summaries).to_excel(writer, sheet_name="TopCashiers", index=False)

def process_file(input_file, top_n_cards=20, top_n_cashiers=20,
                 encrypt=True):
 
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_dir, "TopTransactionsPerMonth")
    os.makedirs(output_folder, exist_ok=True)

    password_log_folder = os.path.join(output_folder, "passwordlogs")
    os.makedirs(password_log_folder, exist_ok=True)
    log_file = os.path.join(password_log_folder, "password_log.txt")

    df = pd.read_excel(input_file)
    schema_type = detect_schema(df)

    month = datetime.now().strftime("%Y-%m")
    output_file = os.path.join(output_folder, f"top_transaction_{month}.xlsx")

    if schema_type == "new":
        process_new_schema(df, output_file, top_n_cards=top_n_cards)
    else:
        process_old_schema(df, output_file, top_n_cards=top_n_cards, top_n_cashiers=top_n_cashiers)

    final_file = output_file
    password = None
    if encrypt:
        password = generate_password()  
        encrypted_target = output_file.replace(".xlsx", "_encrypted.xlsx")
        try:
            final_file = encrypt_excel(output_file, encrypted_target, password)
            os.remove(output_file)
        except Exception as e:
            raise RuntimeError(f"Failed to encrypt '{output_file}': {e}")

    # Logging
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a", encoding="utf-8") as log:
        if encrypt:
            log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | Output: {os.path.basename(final_file)} | Encryption: ENABLED | Password: {password}\n")
        else:
            log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | Output: {os.path.basename(final_file)} | Encryption: DISABLED\n")

    print(f"Saved {'and encrypted ' if encrypt else ''}{final_file}")
    return output_folder, final_file, password
