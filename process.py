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


# -------- Dynamic summarizer --------
def summarize_entities(df, entity_col, date_col="transaction_datetime", top_n=20):
    """Generic summarizer for cards or cashiers based on available columns."""
    summaries = []

    if entity_col not in df.columns or date_col not in df.columns:
        return pd.DataFrame()  # skip if essential columns are missing

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df["YearMonth"] = df[date_col].dt.to_period("M")

    top_entities = df[entity_col].value_counts().head(top_n).index
    for entity in top_entities:
        entity_data = df[df[entity_col] == entity]

        day_counts = entity_data[date_col].dt.date.value_counts()
        peak_day = day_counts.idxmax() if not day_counts.empty else None
        peak_count = int(day_counts.max()) if not day_counts.empty else 0
        low_day = day_counts.idxmin() if not day_counts.empty else None
        low_count = int(day_counts.min()) if not day_counts.empty else 0

        summary = {
            entity_col.capitalize(): str(entity),
            "Month": str(entity_data["YearMonth"].iloc[0]),
            "Total Transactions": len(entity_data),
            "First Transaction": entity_data[date_col].min(),
            "Last Transaction": entity_data[date_col].max(),
            "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
            "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
        }

        # Optional fields
        if "branch_code" in entity_data.columns:
            summary["Distinct Branches"] = entity_data["branch_code"].nunique()
            summary["Branch List"] = ", ".join(entity_data["branch_code"].astype(str).unique())
        elif "branch_name" in entity_data.columns:
            summary["Distinct Branches"] = entity_data["branch_name"].nunique()
            summary["Branch List"] = ", ".join(entity_data["branch_name"].astype(str).unique())

        if "cashier" in entity_data.columns and entity_col != "cashier":
            summary["Distinct Cashiers"] = entity_data["cashier"].nunique()
            summary["Cashier List"] = ", ".join(entity_data["cashier"].astype(str).unique())

        if "register_no" in entity_data.columns:
            summary["Distinct Registers"] = entity_data["register_no"].nunique()
            summary["Register List"] = ", ".join(entity_data["register_no"].astype(str).unique())

        if entity_col == "cashier" and "card_no" in entity_data.columns:
            summary["Distinct Cards"] = entity_data["card_no"].nunique()
            summary["Cards List"] = ", ".join(entity_data["card_no"].astype(str).unique())

        if "trans_total" in entity_data.columns:
            summary["Sum of Transaction Total"] = float(entity_data["trans_total"].sum())
        elif "transaction_amount" in entity_data.columns:
            summary["Sum of Transaction Total"] = float(entity_data["transaction_amount"].sum())

        if "point_earned" in entity_data.columns:
            summary["Total Points"] = float(entity_data["point_earned"].sum())

        summaries.append(summary)

    return pd.DataFrame(summaries)


def process_dynamic_schema(df, output_file, top_n_cards=20, top_n_cashiers=20):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Raw data
        df.to_excel(writer, sheet_name="RawData", index=False)

        # Decide which datetime column to use
        date_col = None
        if "transaction_datetime" in df.columns:
            date_col = "transaction_datetime"
        elif "TransactionDateTime" in df.columns:
            date_col = "TransactionDateTime"

        # Top Cards
        if "card_no" in df.columns and date_col:
            card_summary = summarize_entities(df, "card_no", date_col=date_col, top_n=top_n_cards)
            if not card_summary.empty:
                card_summary.to_excel(writer, sheet_name="TopCards", index=False)

        # Top Cashiers
        if "cashier" in df.columns and date_col:
            cashier_summary = summarize_entities(df, "cashier", date_col=date_col, top_n=top_n_cashiers)
            if not cashier_summary.empty:
                cashier_summary.to_excel(writer, sheet_name="TopCashiers", index=False)


# -------- Main entry point --------
def process_file(input_file, top_n_cards=20, top_n_cashiers=20, encrypt=True):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_dir, "TopTransactionsPerMonth")
    os.makedirs(output_folder, exist_ok=True)

    password_log_folder = os.path.join(output_folder, "passwordlogs")
    os.makedirs(password_log_folder, exist_ok=True)
    log_file = os.path.join(password_log_folder, "password_log.txt")

    df = pd.read_excel(input_file)

    # Determine yearmonth range from transaction date
    date_col = None
    if "transaction_datetime" in df.columns:
        date_col = "transaction_datetime"
    elif "TransactionDateTime" in df.columns:
        date_col = "TransactionDateTime"

    if date_col and date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        yearmonths = df[date_col].dt.to_period("M").dropna().unique()
        if len(yearmonths) > 0:
            start = str(min(yearmonths))
            end = str(max(yearmonths))
            if start == end:
                month_range = start
            else:
                month_range = f"{start}_to_{end}"
        else:
            month_range = datetime.now().strftime("%Y-%m")
    else:
        month_range = datetime.now().strftime("%Y-%m")

    output_file = os.path.join(output_folder, f"top_transaction_{month_range}.xlsx")

    process_dynamic_schema(df, output_file, top_n_cards, top_n_cashiers)

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
            log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | "
                      f"Output: {os.path.basename(final_file)} | "
                      f"Encryption: ENABLED | Password: {password}\n")
        else:
            log.write(f"[{timestamp}] Input: {os.path.basename(input_file)} | "
                      f"Output: {os.path.basename(final_file)} | "
                      f"Encryption: DISABLED\n")

    print(f"Saved {'and encrypted ' if encrypt else ''}{final_file}")
    return output_folder, final_file, password
