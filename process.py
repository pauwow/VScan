import os
import pandas as pd

def process_file(input_file, top_n_cards=20, top_n_cashiers=20, output_folder="TopTransactionsPerMonth"):
    # Create output folder
    os.makedirs(output_folder, exist_ok=True)

    # Read Excel file, ensure card_no is string
    df = pd.read_excel(input_file, dtype={"card_no": str})
    df["TransactionDateTime"] = pd.to_datetime(df["TransactionDateTime"])

    # Extract Year-Month
    df["YearMonth"] = df["TransactionDateTime"].dt.to_period("M")

    # Check if trans_total column exists
    has_trans_total = "trans_total" in df.columns

    for month, month_data in df.groupby("YearMonth"):
        month_data = month_data.sort_values("TransactionDateTime")
        month_data["card_no"] = month_data["card_no"].astype(str)

        output_file = os.path.join(output_folder, f"top_transaction_{month}.xlsx")

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # Raw Data (exclude YearMonth)
            raw_data = month_data.drop(columns=["YearMonth"])
            raw_data.to_excel(writer, sheet_name="RawData", index=False)

            # ---- Top Cards ----
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
                    "Distinct Branches (Count)": card_data["branch_code"].nunique(),
                    "Branch List": ", ".join(map(str, sorted(card_data["branch_code"].unique()))),
                    "Distinct Cashiers (Count)": card_data["cashier"].nunique(),
                    "Cashier List": ", ".join(map(str, sorted(card_data["cashier"].unique()))),
                    "Distinct Registers (Count)": card_data["register_no"].nunique(),
                    "Register List": ", ".join(map(str, sorted(card_data["register_no"].unique()))),
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = card_data["trans_total"].sum()

                card_summaries.append(summary)

            pd.DataFrame(card_summaries).to_excel(writer, sheet_name="TopCards", index=False)

            # ---- Top Cashiers ----
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
                    "Distinct Branches (Count)": cashier_data["branch_code"].nunique(),
                    "Branch List": ", ".join(map(str, sorted(cashier_data["branch_code"].unique()))),
                    "Distinct Cards Handled (Count)": cashier_data["card_no"].nunique(),
                    "Card List": ", ".join(map(str, sorted(cashier_data["card_no"].unique()))),
                    "Distinct Registers (Count)": cashier_data["register_no"].nunique(),
                    "Register List": ", ".join(map(str, sorted(cashier_data["register_no"].unique()))),
                    "Day with Most Transactions": f"{peak_day} ({peak_count})" if peak_day else "N/A",
                    "Day with Fewest Transactions": f"{low_day} ({low_count})" if low_day else "N/A",
                }
                if has_trans_total:
                    summary["Sum of Transaction Total"] = cashier_data["trans_total"].sum()

                cashier_summaries.append(summary)

            pd.DataFrame(cashier_summaries).to_excel(writer, sheet_name="TopCashiers", index=False)

        print(f"Saved {output_file}")

    print(f"\nAll monthly reports created in '{output_folder}' folder.")
