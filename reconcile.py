import pandas as pd
import os
from datetime import datetime

# ── 1. Load the spreadsheets ──────────────────────────────────────────
ccure_df   = pd.read_excel("data/ccure.xlsx")
genetec_df = pd.read_excel("data/genetec.xlsx")

# ── 2. Standardize column names ───────────────────────────────────────
ccure_df.columns   = ccure_df.columns.str.strip().str.lower().str.replace(" ", "_")
genetec_df.columns = genetec_df.columns.str.strip().str.lower().str.replace(" ", "_")

# ── 3. Print column names to verify ──────────────────────────────────
print("C-CURE columns  :", ccure_df.columns.tolist())
print("Genetec columns :", genetec_df.columns.tolist())

# ── 4. Grab card numbers from each sheet ─────────────────────────────
ccure_cards   = set(ccure_df["card_number"].dropna().astype(str))
genetec_cards = set(genetec_df["card_number"].dropna().astype(str))

# ── 5. Run the three comparisons ─────────────────────────────────────
only_in_ccure   = ccure_cards - genetec_cards
only_in_genetec = genetec_cards - ccure_cards
in_both         = ccure_cards & genetec_cards

# ── 6. Filter dataframes for full row details ─────────────────────────
ccure_only_df   = ccure_df[ccure_df["card_number"].astype(str).isin(only_in_ccure)]
genetec_only_df = genetec_df[genetec_df["card_number"].astype(str).isin(only_in_genetec)]
both_df         = ccure_df[ccure_df["card_number"].astype(str).isin(in_both)]

# ── 7. Create output folder if it doesn't exist ──────────────────────
os.makedirs("output", exist_ok=True)

# ── 8. Timestamp the output file ─────────────────────────────────────
timestamp   = datetime.now().strftime("%Y-%m-%d_%H-%M")
output_file = f"output/reconciliation_{timestamp}.xlsx"

# ── 9. Write results into one Excel file with 3 tabs ─────────────────
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    ccure_only_df.to_excel(writer,   sheet_name="Only_in_CCURE",   index=False)
    genetec_only_df.to_excel(writer, sheet_name="Only_in_Genetec", index=False)
    both_df.to_excel(writer,         sheet_name="In_Both",         index=False)

# ── 10. Print summary ─────────────────────────────────────────────────
print("=" * 40)
print("   RECONCILIATION COMPLETE")
print("=" * 40)
print(f"  Only in C-CURE   : {len(only_in_ccure):,}")
print(f"  Only in Genetec  : {len(only_in_genetec):,}")
print(f"  In Both          : {len(in_both):,}")
print(f"\n  Output saved to  : {output_file}")
print("=" * 40)