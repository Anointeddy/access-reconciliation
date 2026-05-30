import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

root = tk.Tk()
root.title("Access Reconciliation Tool")
root.geometry("600x500")
root.configure(bg="#1e1e2e")
root.resizable(False, False)

ccure_path   = tk.StringVar(value="No file selected")
genetec_path = tk.StringVar(value="No file selected")

def select_ccure():
    path = filedialog.askopenfilename(
        title="Select C-CURE File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        ccure_path.set(path)

def select_genetec():
    path = filedialog.askopenfilename(
        title="Select Genetec File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        genetec_path.set(path)

def run_reconciliation():
    if ccure_path.get() == "No file selected":
        messagebox.showerror("Error", "Please select the C-CURE file!")
        return
    if genetec_path.get() == "No file selected":
        messagebox.showerror("Error", "Please select the Genetec file!")
        return

    try:
        status_label.config(text="Running reconciliation...", fg="#f9e2af")
        root.update()

        ccure_df   = pd.read_excel(ccure_path.get())
        genetec_df = pd.read_excel(genetec_path.get())

        ccure_df.columns   = ccure_df.columns.str.strip().str.lower().str.replace(" ", "_")
        genetec_df.columns = genetec_df.columns.str.strip().str.lower().str.replace(" ", "_")

        ccure_cards   = set(ccure_df["card_number"].dropna().astype(str))
        genetec_cards = set(genetec_df["card_number"].dropna().astype(str))

        only_in_ccure   = ccure_cards - genetec_cards
        only_in_genetec = genetec_cards - ccure_cards
        in_both         = ccure_cards & genetec_cards

        ccure_only_df   = ccure_df[ccure_df["card_number"].astype(str).isin(only_in_ccure)]
        genetec_only_df = genetec_df[genetec_df["card_number"].astype(str).isin(only_in_genetec)]
        both_df         = ccure_df[ccure_df["card_number"].astype(str).isin(in_both)]

        os.makedirs("output", exist_ok=True)
        timestamp   = datetime.now().strftime("%Y-%m-%d") + "_" + datetime.now().strftime("%H-%M")
        output_file = "output/reconciliation_" + timestamp + ".xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            ccure_only_df.to_excel(writer,   sheet_name="Only_in_CCURE",   index=False)
            genetec_only_df.to_excel(writer, sheet_name="Only_in_Genetec", index=False)
            both_df.to_excel(writer,         sheet_name="In_Both",         index=False)

        result_text = (
            "RECONCILIATION COMPLETE\n\n"
            "Only in C-CURE   : " + str(len(only_in_ccure)) + "\n"
            "Only in Genetec  : " + str(len(only_in_genetec)) + "\n"
            "In Both          : " + str(len(in_both)) + "\n\n"
            "Output saved to:\n" + output_file
        )

        result_label.config(text=result_text, fg="#a6e3a1")
        status_label.config(text="Done!", fg="#a6e3a1")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        status_label.config(text="Error occurred!", fg="#f38ba8")

tk.Label(root, text="Access Reconciliation Tool",
         font=("Arial", 18, "bold"), bg="#1e1e2e", fg="#cdd6f4").pack(pady=20)

tk.Label(root, text="C-CURE File:",
         font=("Arial", 11), bg="#1e1e2e", fg="#cdd6f4").pack(anchor="w", padx=40)
tk.Label(root, textvariable=ccure_path,
         font=("Arial", 9), bg="#1e1e2e", fg="#89b4fa").pack(anchor="w", padx=40)
tk.Button(root, text="Select C-CURE File", command=select_ccure,
          bg="#89b4fa", fg="#1e1e2e", font=("Arial", 10, "bold"),
          padx=10, pady=5).pack(anchor="w", padx=40, pady=5)

tk.Label(root, text="Genetec File:",
         font=("Arial", 11), bg="#1e1e2e", fg="#cdd6f4").pack(anchor="w", padx=40)
tk.Label(root, textvariable=genetec_path,
         font=("Arial", 9), bg="#1e1e2e", fg="#89b4fa").pack(anchor="w", padx=40)
tk.Button(root, text="Select Genetec File", command=select_genetec,
          bg="#89b4fa", fg="#1e1e2e", font=("Arial", 10, "bold"),
          padx=10, pady=5).pack(anchor="w", padx=40, pady=5)

tk.Button(root, text="RUN RECONCILIATION", command=run_reconciliation,
          bg="#a6e3a1", fg="#1e1e2e", font=("Arial", 13, "bold"),
          padx=20, pady=10).pack(pady=20)

status_label = tk.Label(root, text="Ready",
                         font=("Arial", 10), bg="#1e1e2e", fg="#cdd6f4")
status_label.pack()

result_label = tk.Label(root, text="",
                         font=("Arial", 10), bg="#1e1e2e", fg="#a6e3a1",
                         justify="left")
result_label.pack(pady=10, padx=40, anchor="w")

root.mainloop()