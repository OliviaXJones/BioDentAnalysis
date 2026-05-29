import pandas as pd
import openpyxl
import shutil
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from BioDent_Utils import (
    COLUMNS_TO_AVERAGE, clean_and_average_data, ask_not_found_action
)


# --- 1. MASTER SYNC ---

def sync_to_master(df_source, master_path):
    wb = openpyxl.load_workbook(master_path)
    sync_count = 0
    synced_ids = set()
    all_processed_ids = set(df_source['ID'].astype(str).str.strip())

    all_sheets = [
        s for s in wb.sheetnames
        if s.lower() not in ["summary", "notes", "calculations"]
    ]

    for _, row_data in df_source.iterrows():
        mouse_code  = str(row_data['ID']).strip()
        found_match = False

        while True:
            for sheet_name in all_sheets:
                ws = wb[sheet_name]
                for r in range(2, max(ws.max_row + 1, 50)):
                    cell_val   = ws.cell(row=r, column=1).value
                    clean_cell = str(cell_val).strip().upper() if cell_val else ""
                    if clean_cell == mouse_code.strip().upper():
                        for j, col_name in enumerate(COLUMNS_TO_AVERAGE):
                            ws.cell(row=r, column=2 + j).value = row_data[col_name]
                        sync_count += 1
                        synced_ids.add(mouse_code)
                        found_match = True
                        break
                if found_match:
                    break

            if found_match:
                break

            new_name = ask_not_found_action(mouse_code)
            if new_name:
                mouse_code = new_name
            else:
                break

    wb.save(master_path)
    print(f"Synced {sync_count} subjects to master.")

    unmatched = all_processed_ids - synced_ids
    if unmatched:
        print("The following subjects were not found in the master list:")
        for m in sorted(unmatched):
            print(f"  - {m}")
    else:
        print("All processed subjects were successfully matched in the master list.")


# --- 2. CSV GENERATION ---

def process_master_to_csv(master_path, csv_output_dir, study_name):
    print("Generating CSV summary...")
    xl = pd.ExcelFile(master_path)
    all_data = []

    for sheet in xl.sheet_names:
        if sheet.lower() in ["summary", "notes", "calculations"]:
            continue
        raw_df = pd.read_excel(master_path, sheet_name=sheet)
        subset = raw_df.iloc[:, 0:11].copy()
        subset.columns = ['Subject_ID'] + COLUMNS_TO_AVERAGE
        subset['Group'] = sheet
        all_data.append(subset)

    if not all_data:
        print("No data sheets found in master.")
        return

    final_df = pd.concat(all_data, ignore_index=True).dropna(subset=['Subject_ID'])
    final_df = final_df[['Subject_ID', 'Group'] + COLUMNS_TO_AVERAGE]

    csv_output_dir.mkdir(parents=True, exist_ok=True)
    out_path = csv_output_dir / f"{study_name}_Summary.csv"
    final_df.to_csv(out_path, index=False)
    print(f"Summary saved to {out_path}")


# --- 3. PIPELINE ENTRY POINT ---

def run_pipeline(config):
    study_name    = config["study_name"]
    raw_data_root = Path(config["raw_data_root"])
    master_path   = Path(config["master_file"])
    output_folder = Path(config["output_folder"])

    txt_files       = list(raw_data_root.glob("*.txt"))
    analyzed_folder = raw_data_root / "Analyzed_Files"
    csv_output_dir  = output_folder / f"{study_name}_CSV"

    print(f"Study:      {study_name}")
    print(f"Raw data:   {raw_data_root}")
    print(f"Master:     {master_path}")
    print(f"CSV output: {csv_output_dir}")

    if not txt_files:
        QMessageBox.critical(None, "No Data Files",
                             f"No .txt files found in:\n{raw_data_root}")
        return

    if not master_path.exists():
        QMessageBox.critical(None, "Master File Not Found",
                             f"Could not find:\n{master_path}")
        return

    df_source = clean_and_average_data(txt_files)
    if df_source.empty:
        QMessageBox.critical(None, "No Valid Data",
                             "No valid data was found in the .txt files.")
        return

    try:
        sync_to_master(df_source, master_path)

        analyzed_folder.mkdir(exist_ok=True)
        for file_path in txt_files:
            shutil.move(str(file_path), str(analyzed_folder / file_path.name))
        print(f"Processed files moved to {analyzed_folder}/")

        output_folder.mkdir(parents=True, exist_ok=True)
        process_master_to_csv(master_path, csv_output_dir, study_name)

        QMessageBox.information(
            None, "Pipeline Complete",
            f"All done!\n\n"
            f"CSV summary saved to:\n{csv_output_dir}\n\n"
            f"Raw files archived to:\n{analyzed_folder}"
        )

    except PermissionError:
        QMessageBox.critical(None, "File In Use",
                             "Could not save the master file.\n"
                             "Please close the Excel file and try again.")
    except Exception as e:
        QMessageBox.critical(None, "Unexpected Error", str(e))


