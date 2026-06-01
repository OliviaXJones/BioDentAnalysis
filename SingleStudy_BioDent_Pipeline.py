import re
import pandas as pd
import openpyxl
import shutil
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from BioDent_Utils import (
    COLUMNS_TO_AVERAGE, ManualSelectionDialog,
    ask_small_group_action, ask_new_id, ask_not_found_action
)


# --- 1. GROUP MAP HELPERS ---

def extract_prefix(mouse_id, group_map):
    """Return the longest group_map key that mouse_id starts with, or None."""
    code = str(mouse_id).strip().upper()
    for prefix in sorted(group_map.keys(), key=len, reverse=True):
        if code.startswith(prefix.upper()):
            return prefix
    return None


def get_group_name(mouse_id, group_map):
    """Return the group label for a mouse ID, or None if no prefix matched."""
    prefix = extract_prefix(mouse_id, group_map)
    if prefix is None:
        return None
    val = group_map[prefix]
    return val if isinstance(val, str) else val.get("group", "")


def deduce_sex(mouse_id, group_map):
    """
    Return 'Male', 'Female', or 'Unknown'.
    Priority: explicit sex in group_map entry → M/F found in ID after prefix.
    """
    prefix = extract_prefix(mouse_id, group_map)
    if prefix:
        val = group_map.get(prefix, "")
        if isinstance(val, dict):
            s = val.get("sex", "")
            if s == "Male":
                return "Male"
            if s == "Female":
                return "Female"
    # Scan the remainder of the ID for a sex indicator
    code      = str(mouse_id).strip().upper()
    remainder = code[len(prefix):] if prefix else code
    m = re.search(r'[_\-\.]([MF])[_\-\.\d]', remainder) \
        or re.search(r'([MF])\d+$', remainder)
    if m:
        return "Male" if m.group(1) == "M" else "Female"
    return "Unknown"


# --- 2. DATA CLEANING & AVERAGING ---

def clean_and_average_data(txt_file_paths, group_map):
    """
    Read, clean, and average BioDent .txt files.
    When group_map is non-empty, only rows whose ID starts with a known
    prefix are kept — freeform prefix matching replaces the FKBP5 regex fix.
    """
    all_data = []
    for path in txt_file_paths:
        try:
            temp_df = pd.read_csv(path, sep='\t', encoding='latin1')
            all_data.append(temp_df)
        except Exception as e:
            print(f"Skipping {path.name}: {e}")

    if not all_data:
        return pd.DataFrame()

    df = pd.concat(all_data, ignore_index=True)
    df.columns = [c.replace('µ', 'u') if isinstance(c, str) else c for c in df.columns]
    df = df.rename(columns={'Sample/Location': 'ID'})

    target_val_col = "1st Cycle Indentation Distance (ID 1st) - um"
    df = df.drop_duplicates(subset=['ID', 'Measurement #', target_val_col], keep='first')

    ignore_keywords = ["ignore", "do not use", "disregard", "don't", "ignor"]
    mask = df['Notes'].str.contains('|'.join(ignore_keywords), case=False, na=False)
    df_cleaned = df[~mask].copy()
    df_cleaned = df_cleaned.dropna(subset=['ID'])
    df_cleaned['ID'] = df_cleaned['ID'].astype(str)

    if group_map:
        df_cleaned = df_cleaned[
            df_cleaned['ID'].apply(lambda x: extract_prefix(x, group_map) is not None)
        ]

    if df_cleaned.empty:
        return pd.DataFrame()

    final_averages = []
    unique_ids = sorted(df_cleaned['ID'].unique())
    i = 0
    while i < len(unique_ids):
        mouse_id     = unique_ids[i]
        group        = df_cleaned[df_cleaned['ID'] == mouse_id]
        mouse_id_str = str(mouse_id).strip()

        if len(group) > 4:
            dlg = ManualSelectionDialog(mouse_id_str, group, target_val_col)
            dlg.exec_()
            indices = dlg.get_selected_indices()
            if indices:
                group = group.loc[indices]

        elif 0 < len(group) < 4:
            action = ask_small_group_action(mouse_id_str, len(group))
            if action == 'r':
                new_id = ask_new_id(mouse_id_str)
                if new_id:
                    new_id = re.sub(r"[.\-_]", "", new_id)
                    df_cleaned.loc[group.index, 'ID'] = new_id
                    unique_ids = sorted(df_cleaned['ID'].unique())
                    continue
            elif action == 's':
                i += 1
                continue

        for col in COLUMNS_TO_AVERAGE:
            group[col] = pd.to_numeric(group[col], errors='coerce')

        avg_values        = group[COLUMNS_TO_AVERAGE].mean()
        avg_values['ID']  = mouse_id_str
        final_averages.append(avg_values)
        i += 1

    return pd.DataFrame(final_averages)


# --- 3. MASTER SYNC ---

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


# --- 4. CSV GENERATION ---

def process_master_to_csv(master_path, csv_output_dir, study_name, group_map, sex="", age=""):
    print("Generating CSV summaries...")
    xl       = pd.ExcelFile(master_path)
    all_data = []

    for sheet in xl.sheet_names:
        if sheet.lower() in ["summary", "notes", "calculations"]:
            continue
        raw_df = pd.read_excel(master_path, sheet_name=sheet)
        subset = raw_df.iloc[:, 0:11].copy()
        subset.columns = ['Subject_ID'] + COLUMNS_TO_AVERAGE
        subset['Sheet'] = sheet
        all_data.append(subset)

    if not all_data:
        print("No data sheets found in master.")
        return

    final_df = pd.concat(all_data, ignore_index=True).dropna(subset=['Subject_ID'])

    if group_map:
        final_df['Group'] = final_df['Subject_ID'].apply(
            lambda x: get_group_name(str(x), group_map) or "Unknown"
        )
    else:
        final_df['Group'] = final_df['Sheet']

    if sex == "Mixed" and group_map:
        final_df['Sex'] = final_df['Subject_ID'].apply(
            lambda x: deduce_sex(str(x), group_map)
        )

    csv_output_dir.mkdir(parents=True, exist_ok=True)

    label_parts  = [study_name] + [p for p in [sex, age] if p]
    summary_path = csv_output_dir / f"{'_'.join(label_parts)}_Summary.csv"
    summary_cols = ['Subject_ID', 'Group'] + (['Sex'] if 'Sex' in final_df.columns else []) + COLUMNS_TO_AVERAGE
    final_df[summary_cols].to_csv(summary_path, index=False)
    print(f"Summary saved to {summary_path}")

    metrics_dir = csv_output_dir / "Per_Metric"
    metrics_dir.mkdir(exist_ok=True)

    if sex == "Mixed" and 'Sex' in final_df.columns:
        sex_splits = [("Male", final_df[final_df['Sex'] == "Male"]),
                      ("Female", final_df[final_df['Sex'] == "Female"])]
    else:
        sex_splits = [(None, final_df)]

    for sex_label, df_split in sex_splits:
        out_dir = metrics_dir / sex_label if sex_label else metrics_dir
        out_dir.mkdir(exist_ok=True)
        if df_split.empty:
            print(f"No data for {sex_label}, skipping.")
            continue
        for metric in COLUMNS_TO_AVERAGE:
            clean_metric = re.sub(r'[\-\/\(\)]', '', metric).replace("  ", " ").strip()
            try:
                table = df_split.pivot_table(index='Subject_ID', columns='Group', values=metric)
                table.to_csv(out_dir / f"{clean_metric}.csv")
            except Exception as e:
                print(f"Could not generate pivot for {metric}: {e}")

    print(f"Per-metric tables saved to {metrics_dir}")


# --- 5. PIPELINE ENTRY POINT ---

def run_pipeline(config):
    study_name    = config["study_name"]
    raw_data_root = Path(config["raw_data_root"])
    master_path   = Path(config["master_file"])
    output_folder = Path(config["output_folder"])
    group_map     = config.get("group_map", {})
    sex           = config.get("sex", "")
    age           = config.get("age", "")

    txt_files       = list(raw_data_root.glob("*.txt"))
    analyzed_folder = raw_data_root / "Analyzed_Files"
    csv_output_dir  = output_folder / f"{study_name}_CSV"

    print(f"Study:      {study_name}")
    print(f"Sex:        {sex or '(not set)'}")
    print(f"Age:        {age or '(not set)'}")
    print(f"Raw data:   {raw_data_root}")
    print(f"Master:     {master_path}")
    print(f"Group map:  {group_map}")
    print(f"CSV output: {csv_output_dir}")

    if not txt_files:
        QMessageBox.critical(None, "No Data Files",
                             f"No .txt files found in:\n{raw_data_root}")
        return

    if not master_path.exists():
        QMessageBox.critical(None, "Master File Not Found",
                             f"Could not find:\n{master_path}")
        return

    df_source = clean_and_average_data(txt_files, group_map)
    if df_source.empty:
        QMessageBox.critical(
            None, "No Valid Data",
            "No valid data was found in the .txt files.\n"
            + ("Check that your group_map prefixes match the IDs in your files."
               if group_map else "Check that your .txt files contain valid data.")
        )
        return

    try:
        sync_to_master(df_source, master_path)

        analyzed_folder.mkdir(exist_ok=True)
        for file_path in txt_files:
            shutil.move(str(file_path), str(analyzed_folder / file_path.name))
        print(f"Processed files moved to {analyzed_folder}/")

        output_folder.mkdir(parents=True, exist_ok=True)
        process_master_to_csv(master_path, csv_output_dir, study_name, group_map, sex, age)

        QMessageBox.information(
            None, "Pipeline Complete",
            f"All done!\n\n"
            f"CSV summaries saved to:\n{csv_output_dir}\n\n"
            f"Raw files archived to:\n{analyzed_folder}"
        )

    except PermissionError:
        QMessageBox.critical(None, "File In Use",
                             "Could not save the master file.\n"
                             "Please close the Excel file and try again.")
    except Exception as e:
        QMessageBox.critical(None, "Unexpected Error", str(e))
