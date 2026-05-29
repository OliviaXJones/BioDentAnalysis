import re
import pandas as pd
import openpyxl
import shutil
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from BioDent_Utils import (
    COLUMNS_TO_AVERAGE, get_sex, clean_and_average_data, ask_not_found_action
)

CONFIG_PATH = Path(__file__).parent / "study_config.json"


# --- 1. FKBP5-SPECIFIC HELPERS ---

def parse_squashed_code(code):
    """Parses IDs like 'W12M10', 'H24F5', or '2W10M10' for CSV grouping."""
    code = str(code).strip().upper()
    match = re.match(r"(\d+)?([A-Z])(\d+)([MF])(\d+)", code)
    if match:
        gen_num, gen_char, age, sex_char, mouse_num = match.groups()
        gen_map = {'W': 'Wildtype', 'M': 'Mutant', 'H': 'Heterozygous'}
        genotype = gen_map.get(gen_char, "Unknown")
        sex = "Male" if sex_char == 'M' else "Female"
        if gen_num:
            genotype = f"{gen_num}{genotype}"
        return genotype, age, sex, mouse_num
    return None, None, None, None


def get_genotype_prefix(code):
    if not code:
        return ""
    genotype_full, _, _, _ = parse_squashed_code(code)
    if genotype_full:
        return re.sub(r'\d+', '', genotype_full)
    return ""


# --- 2. MASTER SYNC ---

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
            genotype  = get_genotype_prefix(mouse_code)
            sex       = get_sex(mouse_code)
            start_col = 1 if sex == 'M' else 13

            target_sheets = [s for s in all_sheets if genotype.lower() in s.lower()]
            other_sheets  = [s for s in all_sheets if s not in target_sheets]
            search_order  = target_sheets + other_sheets

            for sheet_name in search_order:
                ws = wb[sheet_name]
                for r in range(2, max(ws.max_row + 1, 50)):
                    cell_val   = ws.cell(row=r, column=start_col).value
                    clean_cell = str(cell_val).strip().upper() if cell_val else ""
                    if clean_cell == mouse_code.strip().upper():
                        for j, col_name in enumerate(COLUMNS_TO_AVERAGE):
                            ws.cell(row=r, column=start_col + 1 + j).value = row_data[col_name]
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
    print(f"Synced {sync_count} mice to master.")

    unmatched = all_processed_ids - synced_ids
    if unmatched:
        print("The following mice were not found in the master list:")
        for m in sorted(unmatched):
            print(f"  - {m}")
    else:
        print("All processed mice were successfully matched in the master list.")


# --- 3. CSV GENERATION ---

def generate_grouped_tables(df, csv_output_dir):
    sort_priority   = ['Wildtype', 'Mutant', 'Heterozygous']
    folder_genotype = csv_output_dir / "Analysis_By_Genotype"
    folder_lineage  = csv_output_dir / "Analysis_By_Lineage"
    folder_genotype.mkdir(parents=True, exist_ok=True)
    folder_lineage.mkdir(parents=True, exist_ok=True)

    for metric in COLUMNS_TO_AVERAGE:
        clean_metric = re.sub(r'[\-\/\(\)]', '', metric).replace("  ", " ").strip()
        for age in df['Age_Extracted'].dropna().unique():
            for sex in ['Male', 'Female']:
                base_subset = df[
                    (df['Sex_Extracted'] == sex) & (df['Age_Extracted'] == age)
                ].copy()
                if base_subset.empty:
                    continue

                gen_subset = base_subset[
                    base_subset['Progeny_Group'].isin(sort_priority)
                ].copy()
                if not gen_subset.empty:
                    table_gen = gen_subset.pivot_table(
                        index='ID_Num', columns='Progeny_Group', values=metric)
                    table_gen = table_gen.reindex(
                        columns=[c for c in sort_priority if c in table_gen.columns])
                    table_gen.to_csv(
                        folder_genotype / f"{sex}_{age}Wks_{clean_metric}.csv")

                table_lin = base_subset.pivot_table(
                    index='ID_Num', columns='Progeny_Group', values=metric)
                sorted_cols = sorted(
                    table_lin.columns,
                    key=lambda x: next(
                        (idx for idx, g in enumerate(sort_priority) if x.startswith(g)), 99)
                )
                table_lin[sorted_cols].to_csv(
                    folder_lineage / f"{sex}_{age}Wks_{clean_metric}.csv")


def process_master_to_csv(master_path, csv_output_dir):
    print("Generating CSV analysis tables...")
    xl = pd.ExcelFile(master_path)
    all_data = []
    for sheet in xl.sheet_names:
        if sheet.lower() in ["summary", "notes", "calculations"]:
            continue
        raw_df = pd.read_excel(master_path, sheet_name=sheet)
        for start_idx in [0, 12]:
            subset = raw_df.iloc[:, start_idx:start_idx + 11].copy()
            subset.columns = ['Mouse Code'] + COLUMNS_TO_AVERAGE
            subset['Progeny_Group'] = sheet
            all_data.append(subset)

    final_df = pd.concat(all_data, ignore_index=True).dropna(subset=['Mouse Code'])
    parsed = final_df['Mouse Code'].apply(lambda x: pd.Series(parse_squashed_code(x)))
    final_df[['Gen_Label', 'Age_Extracted', 'Sex_Extracted', 'ID_Num']] = parsed
    generate_grouped_tables(final_df, csv_output_dir)


# --- 4. PIPELINE ENTRY POINT ---

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
        process_master_to_csv(master_path, csv_output_dir)

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


