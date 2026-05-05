import pandas as pd
import numpy as np

old_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\BioDent\FKBP5_BioDentMasterBackup_042826.xlsx"
new_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\BioDent\FKBP5_BioDentMaster.xlsx"

# Use the sheet names from your error log
sheets_to_check = [
    'Wildtype', 'Wildtype F0', 'Wildtype F1', 'Wildtype F2',
    'Mutant', 'Mutant F0', 'Mutant F1', 'Mutant F2',
    'Heterozygous', 'Heterozygous F0', 'Heterozygous F1', 'Heterozygous F2'
]

for sheet in sheets_to_check:
    print(f"\n--- Checking Sheet: {sheet} ---")
    try:
        # Load sheets
        df_old = pd.read_excel(old_file, sheet_name=sheet)
        df_new = pd.read_excel(new_file, sheet_name=sheet)

        # Standardize: First column is ID, strip spaces from headers
        id_col = df_old.columns[0]
        df_old.columns = df_old.columns.str.strip()
        df_new.columns = df_new.columns.str.strip()

        # Drop rows that are completely empty
        df_old = df_old.dropna(subset=[id_col])
        df_new = df_new.dropna(subset=[id_col])

        # Convert IDs to strings and strip spaces to ensure matching works
        df_old[id_col] = df_old[id_col].astype(str).str.strip()
        df_new[id_col] = df_new[id_col].astype(str).str.strip()

        # Identify numeric columns only
        numeric_cols = df_old.select_dtypes(include=[np.number]).columns.intersection(
            df_new.select_dtypes(include=[np.number]).columns)

        # Create a dictionary for the old data for fast lookup
        # This prevents "broadcasting" errors entirely
        old_data_lookup = df_old.set_index(
            id_col)[numeric_cols].to_dict('index')

        mismatches_found = False

        # Iterate through the NEW file and compare to the lookup
        for _, row in df_new.iterrows():
            m_id = row[id_col]

            if m_id in old_data_lookup:
                diff_cols = []
                for col in numeric_cols:
                    val_new = row[col]
                    val_old = old_data_lookup[m_id][col]

                    # Handle NaNs (treat two NaNs as equal)
                    if pd.isna(val_new) and pd.isna(val_old):
                        continue

                    # Compare with tolerance
                    if not np.isclose(float(val_new or 0), float(val_old or 0), atol=1e-2):
                        diff_cols.append(col)

                if diff_cols:
                    print(f"❌ ID: {m_id}")
                    print(f"   Mismatched Columns: {', '.join(diff_cols)}")
                    mismatches_found = True

        if not mismatches_found:
            print("All matching mice in this sheet have identical numerical data.")

    except Exception as e:
        print(f"Error checking {sheet}: {e}")
