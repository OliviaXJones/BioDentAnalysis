import json
import sys
import pandas as pd
import openpyxl
import shutil
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QDialogButtonBox,
    QMessageBox, QWidget, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from BioDent_Utils import (
    COLUMNS_TO_AVERAGE, clean_and_average_data, ask_not_found_action
)

CONFIG_PATH = Path(__file__).parent / "single_study_config.json"


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


# --- 4. STANDALONE CONFIG DIALOG ---

class ConfigDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BioDent Single Study Pipeline")
        self.setMinimumWidth(640)
        self.setWindowModality(Qt.ApplicationModal)
        self._config = {}

        existing = {}
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH) as f:
                    existing = json.load(f)
            except Exception:
                pass

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(28, 24, 28, 24)

        title = QLabel("Single Study Configuration")
        title.setFont(QFont("Arial", 15, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        div = QFrame()
        div.setFrameShape(QFrame.HLine)
        div.setFrameShadow(QFrame.Sunken)
        layout.addWidget(div)

        form = QFormLayout()
        form.setSpacing(12)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self._study_name = QLineEdit(existing.get("study_name", ""))
        self._study_name.setPlaceholderText("e.g. Cohort2_Treatment")
        form.addRow("Study Name:", self._study_name)

        self._raw_data, raw_row = self._path_row(existing.get("raw_data_root", ""), folder=True)
        form.addRow("Raw Data Folder:", raw_row)

        self._master, master_row = self._path_row(existing.get("master_file", ""), folder=False)
        form.addRow("Master File (.xlsx):", master_row)

        self._output, output_row = self._path_row(existing.get("output_folder", ""), folder=True)
        form.addRow("Output Folder:", output_row)

        layout.addLayout(form)
        layout.addSpacing(8)

        btn_box = QDialogButtonBox()
        run_btn = btn_box.addButton("Run Pipeline", QDialogButtonBox.AcceptRole)
        run_btn.setFixedHeight(36)
        run_btn.setFont(QFont("Arial", 10, QFont.Bold))
        btn_box.addButton("Cancel", QDialogButtonBox.RejectRole).setFixedHeight(36)
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _path_row(self, value, folder):
        edit = QLineEdit(value)
        edit.setPlaceholderText("Click Browse or type a path...")
        edit.setMinimumWidth(360)
        btn = QPushButton("Browse...")
        btn.setFixedWidth(90)
        if folder:
            btn.clicked.connect(lambda: self._pick_folder(edit))
        else:
            btn.clicked.connect(lambda: self._pick_file(edit))
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)
        h.addWidget(edit)
        h.addWidget(btn)
        return edit, row

    def _pick_folder(self, edit):
        p = QFileDialog.getExistingDirectory(self, "Select Folder",
                                             edit.text() or str(Path.home()))
        if p:
            edit.setText(p)

    def _pick_file(self, edit):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select Master Excel File",
            edit.text() or str(Path.home()),
            "Excel Files (*.xlsx *.xlsm)"
        )
        if p:
            edit.setText(p)

    def _on_accept(self):
        name   = self._study_name.text().strip()
        raw    = self._raw_data.text().strip()
        master = self._master.text().strip()
        output = self._output.text().strip()

        if not all([name, raw, master, output]):
            QMessageBox.warning(self, "Missing Fields",
                                "Please fill in all four fields before running.")
            return

        self._config = {
            "study_name":    name,
            "raw_data_root": raw,
            "master_file":   master,
            "output_folder": output,
        }
        with open(CONFIG_PATH, "w") as f:
            json.dump(self._config, f, indent=4)
        self.accept()

    def get_config(self):
        return self._config


# --- 5. STANDALONE ENTRY POINT ---

if __name__ == "__main__":
    app = QApplication(sys.argv)
    dlg = ConfigDialog()
    if dlg.exec_() != QDialog.Accepted:
        sys.exit(0)
    run_pipeline(dlg.get_config())
    sys.exit(0)
