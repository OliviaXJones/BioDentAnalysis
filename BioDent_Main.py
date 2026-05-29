import json
import sys
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QDialogButtonBox,
    QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QWidget, QFrame, QComboBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

STUDIES_CONFIG_PATH = Path(__file__).parent / "studies_config.json"

STUDY_TYPES = ["FKBP5", "Single Study"]


# --- CONFIG HELPERS ---

def load_studies():
    if STUDIES_CONFIG_PATH.exists():
        try:
            with open(STUDIES_CONFIG_PATH) as f:
                return json.load(f).get("studies", [])
        except Exception:
            pass
    return []


def save_studies(studies):
    with open(STUDIES_CONFIG_PATH, "w") as f:
        json.dump({"studies": studies}, f, indent=4)


# --- STUDY EDIT DIALOG ---

class StudyEditDialog(QDialog):
    def __init__(self, study=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Study" if study else "Add Study")
        self.setMinimumWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self._study = {}
        s = study or {}

        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(24, 20, 24, 20)

        title = QLabel("Edit Study" if study else "Add Study")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        div = QFrame()
        div.setFrameShape(QFrame.HLine)
        div.setFrameShadow(QFrame.Sunken)
        layout.addWidget(div)

        form = QFormLayout()
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self._name = QLineEdit(s.get("study_name", ""))
        self._name.setPlaceholderText("e.g. FKBP5 Cohort 2")
        form.addRow("Study Name:", self._name)

        self._type = QComboBox()
        self._type.addItems(STUDY_TYPES)
        stored_type = s.get("study_type", "fkbp5").lower()
        self._type.setCurrentIndex(0 if stored_type == "fkbp5" else 1)
        form.addRow("Study Type:", self._type)

        self._raw, raw_row = self._path_row(s.get("raw_data_root", ""), folder=True)
        form.addRow("Raw Data Folder:", raw_row)

        self._master, master_row = self._path_row(s.get("master_file", ""), folder=False)
        form.addRow("Master File (.xlsx):", master_row)

        self._output, output_row = self._path_row(s.get("output_folder", ""), folder=True)
        form.addRow("Output Folder:", output_row)

        layout.addLayout(form)
        layout.addSpacing(8)

        btn_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _path_row(self, value, folder):
        edit = QLineEdit(value)
        edit.setPlaceholderText("Click Browse or type a path...")
        edit.setMinimumWidth(340)
        btn = QPushButton("Browse...")
        btn.setFixedWidth(85)
        if folder:
            btn.clicked.connect(lambda: self._pick_folder(edit))
        else:
            btn.clicked.connect(lambda: self._pick_file(edit))
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(6)
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
        name   = self._name.text().strip()
        raw    = self._raw.text().strip()
        master = self._master.text().strip()
        output = self._output.text().strip()

        if not all([name, raw, master, output]):
            QMessageBox.warning(self, "Missing Fields", "Please fill in all fields.")
            return

        type_str = "fkbp5" if self._type.currentIndex() == 0 else "single_study"
        self._study = {
            "study_name":    name,
            "study_type":    type_str,
            "raw_data_root": raw,
            "master_file":   master,
            "output_folder": output,
        }
        self.accept()

    def get_study(self):
        return self._study


# --- STUDY MANAGER DIALOG ---

class StudyManagerDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BioDent Analysis Pipeline")
        self.setMinimumSize(700, 400)
        self.setWindowModality(Qt.ApplicationModal)
        self._studies = load_studies()

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 20, 24, 20)

        title = QLabel("BioDent Analysis Pipeline")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Select a study and click Run, or manage your study list below.")
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)

        div = QFrame()
        div.setFrameShape(QFrame.HLine)
        div.setFrameShadow(QFrame.Sunken)
        layout.addWidget(div)

        self._table = QTableWidget()
        self._table.setColumnCount(2)
        self._table.setHorizontalHeaderLabels(["Study Name", "Type"])
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.verticalHeader().setVisible(False)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.setColumnWidth(1, 110)
        self._table.setAlternatingRowColors(True)
        self._table.doubleClicked.connect(self._run_selected)
        layout.addWidget(self._table)

        self._refresh_table()

        mgmt_row = QHBoxLayout()
        add_btn  = QPushButton("Add Study")
        edit_btn = QPushButton("Edit Study")
        del_btn  = QPushButton("Remove Study")
        for b in [add_btn, edit_btn, del_btn]:
            b.setFixedHeight(32)
            mgmt_row.addWidget(b)
        add_btn.clicked.connect(self._add_study)
        edit_btn.clicked.connect(self._edit_study)
        del_btn.clicked.connect(self._remove_study)
        layout.addLayout(mgmt_row)

        run_btn = QPushButton("Run Selected Study")
        run_btn.setFixedHeight(44)
        run_btn.setFont(QFont("Arial", 11, QFont.Bold))
        run_btn.setStyleSheet(
            "QPushButton { background-color: #1976D2; color: white; border-radius: 5px; }"
            "QPushButton:hover { background-color: #1565C0; }"
        )
        run_btn.clicked.connect(self._run_selected)
        layout.addWidget(run_btn)

    def _refresh_table(self):
        self._table.setRowCount(len(self._studies))
        for i, s in enumerate(self._studies):
            self._table.setItem(i, 0, QTableWidgetItem(s.get("study_name", "")))
            type_label = "FKBP5" if s.get("study_type") == "fkbp5" else "Single Study"
            self._table.setItem(i, 1, QTableWidgetItem(type_label))

    def _selected_row(self):
        items = self._table.selectedItems()
        return self._table.currentRow() if items else -1

    def _add_study(self):
        dlg = StudyEditDialog(parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self._studies.append(dlg.get_study())
            save_studies(self._studies)
            self._refresh_table()

    def _edit_study(self):
        row = self._selected_row()
        if row < 0:
            QMessageBox.information(self, "No Selection", "Please select a study to edit.")
            return
        dlg = StudyEditDialog(self._studies[row], parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self._studies[row] = dlg.get_study()
            save_studies(self._studies)
            self._refresh_table()

    def _remove_study(self):
        row = self._selected_row()
        if row < 0:
            QMessageBox.information(self, "No Selection", "Please select a study to remove.")
            return
        name = self._studies[row].get("study_name", "this study")
        reply = QMessageBox.question(
            self, "Confirm Remove",
            f"Remove '{name}' from the list?\n(This does not delete any files.)",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._studies.pop(row)
            save_studies(self._studies)
            self._refresh_table()

    def _run_selected(self):
        row = self._selected_row()
        if row < 0:
            QMessageBox.information(self, "No Selection", "Please select a study to run.")
            return

        study      = self._studies[row]
        study_type = study.get("study_type", "single_study").lower()

        if study_type == "fkbp5":
            from FKBP5_BioDent_Pipeline import run_pipeline
        else:
            from SingleStudy_BioDent_Pipeline import run_pipeline

        run_pipeline(study)


# --- ENTRY POINT ---

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mgr = StudyManagerDialog()
    mgr.exec_()
    sys.exit(0)
