import json
import sys
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QDialogButtonBox,
    QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QWidget, QFrame, QComboBox, QGroupBox
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

        self._sex = QComboBox()
        self._sex.addItems(["Male", "Female", "Mixed"])
        self._sex.setCurrentText(s.get("sex", "Male"))
        self._sex.currentTextChanged.connect(self._on_sex_changed)
        form.addRow("Cohort Sex:", self._sex)

        self._age = QLineEdit(s.get("age", ""))
        self._age.setPlaceholderText("e.g. 16 Weeks")
        form.addRow("Cohort Age:", self._age)

        self._raw, raw_row = self._path_row(s.get("raw_data_root", ""), folder=True)
        form.addRow("Raw Data Folder:", raw_row)

        self._master, master_row = self._path_row(s.get("master_file", ""), folder=False)
        form.addRow("Master File (.xlsx):", master_row)

        self._output, output_row = self._path_row(s.get("output_folder", ""), folder=True)
        form.addRow("Output Folder:", output_row)

        layout.addLayout(form)
        layout.addSpacing(8)

        # --- Group Map (Single Study only) ---
        self._gm_group = QGroupBox("Group Map  (prefix → group name)")
        gm_layout = QVBoxLayout(self._gm_group)

        _init_mixed = s.get("sex", "Male") == "Mixed"
        self._gm_table = QTableWidget(0, 3 if _init_mixed else 2)
        self._gm_table.setMinimumHeight(120)
        self._gm_table.verticalHeader().setVisible(False)
        self._update_gm_headers()
        gm_layout.addWidget(self._gm_table)

        gm_btn_row = QHBoxLayout()
        add_row_btn = QPushButton("+ Add Row")
        add_row_btn.clicked.connect(lambda: self._add_gm_row())
        rm_row_btn = QPushButton("- Remove Selected")
        rm_row_btn.clicked.connect(self._remove_gm_row)
        gm_btn_row.addWidget(add_row_btn)
        gm_btn_row.addWidget(rm_row_btn)
        gm_btn_row.addStretch()
        gm_layout.addLayout(gm_btn_row)

        layout.addWidget(self._gm_group)

        for prefix, val in s.get("group_map", {}).items():
            if isinstance(val, dict):
                self._add_gm_row(prefix, val.get("group", ""), val.get("sex", "Male"))
            else:
                self._add_gm_row(prefix, val)

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

        self._study = {
            "study_name":    name,
            "study_type":    "single_study",
            "raw_data_root": raw,
            "master_file":   master,
            "output_folder": output,
            "sex":           self._sex.currentText(),
            "age":           self._age.text().strip(),
        }

        group_map = {}
        is_mixed  = self._sex.currentText() == "Mixed"
        for row in range(self._gm_table.rowCount()):
            p_item = self._gm_table.item(row, 0)
            prefix = p_item.text().strip() if p_item else ""
            if not prefix:
                continue
            if is_mixed:
                sex_w  = self._gm_table.cellWidget(row, 1)
                sex_val = sex_w.currentText() if sex_w else "Male"
                g_item  = self._gm_table.item(row, 2)
                group_map[prefix] = {"group": g_item.text().strip() if g_item else "", "sex": sex_val}
            else:
                g_item = self._gm_table.item(row, 1)
                group_map[prefix] = g_item.text().strip() if g_item else ""
        self._study["group_map"] = group_map

        self.accept()

    def _update_gm_headers(self):
        if self._gm_table.columnCount() == 3:
            self._gm_table.setHorizontalHeaderLabels(["Prefix", "Sex", "Group Name"])
            self._gm_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self._gm_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            self._gm_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        else:
            self._gm_table.setHorizontalHeaderLabels(["Prefix", "Group Name"])
            self._gm_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self._gm_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)

    def _on_sex_changed(self, sex_text):
        if not hasattr(self, '_gm_table'):
            return
        is_mixed       = sex_text == "Mixed"
        currently_mixed = self._gm_table.columnCount() == 3
        if is_mixed == currently_mixed:
            return
        # Snapshot existing rows before rebuilding
        rows = []
        for row in range(self._gm_table.rowCount()):
            p_item = self._gm_table.item(row, 0)
            prefix = p_item.text().strip() if p_item else ""
            if currently_mixed:
                sex_w   = self._gm_table.cellWidget(row, 1)
                sex_val = sex_w.currentText() if sex_w else "Male"
                g_item  = self._gm_table.item(row, 2)
            else:
                sex_val = "Male"
                g_item  = self._gm_table.item(row, 1)
            rows.append((prefix, g_item.text().strip() if g_item else "", sex_val))
        self._gm_table.setRowCount(0)
        self._gm_table.setColumnCount(3 if is_mixed else 2)
        self._update_gm_headers()
        for prefix, group_name, sex_val in rows:
            self._add_gm_row(prefix, group_name, sex_val)

    def _add_gm_row(self, prefix="", group_name="", sex="Male"):
        row = self._gm_table.rowCount()
        self._gm_table.insertRow(row)
        self._gm_table.setItem(row, 0, QTableWidgetItem(prefix))
        if self._gm_table.columnCount() == 3:
            sex_combo = QComboBox()
            sex_combo.addItems(["Male", "Female"])
            sex_combo.setCurrentText(sex if sex in ("Male", "Female") else "Male")
            self._gm_table.setCellWidget(row, 1, sex_combo)
            self._gm_table.setItem(row, 2, QTableWidgetItem(group_name))
        else:
            self._gm_table.setItem(row, 1, QTableWidgetItem(group_name))

    def _remove_gm_row(self):
        row = self._gm_table.currentRow()
        if row >= 0:
            self._gm_table.removeRow(row)

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
        add_btn        = QPushButton("Add Study")
        self._edit_btn = QPushButton("Edit Study")
        self._del_btn  = QPushButton("Remove Study")
        for b in [add_btn, self._edit_btn, self._del_btn]:
            b.setFixedHeight(32)
            mgmt_row.addWidget(b)
        add_btn.clicked.connect(self._add_study)
        self._edit_btn.clicked.connect(self._edit_study)
        self._del_btn.clicked.connect(self._remove_study)
        self._table.itemSelectionChanged.connect(self._update_mgmt_buttons)
        layout.addLayout(mgmt_row)
        self._update_mgmt_buttons()

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
        from PyQt5.QtGui import QColor
        self._table.setRowCount(len(self._studies))
        for i, s in enumerate(self._studies):
            name_item = QTableWidgetItem(s.get("study_name", ""))
            is_fkbp5  = s.get("study_type") == "fkbp5"
            type_label = "FKBP5 (built-in)" if is_fkbp5 else "Single Study"
            type_item  = QTableWidgetItem(type_label)
            if is_fkbp5:
                muted = QColor("#6c7086")
                name_item.setForeground(muted)
                type_item.setForeground(muted)
            self._table.setItem(i, 0, name_item)
            self._table.setItem(i, 1, type_item)
        self._update_mgmt_buttons()

    def _update_mgmt_buttons(self):
        if not hasattr(self, '_edit_btn'):
            return
        row      = self._selected_row()
        is_fkbp5 = row >= 0 and self._studies[row].get("study_type") == "fkbp5"
        locked   = row < 0 or is_fkbp5
        self._edit_btn.setEnabled(not locked)
        self._del_btn.setEnabled(not locked)

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


# --- DARK THEME ---

DARK_STYLESHEET = """
QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: Arial;
    font-size: 10pt;
}
QDialog {
    background-color: #1e1e2e;
}
QGroupBox {
    border: 1px solid #45475a;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 6px;
    color: #89b4fa;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
}
QLineEdit, QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 4px 6px;
    color: #cdd6f4;
    selection-background-color: #89b4fa;
    selection-color: #1e1e2e;
}
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #89b4fa;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #89b4fa;
    selection-color: #1e1e2e;
}
QPushButton {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 5px 14px;
    color: #cdd6f4;
}
QPushButton:hover {
    background-color: #45475a;
    border-color: #89b4fa;
}
QPushButton:pressed {
    background-color: #89b4fa;
    color: #1e1e2e;
}
QTableWidget {
    background-color: #181825;
    alternate-background-color: #1e1e2e;
    gridline-color: #45475a;
    border: 1px solid #45475a;
    border-radius: 4px;
}
QTableWidget::item {
    padding: 4px;
}
QTableWidget::item:selected {
    background-color: #89b4fa;
    color: #1e1e2e;
}
QHeaderView::section {
    background-color: #313244;
    color: #89b4fa;
    border: none;
    border-right: 1px solid #45475a;
    border-bottom: 1px solid #45475a;
    padding: 5px 8px;
    font-weight: bold;
}
QScrollBar:vertical {
    background: #181825;
    width: 10px;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background: #45475a;
    border-radius: 5px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #89b4fa;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: #181825;
    height: 10px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal {
    background: #45475a;
    border-radius: 5px;
    min-width: 20px;
}
QScrollBar::handle:horizontal:hover { background: #89b4fa; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QLabel {
    color: #cdd6f4;
    background: transparent;
}
QFrame[frameShape="4"], QFrame[frameShape="5"] {
    color: #45475a;
}
QDialogButtonBox QPushButton {
    min-width: 80px;
}
QMessageBox {
    background-color: #1e1e2e;
}
QMessageBox QLabel {
    color: #cdd6f4;
}
"""


# --- ENTRY POINT ---

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    mgr = StudyManagerDialog()
    mgr.exec_()
    sys.exit(0)
