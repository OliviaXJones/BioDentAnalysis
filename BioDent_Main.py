import json
import sys
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QDialogButtonBox,
    QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QWidget, QFrame, QComboBox, QGroupBox, QCheckBox, QTextBrowser,
    QScrollArea
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
        self.setMinimumWidth(640)
        self.setWindowModality(Qt.ApplicationModal)
        self._study = {}
        s = study or {}

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 16, 20, 16)

        # ── Study Details ─────────────────────────────────────────────
        details_group = QGroupBox("Study Details")
        form = QFormLayout(details_group)
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
        self._master.setPlaceholderText("Leave blank to auto-create in Output Folder")
        form.addRow("Master File (.xlsx)\n(optional):", master_row)

        self._output, output_row = self._path_row(s.get("output_folder", ""), folder=True)
        form.addRow("Output Folder:", output_row)

        layout.addWidget(details_group)

        # ── Group Map ─────────────────────────────────────────────────
        gm_group = QGroupBox("Group Map  (prefix → group name)")
        gm_layout = QVBoxLayout(gm_group)

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
        self._deduce_cb = QCheckBox("Deduce sex from ID")
        self._deduce_cb.setVisible(_init_mixed)
        self._deduce_cb.toggled.connect(self._on_deduce_toggled)
        gm_btn_row.addWidget(self._deduce_cb)
        gm_layout.addLayout(gm_btn_row)

        layout.addWidget(gm_group)

        # Detect initial deduce state and load rows
        _init_deduce = any(
            isinstance(v, dict) and v.get("sex") == "Deduce from ID"
            for v in s.get("group_map", {}).values()
        )
        self._deduce_cb.setChecked(_init_deduce)

        for prefix, val in s.get("group_map", {}).items():
            if isinstance(val, dict):
                stored_sex = val.get("sex", "Male")
                self._add_gm_row(prefix, val.get("group", ""),
                                  "Male" if stored_sex == "Deduce from ID" else stored_sex)
            else:
                self._add_gm_row(prefix, val)

        # ── Buttons ───────────────────────────────────────────────────
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    # ------------------------------------------------------------------

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

    def _on_deduce_toggled(self, checked):
        for row in range(self._gm_table.rowCount()):
            w = self._gm_table.cellWidget(row, 2)
            if w:
                w.setEnabled(not checked)

    def _on_accept(self):
        name   = self._name.text().strip()
        raw    = self._raw.text().strip()
        output = self._output.text().strip()

        if not all([name, raw, output]):
            QMessageBox.warning(self, "Missing Fields",
                                "Please fill in Study Name, Raw Data Folder, and Output Folder.")
            return

        self._study = {
            "study_name":    name,
            "study_type":    "single_study",
            "raw_data_root": raw,
            "master_file":   self._master.text().strip(),
            "output_folder": output,
            "sex":           self._sex.currentText(),
            "age":           self._age.text().strip(),
        }

        group_map  = {}
        is_mixed   = self._sex.currentText() == "Mixed"
        deduce_all = is_mixed and self._deduce_cb.isChecked()
        for row in range(self._gm_table.rowCount()):
            p_item = self._gm_table.item(row, 0)
            prefix = p_item.text().strip() if p_item else ""
            if not prefix:
                continue
            if is_mixed:
                g_item  = self._gm_table.item(row, 1)
                sex_w   = self._gm_table.cellWidget(row, 2)
                sex_val = "Deduce from ID" if deduce_all else (
                    sex_w.currentText() if sex_w else "Male")
                group_map[prefix] = {
                    "group": g_item.text().strip() if g_item else "",
                    "sex":   sex_val,
                }
            else:
                g_item = self._gm_table.item(row, 1)
                group_map[prefix] = g_item.text().strip() if g_item else ""
        self._study["group_map"] = group_map

        self.accept()

    def _update_gm_headers(self):
        if self._gm_table.columnCount() == 3:
            self._gm_table.setHorizontalHeaderLabels(["Prefix", "Group Name", "Sex"])
            self._gm_table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.ResizeToContents)
            self._gm_table.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.Stretch)
            self._gm_table.setColumnWidth(2, 90)
            self._gm_table.horizontalHeader().setSectionResizeMode(
                2, QHeaderView.Fixed)
        else:
            self._gm_table.setHorizontalHeaderLabels(["Prefix", "Group Name"])
            self._gm_table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.ResizeToContents)
            self._gm_table.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.Stretch)

    def _on_sex_changed(self, sex_text):
        if not hasattr(self, '_gm_table'):
            return
        is_mixed        = sex_text == "Mixed"
        currently_mixed = self._gm_table.columnCount() == 3
        if is_mixed == currently_mixed:
            return

        if hasattr(self, '_deduce_cb'):
            self._deduce_cb.setVisible(is_mixed)

        rows = []
        for row in range(self._gm_table.rowCount()):
            p_item = self._gm_table.item(row, 0)
            prefix = p_item.text().strip() if p_item else ""
            if currently_mixed:
                g_item  = self._gm_table.item(row, 1)
                sex_w   = self._gm_table.cellWidget(row, 2)
                sex_val = sex_w.currentText() if sex_w else "Male"
                rows.append((prefix, g_item.text().strip() if g_item else "", sex_val))
            else:
                g_item = self._gm_table.item(row, 1)
                rows.append((prefix, g_item.text().strip() if g_item else "", "Male"))

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
            self._gm_table.setItem(row, 1, QTableWidgetItem(group_name))
            sex_combo = QComboBox()
            sex_combo.addItems(["Male", "Female"])
            sex_combo.setCurrentText(sex if sex in ("Male", "Female") else "Male")
            if hasattr(self, '_deduce_cb'):
                sex_combo.setEnabled(not self._deduce_cb.isChecked())
            self._gm_table.setCellWidget(row, 2, sex_combo)
        else:
            self._gm_table.setItem(row, 1, QTableWidgetItem(group_name))

    def _remove_gm_row(self):
        row = self._gm_table.currentRow()
        if row >= 0:
            self._gm_table.removeRow(row)

    def get_study(self):
        return self._study


# --- HELP DIALOG ---

_HELP_HTML = """
<style>
  body  { font-family: Arial, sans-serif; font-size: 10pt; color: #cdd6f4; }
  h1    { font-size: 13pt; color: #89b4fa; margin-bottom: 4px; }
  h2    { font-size: 11pt; color: #89dceb; margin-top: 18px; margin-bottom: 4px; border-bottom: 1px solid #45475a; padding-bottom: 3px; }
  h3    { font-size: 10pt; color: #a6e3a1; margin-top: 12px; margin-bottom: 2px; }
  ul    { margin: 4px 0 8px 18px; padding: 0; }
  li    { margin-bottom: 3px; }
  code  { background: #313244; padding: 1px 5px; border-radius: 3px; font-family: Consolas, monospace; font-size: 9.5pt; }
  .note { color: #f38ba8; font-style: italic; }
  table { border-collapse: collapse; width: 100%; margin-top: 6px; }
  td, th { border: 1px solid #45475a; padding: 5px 8px; text-align: left; vertical-align: top; }
  th    { background: #313244; color: #89b4fa; font-weight: bold; }
  tr:nth-child(even) { background: #1e1e2e; }
</style>

<h1>BioDent Analysis Pipeline &mdash; User Guide</h1>
<p>This guide explains every automatic data-cleaning rule and every interactive prompt
you may encounter while running a study.</p>

<h2>1 &mdash; Setting Up a Study</h2>
<p>Click <b>Add Study</b> to create a new study configuration. All settings are saved
and reused on future runs.</p>

<table>
  <tr><th>Field</th><th>What to enter</th></tr>
  <tr><td><b>Study Name</b></td>
      <td>A short label for this cohort (e.g. <code>SHPvsZA Male</code>).
          Used in output file names &mdash; avoid special characters.</td></tr>
  <tr><td><b>Cohort Sex</b></td>
      <td><code>Male</code>, <code>Female</code>, or <code>Mixed</code>.
          Choosing <b>Mixed</b> splits all CSV outputs into separate Male / Female
          folders and enables per-prefix sex assignment in the Group Map.</td></tr>
  <tr><td><b>Cohort Age</b></td>
      <td>Optional label (e.g. <code>16 Weeks</code>). Appended to output file names
          so you can tell batches apart.</td></tr>
  <tr><td><b>Raw Data Folder</b></td>
      <td>The folder that contains the BioDent <code>.txt</code> export files for this
          run. After processing, every <code>.txt</code> is moved automatically into
          an <code>Analyzed_Files</code> subfolder so it won&rsquo;t be re-processed
          next time.</td></tr>
  <tr><td><b>Master File (.xlsx) &mdash; optional</b></td>
      <td>Point to an existing Excel master to update it, or leave <i>blank</i> to
          auto-create a new master file in the Output Folder.
          The auto-created file is named <code>&lt;StudyName&gt;_Master.xlsx</code>.</td></tr>
  <tr><td><b>Output Folder</b></td>
      <td>Where the master file and all CSV summaries are written.</td></tr>
</table>

<h2>2 &mdash; Group Map</h2>
<p>The Group Map tells the pipeline which ID prefix belongs to which experimental group.
Enter one row per group:</p>
<ul>
  <li><b>Prefix</b> &mdash; the letters at the start of a mouse ID (e.g. <code>ZC</code>,
      <code>ZT</code>, <code>SHP</code>). Matching is case-insensitive and tolerates
      separators (<code>ZC-1M</code>, <code>ZC_1M</code>, and <code>ZC1M</code> all
      match prefix <code>ZC</code>).</li>
  <li><b>Group Name</b> &mdash; the label that appears in CSV column headers
      (e.g. <code>Zinc Control</code>).</li>
  <li><b>Sex</b> (Mixed cohorts only) &mdash; the sex assigned to every animal with
      this prefix. Check <b>&ldquo;Deduce sex from ID&rdquo;</b> to read the trailing
      <code>M</code> or <code>F</code> from each animal&rsquo;s ID automatically
      instead of using the dropdown.</li>
</ul>
<p class="note">IDs that don&rsquo;t match any prefix are excluded from analysis.</p>

<h2>3 &mdash; Automatic Data Cleaning</h2>
<p>The pipeline applies these cleaning steps <i>before</i> any averages are calculated.
You will not be prompted &mdash; these happen silently.</p>

<h3>3a &mdash; &ldquo;Ignore / Do Not Use&rdquo; flags</h3>
<p>Any measurement row whose <b>Notes</b> column contains one of the following words or
phrases is <b>dropped entirely</b>:</p>
<ul>
  <li><code>ignore</code> &nbsp;(or partial: <code>ignor</code>)</li>
  <li><code>do not use</code></li>
  <li><code>disregard</code></li>
  <li><code>don't</code></li>
</ul>
<p><b>Example:</b> Notes = <code>do not use &mdash; probe slipped</code> &rarr;
that row is removed from all calculations.</p>

<h3>3b &mdash; &ldquo;Actually&rdquo; in the ID field</h3>
<p>Sometimes an ID cell was typed with noise after the fact.
The pipeline strips the noise and uses the clean ID:</p>
<ul>
  <li><code>ZC1M-actual</code> &rarr; <code>ZC1M</code></li>
  <li><code>ZC1M actual</code> &rarr; <code>ZC1M</code></li>
  <li><code>Actually ZC1M</code> &rarr; <code>ZC1M</code></li>
</ul>

<h3>3c &mdash; &ldquo;ACTUALLY &lt;ID&gt;&rdquo; in Notes</h3>
<p>If a Notes cell contains <code>ACTUALLY &lt;correct ID&gt;</code>, the pipeline
<b>renames that measurement &mdash; and all other measurements sharing the same
original wrong ID</b> &mdash; to the correct animal. This handles the common case
where measurement #2, #3, and #4 carry the wrong label but only one row has a note.</p>
<p><b>Example:</b> ID = <code>ZC1M</code>, Notes = <code>actually ZC2M</code><br>
&rarr; <i>all</i> rows labelled <code>ZC1M</code> in that file are renamed
<code>ZC2M</code>.</p>
<p class="note">The first correction found wins when the same wrong ID appears in
multiple notes with different claims.</p>

<h2>4 &mdash; Interactive Prompts During Processing</h2>
<p>After cleaning, the pipeline averages each animal&rsquo;s measurements. The
following dialogs appear when the data needs your judgement.</p>

<h3>4a &mdash; Insufficient Measurements (&lt; 4 rows)</h3>
<p>A well-formed BioDent run has exactly 4 measurements per animal.
If an animal has fewer, you are asked:</p>
<table>
  <tr><th>Button</th><th>What it does</th></tr>
  <tr><td><b>Average Anyway</b></td>
      <td>Include this animal in the analysis using however many measurements exist.
          The average is calculated from fewer data points &mdash; flag in your notes
          if this is intentional.</td></tr>
  <tr><td><b>Skip</b></td>
      <td>Exclude this animal entirely. It will not appear in the master file or
          any CSV output.</td></tr>
  <tr><td><b>Rename / Realign</b></td>
      <td>The measurements may belong to a different animal (wrong ID typed at the
          machine). You will be prompted to type the correct ID. All measurements
          are reassigned and the pipeline retries with the new ID.</td></tr>
</table>

<h3>4b &mdash; Too Many Measurements (&gt; 4 rows)</h3>
<p>A dialog shows all rows for the animal with a <b>Keep?</b> checkbox.
Tick exactly the 4 rows you want to keep, then click <b>Confirm Selection</b>.
Unchecked rows are discarded. This usually happens when a measurement run was
accidentally repeated.</p>

<h3>4c &mdash; Mouse Not Found in Master (during sync)</h3>
<p>When writing averages back to the master Excel file, if a processed ID doesn&rsquo;t
match any row in the sheet:</p>
<table>
  <tr><th>Button</th><th>What it does</th></tr>
  <tr><td><b>Rename / Retry</b></td>
      <td>Type the ID <i>exactly</i> as it appears in the master Excel file
          (including any capitalisation or separators). The pipeline retries
          the match and writes the data if found.</td></tr>
  <tr><td><b>Skip</b></td>
      <td>Do not write this animal to the master. Its averaged values are lost
          for this run.</td></tr>
</table>

<h2>5 &mdash; Output Files</h2>
<ul>
  <li><b>&lt;StudyName&gt;_Master.xlsx</b> &mdash; one row per animal with all
      10 averaged metrics. Updated on every run.</li>
  <li><b>&lt;StudyName&gt;_CSVFiles / &lt;StudyName&gt;_Summary.csv</b> &mdash;
      flat summary of every animal with Group (and Sex if Mixed).</li>
  <li><b>CSVFiles / Per_Metric / &lt;metric&gt;.csv</b> &mdash; one pivot table
      per BioDent metric, columns = groups, rows = animal IDs. Ready for
      copy-paste into GraphPad / SPSS.</li>
  <li><b>Analyzed_Files/</b> (inside Raw Data Folder) &mdash; processed
      <code>.txt</code> files are moved here so they&rsquo;re not processed again.</li>
</ul>
"""


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("BioDent User Guide")
        self.setMinimumSize(760, 600)
        self.resize(800, 660)
        self.setWindowModality(Qt.ApplicationModal)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setHtml(_HELP_HTML)
        layout.addWidget(browser)

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(close_btn)
        layout.addLayout(row)


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

        bottom_row = QHBoxLayout()

        help_btn = QPushButton("? Help / User Guide")
        help_btn.setFixedHeight(44)
        help_btn.setFont(QFont("Arial", 10))
        help_btn.clicked.connect(self._show_help)
        bottom_row.addWidget(help_btn)

        run_btn = QPushButton("Run Selected Study")
        run_btn.setFixedHeight(44)
        run_btn.setFont(QFont("Arial", 11, QFont.Bold))
        run_btn.setStyleSheet(
            "QPushButton { background-color: #1976D2; color: white; border-radius: 5px; }"
            "QPushButton:hover { background-color: #1565C0; }"
        )
        run_btn.clicked.connect(self._run_selected)
        bottom_row.addWidget(run_btn, stretch=1)

        layout.addLayout(bottom_row)

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

    def _show_help(self):
        dlg = HelpDialog(parent=self)
        dlg.exec_()

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
