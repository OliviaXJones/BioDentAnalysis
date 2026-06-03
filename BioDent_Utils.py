import re
import pandas as pd

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QHeaderView, QPushButton, QMessageBox, QInputDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

# --- SHARED CONSTANTS ---

COLUMNS_TO_AVERAGE = [
    "1st Cycle Touchdown Distance (TDD 1st) - um",
    "Avg Loading Slope (Avg LS 1st-L) - N/um",
    "Avg Unloading Slope (Avg US 1st-L) - N/um",
    "Avg Energy Dissipated (Avg ED 3rd-L) - uJ",
    "Avg Creep Indentation Distance (Avg CID 1st-L) - um",
    "Indentation Distance Increase (IDI 1st-L) - um",
    "Total Indentation Distance (TID 1st-L) - um",
    "1st Cycle Creep Indentation Distance (CID 1st) - um",
    "1st Cycle Unloading Slope (US 1st) - N/um",
    "1st Cycle Indentation Distance (ID 1st) - um"
]

FKBP5_ID_REGEX = r"2?[WMH][.\-_]?\d+[.\-_]?[MF][.\-_]?\d+"

# --- SHARED HELPERS ---

def get_sex(code):
    return 'F' if 'F' in str(code).upper() else 'M'


def _is_valid_id(candidate, id_pattern=None, group_map=None):
    """Return True if candidate looks like a real mouse ID for this study."""
    if not candidate:
        return False
    if id_pattern and re.search(id_pattern, candidate, re.IGNORECASE):
        return True
    if group_map:
        for prefix in group_map:
            if re.match(re.escape(prefix) + r'[-_.\s]?', candidate, re.IGNORECASE):
                return True
    return False


def _find_id_in_text(text, id_pattern=None, group_map=None):
    """
    Scan arbitrary text for a valid mouse ID.
    Returns a squashed (no separator chars) ID string, or None.
    For id_pattern studies, searches the full text with the regex.
    For group_map studies, looks for a known prefix (with optional separator)
    followed by alphanumerics.
    """
    text_upper = str(text).strip().upper()
    if id_pattern:
        m = re.search(id_pattern, text_upper)
        if m:
            return re.sub(r"[.\-_]", "", m.group(0))
    if group_map:
        for prefix in sorted(group_map.keys(), key=len, reverse=True):
            m = re.search(
                re.escape(prefix.upper()) + r'[-_.\s]?[A-Z0-9]*', text_upper)
            if m:
                return m.group(0)
    return None


def _clean_id_field(raw_id, id_pattern=None, group_map=None):
    """
    Strip 'actual' / 'actually' noise from the ID field itself.
    Handles all three forms: 'ZC1M-actual', 'ZC1M actual', 'Actually ZC1M'.
    Falls back to extracting a valid code embedded in the noisy string.
    """
    s = str(raw_id).strip()
    # Strip leading "Actually ..."
    cleaned = re.sub(r'(?i)^actually[\s\-_]+', '', s).strip()
    # Strip trailing "... actual(ly)"
    cleaned = re.sub(r'(?i)[\s\-_]+actual(?:ly)?$', '', cleaned).strip()
    if cleaned and cleaned.upper() != s.upper():
        return cleaned
    # Fall back: find a valid ID embedded in the noisy field
    found = _find_id_in_text(s, id_pattern=id_pattern, group_map=group_map)
    if found:
        found_squashed = re.sub(r"[.\-_]", "", found.upper())
        s_squashed     = re.sub(r"[.\-_]", "", s.upper())
        if found_squashed != s_squashed:
            return found
    return s


def _apply_note_corrections(df, id_pattern=None, group_map=None):
    """
    Scan Notes for 'ACTUALLY <ID>' (explicit keyword, both studies) or a bare
    FKBP5 pattern match anywhere in Notes (id_pattern studies only, preserving
    the existing behaviour).

    When a row's ID is corrected, ALL rows sharing the same original wrong ID
    are renamed too — this handles #2 / #3 / #4 measurements that carry the
    wrong label but have no note of their own.

    First-found correction wins when the same wrong ID appears in multiple notes
    with conflicting claims (rare, but deterministic).
    """
    corrections = {}  # original_id_upper -> corrected_id

    for _, row in df.iterrows():
        notes_upper   = str(row['Notes']).strip().upper()
        current_upper = str(row['ID']).strip().upper()

        # Explicit "ACTUALLY <ID>" — works for both study types
        m = re.search(r'ACTUALLY\s+([A-Z0-9][A-Z0-9.\-_]*)', notes_upper)
        if m:
            # Strip trailing measurement markers like "#1", "# 2"
            raw_candidate = re.sub(r'[#\s]*\d+\s*$', '', m.group(1)).rstrip('. -_')
            candidate = re.sub(r'[.\-_]', '', raw_candidate)
            if _is_valid_id(candidate, id_pattern, group_map) and candidate != current_upper:
                corrections.setdefault(current_upper, candidate)
                continue

        # Fallback for FKBP5: bare regex anywhere in Notes (existing behaviour)
        if id_pattern:
            found = _find_id_in_text(notes_upper, id_pattern=id_pattern)
            if found and found != current_upper:
                corrections.setdefault(current_upper, found)

    if corrections:
        df = df.copy()
        df['ID'] = df['ID'].apply(
            lambda x: corrections.get(str(x).strip().upper(), x)
        )
    return df


# --- MANUAL SELECTION DIALOG ---

class ManualSelectionDialog(QDialog):
    def __init__(self, mouse_id, group, target_col):
        super().__init__()
        self.setWindowTitle(f"Data Selection: {mouse_id}")
        self.setMinimumSize(960, 480)
        self.setWindowModality(Qt.ApplicationModal)
        self._group_index = list(group.index)
        self._selected = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        label = QLabel(f"Select the 4 rows to keep for <b>{mouse_id}</b>:")
        label.setFont(QFont("Arial", 12))
        layout.addWidget(label)

        self._table = QTableWidget()
        self._table.setColumnCount(5)
        self._table.setHorizontalHeaderLabels(
            ["Index", "Meas #", "ID 1st (um)", "Notes", "Keep?"])
        self._table.setRowCount(len(group))
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionMode(QTableWidget.NoSelection)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self._table.setColumnWidth(0, 70)
        self._table.setColumnWidth(1, 80)
        self._table.setColumnWidth(2, 130)
        self._table.setColumnWidth(4, 80)

        for row_i, (idx, row) in enumerate(group.iterrows()):
            self._table.setItem(row_i, 0, QTableWidgetItem(str(idx)))
            self._table.setItem(row_i, 1, QTableWidgetItem(str(row["Measurement #"])))
            self._table.setItem(row_i, 2, QTableWidgetItem(str(row[target_col])))
            self._table.setItem(row_i, 3, QTableWidgetItem(str(row["Notes"])[:80]))

            keep = QTableWidgetItem()
            keep.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            keep.setCheckState(Qt.Unchecked)
            keep.setTextAlignment(Qt.AlignCenter)
            self._table.setItem(row_i, 4, keep)

        layout.addWidget(self._table)

        confirm_btn = QPushButton("Confirm Selection")
        confirm_btn.setFixedHeight(40)
        confirm_btn.setFont(QFont("Arial", 11, QFont.Bold))
        confirm_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; border-radius: 5px; }"
            "QPushButton:hover { background-color: #43A047; }"
        )
        confirm_btn.clicked.connect(self._on_confirm)
        layout.addWidget(confirm_btn)

    def _on_confirm(self):
        self._selected = [
            self._group_index[r]
            for r in range(self._table.rowCount())
            if self._table.item(r, 4).checkState() == Qt.Checked
        ]
        self.accept()

    def get_selected_indices(self):
        return self._selected


# --- PROMPT HELPERS ---

def ask_small_group_action(mouse_id, count):
    """Returns 'a' (average anyway), 's' (skip), or 'r' (rename)."""
    msg = QMessageBox()
    msg.setWindowTitle("Insufficient Measurements")
    msg.setText(f"<b>{mouse_id}</b> only has {count} measurement(s).")
    msg.setInformativeText("How would you like to proceed?")
    avg_btn    = msg.addButton("Average Anyway",   QMessageBox.AcceptRole)
    skip_btn   = msg.addButton("Skip",             QMessageBox.RejectRole)
    rename_btn = msg.addButton("Rename / Realign", QMessageBox.ActionRole)
    msg.setDefaultButton(avg_btn)
    msg.exec_()
    clicked = msg.clickedButton()
    if clicked == rename_btn:
        return 'r'
    if clicked == skip_btn:
        return 's'
    return 'a'


def ask_new_id(mouse_id, auto_detected=None):
    """Returns a corrected mouse ID string, or None if cancelled."""
    hint = f"Auto-detected ID from notes: {auto_detected}\n\n" if auto_detected else ""
    text, ok = QInputDialog.getText(
        None, "Rename Mouse ID",
        f"{hint}Enter the correct ID for {mouse_id}:",
        text=auto_detected or ""
    )
    if ok and text.strip():
        return text.strip().upper()
    return None


def ask_not_found_action(mouse_code):
    """Returns a new ID string to retry with, or None to skip."""
    msg = QMessageBox()
    msg.setWindowTitle("Mouse Not Found in Master")
    msg.setText(f"'{mouse_code}' was not found in any sheet.")
    msg.setInformativeText("Would you like to rename and retry, or skip this mouse?")
    rename_btn = msg.addButton("Rename / Retry", QMessageBox.AcceptRole)
    skip_btn   = msg.addButton("Skip",           QMessageBox.RejectRole)
    msg.setDefaultButton(skip_btn)
    msg.exec_()

    if msg.clickedButton() == rename_btn:
        text, ok = QInputDialog.getText(
            None, "Enter Mouse ID",
            "Enter the ID exactly as it appears in Excel:"
        )
        if ok and text.strip():
            return text.strip()
    return None


# --- DATA CLEANING & AVERAGING ---

def clean_and_average_data(txt_file_paths, id_pattern=None, group_map=None):
    """
    Read, clean, and average BioDent .txt files.

    id_pattern  – regex string used by FKBP5-style studies (pass FKBP5_ID_REGEX)
    group_map   – prefix→group dict used by single-study pipelines

    Cleaning steps applied to both study types:
      1. Strip 'actual'/'actually' noise from the ID field itself.
      2. Scan Notes for 'ACTUALLY <ID>' and rename ALL rows sharing the same
         wrong original ID (propagates the fix to #2/#3/#4 measurements).
      3. For id_pattern studies, also scan Notes for a bare pattern match
         (existing FKBP5 behaviour).
      4. Drop duplicates, drop rows flagged 'ignore'/'do not use'/etc.
      5. Filter to known group_map prefixes (single-study only).
      6. Per-mouse averaging; >4 → manual selection dialog,
         <4 → average/skip/rename prompt.
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

    # Step 1: clean the ID field itself ("ZC1M-actual" → "ZC1M", etc.)
    df['ID'] = df['ID'].apply(
        lambda x: _clean_id_field(x, id_pattern=id_pattern, group_map=group_map)
    )

    # Step 2+3: correct mislabeled IDs via Notes; bulk-rename same-ID siblings
    df = _apply_note_corrections(df, id_pattern=id_pattern, group_map=group_map)

    target_val_col = "1st Cycle Indentation Distance (ID 1st) - um"
    df = df.drop_duplicates(subset=['ID', 'Measurement #', target_val_col], keep='first')

    ignore_keywords = ["ignore", "do not use", "disregard", "don't", "ignor"]
    mask = df['Notes'].str.contains('|'.join(ignore_keywords), case=False, na=False)
    df_cleaned = df[~mask].copy()
    df_cleaned = df_cleaned.dropna(subset=['ID'])
    df_cleaned['ID'] = df_cleaned['ID'].astype(str)

    # Step 5: keep only IDs that match a known group_map prefix (single-study)
    if group_map:
        def _prefix_matches(mouse_id):
            for prefix in group_map:
                if re.match(re.escape(prefix) + r'[-_.\s]?', str(mouse_id).strip(),
                            re.IGNORECASE):
                    return True
            return False
        df_cleaned = df_cleaned[df_cleaned['ID'].apply(_prefix_matches)]

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
                auto_id = None
                for _, row in group.iterrows():
                    found = _find_id_in_text(
                        str(row['Notes']), id_pattern=id_pattern, group_map=group_map
                    )
                    if found and found.upper() != mouse_id.strip().upper():
                        auto_id = found
                        break
                new_id = ask_new_id(mouse_id_str, auto_id)
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
