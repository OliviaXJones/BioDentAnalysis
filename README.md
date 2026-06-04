# BioDent Analysis Pipeline

Pipeline for processing BioDent indentation data: cleans raw `.txt` exports,
averages measurements per animal, syncs to a master Excel file, and generates
per-metric CSV pivot tables ready for GraphPad Prism / SPSS.

---

## Quick Start

**Double-click `BioDent.bat`** (or run `python BioDent_Main.py`) to open the launcher.

Or download a pre-built `BioDent.exe` from the [Releases](../../releases) page — place it
in the same folder as `studies_config.json` and double-click.

---

## Setting Up a Study

Click **Add Study** to create a new configuration. All settings are saved to
`studies_config.json` and reused on future runs.

| Field | What to enter |
|---|---|
| **Study Name** | Short label for this cohort (e.g. `SHPvsZA Male`). Used in output file names — avoid special characters. |
| **Cohort Sex** | `Male`, `Female`, or `Mixed`. Choosing **Mixed** splits all CSV outputs into separate Male/Female folders and enables per-prefix sex assignment in the Group Map. |
| **Cohort Age** | Optional label (e.g. `16 Weeks`). Appended to output file names. |
| **Raw Data Folder** | Folder containing the BioDent `.txt` export files. After processing, every `.txt` is moved into an `Analyzed_Files` subfolder automatically so it won't be re-processed next time. |
| **Master File (.xlsx) — optional** | Point to an existing Excel master to update it, or leave **blank** to auto-create a new master file in the Output Folder (`<StudyName>_Master.xlsx`). |
| **Output Folder** | Where the master file and all CSV summaries are written. |

---

## Group Map

Maps an ID prefix to an experimental group name for statistical grouping in CSVs.

- **Prefix** — the letters at the start of a mouse ID (e.g. `ZC`, `ZT`, `SHP`). Matching is case-insensitive and tolerates separators (`ZC-1M`, `ZC_1M`, and `ZC1M` all match prefix `ZC`).
- **Group Name** — the label that appears in CSV column headers (e.g. `Zinc Control`).
- **Sex** (Mixed cohorts only) — sex assigned to every animal with this prefix. Check **"Deduce sex from ID"** to read the trailing `M` or `F` from each animal's ID automatically.

> IDs that don't match any prefix are excluded from analysis.

---

## Automatic Data Cleaning

These steps happen silently before any averages are calculated.

### Ignore / Do Not Use flags

Any measurement row whose **Notes** column contains one of the following words is **dropped entirely**:

- `ignore` (or partial: `ignor`)
- `do not use`
- `disregard`
- `don't`

**Example:** Notes = `do not use — probe slipped` → that row is removed from all calculations.

### "Actually" in the ID field

Sometimes an ID cell was typed with noise. The pipeline strips it automatically:

| Raw ID in cell | Cleaned to |
|---|---|
| `ZC1M-actual` | `ZC1M` |
| `ZC1M actual` | `ZC1M` |
| `Actually ZC1M` | `ZC1M` |

### "ACTUALLY \<ID\>" in Notes

If a Notes cell contains `ACTUALLY <correct ID>`, the pipeline **renames that
measurement — and all other measurements sharing the same original wrong ID —** to
the correct animal. This handles the common case where measurements #2, #3, and #4
carry the wrong label but only one row has a note.

**Example:** ID = `ZC1M`, Notes = `actually ZC2M`
→ *all* rows labelled `ZC1M` in that file are renamed `ZC2M`.

> The first correction found wins when the same wrong ID appears in multiple notes with different claims.

---

## Interactive Prompts During Processing

After cleaning, the pipeline averages each animal's measurements. These dialogs appear
when the data needs your judgement.

### Insufficient Measurements (< 4 rows)

A well-formed BioDent run has exactly 4 measurements per animal. If an animal has fewer:

| Button | What it does |
|---|---|
| **Average Anyway** | Include this animal using however many measurements exist. The average is calculated from fewer data points — flag in your notes if intentional. |
| **Skip** | Exclude this animal entirely. It will not appear in the master file or any CSV output. |
| **Rename / Realign** | The measurements may belong to a different animal (wrong ID at the machine). You will be prompted to type the correct ID. All measurements are reassigned and the pipeline retries. |

### Too Many Measurements (> 4 rows)

A dialog shows all rows for the animal with a **Keep?** checkbox.
Tick exactly the 4 rows you want to keep, then click **Confirm Selection**.
Unchecked rows are discarded. This usually happens when a run was accidentally repeated.

### Mouse Not Found in Master (during sync)

When writing averages back to the master Excel file, if a processed ID doesn't match any row:

| Button | What it does |
|---|---|
| **Rename / Retry** | Type the ID *exactly* as it appears in the master Excel file (including capitalisation / separators). The pipeline retries the match and writes the data if found. |
| **Skip** | Do not write this animal to the master. Its averaged values are lost for this run. |

---

## Output Files

| File / Folder | Contents |
|---|---|
| `<StudyName>_Master.xlsx` | One row per animal with all 10 averaged metrics. Updated on every run. |
| `<StudyName>_CSVFiles/<StudyName>_Summary.csv` | Flat summary of every animal with Group (and Sex if Mixed). |
| `CSVFiles/Per_Metric/<metric>.csv` | One pivot table per BioDent metric — columns = groups, rows = animal IDs. Ready for copy-paste into GraphPad / SPSS. |
| `Analyzed_Files/` (inside Raw Data Folder) | Processed `.txt` files moved here so they're not re-processed next time. |

---

## Building a Standalone EXE

Run **`build_exe.bat`** from the project folder. It will install PyInstaller if needed
and produce `dist/BioDent.exe`. Distribute that file together with a
`studies_config.json` (can be empty: `{"studies": []}`) in the same folder.

---

## Dependencies

```
pip install PyQt5 pandas openpyxl
```
