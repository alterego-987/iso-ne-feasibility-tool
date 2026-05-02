# ISO-NE Feasibility Study Tool

A desktop application for performing **N-1 contingency feasibility analysis** on the ISO New England power grid. Given a proposed generation or storage project, the tool runs a redispatch algorithm against an N-1 power flow dataset to determine the maximum project size that keeps all monitored transmission elements within thermal limits (≤ 102% loading).

---

## What It Does

When a new energy project (solar, wind, battery storage, etc.) seeks interconnection to the ISO-NE grid, it must pass a feasibility study. The core question: after adding this project and redispatching the grid to compensate, does any transmission element become overloaded under an N-1 contingency?

This tool automates that analysis:

1. Loads an N-1 power flow Excel file (containing Flows and Dispatch tables with DFAX factors)
2. Applies the project at a specified bus with a specified MW size
3. Runs a redispatch algorithm that iteratively reduces generation at other buses to relieve overloads
4. Outputs a new Excel file with the updated dispatch and flow results
5. If no feasible dispatch exists at the full project size, the tool steps down in 5 MW increments until a feasible solution is found

Supports both **charging** (load increase) and **discharging** (generation injection) project modes for battery storage interconnection studies.

---

## Project Structure

```
iso-ne-feasibility-tool/
├── src/
│   ├── main.py          # PyQt5 GUI — entry point
│   ├── core_logic.py    # excelExtract, tableReformation, redispatch
│   ├── excel_writer.py  # writeExcel — runs the full study and writes output
│   └── config.py        # BOSTON_ZONES, EXCLUDED_PLANTS constants
├── ui/
│   └── feasibilitytool.ui   # Qt Designer UI file (reference)
├── archive/
│   └── redispatch_function.txt  # Original standalone draft
├── NE FEASIBILITY TOOL - Benchmark and Test Results/
│   └── ...              # Benchmark and test run folder structure (data excluded)
├── requirements.txt
└── README.md
```

---

## Setup

**Requirements:** Python 3.9+

```bash
# Clone the repo
git clone https://github.com/alterego-987/iso-ne-feasibility-tool.git
cd iso-ne-feasibility-tool

# Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate        # macOS/Linux
.venv\Scripts\activate           # Windows

# Install dependencies
pip install -r requirements.txt
```

---

## Running the Tool

```bash
python src/main.py
```

### Usage

1. **Enter Bus Number** — the POI (Point of Interconnection) bus number for the project
2. **Enter Project Size (MW)** — the nameplate capacity of the project
3. **Select Operation Mode** — Charging (load) or Discharging (generation)
4. **Load Excel File** — select the N-1 power flow `.xlsx` file
5. **Run Redispatch** — the tool processes all sheets and writes a `redispatch_<filename>.xlsx` output file to the same directory as the input

The output file contains the updated dispatch table reflecting the feasible redispatch solution. An "Open in Folder" button appears on success for quick access.

---

## Input File Format

The input `.xlsx` file must follow the ISO-NE power flow export format with:
- A **Flows** block: transmission element flows, limits, and DFAX factors
- A **Dispatch** block: bus-level generation dispatch, bus numbers, zones, and DFAX factors

Each worksheet in the file represents a separate N-1 contingency case.

---

## Algorithm Overview

The redispatch algorithm (`src/core_logic.py`) works in two passes:

**Pass 1 — Primary redispatch:**
- Sorts generators by DFAX factor (highest impact first)
- For **discharging** projects: reduces generation at other buses to offset the MW injection
- For **charging** projects: increases generation at other buses to offset the load increase
- Skips excluded plants (nuclear, imports: Seabrook, Millstone, NYNE, NYPA, NBNE)
- For projects in Boston-area zones (`BOSTON_ZONES`), restricts redispatch to in-zone generators first

**Pass 2 — Residual delta cleanup:**
- If remaining redispatch delta > 0 after Pass 1, sweeps out-of-zone generators with low DFAX impact (|DFAX| ≤ 0.01)

Loading is evaluated as `FlowResult / Limit`. The study is feasible when all elements are ≤ 102%.

---

## Sample Data

A synthetic N-1 study file is included in `sample_data/` for testing and demonstration. All bus numbers, plant names, and transmission elements are fictional — no CEII or client data.

The sample workbook is formatted to resemble the compact TARA redispatch export structure: `PRDcalc` worksheets, five monitored flow rows, `Dispatch` summary row, five DFAX columns, formula-driven impact columns, autofilter, freeze pane, bordered tables, and highlighted editable `Pnew` cells.

```
sample_data/
├── Sample_N-1_Study.xlsx   # Ready-to-use input file
└── generate_sample.py      # Script to regenerate the file
```

**Run the tool with these inputs against the sample file:**

| Field | Value |
|-------|-------|
| Bus Number | `800001` |
| Project Size | `50` |
| Mode | Discharging |

**Expected behavior:**
- **Sheet 1 (LINE_ALP-BET_TRIP):** 50 MW raises flow loading to 1.032, triggering redispatch. `ALPHA_GAS_1` is reduced by 50 MW, bringing loading back to ~0.987.
- **Sheet 2 (LINE_GAM-DEL_TRIP):** 50 MW raises loading to 1.006 — within the 1.02 limit. No intermediate files generated.

Output file `redispatch_Sample_N-1_Study.xlsx` is written to the same folder as the input.

---

## Dependencies

| Package | Version |
|---------|---------|
| PyQt5 | 5.15.11 |
| pandas | 2.3.3 |
| numpy | 2.0.2 |
| openpyxl | 3.1.5 |

---

## Notes

- Benchmark and test result data files (`.xlsx`, project detail `.txt`) are excluded from this repository as they contain client-specific information. The folder structure is preserved for reference.
- The tool was developed for internal use during ISO-NE interconnection feasibility studies at Daymark Energy Advisors.
