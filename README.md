# BOACTAR Runbook

BOA CTA Report data processing pipeline. Extracts trend analysis data from BOA Excel reports covering Equities, Bonds, Commodities, and FX instruments, and produces standardized DATA, META, and ZIP output files.

## Project Structure

```
BOACTAR/
├── orchestrator.py       # Main entry point - runs the full pipeline
├── config.py             # Central configuration (paths, mappings, settings)
├── data_loader.py        # Loads and cleans input Excel files
├── parser.py             # Parses data and merges with master file
├── file_generator.py     # Generates DATA, META, ZIP output files
├── logger_setup.py       # Logging configuration
├── requirements.txt      # Python dependencies
├── input/                # Place input .xlsm files here
├── output/               # Generated output (timestamped + latest)
├── Master_file/          # Cumulative master data file
├── logs/                 # Timestamped log files
└── project_information/  # Reference docs and sample files
```

## Setup

```bash
pip install -r requirements.txt
```

**Dependencies:** openpyxl, pandas, python-dateutil

## Usage

1. Place one or more input `.xlsm` files in the `input/` folder
2. Run the pipeline:

```bash
python orchestrator.py
```

The pipeline will:
- Scan the `input/` folder for all `.xlsm` / `.xlsx` files
- Sort files by data date (oldest first)
- Process each file in chronological order
- Skip duplicate dates automatically
- Merge new data into the cumulative master file
- Generate output files in `output/<timestamp>/` and `output/latest/`

## Input Files

- Excel files (`.xlsm`) from the BOA CTA Report
- Must contain a **Summary** tab with trend analysis data
- Date is read from cell C2
- Data starts at row 6 with columns for asset class, underlying, trend strengths, projections, trigger ranges, and current price

## Output Files

Each run produces three files:

| File | Description |
|------|-------------|
| `BOACTAR_DATA_<timestamp>.xlsx` | Row 1: column codes, Row 2: descriptions, Row 3+: data |
| `BOACTAR_META_<timestamp>.xlsx` | Time series metadata (frequency, unit type, source, etc.) |
| `BOACTAR_<timestamp>.zip` | Archive containing DATA and META files |

## Asset Coverage

- **Equities:** 18 indices (S&P 500, NASDAQ-100, DAX, Nikkei, etc.)
- **Bonds:** 15 futures (US Treasuries, Bunds, Gilts, JGBs, etc.)
- **Commodities:** 20 instruments (Gold, Oil, Copper, Corn, etc.)
- **FX:** 26 currency pairs (EUR/USD, USD/JPY, EM currencies, etc.)

**Total output columns:** 1,157 (17 fields per asset x ~68 assets + SPOT prices)

## Data Fields Per Asset

| Field | Type |
|-------|------|
| Current Price (SPOT) | Price level |
| Short/Medium/Long Term Trend Strength | Percentage (x100) |
| 5-day Trend Projection (Bearish/Median/Bullish) | Percentage (x100) |
| Sell Trigger Range (Low/High) | Price level |
| Buy Trigger Range (Low/High) | Price level |

## Configuration

All settings are in `config.py`:
- **ASSET_MAPPING** - Maps input underlying names to output column codes
- **INPUT_COLUMNS** - Maps field names to input Excel column numbers
- **PERCENTAGE_FIELDS** - Fields that get multiplied by 100 during processing
- **ASSET_NAME_ALIASES** - Handles name variations in input files
