"""
BOACTAR Runbook Configuration
=============================
Central configuration file for the BOA CTA Report data processing pipeline.

This runbook processes Excel files from BOA containing trend analysis data
for Equities, Bonds, Commodities, and FX instruments.
"""

import os
from datetime import datetime

# =============================================================================
# PATHS CONFIGURATION
# =============================================================================

# Base directory (where this script is located)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Input/Output directories
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
MASTER_DIR = os.path.join(BASE_DIR, 'Master_file')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')

# Master data file (your existing master file)
MASTER_DATA_FILE = os.path.join(MASTER_DIR, 'BOACTAR_DATA_20251229 (27).xlsx')

# Historical data file (same as master for initialization)
HISTORICAL_DATA_FILE = os.path.join(MASTER_DIR, 'BOACTAR_DATA_20251229 (27).xlsx')

# =============================================================================
# INPUT FILE CONFIGURATION
# =============================================================================

# Sheet name to read from input files
INPUT_SHEET_NAME = 'Summary'

# Input file structure - cell locations
INPUT_DATE_CELL = (2, 3)  # Row 2, Column C (0-indexed: row 1, col 2)

# Header row locations (1-indexed as in Excel)
INPUT_HEADER_ROW_1 = 4  # Contains "5-day Trend Projection", etc.
INPUT_HEADER_ROW_2 = 5  # Contains column names like "Asset Class", "Underlying", etc.
INPUT_DATA_START_ROW = 6  # First row of actual data

# Input column mapping (1-indexed column numbers)
INPUT_COLUMNS = {
    'ASSET_CLASS': 2,           # Column B
    'UNDERLYING': 3,            # Column C
    'SHORT_TERM_STRENGTH': 4,   # Column D
    'SHORT_BEARISH': 6,         # Column F
    'SHORT_MEDIAN': 7,          # Column G
    'SHORT_BULLISH': 8,         # Column H
    'MEDIUM_TERM_STRENGTH': 10, # Column J
    'MEDIUM_BEARISH': 12,       # Column L
    'MEDIUM_MEDIAN': 13,        # Column M
    'MEDIUM_BULLISH': 14,       # Column N
    'LONG_TERM_STRENGTH': 16,   # Column P
    'LONG_BEARISH': 18,         # Column R
    'LONG_MEDIAN': 19,          # Column S
    'LONG_BULLISH': 20,         # Column T
    'CURRENT_PRICE': 21,        # Column U
    'SELL_TRIGGER_LOW': 23,     # Column W
    'SELL_TRIGGER_HIGH': 24,    # Column X
    'BUY_TRIGGER_LOW': 26,      # Column Z
    'BUY_TRIGGER_HIGH': 27,     # Column AA
}

# =============================================================================
# OUTPUT FIELD DEFINITIONS
# =============================================================================

# Field types for trend data (in order for each asset)
TREND_FIELDS = [
    ('SHORTTERMTRENDSTRENGTH', 'Short Term Trend Strength'),
    ('SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH', 'Short Term Trend Strength 5-day Trend Projection, Bearish'),
    ('SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN', 'Short Term Trend Strength 5-day Trend Projection, Median'),
    ('SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH', 'Short Term Trend Strength 5-day Trend Projection, Bullish'),
    ('MEDIUMTERMTRENDSTRENGTH', 'Medium Term Trend Strength'),
    ('MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH', 'Medium Term Trend Strength 5-day Trend Projection, Bearish'),
    ('MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN', 'Medium Term Trend Strength 5-day Trend Projection, Median'),
    ('MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH', 'Medium Term Trend Strength 5-day Trend Projection, Bullish'),
    ('LONGTERMTRENDSTRENGTH', 'Long Term Trend Strength'),
    ('LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH', 'Long Term Trend Strength 5-day Trend Projection, Bearish'),
    ('LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN', 'Long Term Trend Strength 5-day Trend Projection, Median'),
    ('LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH', 'Long Term Trend Strength 5-day Trend Projection, Bullish'),
    ('SELLTRIGGERRANGE.LOW', 'Sell Trigger Range, Low'),
    ('SELLTRIGGERRANGE.HIGH', 'Sell Trigger Range, High'),
    ('BUYTRIGGERRANGE.LOW', 'Buy Trigger Range, Low'),
    ('BUYTRIGGERRANGE.HIGH', 'Buy Trigger Range, High'),
]

# Mapping from input column key to trend field code
INPUT_TO_TREND_FIELD = {
    'SHORT_TERM_STRENGTH': 'SHORTTERMTRENDSTRENGTH',
    'SHORT_BEARISH': 'SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH',
    'SHORT_MEDIAN': 'SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN',
    'SHORT_BULLISH': 'SHORTTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH',
    'MEDIUM_TERM_STRENGTH': 'MEDIUMTERMTRENDSTRENGTH',
    'MEDIUM_BEARISH': 'MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH',
    'MEDIUM_MEDIAN': 'MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN',
    'MEDIUM_BULLISH': 'MEDIUMTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH',
    'LONG_TERM_STRENGTH': 'LONGTERMTRENDSTRENGTH',
    'LONG_BEARISH': 'LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BEARISH',
    'LONG_MEDIAN': 'LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.MEDIAN',
    'LONG_BULLISH': 'LONGTERMTRENDSTRENGTH5DAYTRENDPROJECTION.BULLISH',
    'SELL_TRIGGER_LOW': 'SELLTRIGGERRANGE.LOW',
    'SELL_TRIGGER_HIGH': 'SELLTRIGGERRANGE.HIGH',
    'BUY_TRIGGER_LOW': 'BUYTRIGGERRANGE.LOW',
    'BUY_TRIGGER_HIGH': 'BUYTRIGGERRANGE.HIGH',
}

# =============================================================================
# ASSET MAPPING CONFIGURATION
# =============================================================================

# Mapping from input underlying name to output codes
# Format: 'Input Underlying Name': ('CATEGORY', 'SPOT_CODE', 'TREND_CODE', 'Description')
# Note: SPOT_CODE is used for Current Price column
#       TREND_CODE is used for all trend-related columns
ASSET_MAPPING = {
    # EQUITIES
    'S&P 500 (ES)': ('EQUITIES', 'SNP500', 'SP500ES', 'S&P 500 (ES)'),
    'NASDAQ-100 (NQ)': ('EQUITIES', 'NASDAQ100', 'NASDAQ100NQ', 'NASDAQ-100 (NQ)'),
    'Russell 2000 (RTY)': ('EQUITIES', 'RUSSELL2000', 'RUSSELL2000RTY', 'Russell 2000 (RTY)'),
    'EURO STOXX 50 (VG)': ('EQUITIES', 'SX5E', 'EUROSTOXX50VG', 'EURO STOXX 50 (VG)'),
    'EURO STOXX Banks (CA)': ('EQUITIES', 'SX7E', 'EUROSTOXXBANKSCA', 'EURO STOXX Banks (CA)'),
    'DAX (GX)': ('EQUITIES', 'DAX', 'DAXGX', 'DAX (GX)'),
    'Nikkei 225 (NK)': ('EQUITIES', 'NIKKEI225', 'NIKKEI225NK', 'Nikkei 225 (NK)'),
    'TOPIX (TP)': ('EQUITIES', 'TOPIX', 'TOPIXTP', 'TOPIX (TP)'),
    'MSCI EM (MES)': ('EQUITIES', 'MSCIEM', 'MSCIEMMES', 'MSCI EM (MES)'),
    'KOSPI 200 (KM)': ('EQUITIES', 'KOSPI200', 'KOSPI200KM', 'KOSPI 200 (KM)'),
    'Hang Seng (HI)': ('EQUITIES', 'HANGSENG', 'HANGSENGHI', 'Hang Seng (HI)'),
    'HSCEI (HC)': ('EQUITIES', 'HANGSENGCHINAENTERPRISES', 'HSCEIHC', 'HSCEI (HC)'),
    'TSWE (FT)': ('EQUITIES', 'TWSE', 'TSWEFT', 'TSWE (FT)'),
    'NIFTY (JGS)': ('EQUITIES', 'NIFTY', 'NIFTYJGS', 'NIFTY (JGS)'),
    'MSCI SG (QZ)': ('EQUITIES', 'MSCISG', 'MSCISGQZ', 'MSCI SG (QZ)'),
    'FTSE/JSE Top40 (AI)': ('EQUITIES', 'FTSEJSETOP40', 'FTSEJSETOP40AI', 'FTSE/JSE Top40 (AI)'),
    'Ibovespa (BZ)': ('EQUITIES', 'IBOVESPA', 'IBOVESPABZ', 'Ibovespa (BZ)'),
    'MEXBOL (IS)': ('EQUITIES', 'MEXBOL', 'MEXBOLIS', 'MEXBOL (IS)'),

    # BONDS
    '2yr Tsy Futures (TU)': ('BONDS', '2YRTSYFUTURESTU', '2YRTSYFUTURESTU', '2yr Tsy Futures (TU)'),
    '5yr Tsy Futures (FV)': ('BONDS', '5YRTSYFUTURESFV', '5YRTSYFUTURESFV', '5yr Tsy Futures (FV)'),
    '10yr Tsy Futures (TY)': ('BONDS', '10YRTSYFUTURESTY', '10YRTSYFUTURESTY', '10yr Tsy Futures (TY)'),
    '10yr Ultra Tsy Futures (UXY)': ('BONDS', '10YRULTRATSYFUTURESUXY', '10YRULTRATSYFUTURESUXY', '10yr Ultra Tsy Futures (UXY)'),
    'US Long Bond Futures (US)': ('BONDS', 'USLONGBONDFUTURESUS', 'USLONGBONDFUTURESUS', 'US Long Bond Futures (US)'),
    'Ultra Long Term US Bond Futures (WN)': ('BONDS', 'ULTRALONGTERMUSBONDFUTURESWN', 'ULTRALONGTERMUSBONDFUTURESWN', 'Ultra Long Term US Bond Futures (WN)'),
    'Canada 10yr Bond Futures (CN)': ('BONDS', 'CANADA10YRBONDFUTURES', 'CANADA10YRBONDFUTURESCN', 'Canada 10yr Bond Futures (CN)'),
    '2yr Schatz Futures (DU)': ('BONDS', '2YRSCHATZFUTURESDU', '2YRSCHATZFUTURESDU', '2yr Schatz Futures (DU)'),
    '5yr OBL Futures (OE)': ('BONDS', '5YROBLFUTURESOE', '5YROBLFUTURESOE', '5yr OBL Futures (OE)'),
    'Bund Futures (RX)': ('BONDS', 'BUNDFUTURESRX', 'BUNDFUTURESRX', 'Bund Futures (RX)'),
    '30yr Buxl Futures (UB)': ('BONDS', '30YRBUXLFUTURESUB', '30YRBUXLFUTURESUB', '30yr Buxl Futures (UB)'),
    'Gilt Futures (G )': ('BONDS', 'GILTFUTURESG', 'GILTFUTURESG', 'Gilt Futures (G)'),
    '10Yr JGB Futures (JB)': ('BONDS', '10YRJGBFUTURES', '10YRJGBFUTURESJB', '10Yr JGB Futures (JB)'),
    'Korea Treasury Bond (KAA)': ('BONDS', 'KOREATREASURYBONDKAA', 'KOREATREASURYBONDKAA', 'Korea Treasury Bond (KAA)'),
    'China 10yr Bond  Futures (TFT)': ('BONDS', 'CHINA10YRBONDFUTURESTFT', 'CHINA10YRBONDFUTURESTFT', 'China 10yr Bond Futures (TFT)'),

    # COMMODITIES
    'Gold (GC)': ('COMMODITIES', 'GOLD', 'GOLDGC', 'Gold (GC)'),
    'Oil (CL)': ('COMMODITIES', 'OIL', 'OILCL', 'Oil (CL)'),
    'Silver (SI)': ('COMMODITIES', 'SILVER', 'SILVERSI', 'Silver (SI)'),
    'Aluminum (LA)': ('COMMODITIES', 'ALUMINUM', 'ALUMINUMLA', 'Aluminum (LA)'),
    'Copper (BCOM, HG)': ('COMMODITIES', 'COPPERBCOMHG', 'COPPERBCOMHG', 'Copper (BCOM, HG)'),
    'Copper (GSCI, LP)': ('COMMODITIES', 'COPPERGSCILP', 'COPPERGSCILP', 'Copper (GSCI, LP)'),
    'Gasoline (XB)': ('COMMODITIES', 'GASOLINE', 'GASOLINEXB', 'Gasoline (XB)'),
    'Heating Oil (HO)': ('COMMODITIES', 'HEATINGOIL', 'HEATINGOILHO', 'Heating Oil (HO)'),
    'Gasoil (QS)': ('COMMODITIES', 'GASOIL', 'GASOILQS', 'Gasoil (QS)'),
    'Natural Gas (NG)': ('COMMODITIES', 'NATURALGAS', 'NATURALGASNG', 'Natural Gas (NG)'),
    'Soybeans (S )': ('COMMODITIES', 'SOYBEANS', 'SOYBEANSS', 'Soybeans (S )'),
    'Soybean Oil (BO)': ('COMMODITIES', 'SOYBEANOIL', 'SOYBEANOILBO', 'Soybean Oil (BO)'),
    'Soybean Meal (SM)': ('COMMODITIES', 'SOYBEANMEAL', 'SOYBEANMEALSM', 'Soybean Meal (SM)'),
    'Corn (C )': ('COMMODITIES', 'CORN', 'CORNC', 'Corn (C )'),
    'Wheat (W )': ('COMMODITIES', 'WHEAT', 'WHEATW', 'Wheat (W )'),
    'Sugar #11 (SB)': ('COMMODITIES', 'SUGAR11', 'SUGAR11SB', 'Sugar #11 (SB)'),
    'Cocoa (CC)': ('COMMODITIES', 'COCOA', 'COCOACC', 'Cocoa (CC)'),
    'Cotton (CT)': ('COMMODITIES', 'COTTON', 'COTTONCT', 'Cotton (CT)'),
    'Lean Hogs (LH)': ('COMMODITIES', 'LEANHOGS', 'LEANHOGSLH', 'Lean Hogs (LH)'),
    'Iron Ore (SCO)': ('COMMODITIES', 'IRONORESCO', 'IRONORESCO', 'Iron Ore (SCO)'),

    # FX
    'EUR/USD': ('FX', 'EURUSD', 'EURUSD', 'EUR/USD'),
    'GBP/USD': ('FX', 'GBPUSD', 'GBPUSD', 'GBP/USD'),
    'AUD/USD': ('FX', 'AUDUSD', 'AUDUSD', 'AUD/USD'),
    'CHF/USD': ('FX', 'CHFUSD', 'CHFUSD', 'CHF/USD'),
    'USD/JPY': ('FX', 'USDJPY', 'USDJPY', 'USD/JPY'),
    'USD/MXN': ('FX', 'USDMXN', 'USDMXN', 'USD/MXN'),
    'USD/CAD': ('FX', 'USDCAD', 'USDCAD', 'USD/CAD'),
    'USD/BRL': ('FX', 'USDBRL', 'USDBRL', 'USD/BRL'),
    'USD/CLP': ('FX', 'USDCLP', 'USDCLP', 'USD/CLP'),
    'USD/COP': ('FX', 'USDCOP', 'USDCOP', 'USD/COP'),
    'USD/PEN': ('FX', 'USDPEN', 'USDPEN', 'USD/PEN'),
    'USD/TRY': ('FX', 'USDTRY', 'USDTRY', 'USD/TRY'),
    'USD/ZAR': ('FX', 'USDZAR', 'USDZAR', 'USD/ZAR'),
    'USD/PLN': ('FX', 'USDPLN', 'USDPLN', 'USD/PLN'),
    'USD/CNH': ('FX', 'USDCNH', 'USDCNH', 'USD/CNH'),
    'USD/HKD': ('FX', 'USDHKD', 'USDHKD', 'USD/HKD'),
    'USD/SGD': ('FX', 'USDSGD', 'USDSGD', 'USD/SGD'),
    'USD/KRW': ('FX', 'USDKRW', 'USDKRW', 'USD/KRW'),
    'USD/INR': ('FX', 'USDINR', 'USDINR', 'USD/INR'),
    'USD/TWD': ('FX', 'USDTWD', 'USDTWD', 'USD/TWD'),
    'USD/IDR': ('FX', 'USDIDR', 'USDIDR', 'USD/IDR'),
    'USD/THB': ('FX', 'USDTHB', 'USDTHB', 'USD/THB'),
    'USD/PHP': ('FX', 'USDPHP', 'USDPHP', 'USD/PHP'),
    'USD/MYR': ('FX', 'USDMYR', 'USDMYR', 'USD/MYR'),
    'USD/CZK': ('FX', 'USDCZK', 'USDCZK', 'USD/CZK'),
    'USD/HUF': ('FX', 'USDHUF', 'USDHUF', 'USD/HUF'),
    'MEX/USD': ('FX', 'MEXUSD', 'MEXUSD', 'MEX/USD'),
}

# Additional name variations that may appear in input files
# These map to the canonical names in ASSET_MAPPING
ASSET_NAME_ALIASES = {
    # Variations with trailing spaces or different formatting
    'Ibovespa': 'Ibovespa (BZ)',
    'Ibovespa ': 'Ibovespa (BZ)',
    'NIKKEI 225 (NK)': 'Nikkei 225 (NK)',
    'China 10yr Bond Futures (TFT)': 'China 10yr Bond  Futures (TFT)',
    'Gilt Futures (G)': 'Gilt Futures (G )',
}

# =============================================================================
# OUTPUT COLUMN ORDER
# =============================================================================

# The exact order of output columns (loaded from master file)
# This will be populated dynamically by reading the master file
# to ensure exact column order matching
OUTPUT_COLUMN_ORDER = None  # Set by file_generator.py from master file

# =============================================================================
# DATA PROCESSING CONFIGURATION
# =============================================================================

# Percentage fields that need to be multiplied by 100
# These are the trend strength and projection fields (NOT trigger ranges or current price)
PERCENTAGE_FIELDS = [
    'SHORT_TERM_STRENGTH',
    'SHORT_BEARISH',
    'SHORT_MEDIAN',
    'SHORT_BULLISH',
    'MEDIUM_TERM_STRENGTH',
    'MEDIUM_BEARISH',
    'MEDIUM_MEDIAN',
    'MEDIUM_BULLISH',
    'LONG_TERM_STRENGTH',
    'LONG_BEARISH',
    'LONG_MEDIAN',
    'LONG_BULLISH',
]

# Fields that should NOT be converted (keep as-is)
NON_PERCENTAGE_FIELDS = [
    'CURRENT_PRICE',
    'SELL_TRIGGER_LOW',
    'SELL_TRIGGER_HIGH',
    'BUY_TRIGGER_LOW',
    'BUY_TRIGGER_HIGH',
]

# Value representing missing data in output
NA_VALUE = 'NA'

# Values in input that should be treated as NA
NA_INPUT_VALUES = ['-', '--', 'N/A', 'NA', '', None]

# =============================================================================
# DATE CONFIGURATION
# =============================================================================

# Output date format
DATE_FORMAT_OUTPUT = '%Y-%m-%d'

# Input date formats to try (in order)
INPUT_DATE_FORMATS = [
    '%Y-%m-%d %H:%M:%S',  # Excel datetime format
    '%Y-%m-%d',
    '%d/%m/%Y',
    '%m/%d/%Y',
    '%d-%m-%Y',
    '%m-%d-%Y',
]

# Excel date epoch (for serial number conversion)
EXCEL_DATE_EPOCH = datetime(1899, 12, 30)

# =============================================================================
# FILE NAMING CONFIGURATION
# =============================================================================

# File naming patterns
DATA_FILE_PREFIX = 'BOACTAR_DATA'
META_FILE_PREFIX = 'BOACTAR_META'
ZIP_FILE_PREFIX = 'BOACTAR'
MASTER_FILE_NAME = 'BOACTAR_MASTER_DATA.xlsx'

# Timestamp format for output folders/files
TIMESTAMP_FORMAT = '%Y%m%d_%H%M%S'

# Latest folder name
LATEST_FOLDER = 'latest'

# =============================================================================
# METADATA CONFIGURATION
# =============================================================================

# Default metadata values for META file
METADATA_DEFAULTS = {
    'FREQUENCY': 'B',  # Business daily
    'UNIT_TYPE': 'Index',
    'DATA_TYPE': 'Level',
    'SOURCE': 'BOA',
    'CATEGORY': 'CTA',
}

# =============================================================================
# LOGGING CONFIGURATION
# =============================================================================

# Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
LOG_LEVEL = 'INFO'

# Debug mode (enables more verbose logging)
DEBUG_MODE = False

# Log format
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

# =============================================================================
# PROCESSING OPTIONS
# =============================================================================

# Whether to rebuild master from historical file (for initialization)
REBUILD_MASTER = False

# Whether to continue processing if some assets fail
CONTINUE_ON_ERROR = True

# Whether to require all expected assets in input file
REQUIRE_ALL_ASSETS = False

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_timestamp():
    """Get current timestamp string for file naming."""
    return datetime.now().strftime(TIMESTAMP_FORMAT)


def get_spot_code(underlying_name):
    """Get the SPOT column code for an underlying."""
    canonical_name = ASSET_NAME_ALIASES.get(underlying_name, underlying_name)
    if canonical_name in ASSET_MAPPING:
        category, spot_code, _, _ = ASSET_MAPPING[canonical_name]
        return f'BOACTAR.{category}.{spot_code}.SPOT.B'
    return None


def get_trend_code(underlying_name, field_code):
    """Get the trend column code for an underlying and field."""
    canonical_name = ASSET_NAME_ALIASES.get(underlying_name, underlying_name)
    if canonical_name in ASSET_MAPPING:
        category, _, trend_code, _ = ASSET_MAPPING[canonical_name]
        return f'BOACTAR.{category}.{trend_code}.{field_code}.B'
    return None


def get_description(underlying_name, field_name):
    """Get the full description for an output column."""
    canonical_name = ASSET_NAME_ALIASES.get(underlying_name, underlying_name)
    if canonical_name in ASSET_MAPPING:
        category, _, _, display_name = ASSET_MAPPING[canonical_name]
        category_display = category.title()
        if category == 'FX':
            category_display = 'FX'
        return f'{category_display}, {display_name}, {field_name}'
    return None


def normalize_underlying_name(name):
    """Normalize an underlying name to match the canonical mapping."""
    if name is None:
        return None
    # Strip whitespace
    name = str(name).strip()
    # Check aliases first
    if name in ASSET_NAME_ALIASES:
        return ASSET_NAME_ALIASES[name]
    # Check if it's already a canonical name
    if name in ASSET_MAPPING:
        return name
    return name


def is_percentage_field(field_key):
    """Check if a field should be treated as a percentage (multiplied by 100)."""
    return field_key in PERCENTAGE_FIELDS


if __name__ == '__main__':
    # Print configuration summary
    print('BOACTAR Configuration Summary')
    print('=' * 50)
    print(f'Base Directory: {BASE_DIR}')
    print(f'Input Directory: {INPUT_DIR}')
    print(f'Output Directory: {OUTPUT_DIR}')
    print(f'Master Directory: {MASTER_DIR}')
    print(f'Total Assets Mapped: {len(ASSET_MAPPING)}')
    print(f'Total Trend Fields: {len(TREND_FIELDS)}')
    print(f'Percentage Fields: {len(PERCENTAGE_FIELDS)}')
