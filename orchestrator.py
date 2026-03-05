"""
BOACTAR Orchestrator
====================
Main entry point for the BOACTAR runbook pipeline.

This script orchestrates the complete data processing workflow:
1. Load and clean input Excel files
2. Parse data and map to output columns
3. Merge with historical master data
4. Generate output DATA, META, and ZIP files

Usage:
    python orchestrator.py
"""

import os
import sys
import logging
from datetime import datetime

import config
from logger_setup import setup_logging
from data_loader import BOACTARDataLoader
from parser import BOACTARParser
from file_generator import BOACTARFileGenerator

logger = logging.getLogger(__name__)


def print_banner():
    """Print welcome banner."""
    banner = """
====================================================================
                      BOACTAR RUNBOOK
              BOA CTA Report Data Processing
====================================================================
    """
    print(banner)


def print_configuration():
    """Print current configuration settings."""
    print('\nConfiguration:')
    print(f'  Input Directory:  {config.INPUT_DIR}')
    print(f'  Output Directory: {config.OUTPUT_DIR}')
    print(f'  Master Directory: {config.MASTER_DIR}')
    print(f'  Debug Mode:       {config.DEBUG_MODE}')
    print(f'  Rebuild Master:   {config.REBUILD_MASTER}')
    print()


def print_summary(result, all_loaded_data, combined_df):
    """
    Print execution summary.

    Args:
        result: Dictionary with output file paths.
        all_loaded_data: List of loaded data dictionaries.
        combined_df: Final combined DataFrame.
    """
    print('\n' + '=' * 60)
    print('EXECUTION SUMMARY')
    print('=' * 60)

    print(f'\nInput Files Processed: {len(all_loaded_data)}')
    for i, loaded_data in enumerate(all_loaded_data, 1):
        print(f"  {i}. {loaded_data['file_name']} ({loaded_data['date_str']})")

    print(f'\nOutput generated:')
    print(f"  Total rows in output: {len(combined_df)}")
    print(f"  Timestamp: {result['timestamp']}")

    print(f'\nFiles created:')
    print(f"  DATA: {os.path.basename(result['data_file'])}")
    print(f"  META: {os.path.basename(result['meta_file'])}")
    print(f"  ZIP:  {os.path.basename(result['zip_file'])}")

    print(f'\nOutput directory: {result["output_dir"]}')
    print(f'Master file: {result["master_file"]}')

    print('\n' + '=' * 60)
    print('COMPLETED SUCCESSFULLY')
    print('=' * 60)


def main():
    """
    Main orchestration function.

    Returns:
        int: Exit code (0 for success, 1 for failure).
    """
    # Setup logging
    timestamp = config.get_timestamp()
    log_file = setup_logging(timestamp)

    print_banner()
    print_configuration()

    try:
        # ===================================================================
        # STEP 1: Load Input Data
        # ===================================================================
        print('STEP 1: Loading input data...')
        logger.info('=' * 50)
        logger.info('STEP 1: Loading input data')
        logger.info('=' * 50)

        loader = BOACTARDataLoader()
        all_loaded_data = loader.load_all_files(sort_by_date=True)

        if not all_loaded_data:
            logger.error('No input files found to process')
            print('ERROR: No input files found in the input directory')
            print(f'Please place input files in: {config.INPUT_DIR}')
            return 1

        print(f'  Found {len(all_loaded_data)} input file(s):')
        for data in all_loaded_data:
            print(f"    - {data['file_name']} (Date: {data['date_str']}, Rows: {len(data['data'])})")

        # ===================================================================
        # STEP 2: Parse and Merge Data (process all files in chronological order)
        # ===================================================================
        print('\nSTEP 2: Parsing and merging data...')
        logger.info('=' * 50)
        logger.info('STEP 2: Parsing and merging data')
        logger.info('=' * 50)

        parser = BOACTARParser()
        combined_df = None

        for i, loaded_data in enumerate(all_loaded_data, 1):
            print(f"\n  Processing file {i}/{len(all_loaded_data)}: {loaded_data['file_name']}")
            logger.info(f"Processing file {i}/{len(all_loaded_data)}: {loaded_data['file_name']}")

            # Pass existing DataFrame to maintain state between files
            combined_df = parser.parse_and_merge(loaded_data, existing_df=combined_df)

            if combined_df.empty:
                logger.warning(f"No data after processing {loaded_data['file_name']}")
                print(f"    WARNING: No data produced")
            else:
                print(f"    Date: {loaded_data['date_str']}, Total rows now: {len(combined_df)}")

        if combined_df is None or combined_df.empty:
            logger.error('No data after parsing and merging all files')
            print('ERROR: No data produced after parsing')
            return 1

        column_order = parser.get_column_order()

        print(f"\n  Final result:")
        print(f"    Total rows:    {len(combined_df)}")
        print(f"    Total columns: {len(column_order) if column_order else 'N/A'}")
        print(f"    Date range:    {combined_df['date'].min()} to {combined_df['date'].max()}")

        # ===================================================================
        # STEP 3: Generate Output Files
        # ===================================================================
        print('\nSTEP 3: Generating output files...')
        logger.info('=' * 50)
        logger.info('STEP 3: Generating output files')
        logger.info('=' * 50)

        generator = BOACTARFileGenerator(column_order=column_order)
        generator.timestamp = timestamp  # Use same timestamp as logging
        result = generator.generate_files(combined_df)

        print(f"  Output folder: {result['timestamp']}")
        print(f"  DATA file:     {os.path.basename(result['data_file'])}")
        print(f"  META file:     {os.path.basename(result['meta_file'])}")
        print(f"  ZIP file:      {os.path.basename(result['zip_file'])}")

        # ===================================================================
        # STEP 4: Summary
        # ===================================================================
        logger.info('=' * 50)
        logger.info('STEP 4: Execution complete')
        logger.info('=' * 50)

        print_summary(result, all_loaded_data, combined_df)

        logger.info('Pipeline completed successfully')

        return 0

    except KeyboardInterrupt:
        logger.warning('Execution interrupted by user')
        print('\n\nExecution interrupted by user')
        return 130

    except Exception as e:
        logger.exception(f'Pipeline failed with error: {e}')
        print(f'\nERROR: {e}')
        print(f'See log file for details: {log_file}')
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
