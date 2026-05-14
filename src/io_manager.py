import pandas as pd
import logging
from pathlib import Path
from logging.handlers import TimedRotatingFileHandler


def setup_logging():
    log_folder = Path("logs")
    log_folder.mkdir(exist_ok=True)
    
    # FIX 1: Harmonized name to "ExcelAutomation" to match main.py exactly
    logger = logging.getLogger("ExcelAutomation")
    logger.setLevel(logging.INFO)

    # FIX 2: Force clear any lingering handlers to release Windows file locks
    if logger.hasHandlers():
        logger.handlers.clear()

    # Re-initialize cleanly without risk of duplication or file locks
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = TimedRotatingFileHandler(log_folder / "automation.log", when="midnight", backupCount=7)
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def merge_excel_files(files, output_name, logger):
    df_list = []
    for file in files:
        file_path = Path(file)
        if file_path.name == output_name:
            continue
        try:
            df = pd.read_excel(file_path)
            df_list.append(df)
            logger.info(f"Loaded: {file_path.name}")
        except Exception as e:
            logger.error(f"Error reading {file_path.name}: {e}")

    if not df_list:
        raise ValueError(f"No valid Excel files were successfully loaded to merge into '{output_name}'.")

    return pd.concat(df_list, ignore_index=True)


def save_to_excel(df, output_path, logger):
    try:
        df.to_excel(output_path, index=False)
        logger.info(f"File saved to {output_path}")
    except PermissionError:
        logger.error(f"Permission Denied: Close {output_path} and retry.")


def save_analysis_to_excel(data_frames_dict, output_file, logger):
    """This function creates ONE file with MULTIPLE tabs."""
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, data in data_frames_dict.items():
                if isinstance(data, pd.Series):
                    df_to_save = data.reset_index()
                else:
                    df_to_save = data

                show_idx = True if "Location" in sheet_name else False
                # Sheet names are sliced to 31 characters to meet Excel standards
                df_to_save.to_excel(writer, sheet_name=sheet_name[:31], index=show_idx)
                logger.info(f"Successfully added sheet: {sheet_name}")

        logger.info(f"DONE: All metrics saved to {output_file}")
    except Exception as e:
        logger.error(f"Excel Export Error: {e}")


# ==========================================
# PIPELINE EXECUTION ENGINE (LOCAL TESTING)
# ==========================================
if __name__ == "__main__":
    # 1. Initialize logging
    logger = setup_logging()
    logger.info("Starting data pipeline execution...")

    # 2. Define input and output directories
    input_folder = Path("input")
    input_folder.mkdir(exist_ok=True)

    output_filename = "final_report.xlsx"
    output_path = Path("output") / output_filename
    output_path.parent.mkdir(exist_ok=True)

    # 3. Dynamically scan the folder for all Excel files (.xlsx and .xls)
    excel_files = list(input_folder.glob("*.xlsx")) + list(input_folder.glob("*.xls"))
    logger.info(f"Found {len(excel_files)} total file(s) inside '{input_folder}'")

    try:
        # 4. Run the merge routine
        merged_df = merge_excel_files(excel_files, output_filename, logger)

        # 5. Run analysis to build multi-tab reports
        analysis_payload = {
            "Raw Merged Data": merged_df,
            "Sales by Location": merged_df.groupby("Location")["Revenue"].sum() if "Location" in merged_df.columns and "Revenue" in merged_df.columns else merged_df.head(10),
            "Summary Metrics": merged_df.describe()
        }

        # 6. Save final multi-tab workbooks
        save_analysis_to_excel(analysis_payload, output_path, logger)
        logger.info("Pipeline executed successfully without critical crashes.")

    except Exception as pipeline_err:
        logger.critical(f"Pipeline crashed during execution: {pipeline_err}")

