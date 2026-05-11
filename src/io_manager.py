import pandas as pd
import glob
import logging
from pathlib import Path
from logging.handlers import TimedRotatingFileHandler


def setup_logging():
    log_folder = Path("logs")
    log_folder.mkdir(exist_ok=True)
    logger = logging.getLogger("Automation")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler = TimedRotatingFileHandler(log_folder / "automation.log", when="midnight", backupCount=7)
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    return logger


def merge_excel_files(input_dir, output_name, logger):
    input_folder = Path(input_dir)
    files = glob.glob(str(input_folder / "*.xlsx"))
    df_list = []

    for file in files:
        if Path(file).name == output_name: continue
        try:
            df = pd.read_excel(file)
            df_list.append(df)
            logger.info(f"Loaded: {Path(file).name}")
        except Exception as e:
            logger.error(f"Error reading {Path(file).name}: {e}")

    if not df_list:
        raise ValueError(f"No Excel files found in {input_dir}")

    return pd.concat(df_list, ignore_index=True)


def save_to_excel(df, output_path, logger):
    try:
        df.to_excel(output_path, index=False)
        logger.info(f"File saved to {output_path}")
    except PermissionError:
        logger.error(f"Permission Denied: Close {output_path} and retry.")


import pandas as pd


def save_analysis_to_excel(data_frames_dict, output_file, logger):
    """
    This function creates ONE file with MULTIPLE tabs.
    """
    try:
        # 1. Open the file 'bubble'
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

            # 2. Loop through every metric in your dictionary
            for sheet_name, data in data_frames_dict.items():

                # 3. Convert to DataFrame if it isn't one already (like a Series)
                if isinstance(data, pd.Series):
                    df_to_save = data.reset_index()
                else:
                    df_to_save = data

                # 4. Save to its own tab
                # We show the index only for 'Sales by Location'
                show_idx = True if "Location" in sheet_name else False
                df_to_save.to_excel(writer, sheet_name=sheet_name[:31], index=show_idx)

                logger.info(f"Successfully added sheet: {sheet_name}")

        logger.info(f"DONE: All metrics saved to {output_file}")

    except Exception as e:
        logger.error(f"Excel Export Error: {e}")

