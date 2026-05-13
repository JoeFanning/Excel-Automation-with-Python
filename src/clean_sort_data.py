import pandas as pd

def clean_data(df, logger):
    # Text cleaning
    object_cols = df.select_dtypes(include='object').columns
    for col in object_cols:
        df[col] = df[col].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()

    to_replace = ['', 'nan', 'NaN', 'NAN', 'null', 'None', 'N/A', 'n/a']
    df = df.replace(r'^\s*$', 'Unknown', regex=True).replace(to_replace, 'Unknown')

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.drop_duplicates(keep='first').reset_index(drop=True)

    logger.info("Data cleaning completed.")
    return df

def sort_data(df, logger):
    # Fixed: Passing 'df' into the function so it has data to work with
    df = df.sort_values(
        by=['Date', 'Time'],
        key=lambda x: pd.to_datetime(x, format='%H:%M') if x.name == 'Time' else x
    ).reset_index(drop=True)
    logger.info("Data sorting completed.")
    return df
