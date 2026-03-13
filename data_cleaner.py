# data_cleaner.py — Cleans messy numbers into usable data

import pandas as pd

def clean_number(val):
    if pd.isna(val):
        return None
    s = str(val).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except:
        return None

def extract_table(company_data, section_name):
    for key in company_data:
        if section_name.lower() in key.lower():
            df = company_data[key].copy()
            label_col = df.columns[0]
            df = df.set_index(label_col)
            df = df.applymap(clean_number)
            df = df.dropna(how="all")
            return df
    return pd.DataFrame()

def clean_all(raw_data):
    cleaned = {}
    for company_name, company_data in raw_data.items():
        print(f"  Cleaning: {company_name}")
        cleaned[company_name] = {
            "pnl":           extract_table(company_data, "Profit & Loss"),
            "balance_sheet": extract_table(company_data, "Balance Sheet"),
            "cashflow":      extract_table(company_data, "Cash Flow"),
            "ratios":        extract_table(company_data, "Ratios"),
        }
    return cleaned