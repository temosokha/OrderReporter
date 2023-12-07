import pandas as pd
import logging
from typing import List
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
pd.options.mode.chained_assignment = None


def read_excel_file(file_path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            logging.warning(f"The file {file_path} is empty. There is no data to process.")
        return df
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")
        raise


def filter_orders(orders_df: pd.DataFrame, affiliate_ids: List[int]) -> pd.DataFrame:
    orders_df = orders_df.drop_duplicates(subset='Order Number', keep='first')
    orders_df['Order Date'] = pd.to_datetime(orders_df['Order Date'], format='%m/%d/%Y %I:%M:%S %p', errors='coerce')
    orders_df = orders_df[orders_df['Order Date'].notna()]
    orders_df = orders_df[pd.to_numeric(orders_df['Order Amount'], errors='coerce').notna()]
    orders_df = orders_df[orders_df['Order Status'].isin(['Refunded', 'Chargeback', 'Completed'])]
    orders_df = orders_df[orders_df['Currency'].isin(['GBP', 'USD', 'EUR'])]
    orders_df = orders_df[pd.to_numeric(orders_df['Affiliate ID'], errors='coerce').notna()]
    orders_df = orders_df[orders_df['Affiliate ID'].isin(affiliate_ids)]

    return orders_df


def save_filtered_data(df: pd.DataFrame, file_path: str) -> None:
    try:
        df.to_excel(file_path, index=False)
    except Exception as e:
        logging.error(f"Error saving file {file_path}: {e}")
        raise


def run_filter(orders_file_path: str, affiliate_rates_file_path: str) -> None:
    orders_df = read_excel_file(orders_file_path)
    affiliate_rates_df = read_excel_file(affiliate_rates_file_path)
    filtered_orders_df = filter_orders(orders_df, affiliate_rates_df['Affiliate ID'].tolist())
    save_filtered_data(filtered_orders_df, orders_file_path)
