import pandas as pd
import logging
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_exchange_rate(currency_rates, date, currency):
    rates_on_date = currency_rates[currency_rates['date'] == date]
    if not rates_on_date.empty:
        return rates_on_date.iloc[0][currency]
    return 1


def calculate_fees(row, affiliates_rates):
    applicable_rates = affiliates_rates[(affiliates_rates['Affiliate ID'] == row['Affiliate ID']) &
                                        (affiliates_rates['Start Date'] <= row['Order Date'])]
    if not applicable_rates.empty:
        applicable_rate = applicable_rates.sort_values('Start Date', ascending=False).iloc[0]
    else:
        return 0, 0, 0

    processing_fee = round(row['Order Amount (EUR)'] * applicable_rate['Processing Rate'], 2)
    refund_fee = round(applicable_rate['Refund Fee'], 2) if row['Order Status'] == 'Refunded' else 0
    chargeback_fee = round(applicable_rate['Chargeback Fee'], 2) if row['Order Status'] == 'Chargeback' else 0

    return processing_fee, refund_fee, chargeback_fee


def generate_reports(orders_path, affiliates_path, currency_path):
    logging.info("Starting report generation process.")

    logging.info("Loading Excel files.")
    orders = pd.read_excel(orders_path)
    affiliates = pd.read_excel(affiliates_path)
    currency_rates = pd.read_excel(currency_path)

    logging.info("Converting date columns to datetime.")
    orders['Order Date'] = pd.to_datetime(orders['Order Date'])
    affiliates['Start Date'] = pd.to_datetime(affiliates['Start Date'])
    currency_rates['date'] = pd.to_datetime(currency_rates['date'])

    logging.info("Converting 'Order Amount' to EUR.")
    orders['Order Amount (EUR)'] = orders.apply(
        lambda row: round(row['Order Amount'] / get_exchange_rate(currency_rates, row['Order Date'], row['Currency']),
                          2)
        if row['Currency'] != 'EUR' else round(row['Order Amount'], 2),
        axis=1
    )

    logging.info("Calculating fees.")
    orders[['Processing Fee', 'Refund Fee', 'Chargeback Fee']] = orders.apply(
        lambda row: calculate_fees(row, affiliates), axis=1, result_type='expand'
    )

    logging.info("Aggregating data on a weekly basis.")
    orders['Week'] = orders['Order Date'].dt.to_period('W-SUN').apply(
        lambda r: f'{r.start_time.strftime("%d-%m-%Y")} - {r.end_time.strftime("%d-%m-%Y")}')
    weekly_reports = orders.groupby(['Affiliate ID', 'Week']).agg(
        Number_of_Orders=('Order Number', 'count'),
        Total_Order_Amount_EUR=('Order Amount (EUR)', 'sum'),
        Total_Processing_Fee=('Processing Fee', 'sum'),
        Total_Refund_Fee=('Refund Fee', 'sum'),
        Total_Chargeback_Fee=('Chargeback Fee', 'sum')
    ).reset_index()

    logging.info("Creating directories and saving weekly report Excel files.")
    base_path = 'affiliate_reports'
    if not os.path.exists(base_path):
        os.makedirs(base_path)

    for affiliate_id, group in weekly_reports.groupby('Affiliate ID'):
        affiliate_name = affiliates[affiliates['Affiliate ID'] == affiliate_id]['Affiliate Name'].iloc[0]
        affiliate_dir = os.path.join(base_path, affiliate_name)
        if not os.path.exists(affiliate_dir):
            os.makedirs(affiliate_dir)
        filename = os.path.join(affiliate_dir, f'{affiliate_name}.xlsx')
        group.to_excel(filename, index=False)
        logging.info(f"Excel report generated successfully: {filename}")

    logging.info("Report generation process completed.")
