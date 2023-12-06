from FilterData.data_filter import run_filter
from report_generator.generate_reports import generate_reports

if __name__ == "__main__":
    run_filter('./data/test-orders.xlsx', './data/test-affiliate-rates.xlsx')
    generate_reports('./data/test-orders.xlsx', './data/test-affiliate-rates.xlsx', './data/test-currency-rates.xlsx')
