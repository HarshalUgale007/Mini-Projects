import pandas as pd
import yfinance as yf
from nselib import capital_market
from datetime import datetime, timedelta
import warnings

# Suppress future warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def fetch_index_data(index_name, from_date, to_date):
    try:
        df_index_data = capital_market.index_data(index_name, from_date=from_date, to_date=to_date)
        return df_index_data
    except Exception as e:
        print(f"Error fetching data for {index_name}: {e}")
        return None

def get_multi_index_price_dataframe(indices):
    to_date = datetime.now().strftime("%d-%m-%Y")
    from_date = (datetime.now() - timedelta(days=6*30)).strftime("%d-%m-%Y")

    result_data = []

    for serial_no, index in enumerate(indices, start=1):
        print(f"Fetching data for {index}...")
        df_index_data = fetch_index_data(index, from_date, to_date)
        
        if df_index_data is not None and not df_index_data.empty:
            try:
                # Use 'TIMESTAMP' instead of 'Date'
                df_index_data['Date'] = pd.to_datetime(df_index_data['TIMESTAMP'], errors='coerce')

                today_date = datetime.now().strftime("%d-%m-%Y")
                one_week_ago = (datetime.now() - timedelta(days=7)).strftime("%d-%m-%Y")
                one_month_ago = (datetime.now() - timedelta(days=30)).strftime("%d-%m-%Y")
                three_months_ago = (datetime.now() - timedelta(days=3*30)).strftime("%d-%m-%Y")
                six_months_ago = (datetime.now() - timedelta(days=6*30)).strftime("%d-%m-%Y")

                def find_closest_trading_day(date_str):
                    date = pd.to_datetime(date_str, format='%d-%m-%Y', errors='coerce')
                    closest_date = min(df_index_data['Date'].dropna(), key=lambda x: abs(x - date))
                    return closest_date

                today_date = find_closest_trading_day(today_date)
                one_week_ago = find_closest_trading_day(one_week_ago)
                one_month_ago = find_closest_trading_day(one_month_ago)
                three_months_ago = find_closest_trading_day(three_months_ago)
                six_months_ago = find_closest_trading_day(six_months_ago)

                today_close = df_index_data.loc[df_index_data['Date'] == today_date, 'CLOSE_INDEX_VAL'].values
                one_week_ago_close = df_index_data.loc[df_index_data['Date'] == one_week_ago, 'CLOSE_INDEX_VAL'].values
                one_month_ago_close = df_index_data.loc[df_index_data['Date'] == one_month_ago, 'CLOSE_INDEX_VAL'].values
                three_months_ago_close = df_index_data.loc[df_index_data['Date'] == three_months_ago, 'CLOSE_INDEX_VAL'].values
                six_months_ago_close = df_index_data.loc[df_index_data['Date'] == six_months_ago, 'CLOSE_INDEX_VAL'].values

                result_data.append({
                    # 'Serial No.': serial_no,
                    'Benchmark': index,
                    'Date': datetime.now().strftime("%d-%m-%Y"),
                    'Closing_Price_Today': today_close[0] if len(today_close) > 0 else None,
                    '1W_Price': one_week_ago_close[0] if len(one_week_ago_close) > 0 else None,
                    '1M_Price': one_month_ago_close[0] if len(one_month_ago_close) > 0 else None,
                    '3M_Price': three_months_ago_close[0] if len(three_months_ago_close) > 0 else None,
                    '6M_Price': six_months_ago_close[0] if len(six_months_ago_close) > 0 else None,
                })
            except Exception as e:
                print(f"Error processing data for {index}: {e}")
        else:
            print(f"No data available for {index}, skipping.")

    result_df = pd.DataFrame(result_data)
    return result_df

# List of indices
indices = [
    "NIFTY 50",
    "NIFTY NEXT 50",
    "NIFTY 100",
    "NIFTY 200",
    "NIFTY 500",
    "NIFTY MIDCAP 50",
    "NIFTY MIDCAP 100",
    "NIFTY SMALLCAP 100",
    "INDIA VIX",
    "NIFTY MIDCAP 150",
    "NIFTY SMALLCAP 50",
    "NIFTY SMALLCAP 250",
    "NIFTY MIDSMALLCAP 400",
    "NIFTY500 MULTICAP 50:25:25",
    "NIFTY LARGEMIDCAP 250",
    "NIFTY MIDCAP SELECT",
    "NIFTY TOTAL MARKET",
    "NIFTY MICROCAP 250",
    "NIFTY BANK",
    "NIFTY AUTO",
    "NIFTY FINANCIAL SERVICES",
    "NIFTY FINANCIAL SERVICES 25/50",
    "NIFTY FMCG",
    "NIFTY IT",
    "NIFTY MEDIA",
    "NIFTY METAL",
    "NIFTY PHARMA",
    "NIFTY PSU BANK",
    "NIFTY PRIVATE BANK",
    "NIFTY REALTY",
    "NIFTY HEALTHCARE INDEX",
    "NIFTY CONSUMER DURABLES",
    "NIFTY OIL & GAS",
    "NIFTY MIDSMALL HEALTHCARE",
    "NIFTY DIVIDEND OPPORTUNITIES 50",
    "NIFTY GROWTH SECTORS 15",
    "NIFTY100 QUALITY 30",
    "NIFTY50 VALUE 20",
    "NIFTY50 TR 2X LEVERAGE",
    "NIFTY50 PR 2X LEVERAGE",
    "NIFTY50 TR 1X INVERSE",
    "NIFTY50 PR 1X INVERSE",
    "NIFTY50 DIVIDEND POINTS",
    "NIFTY ALPHA 50",
    "NIFTY50 EQUAL WEIGHT",
    "NIFTY100 EQUAL WEIGHT",
    "NIFTY100 LOW VOLATILITY 30",
    "NIFTY200 QUALITY 30",
    "NIFTY ALPHA LOW-VOLATILITY 30",
    "NIFTY200 MOMENTUM 30",
    "NIFTY MIDCAP150 QUALITY 50",
    "NIFTY200 ALPHA 30",
    "NIFTY MIDCAP150 MOMENTUM 50",
    "NIFTY COMMODITIES",
    "NIFTY INDIA CONSUMPTION",
    "NIFTY CPSE",
    "NIFTY ENERGY",
    "NIFTY INFRASTRUCTURE",
    "NIFTY100 LIQUID 15",
    "NIFTY MIDCAP LIQUID 15",
    "NIFTY MNC",
    "NIFTY PSE",
    "NIFTY SERVICES SECTOR",
    "NIFTY100 ESG SECTOR LEADERS",
    "NIFTY INDIA DIGITAL",
    "NIFTY100 ESG",
    "NIFTY INDIA MANUFACTURING",
    "NIFTY INDIA CORPORATE GROUP INDEX - TATA GROUP 25% CAP",
    "NIFTY500 MULTICAP INDIA MANUFACTURING 50:30:20",
    "NIFTY500 MULTICAP INFRASTRUCTURE 50:30:20",
    "NIFTY 8-13 YR G-SEC",
    "NIFTY 10 YR BENCHMARK G-SEC",
    "NIFTY 10 YR BENCHMARK G-SEC (CLEAN PRICE)",
    "NIFTY 4-8 YR G-SEC INDEX",
    "NIFTY 11-15 YR G-SEC INDEX",
    "NIFTY 15 YR AND ABOVE G-SEC INDEX",
    "NIFTY COMPOSITE G-SEC INDEX"
]

# Fetch the data for the indices
result_df = get_multi_index_price_dataframe(indices)

# Export the result DataFrame to a CSV file
result_df.to_csv(r"C:\Users\HP\Desktop\Report\Automation\Index_Price_Data.csv", index=False)

# Print confirmation
print("Data has been exported to Index_Price_Data.csv")
