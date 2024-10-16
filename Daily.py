import requests
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf
import datetime
import math

# URL for fetching stock exchange data
nasdaq_stock_exchange_overall_url = 'https://nasdaqbaltic.com/statistics/lt/charts/charts_data_json?filter=1&indexes%5B%5D=OMXVGI&start=2024-01-02_=1724150770879'

def round_half_up(n, decimals=0):
    """Round a number to the nearest value with a specific number of decimal places."""
    multiplier = 10 ** decimals
    return math.floor(n * multiplier + 0.5) / multiplier

def nasdaq_get_stock_exchange_overall(nasdaq_stock_exchange_overall_url):
    """Fetch stock exchange data from the provided API URL."""
    response = requests.get(nasdaq_stock_exchange_overall_url, verify=False)
    data = response.json()

    charts_data = data['data']['charts'][0]['data']
    date_value_list = []

    for item in charts_data[1:]:
        timestamp_ms = item[0]
        value = item[1]
        
        # Convert timestamp to date in YYYY-MM-DD format
        date_str = datetime.datetime.fromtimestamp(timestamp_ms / 1000.0).strftime('%Y-%m-%d')
        
        date_value_list.append((date_str, value))
    
    print(date_value_list)
    return date_value_list

def update_excel_value(ws, row, value, red_fill):
    """Update the Excel sheet with the provided value at the specified row."""
    print(f"Updating value for row {row} with value {value}")
    ws.write(row, 6, value, red_fill)  # Write new value and apply red fill

def compare_stock_overall(date_value_list, ws, df):
    """Compare API data with Excel data and update accordingly."""
    # Ensure dates in DataFrame are in datetime format for comparison
    df['A'] = pd.to_datetime(df['A'], format='%Y-%m-%d', errors='coerce')

    # Find the newest date from the API data
    newest_date = max(datetime.datetime.strptime(date_str, '%Y-%m-%d') for date_str, _ in date_value_list)
    print(f"Newest date from API: {newest_date}")

    # Create a set of API dates for quick lookup
    api_dates = {date_str for date_str, _ in date_value_list}

    # Iterate over API data for comparison
    for row in date_value_list:
        date_str = row[0]
        value = row[1]

        try:
            float_value = float(value)
            formatted_value = round_half_up(float_value, decimals=3)
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        # Convert column A to YYYY-MM-DD format for matching
        match = df[df['A'] == date_str]

        if not match.empty:
            # If a matching date is found, check if the value needs updating
            index = match.index[0]
            excel_value = str(df.at[index, 'G']).replace(',', '.').strip().lower()

            if excel_value == 'wh':
                # If the Excel value is 'wh', replace it with the API value
                print(f"Replacing 'wh' with {formatted_value} for date {date_str}.")
                update_excel_value(ws, index + 3, formatted_value, easyxf('pattern: pattern solid, fore_color red;'))
            else:
                # Check if the excel_value is numeric
                if not is_number(excel_value):
                    print(f"Warning: Existing value '{excel_value}' cannot be converted to float. Skipping.")
                    continue

                try:
                    excel_float_value = float(excel_value)
                except ValueError:
                    print(f"Warning: Existing value '{excel_value}' cannot be converted to float. Skipping.")
                    continue

                formatted_excel_value = f"{excel_float_value:.3f}"

                # If the existing value in Excel is different, update it
                if float(formatted_value) != float(formatted_excel_value):
                    update_excel_value(ws, index + 3, formatted_value, easyxf('pattern: pattern solid, fore_color red;'))
                else:
                    print(f"Value for {date_str} is already up to date: {formatted_excel_value}")
        else:
            # If no matching date is found, add a new row
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            print(f"Adding new row for date {date_str} with value {formatted_value}")
            ws.write(new_row_index, 0, date_str)
            update_excel_value(ws, new_row_index, formatted_value, easyxf('pattern: pattern solid, fore_color red;'))

    # After comparisons, write "wh" for any dates before the newest date that don't have values from the API
    write_wh_for_invalid_cells(ws, df, api_dates, newest_date)


def write_wh_for_invalid_cells(ws, df, api_dates, newest_date):
    """Write 'wh' for cells that do not have values from the API and are before the newest API date."""
    for index, row in df.iterrows():
        excel_date = row['A']
        excel_value = str(row['G']).strip().lower()

        # Check if the date is before the newest date and there's no corresponding API value
        if excel_date < newest_date and excel_date.strftime('%Y-%m-%d') not in api_dates:
            if excel_value != 'wh':  # Only write 'wh' if it isn't already there
                print(f"Setting value for {excel_date} to 'wh'.")
                ws.write(index + 3, 6, 'wh', easyxf('pattern: pattern solid, fore_color red;'))  # Write 'wh' and apply red fill
            else:
                print(f"Value for {excel_date} is already 'wh', skipping.")

# Helper function to determine if a string can be converted to a float
def is_number(s):
    """Check if the provided string can be converted to a float."""
    try:
        float(s)
        return True
    except ValueError:
        return False

def process_excel_file(excel_file, date_value_list):
    """Process the Excel file and update it with data from the API."""
    try:
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        wb = copy(rb)
        ws = wb.get_sheet('Daily')

        df = pd.read_excel(
            excel_file, 
            sheet_name='Daily', 
            usecols=[0, 6],
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'G']

        compare_stock_overall(date_value_list, ws, df)
        wb.save("naujas_combined.xls")

        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Fetch the latest stock exchange data
date_value_list = nasdaq_get_stock_exchange_overall(nasdaq_stock_exchange_overall_url)

# Process the Excel file with the fetched data
process_excel_file('laikinas.xls', date_value_list)
