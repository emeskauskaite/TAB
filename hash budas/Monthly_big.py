import requests
import re
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf
#mazesnis 357bde94-fa9f-40d6-89fa-782098fa6667
#https://osp.stat.gov.lt/lt/statistiniu-rodikliu-analize?hash=cd52482a-1fac-48d8-a0df-49aa60707c19#/

main_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/221b6516-48a7-4964-906a-14622a8ab3e7'
import re

def sanitize_value(value):
    if value is None:
        return None
    
    # Convert value to string, remove leading/trailing whitespace, and format it
    value = str(value).strip()  # First, strip whitespace from the beginning and end
    value = re.sub(r'\s+', '', value)  # Remove all spaces
    value = value.replace(',', '.')  # Replace commas with dots
    
    # Match valid float values (positive or negative)
    if re.match(r'^-?\d+(\.\d+)?$', value):  
        return value  # Return the sanitized value as a string
    else:
        print(f"Warning: Value '{value}' is not a valid number and will be skipped.")
        return None


def vda_get_data_matrix_one_indicator(api_url, verify=False):
    response = requests.get(api_url, verify=verify)
    data_matrix = []

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'itemMatrix' in data['data']:
            item_matrix = data['data']['itemMatrix']

            for listed in item_matrix:
                for item in listed:
                    # Check if item is not None
                    if item is None:
                        continue  # Skip if item is None

                    # Check if 'value' exists in item
                    value = sanitize_value(item['value']) if 'value' in item else None

                    timeperiod = None
                    # Check if 'coordinate' exists in item and iterate if it's a list
                    if 'coordinate' in item and isinstance(item['coordinate'], list):
                        for coordinate in item['coordinate']:
                            if isinstance(coordinate, str) and coordinate.startswith('L#LAIKOTARPIS'):
                                year_month_match = re.findall(r'\d+', coordinate.split('#')[2])
                                if len(year_month_match) >= 2:
                                    year_str = year_month_match[0]
                                    month_str = year_month_match[1]
                                    timeperiod = int(year_str + month_str)
                                break

                    if timeperiod is not None:
                        data_matrix.append([timeperiod, value])
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

    return data_matrix


def vki_compare(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Ensure data_matrix is a list

    for row in data_matrix:
        num = row[0]
        if row[1] is None:
            return 0 
        value = row[1].replace(',', '.').replace(' ', '')
        
        try:
            float_value = float(value)
            formatted_value = f"{float_value:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AB']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.4f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 27, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 27, formatted_value)

def hicp_compare(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Ensure data_matrix is a list

    for row in data_matrix:
        if len(row) < 3:  # Ensure there are enough elements in the row
            print(f"Warning: Expected at least 3 values in row but got {len(row)}: {row}")
            continue
        if row[1] is None:
            return 0 
        if row[2] is None:
            return 0 
        num = row[0]
        value1 = str(row[1]).replace(',', '.').replace(' ', '')
        value2 = str(row[2]).replace(',', '.').replace(' ', '')
        

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.1f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.1f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]
        if not match.empty:
            index = match.index[0]
            excel_value1 = str(df.at[index, 'AD']).replace(',', '.').replace(' ', '')
            excel_value2 = str(df.at[index, 'AE']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value1 = float(excel_value1)
                formatted_excel_value1 = f"{excel_float_value1:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                continue

            try:
                excel_float_value2 = float(excel_value2)
                formatted_excel_value2 = f"{excel_float_value2:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                continue

            cell_row = index + 3  # Adjust row index based on header rows

            if float(formatted_excel_value1) != float(formatted_value1):
                ws.write(cell_row, 29, formatted_value1, red_fill)  # Write new value and apply red fill
            if float(formatted_excel_value2) != float(formatted_value2):
                ws.write(cell_row, 30, formatted_value2, red_fill)  # Write new value and apply red fill
        else:
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 29, formatted_value1)
            ws.write(new_row_index, 30, formatted_value2)


def gki_compare(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Ensure data_matrix is a list
        

    for row in data_matrix:
        num = row[0]
        if row[1] is None:
            return 0 
        value = row[1].replace(',', '.').replace(' ', '')
        

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to one decimal place
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AF']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.1f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 31, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 31, formatted_value)
            print(f"Added new value for {num}: {formatted_value}")



def pram_c_compare(data_matrix, ws, df,  red_fill):
     for row in data_matrix:
        num = row[0]
        if row[1] is None:
            return 0 
        value = row[1].replace(',', '.').replace(' ', '')

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to one decimal place
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AM']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.1f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 38, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 38, formatted_value)
            print(f"Added new value for {num}: {formatted_value}")


def nsa_compare(data_input, ws, df, red_fill):
    # If the input is a single list, treat it as a single time period with its values
    if isinstance(data_input, list) and len(data_input) == 5:
        data_dict = {data_input[0]: data_input[1:]}
    elif isinstance(data_input, dict):
        data_dict = data_input

    for timeperiod, values in data_dict.items():
        num = timeperiod
        
        # Ensure values is a list and has exactly 4 elements
        if not isinstance(values, list) or len(values) != 4:
            print(f"Skipping time period {timeperiod}: Expected 4 values but got {len(values) if isinstance(values, list) else 'non-list'}")
            continue

        # Convert and format the values, replacing None or missing values with '0'
        formatted_values = []
        for value in values:
        
            if value is None or pd.isna(value):  
                formatted_values.append('')   
            else:
                try:
                    cleaned_value = re.sub(r'\s+', '', str(value)).replace(',', '.').strip()
                    float_value = float(cleaned_value)  
   
                    formatted_values.append(f"{float_value:.1f}")  
                except ValueError:
                    print(f"Warning: Value '{value}' is not a valid number and will be skipped.")
                    continue

        # Check if we have all formatted values
        if len(formatted_values) < 4:
            print(f"Not all values were converted for time period {timeperiod}. Skipping.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]

            # Helper function to safely get float value from a cell, returning 0 for None or NaN
            def get_float(value):
                if value is None or pd.isna(value):
                    return 0  # Default to 0 if value is None or NaN
                try:
                    # Remove spaces and commas before converting
                    cleaned_value = str(value).replace(' ', '').replace(',', '.').strip()
                    return float(cleaned_value)  # Convert to float
                except ValueError:
                    return 0  # Return 0 if conversion fails

            # Retrieve existing values
            excel_value1 = get_float(df.at[index, 'AZ']) if 'AZ' in df.columns else 0
            excel_value2 = get_float(df.at[index, 'BA']) if 'BA' in df.columns else 0
            excel_value3 = get_float(df.at[index, 'BB']) if 'BB' in df.columns else 0
            excel_value4 = get_float(df.at[index, 'BC']) if 'BC' in df.columns else 0

            cell_row = index + 3

            # Apply formatting and update if values differ
            if excel_value1 != float(formatted_values[0]):
                ws.write(cell_row, 51, formatted_values[0], red_fill)  # Column AZ

            if excel_value2 != float(formatted_values[1]):
                ws.write(cell_row, 52, formatted_values[1], red_fill)  # Column BA

            if excel_value3 != float(formatted_values[2]):
                ws.write(cell_row, 53, formatted_values[2], red_fill)  # Column BB

            if excel_value4 != float(formatted_values[3]):
                ws.write(cell_row, 54, formatted_values[3], red_fill)  # Column BC

        else:
            # Adding new rows for missing time periods
            new_row_index = len(df) + 3
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 51, formatted_values[0])  # Column AZ
            ws.write(new_row_index, 52, formatted_values[1])  # Column BA
            ws.write(new_row_index, 53, formatted_values[2])  # Column BB
            ws.write(new_row_index, 54, formatted_values[3])  # Column BC


def eki_compare(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Ensure data_matrix is a list

    for row in data_matrix:
        num = row[0]
        if row[1] is None:
            return 0 
        value = row[1].replace(',', '.').replace(' ', '')


        try:
            float_value1 = float(value)
            formatted_value1 = f"{float_value1:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value1 = str(df.at[index, 'AG']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value1 = float(excel_value1)
                formatted_excel_value1 = f"{excel_float_value1:.4f}"  # Format existing value
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                formatted_excel_value1 = None

            cell_row = index + 3  # Adjust row index based on header rows

            if formatted_excel_value1 != formatted_value1:
                ws.write(cell_row, 32, formatted_value1, red_fill)  # Write new value and apply red fill to AG
        else:
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 32, formatted_value1)  # AG column

def iki_compare(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Ensure data_matrix is a list

    for row in data_matrix:
        num = row[0]
        if row[1] is None:
            return 0 
        value = row[1].replace(',', '.').replace(' ', '')



        try:
            float_value2 = float(value)
            formatted_value2 = f"{float_value2:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value2 = str(df.at[index, 'AO']).replace(',', '.').replace(' ', '')

            try:
                excel_float_value2 = float(excel_value2)
                formatted_excel_value2 = f"{excel_float_value2:.4f}"  # Format existing value
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                formatted_excel_value2 = None

            cell_row = index + 3  # Adjust row index based on header rows

            if formatted_excel_value2 != formatted_value2:
                ws.write(cell_row, 40, formatted_value2, red_fill)  # Write new value and apply red fill to AO
        else:
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 40, formatted_value2)  # AO column

def do_nothing(data_matrix, ws, df, red_fill):
    return 0

def big_compare(excel_file, big_data_matrix):
    try:
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Monthly')
        wb = copy(rb)
        ws = wb.get_sheet('Monthly')

        df = pd.read_excel(
            excel_file, 
            sheet_name='Monthly', 
            usecols=[0, 27, 29, 30, 31, 32, 38, 40, 51, 52, 53, 54], 
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'AB', 'AD', 'AE', 'AF', 'AG', 'AM', 'AO', 'AZ', 'BA', 'BB', 'BC']

        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        # Clear existing cell formatting
        for i in range(4, rs.nrows):
            for j in range(2, 50):
                cell_value = rs.cell_value(i, j)
                ws.write(i, j, cell_value, no_fill)

        # Mapping function keys to their corresponding functions
        compare_functions = {
            0: vki_compare,
            1: hicp_compare,
            3: gki_compare, 
            4: eki_compare,
            5: pram_c_compare,
            6: iki_compare,
            7: nsa_compare,
            10: do_nothing
        }
        
        current_index = 0  # This will track the current index
        current_year = None
        current_month = None
        start_index = 0  # Track the starting index for each year

        while current_index < len(big_data_matrix):
            # Extract year and month from the current entry
            time_period = big_data_matrix[current_index][0]
            entry_year = int(str(time_period)[:4])
            entry_month = int(str(time_period)[4:])

            # Check if we need to reset for a new year
            if current_year is None or current_month is None:
                current_year = entry_year
                current_month = entry_month
            elif entry_year < current_year or (entry_year == current_year and entry_month < current_month):
                print(f"Year changed from {current_year}-{current_month} to {entry_year}-{entry_month}. Resetting index.")
                # Resetting the current index to the start index for the new year
                current_index = start_index  # Go back to the start of the new year
                current_year = entry_year
                current_month = entry_month  # Update current month/year
                continue  # Restart the process from the current index
            
            # Update the start_index to the current_index for the new year
            start_index = current_index

            # Process each entry according to the mapping of functions
            for key in sorted(compare_functions.keys()):
                if current_index >= len(big_data_matrix):
                    break

                compare_function = compare_functions[key]

                # Handle each compare function's specific data requirements
                if key == 1:  # HICP expects pairs of values
                    if current_index + 1 < len(big_data_matrix):
                        values = [
                            big_data_matrix[current_index][0], 
                            big_data_matrix[current_index][1], 
                            big_data_matrix[current_index + 1][1]
                        ]
                        current_index += 2
                    else:
                        print(f"Insufficient data for hicp_compare at index {current_index}. Moving to the next period.")
                        break  # Exit if not enough values
                
                elif key == 7:  # NSA expects four values
                    if current_index + 3 < len(big_data_matrix):
                        values = [
                            big_data_matrix[current_index][0], 
                            big_data_matrix[current_index][1], 
                            big_data_matrix[current_index + 1][1], 
                            big_data_matrix[current_index + 2][1], 
                            big_data_matrix[current_index + 3][1]
                        ]
                        current_index += 3
                    else:
                        print(f"Insufficient data for nsa_compare at index {current_index}. Moving to the next period.")
                        break  # Exit if not enough values
                
                else:  # Other functions expect single values
                    values = [big_data_matrix[current_index]]
                    current_index += 1
                
                try:
                    if key == 1:
                        compare_function([values], ws, df, red_fill)  # Pass as a list
                    else:
                        compare_function(values, ws, df, red_fill)
                except Exception as e:
                    print(f"Error in function {compare_function.__name__}: {e}")

            # Increment the index to move to the next time period for the next iteration
            if current_index >= len(big_data_matrix):  # Check if we have reached the end
                break

        # Save the modified Excel file
        wb.save("naujas_combined.xls")
        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

big_data_matrix = vda_get_data_matrix_one_indicator(main_url)
print(big_data_matrix)

# Compare data
if big_data_matrix:
    big_compare('laikinas.xls',big_data_matrix)