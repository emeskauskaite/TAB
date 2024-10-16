#https://osp.stat.gov.lt/lt/statistiniu-rodikliu-analize?hash=414a78a9-0368-4e1c-bae2-123595dd5b5a
import requests
import re
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf

# API URL
main_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/414a78a9-0368-4e1c-bae2-123595dd5b5a'

def sanitize_value(value):
    # Remove non-breaking spaces, replace commas with dots, and strip leading/trailing spaces
    return value.replace('\xa0', '').strip()


def compare_vda_HPI_Q(data_list, ws, df, red_fill):
    index = 0
    while index < len(data_list):
        if index + 1 >= len(data_list):
            print(f"Warning: Not enough values to process for timeperiod '{data_list[index][0]}'.")
            break

        # Extract the timeperiod and values
        timeperiod1 = data_list[index][0]
        value1 = sanitize_value(str(data_list[index][1])).replace(',', '.')
        timeperiod2 = data_list[index + 1][0]
        value2 = sanitize_value(str(data_list[index + 1][1])).replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.2f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.2f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        match = df[df['A'] == timeperiod1]  # Use `timeperiod1` for matching

        if not match.empty:
            match_index = match.index[0]
            excel_value1 = sanitize_value(str(df.at[match_index, 'G'])).replace(',', '.')
            excel_value2 = sanitize_value(str(df.at[match_index, 'H'])).replace(',', '.')

            try:
                excel_float_value1 = float(excel_value1)
                formatted_excel_value1 = f"{excel_float_value1:.2f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                excel_float_value1 = None  # Set to None if conversion fails

            try:
                excel_float_value2 = float(excel_value2)
                formatted_excel_value2 = f"{excel_float_value2:.2f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                excel_float_value2 = None  # Set to None if conversion fails

            cell_row = match_index + 3  # Adjust row index based on header rows

            # Write new value and apply red fill if different
            if excel_float_value1 is not None and float(formatted_value1) != excel_float_value1:
                ws.write(cell_row, 6, formatted_value1, red_fill)
            if excel_float_value2 is not None and float(formatted_value2) != excel_float_value2:
                ws.write(cell_row, 7, formatted_value2, red_fill)
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, timeperiod1)
            ws.write(new_row_index, 6, formatted_value1)
            ws.write(new_row_index, 7, formatted_value2)

        index += 2  # Move to the next pair of values



def compare_vda_OSP(data_list, ws, df, red_fill):
    index = 0
    while index < len(data_list):
        if index + 1 >= len(data_list):
            print(f"Warning: Not enough values to process for timeperiod '{data_list[index][0]}'.")
            break

        # Extract the timeperiod and values
        num = data_list[index][0]
        value1 = sanitize_value(str(data_list[index][1])).replace(',', '.')
        value2 = sanitize_value(str(data_list[index + 1][1])).replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.1f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.1f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        match = df[df['A'] == num]  # Ensure `num` is correct identifier for matching

        if not match.empty:
            match_index = match.index[0]
            excel_value1 = sanitize_value(str(df.at[match_index, 'Z'])).replace(',', '.')
            excel_value2 = sanitize_value(str(df.at[match_index, 'AA'])).replace(',', '.')

            try:
                excel_float_value1 = float(excel_value1)
                formatted_excel_value1 = f"{excel_float_value1:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                excel_float_value1 = None  # Set to None if conversion fails

            try:
                excel_float_value2 = float(excel_value2)
                formatted_excel_value2 = f"{excel_float_value2:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                excel_float_value2 = None  # Set to None if conversion fails

            cell_row = match_index + 3  # Adjust row index based on header rows

            # Write new value and apply red fill if different
            if excel_float_value1 is not None and float(formatted_value1) != excel_float_value1:
                ws.write(cell_row, 25, formatted_value1, red_fill)
            if excel_float_value2 is not None and float(formatted_value2) != excel_float_value2:
                ws.write(cell_row, 26, formatted_value2, red_fill)
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 25, formatted_value1)
            ws.write(new_row_index, 26, formatted_value2)

        index += 2  # Move to the next pair of values

def compare_vda_NAC_S_Q(data_matrix, ws, df, red_fill):
    for row in data_matrix:
        num = row[0]
        value = str(row[1].replace(',', '.'))

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AJ']).replace(',', '.')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.4f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 35, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 35, formatted_value)

def compare_vda_NL_UG(data_list, ws, df, red_fill):
    index = 0
    while index < len(data_list):
        if index + 1 >= len(data_list):
            print(f"Warning: Not enough values to process for timeperiod '{data_list[index][0]}'.")
            break

        # Extract the timeperiod and values
        num = data_list[index][0]
        value1 = sanitize_value(str(data_list[index][1])).replace(',', '.')
        value2 = sanitize_value(str(data_list[index + 1][1])).replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.1f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.1f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            index += 2  # Move to next pair
            continue

        match = df[df['A'] == num]  # Ensure `num` is the correct identifier for matching

        if not match.empty:
            match_index = match.index[0]
            excel_value1 = sanitize_value(str(df.at[match_index, 'AK'])).replace(',', '.')
            excel_value2 = sanitize_value(str(df.at[match_index, 'AL'])).replace(',', '.')

            try:
                excel_float_value1 = float(excel_value1)
                formatted_excel_value1 = f"{excel_float_value1:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                excel_float_value1 = None  # Set to None if conversion fails

            try:
                excel_float_value2 = float(excel_value2)
                formatted_excel_value2 = f"{excel_float_value2:.1f}"
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                excel_float_value2 = None  # Set to None if conversion fails

            cell_row = match_index + 3  # Adjust row index based on header rows

            # Write new value and apply red fill if different
            if excel_float_value1 is not None and float(formatted_value1) != excel_float_value1:
                ws.write(cell_row, 36, formatted_value1, red_fill)
            if excel_float_value2 is not None and float(formatted_value2) != excel_float_value2:
                ws.write(cell_row, 37, formatted_value2, red_fill)
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 36, formatted_value1)
            ws.write(new_row_index, 37, formatted_value2)

        index += 2  # Move to the next pair of values

def compare_vda_bvp_ind_2015(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Convert single item to list

    for row in data_matrix:
        num = row[0]
        value = row[1].replace(',', '.')

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AC']).replace(',', '.')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.4f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 28, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 28, formatted_value)



def compare_vda_DU_Q(data_matrix, ws, df, red_fill):
    if not isinstance(data_matrix, list):
        data_matrix = [data_matrix]  # Convert single item to list

    for row in data_matrix:
        num = row[0]
        value = str(row[1]).replace(',', '.')

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to one decimal place
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AM']).replace(',', '.')

            try:
                excel_float_value = float(excel_value)
                formatted_excel_value = f"{excel_float_value:.1f}"  # Format existing value
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            # Write new value and apply red fill if different
            if float(formatted_value) != float(formatted_excel_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 38, formatted_value, red_fill)
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 38, formatted_value)


def compare_vda_PRAM_C_2021_Q(data_matrix, ws, df, red_fill):
    for row in data_matrix:
        num = row[0]
        value = str(row[1].replace(',', '.'))

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AY']).replace(',', '.')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                print(f"Warning: Existing value '{excel_value}' cannot be converted to float.")
                continue

            formatted_excel_value = f"{excel_float_value:.1f}"  # Format existing value

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 50, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 50, formatted_value)


def vda_get_data_matrix_one_indicator(api_url, verify=False):
    response = requests.get(api_url, verify=verify)
    data_matrix = []

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'itemMatrix' in data['data']:
            item_matrix = data['data']['itemMatrix']

            for listed in item_matrix:
                for item in listed:
                    if 'value' in item:
                        value = sanitize_value(item['value'])
                    timeperiod = None
                    for coordinate in item['coordinate']:
                        if coordinate.startswith('L#LAIKOTARPIS'):
                            year_str = re.findall(r'\d+', coordinate.split('#')[2])[0]
                            month_str = re.findall(r'\d+', coordinate.split('#')[2])[1]
                            timeperiod = int(year_str + month_str)
                            break

                    if timeperiod is not None:
                        data_matrix.append([timeperiod, value])
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

    return data_matrix

# Define your comparison functions for each indicator
# These should already be defined: compare_vda_HPI_Q, compare_vda_OSP, etc.

def big_compare(excel_file, big_data_matrix):
    try:
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Quarterly')
        wb = copy(rb)
        ws = wb.get_sheet('Quarterly')

        df = pd.read_excel(
            excel_file, 
            sheet_name='Quarterly', 
            usecols=[0, 5, 6, 7,25,26,28,35,36,37,38,50],  
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'F','G','H','Z','AA','AC','AJ','AK','AL','AM','AY']

        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        # Clear existing cell formatting
        for i in range(4, rs.nrows):
            for j in range(2, 50):
                cell_value = rs.cell_value(i, j)
                ws.write(i, j, cell_value, no_fill)

        # Mapping function keys to their corresponding functions
        compare_functions = {
            0: compare_vda_HPI_Q,
            2: compare_vda_OSP,
            4: compare_vda_bvp_ind_2015, 
            5: compare_vda_NAC_S_Q,
            6: compare_vda_NL_UG,
            8: compare_vda_DU_Q,
            9: compare_vda_PRAM_C_2021_Q
        }

        index = 0
        while index < len(big_data_matrix):
            for key in sorted(compare_functions.keys()):
                if index >= len(big_data_matrix):
                    break

                if key in {0, 2, 6}:  # Expecting pairs of values
                    if index + 1 < len(big_data_matrix):
                        values = [big_data_matrix[index], big_data_matrix[index + 1]]
                        index += 2
                    else:
                        values = [big_data_matrix[index]]
                        index += 1
                elif key == 4:  # Special case for BVP function expecting a list of lists
                    values = big_data_matrix[index]
                    index += 1
                else:  # Expecting single values for other functions
                    values = [big_data_matrix[index]]
                    index += 1

                print(f"Calling function {compare_functions[key]} with values: {values}")

                compare_function = compare_functions[key]
                try:
                    # Adjust call if the function expects multiple values
                    if key == 4:  # Special case for BVP
                        compare_function([values], ws, df, red_fill)  # Ensure it's passed as a list
                    else:
                        compare_function(values, ws, df, red_fill)
                except Exception as e:
                    print(f"Error in function {compare_functions[key]}: {e}")

        wb.save("naujas_combined.xls")
        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def sanitize_value(value):
    # Remove non-breaking spaces, replace commas with dots, and strip leading/trailing spaces
    return value.replace('\xa0', '').strip()

def vda_get_data_matrix_one_indicator(api_url, verify=False):
    response = requests.get(api_url, verify=verify)
    data_matrix = []

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'itemMatrix' in data['data']:
            item_matrix = data['data']['itemMatrix']
            counter = 0

            for listed in item_matrix:
                for item in listed:
                    if item is not None and 'value' in item:  # Check if item is not None before accessing it
                        value = sanitize_value(item['value'])
                    else:
                        value = None

                    timeperiod = None
                    if item is not None:  # Check if item is not None before accessing coordinate
                        for coordinate in item['coordinate']:
                            if coordinate.startswith('L#LAIKOTARPIS'):
                                year_str = re.findall(r'\d+', coordinate.split('#')[2])[0]
                                month_str = re.findall(r'\d+', coordinate.split('#')[2])[1]
                                timeperiod = int(year_str + month_str)
                                break

                    if timeperiod is not None:
                        data_matrix.append([timeperiod, value])
                        counter += 1
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

    return data_matrix

# Fetch data
big_data_matrix = vda_get_data_matrix_one_indicator(main_url)
print(big_data_matrix)

# Compare data
if big_data_matrix:
    big_compare('laikinas.xls',big_data_matrix)