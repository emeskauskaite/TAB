import requests
import re
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf
#------------------------------MECHANINIS DARBAS------------------------------
#Padarykite norimo excel failo kopiją, su kuriuo norite lyginti. Užvadinkite jį "laikinas.xls"


#https://osp.stat.gov.lt/lt/statistiniu-rodikliu-analize?hash=357bde94-fa9f-40d6-89fa-782098fa6667



#Vartotojų kainų indeksas. Pakeisti tik kas eina po 'hash/'. pvz., c1980573-e9a6-4003-a5ee-071886b35114
vki_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/c1980573-e9a6-4003-a5ee-071886b35114'

#Suderinti vartotojų kainų indeksai. Pakeisti tik kas eina po 'hash/'.
core_hicp_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/3ea702c4-7be6-43f1-9b3a-b8c0d8875d9d'

#Gamintojų parduotos pramonės produkcijos kainų indeksai. Pakeisti tik kas eina po 'hash/'.
gki_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/8dfbaf3b-bd4e-4783-92c6-b9aef33520c3'

#Eksportuotų prekių kainų indeksai, Importuotų prekių kainų indeksai. Pakeisti tik kas eina po 'hash/'.
eki_iki_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/09ee75cd-5841-4bcb-b111-df82f59ffaf4'

#Pramonės produkcijos (be PVM ir akcizo) indeksai. Pakeisti tik kas eina po 'hash/'.
pram_c_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/184acd2d-c01e-416a-b976-6518217733c7'

#Eksportas, Importas, Eksportas, pašalinus sezono įtaką, Importas, pašalinus sezono įtaką. Pakeisti tik kas eina po 'hash/'.
nsa_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/816bb9a1-5ee4-4dcd-b74f-f1620ff428d0'


#------------------------------MECHANINIS DARBAS BAIGIASI------------------------------
def switch_columns(data_dict, index1, index2):
    # Create a new dictionary to store the modified data
    switched_dict = {}
    for key, values in data_dict.items():
        if len(values) > max(index1, index2):
            # Swap the specified columns
            values[index1], values[index2] = values[index2], values[index1]
        switched_dict[key] = values
    return switched_dict


def vki_compare(data_matrix, ws, df, no_fill, red_fill):
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
            excel_value = str(df.at[index, 'AB']).replace(',', '.')

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

def hicp_compare(data_dict, ws, df, no_fill, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        value1 = values[0].replace(',', '.')
        value2 = values[1].replace(',', '.')

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
            excel_value1 = str(df.at[index, 'AD']).replace(',', '.')
            excel_value2 = str(df.at[index, 'AE']).replace(',', '.')

            try:
                excel_float_value1 = float(excel_value1)
            except ValueError:
                print(f"Warning: Existing value '{excel_value1}' cannot be converted to float.")
                continue

            try:
                excel_float_value2 = float(excel_value2)
            except ValueError:
                print(f"Warning: Existing value '{excel_value2}' cannot be converted to float.")
                continue

            formatted_excel_value1 = f"{excel_float_value1:.1f}"  # Format existing value
            formatted_excel_value2 = f"{excel_float_value2:.1f}"  # Format existing value

            cell_row = index + 3  # Adjust row index based on header rows

            if float(formatted_excel_value1) != float(formatted_value1):
                ws.write(cell_row, 29, formatted_value1, red_fill)  # Write new value and apply red fill
            if float(formatted_excel_value2) != float(formatted_value2):
                ws.write(cell_row, 30, formatted_value2, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 29, formatted_value1)
            ws.write(new_row_index, 30, formatted_value2)

def gki_compare(data_matrix, ws, df, no_fill, red_fill):
    for row in data_matrix:
        num = row[0]
        value = row[1].replace(',', '.')

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to one decimal place
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'AF']).replace(',', '.')

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



def pram_c_compare(data_matrix, ws, df, no_fill, red_fill):
     for row in data_matrix:
        num = row[0]
        value = row[1].replace(',', '.')

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


def nsa_compare(data_dict, ws, df, no_fill, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        # Convert and format the values
        value1 = values[0].replace(',', '.')
        value2 = values[1].replace(',', '.')
        value3 = values[2].replace(',', '.')
        value4 = values[3].replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.1f}"
        except ValueError:
            print(f"Error converting value '{value1}' to float.")
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.1f}"
        except ValueError:
            print(f"Error converting value '{value2}' to float.")
            continue

        try:
            float_value3 = float(value3)
            formatted_value3 = f"{float_value3:.1f}"
        except ValueError:
            print(f"Error converting value '{value3}' to float.")
            continue

        try:
            float_value4 = float(value4)
            formatted_value4 = f"{float_value4:.1f}"
        except ValueError:
            print(f"Error converting value '{value4}' to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]

            # Function to safely get float value from a cell
            def get_float(value):
                try:
                    return float(value.replace(',', '.'))
                except (ValueError, AttributeError):
                    return float('nan')

            # Retrieve existing values
            excel_value1 = get_float(str(df.at[index, 'AZ'])) if 'AZ' in df.columns else float('nan')
            excel_value2 = get_float(str(df.at[index, 'BA'])) if 'BA' in df.columns else float('nan')
            excel_value3 = get_float(str(df.at[index, 'BB'])) if 'BB' in df.columns else float('nan')
            excel_value4 = get_float(str(df.at[index, 'BC'])) if 'BC' in df.columns else float('nan')

            cell_row = index + 3

            # Apply formatting based on comparisons
            if pd.isna(excel_value1) or excel_value1 != float(formatted_value1):
                ws.write(cell_row, 51, formatted_value1, red_fill)  # Column AZ
            else:
                ws.write(cell_row, 51, formatted_value1, no_fill)  # Column AZ

            if pd.isna(excel_value2) or excel_value2 != float(formatted_value2):
                ws.write(cell_row, 52, formatted_value2, red_fill)  # Column BA
            else:
                ws.write(cell_row, 52, formatted_value2, no_fill)  # Column BA

            if pd.isna(excel_value3) or excel_value3 != float(formatted_value3):
                ws.write(cell_row, 53, formatted_value3, red_fill)  # Column BB
            else:
                ws.write(cell_row, 53, formatted_value3, no_fill)  # Column BB

            if pd.isna(excel_value4) or excel_value4 != float(formatted_value4):
                ws.write(cell_row, 54, formatted_value4, red_fill)  # Column BC
            else:
                ws.write(cell_row, 54, formatted_value4, no_fill)  # Column BC
        else:
            # Adding new rows for missing time periods
            new_row_index = len(df) + 3
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 51, formatted_value1)  # Column AZ
            ws.write(new_row_index, 52, formatted_value2)  # Column BA
            ws.write(new_row_index, 53, formatted_value3)  # Column BB
            ws.write(new_row_index, 54, formatted_value4)  # Column BC



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

def vda_get_data_matrix_two_indicators(api_url, verify=False):
    data_matrix = vda_get_data_matrix_one_indicator(api_url)
    data_dict = {}
    for entry in data_matrix:
        timeperiod = entry[0]
        value = entry[1]
        if timeperiod not in data_dict:
            data_dict[timeperiod] = []
        data_dict[timeperiod].append(value)

    return data_dict

def vda_get_data_matrix_four_indicators(api_url, verify=False):
    response = requests.get(api_url, verify=verify)
    data_dict = {}

    if response.status_code == 200:
        data = response.json()

        if 'data' in data and 'itemMatrix' in data['data']:
            item_matrix = data['data']['itemMatrix']

            for listed in item_matrix:
                for item in listed:
                    values = []
                    if 'value' in item:
                        raw_values = item['value']
                        values = [sanitize_value(v) for v in raw_values]
                    
                    timeperiod = None
                    for coordinate in item['coordinate']:
                        if coordinate.startswith('L#LAIKOTARPIS'):
                            year_str = re.findall(r'\d+', coordinate.split('#')[2])[0]
                            month_str = re.findall(r'\d+', coordinate.split('#')[2])[1]
                            timeperiod = int(year_str + month_str)
                            break

                    if timeperiod is not None:
                        data_dict[timeperiod] = values
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

    return data_dict



def eki_compare(data_dict, ws, df, no_fill, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        if len(values) > 1:
            value1 = values[0].replace(',', '.')
        else:
            print(f"Warning: Not enough values for timeperiod {timeperiod}.")
            continue

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value1 = str(df.at[index, 'AG']).replace(',', '.')

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

def iki_compare(data_dict, ws, df, no_fill, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        if len(values) > 1:
            value2 = values[1].replace(',', '.')
        else:
            print(f"Warning: Not enough values for timeperiod {timeperiod}.")
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.4f}"  # Format number to four decimal places
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value2 = str(df.at[index, 'AO']).replace(',', '.')

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

def process_excel_file(excel_file, data_matrix_single, data_matrix_double, gki_data_matrix_single,pram_c_matrix_single, eki_iki_data_matrix_double, nsa_matrix_quatruple):
    try:
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Monthly')
        wb = copy(rb)
        ws = wb.get_sheet('Monthly')

        df = pd.read_excel(
            excel_file, 
            sheet_name='Monthly', 
            usecols=[0, 27, 29, 30, 31, 32, 38,40, 51, 52, 53, 54],
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'AB', 'AD', 'AE', 'AF', 'AG', 'AM', 'AO', 'AZ', 'BA', 'BB', 'BC']
        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        for i in range(4, rs.nrows):
            for j in range(27, 54):  
                cell_value = rs.cell_value(i, j)
                ws.write(i, j, cell_value, no_fill)

        insert_column_am(ws, df)

        vki_compare(data_matrix_single, ws, df, no_fill, red_fill)
        hicp_compare(data_matrix_double, ws, df, no_fill, red_fill)
        gki_compare(gki_data_matrix_single, ws, df, no_fill, red_fill)
        pram_c_compare(pram_c_matrix_single, ws, df, no_fill, red_fill)
        eki_compare(eki_iki_data_matrix_double, ws, df, no_fill, red_fill)
       
        iki_compare(eki_iki_data_matrix_double, ws, df, no_fill, red_fill)
        nsa_compare(nsa_matrix_quatruple, ws, df, no_fill, red_fill)
        
        wb.save("naujas_combined.xls")

        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


def insert_column_am(ws, df):
    for index, row in enumerate(df.itertuples(index=False), start=1):
        ws.write(index + 2, 55, index)


vki_data_matrix_single = vda_get_data_matrix_one_indicator(vki_url)
core_hicp_data_matrix_double = vda_get_data_matrix_two_indicators(core_hicp_url)
gki_data_matrix_single = vda_get_data_matrix_one_indicator(gki_url)
eki_iki_matrix_double = vda_get_data_matrix_two_indicators(eki_iki_url)
eki_iki_matrix_double = switch_columns(eki_iki_matrix_double, 0, 1)
pram_c_matrix_single = vda_get_data_matrix_one_indicator(pram_c_url)
nsa_matrix_quatruple=vda_get_data_matrix_two_indicators(nsa_url)

process_excel_file('laikinas.xls', vki_data_matrix_single, core_hicp_data_matrix_double, gki_data_matrix_single, pram_c_matrix_single,eki_iki_matrix_double, nsa_matrix_quatruple)
