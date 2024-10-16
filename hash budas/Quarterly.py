import requests
import re
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf

#Būstų kainų indeksai. NIEKO NEKEISTI
rc_index_bustai_url = 'https://www.registrucentras.lt/statdata/indeks_bustai.json'

#https://osp.stat.gov.lt/lt/statistiniu-rodikliu-analize?hash=c8af8a3d-eda3-47b2-a694-2286dae106a3#/

#------------------------------MECHANINIS DARBAS------------------------------
#Padarykite norimo excel failo kopiją, su kuriuo norite lyginti. Užvadinkite jį "laikinas.xls"

#Būsto kainų indeksai. Pakeisti tik kas eina po 'hash/'. pvz., c1980573-e9a6-4003-a5ee-071886b35114
vda_HPI_Q_bustu_isigijimas_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/c6b03430-9f48-489c-9b2b-91a924f9e663'

#Valdžios sektoriaus skola. Pakeisti tik kas eina po 'hash/'.
vda_OSP_valdzios_skolos_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/cf18c9b6-243d-44d8-a5a2-1d48daad9de9'

#BVP indeksai. Pakeisti tik kas eina po 'hash/'.
vda_bvp_ind_2015_Q_nepasalinus_sezono_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/5ccba46a-3b3d-49ed-a322-151755987e8e'

#Nacionalinės pajamos, santaupos, grynasis skolinimas / skolinimasis. Pakeisti tik kas eina po 'hash/'.
vda_NAC_S_Q_grynosios_santaupos_nepasalinus_sezono_url = 'https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/3a76f372-0bb6-4c3c-8f9c-23974d50afad'

#Nedarbo lygis, Užimti gyventojai. Pakeisti tik kas eina po 'hash/'.
vda_NL_UG_nedarbo_lygis_uzimti_gyventojai_url='https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/f9325464-e347-4b41-949a-73eb1bc63c82'

#Darbo užmokestis (mėnesinis). Pakeisti tik kas eina po 'hash/'.
vda_DU_Q_darbo_uzmokestis_url='https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/fb1aa404-22a2-48ab-9824-b543541b6bfd'

#Pramonės produkcijos (be PVM ir akcizo) indeksai. Pakeisti tik kas eina po 'hash/'.
vda_PRAM_C_2021_Q_pramones_produkcijos_indeksas_url='https://osp.stat.gov.lt/analysis-portlet/services/api/v1/data/generate/table/hash/8c691d8a-7717-4c95-9532-3f272373e038'


#------------------------------MECHANINIS DARBAS BAIGIASI------------------------------

def sanitize_value(value):
    return value.replace('\xa0', '').strip()

def switch_columns(data_dict, index1, index2):
    # Create a new dictionary to store the modified data
    switched_dict = {}
    for key, values in data_dict.items():
        if len(values) > max(index1, index2):
            # Swap the specified columns
            values[index1], values[index2] = values[index2], values[index1]
        switched_dict[key] = values
    return switched_dict

def rc_get_data_matrix_one_indicator(api_url, verify=False):
    response = requests.get(api_url, verify=verify)
    data_matrix = []

    if response.status_code == 200:
        data = response.json()
        for item in data:
            if item.get('kur') == 'Lietuvoje' and item.get('tipas') == 'Viso fondo':
                metket = item.get('metket').strip()
                metket_short = re.findall(r'\d+', metket)[0]  # Extract the year

                if 'IV ketv.' in metket:
                    metket_short += '4'
                elif 'III ketv.' in metket:
                    metket_short += '3'
                elif 'II ketv.' in metket:
                    metket_short += '2'
                elif 'I ketv.' in metket:
                    metket_short += '1'

                filtered_item = {'metket': metket_short, 'vidprc': item.get('vidprc')}
                data_matrix.append(filtered_item)

    return data_matrix

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
def compare_rc_index_bustai(data_matrix, ws, df, red_fill):
    for row in data_matrix:
        num = row['metket']  # Get the metket value from API
        value = row['vidprc'].replace(',', '.')
        
        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to one decimal place
        except ValueError:
            print(f"Warning: Value '{value}' cannot be converted to float.")
            continue

        match = df[df['A'].astype(str).str.strip() == num]

        if not match.empty:
            index = match.index[0]
            excel_value = str(df.at[index, 'F']).replace(',', '.')

            try:
                excel_float_value = float(excel_value)
            except ValueError:
                excel_float_value = None  # Handle the case where the cell is empty

            if excel_float_value is None or float(excel_float_value) != float(formatted_value):
                cell_row = index + 3  # Adjust row index based on header rows
                ws.write(cell_row, 5, formatted_value, red_fill)  # Write new value and apply red fill
        else:
            print(f"No matching row found in Excel for metket: {num}")

def compare_vda_bustu_isigijimas(data_dict, ws, df, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        value1 = values[0].replace(',', '.')
        value2 = values[1].replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.2f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.2f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value1 = str(df.at[index, 'G']).replace(',', '.')
            excel_value2 = str(df.at[index, 'H']).replace(',', '.')

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

            formatted_excel_value1 = f"{excel_float_value1:.2f}"  # Format existing value
            formatted_excel_value2 = f"{excel_float_value2:.2f}"  # Format existing value

            cell_row = index + 3  # Adjust row index based on header rows

            if float(formatted_excel_value1) != float(formatted_value1):
                ws.write(cell_row, 6, formatted_value1, red_fill)  # Write new value and apply red fill
            if float(formatted_excel_value2) != float(formatted_value2):
                ws.write(cell_row, 7, formatted_value2, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 6, formatted_value1)
            ws.write(new_row_index, 7, formatted_value2)
def compare_vda_valdzios_skola(data_dict, ws, df, red_fill):
    for timeperiod, values in data_dict.items():
        num = timeperiod

        value1 = values[0].replace(',', '.')
        value2 = values[1].replace(',', '.')

        try:
            float_value1 = float(value1)
            formatted_value1 = f"{float_value1:.0f}"
        except ValueError:
            print(f"Warning: Value '{value1}' cannot be converted to float.")
            continue

        try:
            float_value2 = float(value2)
            formatted_value2 = f"{float_value2:.0f}"
        except ValueError:
            print(f"Warning: Value '{value2}' cannot be converted to float.")
            continue

        match = df[df['A'] == num]

        if not match.empty:
            index = match.index[0]
            excel_value1 = str(df.at[index, 'Z']).replace(',', '.')
            excel_value2 = str(df.at[index, 'AA']).replace(',', '.')

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

            formatted_excel_value1 = f"{excel_float_value1:.0f}"  # Format existing value
            formatted_excel_value2 = f"{excel_float_value2:.0f}"  # Format existing value

            cell_row = index + 3  # Adjust row index based on header rows

            if float(formatted_excel_value1) != float(formatted_value1):
                ws.write(cell_row, 25, formatted_value1, red_fill)  # Write new value and apply red fill
            if float(formatted_excel_value2) != float(formatted_value2):
                ws.write(cell_row, 26, formatted_value2, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 25, formatted_value1)
            ws.write(new_row_index, 26, formatted_value2)

def compare_vda_vda_bvp_nepasalinus_sezono(data_matrix, ws, df, red_fill):
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

def compare_vda_nedarbo_lygis_uzimti_gyventojai(data_dict, ws, df, red_fill):
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
            excel_value1 = str(df.at[index, 'AK']).replace(',', '.')
            excel_value2 = str(df.at[index, 'AL']).replace(',', '.')

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
                ws.write(cell_row, 36, formatted_value1, red_fill)  # Write new value and apply red fill
            if float(formatted_excel_value2) != float(formatted_value2):
                ws.write(cell_row, 37, formatted_value2, red_fill)  # Write new value and apply red fill
        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, num)
            ws.write(new_row_index, 36, formatted_value1)
            ws.write(new_row_index, 37, formatted_value2)
def compare_vda_grynosios_santaupos_nepasalinus_sezono(data_matrix, ws, df, red_fill):
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

def compare_vda_darbo_uzmokestis(data_matrix, ws, df, red_fill):
    for row in data_matrix:
        num = row[0]
        value = row[1].replace(',', '.')

        try:
            float_value = float(value)
            formatted_value = f"{float_value:.1f}"  # Format number to four decimal places
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
def compare_vda_pramones_produkcijos_indeksas(data_matrix, ws, df, red_fill):
    for row in data_matrix:
        num = row[0]
        value = row[1].replace(',', '.')

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


def process_excel_file(excel_file, data_matrix_single,bustu_isijgijimas_double_data_matrix,valdzios_skola_double_data_matrix,vda_bvp_nepasalinus_sezono_data_matrix,vda_grynosios_santaupos_nepasalinus_sezono_data_matrix,vda_nedarbo_lygis_data_matrix,vda_darbo_uzmokestis_data_matrix,vda_pramones_produkcijos_indeksas_data_matrix):
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
        df.columns = ['A', 'F','G','H','Z','AA','AC','AJ','AK','AL','AM','AY']  # Rename columns to something easier to work with

        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        for i in range(4, rs.nrows):
            for j in range(2, 50):
                cell_value = rs.cell_value(i, j)
                ws.write(i, j, cell_value, no_fill)

        compare_rc_index_bustai(data_matrix_single, ws, df, red_fill)
        compare_vda_bustu_isigijimas(bustu_isijgijimas_double_data_matrix, ws, df, red_fill)
        compare_vda_valdzios_skola(valdzios_skola_double_data_matrix, ws, df, red_fill)
        compare_vda_vda_bvp_nepasalinus_sezono(vda_bvp_nepasalinus_sezono_data_matrix, ws, df, red_fill)
        compare_vda_grynosios_santaupos_nepasalinus_sezono(vda_grynosios_santaupos_nepasalinus_sezono_data_matrix, ws, df, red_fill)
        compare_vda_nedarbo_lygis_uzimti_gyventojai(vda_nedarbo_lygis_data_matrix, ws, df, red_fill)
        compare_vda_darbo_uzmokestis(vda_darbo_uzmokestis_data_matrix, ws, df, red_fill)
        compare_vda_pramones_produkcijos_indeksas(vda_pramones_produkcijos_indeksas_data_matrix, ws, df, red_fill)

        wb.save("naujas_combined.xls")

        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Fetch data from the API
rc_bustai_matrix = rc_get_data_matrix_one_indicator(rc_index_bustai_url)
vda_bustu_isigijimas_two_indicators_matrix = vda_get_data_matrix_two_indicators(vda_HPI_Q_bustu_isigijimas_url)
vda_valdzios_skolos_two_indicators_matrix = vda_get_data_matrix_two_indicators(vda_OSP_valdzios_skolos_url)
vda_bvp_nepasalinus_sezono_data_matrix = vda_get_data_matrix_one_indicator(vda_bvp_ind_2015_Q_nepasalinus_sezono_url)
vda_grynosios_santaupos_nepasalinus_sezono_data_matrix = vda_get_data_matrix_one_indicator(vda_NAC_S_Q_grynosios_santaupos_nepasalinus_sezono_url)
vda_nedarbo_lygis_uzimti_gyventojai_double_data_matrix = vda_get_data_matrix_two_indicators(vda_NL_UG_nedarbo_lygis_uzimti_gyventojai_url)
vda_nedarbo_lygis_uzimti_gyventojai_double_data_matrix = switch_columns(vda_nedarbo_lygis_uzimti_gyventojai_double_data_matrix, 0, 1)
vda_darbo_uzmokestis_data_matrix = vda_get_data_matrix_one_indicator(vda_DU_Q_darbo_uzmokestis_url)
vda_pramones_produkcijos_indeksas_data_matrix = vda_get_data_matrix_one_indicator(vda_PRAM_C_2021_Q_pramones_produkcijos_indeksas_url)
# Process the Excel file
process_excel_file('laikinas.xls', rc_bustai_matrix,vda_bustu_isigijimas_two_indicators_matrix,vda_valdzios_skolos_two_indicators_matrix,vda_bvp_nepasalinus_sezono_data_matrix,vda_grynosios_santaupos_nepasalinus_sezono_data_matrix,vda_nedarbo_lygis_uzimti_gyventojai_double_data_matrix,vda_darbo_uzmokestis_data_matrix,vda_pramones_produkcijos_indeksas_data_matrix)
