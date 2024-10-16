from decimal import Decimal
import requests
import pandas as pd
import xlrd
import re
from xlutils.copy import copy
from xlwt import easyxf
from lxml import etree
import math


def is_numeric(value):
    """Check if a string can be converted to a float."""
    try:
        float(value)
        return True
    except ValueError:
        return False
    
def round_half_up(n, decimals=0):
    multiplier = 10**decimals
    return math.floor(n * multiplier + 0.5) / multiplier
    
def round_half_away_from_zero(n, decimals=0):
    rounded_abs = round_half_up(abs(n), decimals)
    return math.copysign(rounded_abs, n)

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

def hpi_q_api_process(hpi_q_url):
    response = requests.get(hpi_q_url, verify=False)
    
    if response.status_code == 200:

        root = etree.fromstring(response.content)

        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            savivaldybe = obs.xpath('./g:ObsKey/g:Value[@id="savivaldybeM2021001"]/@value', namespaces=namespaces)
            busto_tipas = obs.xpath('./g:ObsKey/g:Value[@id="bustotipasM2021001"]/@value', namespaces=namespaces)
            bustas = obs.xpath('./g:ObsKey/g:Value[@id="bustasM2021"]/@value', namespaces=namespaces)
            matvnt = obs.xpath('./g:ObsKey/g:Value[@id="MATVNT"]/@value', namespaces=namespaces)

            if (savivaldybe and (savivaldybe[0] == "00" or savivaldybe[0] == "13") and
                busto_tipas and busto_tipas[0] == "11" and
                bustas and bustas[0] == "H1" and
                matvnt and matvnt[0] == "nera" and
                obs_value):

                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0],
                    'Savivaldybe': savivaldybe[0],
                })
        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None


def hpi_q_compare(df_filtered, ws, df, red_fill):
    for index, row in df_filtered.iterrows():
        laikotarpis = row['LAIKOTARPIS'].replace('K', '')  
        obs_value = row['ObsValue'].replace(',', '.') 
        savivaldybe = row['Savivaldybe'] 

        if not laikotarpis or not obs_value or not savivaldybe:
            print(f"Invalid entry data: {laikotarpis}, {obs_value}, {savivaldybe}")
            continue

        try:
            formatted_value = float(obs_value)
            formatted_value = round_half_up(formatted_value, 2)
        except ValueError:
            print(f"Skipping invalid obs_value: {obs_value}")
            continue

        df['A'] = df['A'].astype(str)
        match = df[df['A'] == laikotarpis] 

        if not match.empty:
            match_index = match.index[0]

            if savivaldybe == '00':
                excel_value = str(df.at[match_index, 'G']).replace(',', '.')

                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_up(excel_value, 2)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 6, f"{formatted_value:.2f}".replace(',', '.'), red_fill)

            elif savivaldybe == '13':
                excel_value = str(df.at[match_index, 'H']).replace(',', '.')
                
                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_up(excel_value, 2)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 7, f"{formatted_value:.2f}".replace(',', '.'), red_fill)

        else:
            new_row_index = len(df) + 3
            ws.write(new_row_index, 0, laikotarpis)

            if savivaldybe == '00':
                ws.write(new_row_index, 6, f"{formatted_value:.2f}".replace(',', '.'), red_fill)
            elif savivaldybe == '13':
                ws.write(new_row_index, 7, f"{formatted_value:.2f}".replace(',', '.'), red_fill)


def osp_api_process(osp_url):
    response = requests.get(osp_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)

        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }
        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            skola = obs.xpath('./g:ObsKey/g:Value[@id="skolaM2040104"]/@value', namespaces=namespaces)
            matvnt = obs.xpath('./g:ObsKey/g:Value[@id="MATVNT"]/@value', namespaces=namespaces)

            if (skola and (skola[0] == "ggd20" or skola[0] == "ggd21") and
                matvnt and matvnt[0] == "mln_euru" and
                obs_value):

                # Append the filtered data
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0],
                    'Skola': skola[0],
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None


def osp_compare(df_filtered, ws, df, red_fill):
    for index, row in df_filtered.iterrows():
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 
        obs_value = row['ObsValue'].replace(',', '.') 
        skola = row['Skola'] 
        if not laikotarpis or not obs_value or not skola:
            print(f"Invalid entry data: {laikotarpis}, {obs_value}, {skola}")
            continue

        try:
            formatted_value = float(obs_value)
            formatted_value = round_half_away_from_zero(formatted_value)
        except ValueError:
            print(f"Skipping invalid obs_value: {obs_value}")
            continue
        
        df['A'] = df['A'].astype(str)
        match = df[df['A'] == laikotarpis]  

        if not match.empty:
            match_index = match.index[0]

            if skola == 'ggd20':
                excel_value = str(df.at[match_index, 'Z']).replace(',', '.')

                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_away_from_zero(excel_value)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 25, f"{formatted_value:.0f}".replace(',', '.'), red_fill)

            elif skola == 'ggd21':
                excel_value = str(df.at[match_index, 'AA']).replace(',', '.')
                
                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_away_from_zero(excel_value)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 26, f"{formatted_value:.0f}".replace(',', '.'), red_fill)

        else:
            new_row_index = len(df) + 3
            ws.write(new_row_index, 0, laikotarpis)

            if skola == 'ggd20':
                ws.write(new_row_index, 25, f"{formatted_value:.0f}".replace(',', '.'), red_fill)
            elif skola == 'ggd21':
                ws.write(new_row_index, 26, f"{formatted_value:.0f}".replace(',', '.'), red_fill)


def bvp_ind_api_process(bvp_ind_api_url):
    response = requests.get(bvp_ind_api_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)


        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            palyg= obs.xpath('./g:ObsKey/g:Value[@id="LYGINIMAS"]/@value', namespaces=namespaces)
            islyg = obs.xpath('./g:ObsKey/g:Value[@id="Islyginimas_indeksai"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and palyg and palyg[0]=="palyg_2021" and islyg and islyg[0]=="bendras":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    

def bvp_ind_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():

        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AC']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 28, f"{formatted_value:.1f}".replace(',', '.'), red_fill)  

        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 28, f"{formatted_value:.1f}".replace(',', '.'))  


def nac_s_q_api_process(nac_s_q_api_url):
    response = requests.get(nac_s_q_api_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)


        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            pajamos= obs.xpath('./g:ObsKey/g:Value[@id="nacpajamosM2110109"]/@value', namespaces=namespaces)
            islyg = obs.xpath('./g:ObsKey/g:Value[@id="Islyginimas_indeksai"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and pajamos and pajamos[0]=="b8n" and islyg and islyg[0]=="bendras":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def nac_s_q_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():

        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value, 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AJ']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 35, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 

        else:
            new_row_index = len(df) + 3 
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 35, f"{formatted_value:.1f}".replace(',', '.'))  


def ug_api_process(ug_api_url):
    response = requests.get(ug_api_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)


        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            vietove= obs.xpath('./g:ObsKey/g:Value[@id="Vietove"]/@value', namespaces=namespaces)
            lytis = obs.xpath('./g:ObsKey/g:Value[@id="Lytis"]/@value', namespaces=namespaces)
            amzius = obs.xpath('./g:ObsKey/g:Value[@id="AmziusM2111"]/@value', namespaces=namespaces)
            if obs_value and laikotarpis and amzius and amzius[0]=="0" and lytis and lytis[0]=="0" and vietove and vietove[0]=="0":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def ug_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():

        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value, 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AK']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 36, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 

        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 36, f"{formatted_value:.1f}".replace(',', '.'))  

def nl_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():

        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value, 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AL']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 37, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 

        else:
            new_row_index = len(df) + 3 
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 37, f"{formatted_value:.1f}".replace(',', '.')) 

def du_q_api_process(ug_api_url):
    response = requests.get(ug_api_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)


        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            ekosekt= obs.xpath('./g:ObsKey/g:Value[@id="Ekon_sektorM2040803"]/@value', namespaces=namespaces)
            lytis = obs.xpath('./g:ObsKey/g:Value[@id="Lytis"]/@value', namespaces=namespaces)
            darbo =obs.xpath('./g:ObsKey/g:Value[@id="darboM3060321"]/@value', namespaces=namespaces)
            evrk = obs.xpath('./g:ObsKey/g:Value[@id="EVRK2M3060207"]/@value', namespaces=namespaces)
            if obs_value and laikotarpis and ekosekt[0]=="0ex" and lytis[0]=="0" and darbo[0] == "bruto" and evrk[0] == "TOTAL":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def du_q_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value, 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AM']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 38, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:

            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 38, f"{formatted_value:.1f}".replace(',', '.')) 

def pram_c_2021_q_api_process(ug_api_url):
    response = requests.get(ug_api_url, verify=False)
    
    if response.status_code == 200:
        root = etree.fromstring(response.content)


        filtered_data = []

        namespaces = {
            'mes': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message',
            'g': 'http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'
        }

        obs_elements = root.xpath('.//g:Obs', namespaces=namespaces)
        
        for obs in obs_elements:
            obs_value = obs.xpath('g:ObsValue/@value', namespaces=namespaces)
            laikotarpis = obs.xpath('./g:ObsKey/g:Value[@id="LAIKOTARPIS"]/@value', namespaces=namespaces)
            islyg = obs.xpath('./g:ObsKey/g:Value[@id="Islyginimas_indeksai"]/@value', namespaces=namespaces)
            lyg =obs.xpath('./g:ObsKey/g:Value[@id="LYGINIMAS"]/@value', namespaces=namespaces)
            evrk = obs.xpath('./g:ObsKey/g:Value[@id="EVRKM4050107"]/@value', namespaces=namespaces)
            if obs_value and laikotarpis and lyg[0]=="palyg_2021" and islyg[0]=="bendras" and evrk[0] == "C":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def pram_c_2021_q_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('K', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value, 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AY']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 50, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:

            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 50, f"{formatted_value:.1f}".replace(',', '.')) 

def round_excel_column_to_precision(ws, rs, start_row, column_index, precision=1):
    """Round values in a specified column of the Excel sheet to a given precision."""
    for row_index in range(start_row, rs.nrows):  # Use rs.nrows to get the number of rows
        cell_value = rs.cell_value(row_index, column_index)
        
        if isinstance(cell_value, str) and cell_value.strip() == '':
            continue  # Skip empty string values
        
        try:
            if pd.notna(cell_value):  # Check if the value is not NaN
                # Convert cell value to float and round
                rounded_value = round_half_up(float(cell_value), precision)
                # Write the rounded value back to the corresponding cell in Excel
                ws.write(row_index, column_index, f"{rounded_value:.1f}".replace(',', '.'))
        except ValueError:
            print(f"Warning: Could not convert '{cell_value}' to float.")

def process_excel_file(excel_file):
    try:
        print(f"Opening Excel file: {excel_file}")
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Quarterly')
        wb = copy(rb)
        ws = wb.get_sheet('Quarterly')

        print("Reading data from Excel sheet...")
        df = pd.read_excel(
            excel_file, 
            sheet_name='Quarterly', 
            usecols=[0, 5, 6, 7, 25, 26, 28, 35, 36, 37, 38, 50],  
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'F', 'G', 'H', 'Z', 'AA', 'AC', 'AJ', 'AK', 'AL', 'AM', 'AY']
        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        # Clear formatting
        for i in range(4, rs.nrows):
            for j in range(27, 54):  
                cell_value = rs.cell_value(i, j)
                ws.write(i, j, cell_value, no_fill)

        # Round values in AJ column (index 35) to 0.1 precision
        round_excel_column_to_precision(ws, rs, 4, 35)  # Use rs to get the number of rows

        rc_index_bustai_url = 'https://www.registrucentras.lt/statdata/indeks_bustai.json'
        rc_bustai_matrix = rc_get_data_matrix_one_indicator(rc_index_bustai_url)
        compare_rc_index_bustai(rc_bustai_matrix, ws, df, red_fill)

        hpi_q_api_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R251_M2021001_1/?startPeriod=2006-Q1'
        hpi_q_filtered = hpi_q_api_process(hpi_q_api_url)
        hpi_q_compare(hpi_q_filtered, ws, df, red_fill)

        osp_api_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R270_M2040104/?startPeriod=2003-Q4'
        osp_filtered = osp_api_process(osp_api_url)
        osp_compare(osp_filtered, ws, df, red_fill)

        bvp_ind_api_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R201_M2110101_6/?startPeriod=1995-Q1'
        bvp_ind_filtered = bvp_ind_api_process(bvp_ind_api_url)
        bvp_ind_compare(bvp_ind_filtered, ws, df, red_fill)

        nac_s_q_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R217_M2110109/?startPeriod=1999-Q1'
        nac_s_q_filtered = nac_s_q_api_process(nac_s_q_url)
        nac_s_q_compare(nac_s_q_filtered, ws, df, red_fill)

        ug_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S3R532_M3030101_2/?startPeriod=1998-Q2'
        ug_filtered = ug_api_process(ug_url)
        ug_compare(ug_filtered, ws, df, red_fill)

        nl_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S3R347_M3030101_1/?startPeriod=1998-Q2'
        nl_filtered = ug_api_process(nl_url)
        nl_compare(nl_filtered, ws, df, red_fill)

        # Very large file
        du_q_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S3R0050_M3060315_1/?startPeriod=1995-Q1'
        du_q_filtered = du_q_api_process(du_q_url)
        du_q_compare(du_q_filtered, ws, df, red_fill)

        pram_c_2021_q_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S8R918_M4050208_5/?startPeriod=1998-Q1'
        pram_c_2021_q_filtered = pram_c_2021_q_api_process(pram_c_2021_q_url)
        pram_c_2021_q_compare(pram_c_2021_q_filtered, ws, df, red_fill)

        wb.save("naujas_combined.xls")
        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Call the function to process the Excel file
process_excel_file('laikinas.xls')