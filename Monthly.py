import requests
import pandas as pd
import xlrd
from xlutils.copy import copy
from xlwt import easyxf
from lxml import etree
import math
    
def round_half_up(n, decimals=0):
    multiplier = 10**decimals
    return math.floor(n * multiplier + 0.5) / multiplier


def vki_api_process(vki_api_url):
    response = requests.get(vki_api_url, verify=False)
    
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
            coicop_cpi = obs.xpath('./g:ObsKey/g:Value[@id="CoicopCPI"]/@value', namespaces=namespaces)

            # Check if the necessary elements exist and CoicopCPI is "00"
            if obs_value and laikotarpis and coicop_cpi and coicop_cpi[0] == "00":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None

def vki_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():

        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')  # Remove the 'M' character

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,4)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AB']).replace(',', '.')

            # Convert existing value to float and format it
            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 4)

            # Compare the ObsValue with the existing value in column 'AB'
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 27, f"{formatted_value:.4f}".replace(',', '.'), red_fill)  # Write new value and apply red fill

        else:
            # If no match is found, add new value to the excel
            new_row_index = len(df) + 3  # Adjust new row index based on header rows
            ws.write(new_row_index, 0, laikotarpis)  # Write LAIKOTARPIS to column A
            ws.write(new_row_index, 27, f"{formatted_value:.4f}".replace(',', '.'))  # Write ObsValue to column 27


def core_hicp_api_process(core_hicp_url):
    response = requests.get(core_hicp_url, verify=False)
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
            spec_pr_pM6040212 = obs.xpath('./g:ObsKey/g:Value[@id="spec_pr_pM6040212"]/@value', namespaces=namespaces)
            if obs_value and laikotarpis and spec_pr_pM6040212:

                if spec_pr_pM6040212[0] in ["v_be_en_m_al_t", "v_be_en"]:
                    filtered_data.append({
                        'ObsValue': obs_value[0],
                        'LAIKOTARPIS': laikotarpis[0],
                        'Indicator': spec_pr_pM6040212[0] 
                    })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None

def core_hicp_compare(df_filtered, ws, df, red_fill):
    for index, row in df_filtered.iterrows():
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')
        obs_value = row['ObsValue'].replace(',', '.')  
        indicator = row['Indicator']  

        # Validate the extracted values
        if not laikotarpis or not obs_value or not indicator:
            print(f"Invalid entry data: {laikotarpis}, {obs_value}, {indicator}")
            continue

        try:
            formatted_value = float(obs_value)
            formatted_value = round_half_up(formatted_value, 1)
        except ValueError:
            print(f"Skipping invalid obs_value: {obs_value}")
            continue

        df['A'] = df['A'].astype(str)
        match = df[df['A'] == laikotarpis] 

        if not match.empty:
            match_index = match.index[0]
            
            # For 'v_be_en_m_al_t' indicator
            if indicator == 'v_be_en_m_al_t':
                excel_value = str(df.at[match_index, 'AD']).replace(',', '.')

                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_up(excel_value, 1)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 29, f"{formatted_value:.1f}".replace(',', '.'), red_fill)
            elif indicator == 'v_be_en':
                excel_value = str(df.at[match_index, 'AE']).replace(',', '.')
                
                try:
                    excel_value = float(excel_value)
                    excel_value = round_half_up(excel_value, 1)
                except (ValueError, TypeError):
                    excel_value = None

                if excel_value is None or float(excel_value) != float(formatted_value):
                    ws.write(match_index + 3, 30, f"{formatted_value:.1f}".replace(',', '.'), red_fill)

        else:
            # No match found, so we add a new row in the Excel sheet
            new_row_index = len(df) + 3
            ws.write(new_row_index, 0, laikotarpis)

            if indicator == 'v_be_en_m_al_t':
                ws.write(new_row_index, 29, f"{formatted_value:.1f}".replace(',', '.'), red_fill)
            elif indicator == 'v_be_en':
                ws.write(new_row_index, 30, f"{formatted_value:.1f}".replace(',', '.'), red_fill)

def gki_c_m_api_process(gki_c_m_api_url):
    response = requests.get(gki_c_m_api_url, verify=False)
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
            evrk = obs.xpath('./g:ObsKey/g:Value[@id="EVRK2M2020313"]/@value', namespaces=namespaces)
            rinka= obs.xpath('./g:ObsKey/g:Value[@id="rinka1"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and evrk and evrk[0] == "C" and rinka and rinka[0]=="1":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })


        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def gki_c_m_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)

        df['A'] = df['A'].astype(str)
        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AF']).replace(',', '.')
            
            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 31, f"{formatted_value:.1f}".replace(',', '.'), red_fill)  
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 31, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  

def pram_c_api_process(pram_c_api_url):
    response = requests.get(pram_c_api_url, verify=False)
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
            evrk = obs.xpath('./g:ObsKey/g:Value[@id="EVRKM4050107"]/@value', namespaces=namespaces)
            palyg= obs.xpath('./g:ObsKey/g:Value[@id="LYGINIMAS"]/@value', namespaces=namespaces)
            islyg = obs.xpath('./g:ObsKey/g:Value[@id="Islyginimas_indeksai"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and evrk and evrk[0] == "C" and palyg and palyg[0]=="palyg_2021" and islyg and islyg[0]=="bendras":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    

def pram_c_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)

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
            ws.write(new_row_index, 38, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  

def iki_api_process(iki_api_url):
    response = requests.get(iki_api_url, verify=False)
    
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
            cpam = obs.xpath('./g:ObsKey/g:Value[@id="CPAM2020601"]/@value', namespaces=namespaces)
            rinka= obs.xpath('./g:ObsKey/g:Value[@id="rinka1"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and cpam and cpam[0] == "TOTAL" and rinka and rinka[0]=="1":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    

def iki_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,4)

        df['A'] = df['A'].astype(str)
        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AG']).replace(',', '.')
            
            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 4)

            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                ws.write(cell_row, 32, f"{formatted_value:.4f}".replace(',', '.'), red_fill)  
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 32, f"{formatted_value:.4f}".replace(',', '.'),red_fill)  


def eki_api_process(eki_api_url):
    response = requests.get(eki_api_url, verify=False)
    
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
            cpam = obs.xpath('./g:ObsKey/g:Value[@id="CPAM2020601"]/@value', namespaces=namespaces)

            if obs_value and laikotarpis and cpam and cpam[0] == "TOTAL" :
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })
        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def eki_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '') 

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,4)
        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AO']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 4)
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  

                ws.write(cell_row, 40, f"{formatted_value:.4f}".replace(',', '.'), red_fill) 
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 40, f"{formatted_value:.4f}".replace(',', '.'),red_fill)  


def eksportas_api_process(eki_api_url):
    response = requests.get(eki_api_url, verify=False)
    
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

            if obs_value and laikotarpis :
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    

def eksportas_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')  

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)
        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'AZ']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  

                ws.write(cell_row, 51, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 51, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  

def importas_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')  

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)
        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'BA']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  

                ws.write(cell_row, 52, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 52, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  


def eksportas_pasalinus_sezona_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')  

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)
        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'BB']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  

                ws.write(cell_row, 53, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 53, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  


def importas_pasalinus_sezona_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS'].replace('M', '')  

        formatted_value = float(obs_value)
        formatted_value = round_half_up(formatted_value,1)
        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'BC']).replace(',', '.')

            if excel_value != 'nan':
                excel_value= float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  

                ws.write(cell_row, 54, f"{formatted_value:.1f}".replace(',', '.'), red_fill) 
        else:
            new_row_index = len(df) + 3  
            ws.write(new_row_index, 0, laikotarpis)  
            ws.write(new_row_index, 54, f"{formatted_value:.1f}".replace(',', '.'),red_fill)  

def process_excel_file(excel_file):
    try:
        print(f"Opening Excel file: {excel_file}")
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Monthly')
        wb = copy(rb)
        ws = wb.get_sheet('Monthly')

        print("Reading data from Excel sheet...")
        df = pd.read_excel(
            excel_file, 
            sheet_name='Monthly', 
            usecols=[0, 27, 29, 30, 31, 32, 38, 40, 51, 52, 53, 54], 
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'AB', 'AD', 'AE', 'AF', 'AG', 'AM', 'AO', 'AZ', 'BA', 'BB', 'BC']
        red_fill = easyxf('pattern: pattern solid, fore_color red;')

        vki_api_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R260_M2020121/?startPeriod=1997-01'
        df_filtered = vki_api_process(vki_api_url)
        vki_compare(df_filtered, ws, df, red_fill)

        core_hicp_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R246_M2020214/?startPeriod=1997-01'
        df_core_hicp_filtered = core_hicp_api_process(core_hicp_url)
        core_hicp_compare(df_core_hicp_filtered, ws, df, red_fill)


        gki_c_m_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R130_M2020330/?startPeriod=1998-01'
        gki_c_m_filtered = gki_c_m_api_process(gki_c_m_url)
        gki_c_m_compare(gki_c_m_filtered, ws, df, red_fill)


        iki_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R289_M2020619_2/?startPeriod=2006-01'
        iki_filtered = iki_api_process(iki_url)
        iki_compare(iki_filtered,ws,df,red_fill)

        pram_c_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S8R918_M4050113_5/?startPeriod=1998-01'
        pram_c_filtered = pram_c_api_process(pram_c_url)
        pram_c_compare(pram_c_filtered,ws,df,red_fill)

        eki_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R288_M2020519_2/?startPeriod=2006-01'
        eki_filtered = eki_api_process(eki_url)
        eki_compare(eki_filtered,ws,df,red_fill)

        eksportas_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S6R005_M6050101/?startPeriod=2020-01'
        eksportas_filtered = eksportas_api_process(eksportas_url)
        eksportas_compare(eksportas_filtered, ws, df, red_fill)

        importas_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S6R007_M6050101/?startPeriod=2020-01'
        importas_filtered = eksportas_api_process(importas_url)
        importas_compare(importas_filtered, ws, df, red_fill)

        eksportas_pasalinus_sezona_url='https://osp-rs.stat.gov.lt/rest_xml/data/S6R006_M6050101/?startPeriod=2020-01'
        eksportas_pasalinus_sezona_filtered = eksportas_api_process(eksportas_pasalinus_sezona_url)
        eksportas_pasalinus_sezona_compare(eksportas_pasalinus_sezona_filtered, ws, df, red_fill)

        importas_pasalinus_sezona_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S6R008_M6050101/?startPeriod=2020-01'
        importas_pasalinus_sezona_filtered = eksportas_api_process(importas_pasalinus_sezona_url)
        importas_pasalinus_sezona_compare(importas_pasalinus_sezona_filtered, ws, df, red_fill)

            
        wb.save("naujas_combined.xls")
        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

process_excel_file('laikinas.xls')
