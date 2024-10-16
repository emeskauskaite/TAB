from decimal import Decimal
import requests
import pandas as pd
import xlrd
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

def nac_s_a_api_process(nac_s_a_api_url):
    response = requests.get(nac_s_a_api_url, verify=False)
    
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
            pajamos = obs.xpath('./g:ObsKey/g:Value[@id="nacpajamosM2110109"]/@value', namespaces=namespaces)
            
            if obs_value and laikotarpis and islyg[0] == "bendras" and pajamos[0] == "b8n":
                filtered_data.append({
                    'ObsValue': obs_value[0],
                    'LAIKOTARPIS': laikotarpis[0]
                })

        df_filtered = pd.DataFrame(filtered_data)
        return df_filtered

    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None
    
def nac_s_a_compare(data_matrix, ws, df, red_fill):
    for index, row in data_matrix.iterrows():
        obs_value = row['ObsValue']
        laikotarpis = row['LAIKOTARPIS']

        # Make sure it's a float and round to 1 decimal place
        formatted_value = round_half_up(float(obs_value), 1)

        df['A'] = df['A'].astype(str)

        match = df[df['A'] == laikotarpis]

        if not match.empty:
            index_in_df = match.index[0]
            excel_value = str(df.at[index_in_df, 'P']).replace(',', '.')

            if excel_value != 'nan':
                # Convert the Excel value to float and round to 1 decimal place
                excel_value = float(excel_value)
                formatted_excel_value = round_half_up(excel_value, 1)

            # Ensure both API and Excel values are rounded to 1 decimal place
            if float(formatted_excel_value) != float(formatted_value):
                cell_row = index_in_df + 3  
                # Write both the API value and Excel value with 1 decimal place precision
                ws.write(cell_row, 15, f"{formatted_value:.1f}".replace(',', '.'), red_fill)
        else:
            new_row_index = len(df) + 3  
            # Write the new row with 1 decimal place precision
            ws.write(new_row_index, 0, laikotarpis) 
            ws.write(new_row_index, 15, f"{formatted_value:.1f}".replace(',', '.'),red_fill)

def process_excel_file(excel_file):
    try:
        print(f"Opening Excel file: {excel_file}")
        rb = xlrd.open_workbook(excel_file, formatting_info=True)
        rs = rb.sheet_by_name('Annual')
        wb = copy(rb)
        ws = wb.get_sheet('Annual')

        print("Reading data from Excel sheet...")
        df = pd.read_excel(
            excel_file, 
            sheet_name='Annual', 
            usecols=[0, 15],  
            skiprows=3, 
            header=None
        )
        df.columns = ['A', 'P']
        red_fill = easyxf('pattern: pattern solid, fore_color red;')
        no_fill = easyxf('')

        nac_s_a_url = 'https://osp-rs.stat.gov.lt/rest_xml/data/S7R217_M2110209/?startPeriod=1999'


        # Now, let's ensure that Excel values are also written with 0.1 precision
        for idx, row in df.iterrows():
            excel_value = row['P']

            # Check for NaN values
            if pd.notna(excel_value) and is_numeric(excel_value):
                excel_value = str(excel_value).replace(',', '.')
                
                # Convert to float, round to 1 decimal place, and write back to Excel
                rounded_excel_value = round_half_up(float(excel_value), 1)
                ws.write(idx + 3, 15, f"{rounded_excel_value:.1f}".replace(',', '.'))  # 0.1 precision
        
        nac_s_a_filtered = nac_s_a_api_process(nac_s_a_url)
        nac_s_a_compare(nac_s_a_filtered, ws, df, red_fill)

        wb.save("naujas_combined.xls")
        print(f"File updated and saved to naujas_combined.xls")

    except PermissionError as e:
        print(f"Permission Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

process_excel_file('laikinas.xls')
