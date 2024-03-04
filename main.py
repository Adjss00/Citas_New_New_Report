import os
import time
from controllers.DataExporter import ExcelReader
from helpers.ObjectExtractor import SalesforceDataExporter

# Verificar si el directorio 'data' existe
if not os.path.exists('data'):
    # Si no existe, crearlo
    os.makedirs('data')

# Verificar si el directorio 'data' existe
if not os.path.exists('out'):
    # Si no existe, crearlo
    os.makedirs('out')

if __name__ == "__main__":
   
    username = 'jesus.sanchez@engen.com.mx'
    password = '21558269Antonio#'
    security_token = 'WqRoCDbMwhMPZ62iWUcXnmbmg'

    exporter = SalesforceDataExporter(username, password, security_token)

    objetos_sf = ['account', 'event']

    exporter.exportar_datos_multiple(objetos_sf)

    time.sleep(2)
    
    excel_info = ("data/event_data.xlsx", "Sheet1")
    reader = ExcelReader(excel_info)
    reader.read_sheet()
    reader.calculate_date_difference()
    reader.assign_weeks()
    reader.export_to_xlsx()
    reader.filter_and_export_swatt_data()
    # Llama a new_and_swatt con los DataFrames que has leído
    reader.new_and_swatt(input_path="out/new_new_meets.xlsx", output_path="out/swatt_data.xlsx")