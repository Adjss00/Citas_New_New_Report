import os
import sys
from simple_salesforce import Salesforce
import pandas as pd

# Obtener la ruta del directorio actual del script
current_dir = os.path.dirname(os.path.abspath(__file__))

# Obtener la ruta al directorio "Public" desde el directorio actual
public_dir = os.path.join(current_dir, '..', 'public')

# Agregar la ruta al directorio "Public" al sys.path
sys.path.append(public_dir)

# Ahora puedes importar el módulo "originators_data" desde cualquier ubicación
from originators_data import data

class SalesforceDataExporter:
    def __init__(self, username, password, security_token, domain='login'):
        self.sf = Salesforce(username=username, password=password, security_token=security_token, domain=domain)

    def asignar_region_a_evento(self, df_evento):
        # Función para asignar la región según el FullName del propietario
        def asignar_region(row):
            for d in data:
                if row['OwnerName__c'] == d['FullName']:
                    return d['Region']
            return None

        # Aplica la función y crea la nueva columna 'Region' en el DataFrame
        df_evento['Region'] = df_evento.apply(asignar_region, axis=1)

    def extraer_y_exportar_objeto_sf(self, objeto_sf, nombre_archivo):
        fields_mapping = {
            'account': ['Id', 'Name', 'ParentId', 'ACC_tx_Account_Status__c'],
            'event': ['Id', 'ActivityDate', 'AccountId', 'OwnerId', 'OwnerName__c']
        }

        fields = fields_mapping.get(objeto_sf.lower())
        if not fields:
            print(f"Fields not defined for object: {objeto_sf}")
            return

        query = f"SELECT {', '.join(fields)} FROM {objeto_sf}"

        resultados = self.sf.query_all(query)['records']

        # Convert the results to a DataFrame
        df = pd.DataFrame(resultados)

        if 'attributes' in df.columns:
            df = df.drop(columns=['attributes'])

        # Export the DataFrame to an Excel file
        df.to_excel(nombre_archivo, index=False, engine='openpyxl')

        if objeto_sf.lower() == 'event':
            # Verifica si 'AccountId' está presente en los resultados de la consulta
            if 'AccountId' not in df.columns:
                print("AccountId not present in the query results for 'event'")
                return

            account_data = pd.DataFrame(self.sf.query_all("SELECT Id, Name, ParentId, ACC_tx_Account_Status__c FROM Account")['records'])
            account_data = account_data.rename(columns={'Id': 'AccountId', 'Name': 'Account Legal Name', 'ParentId': 'Top Parent Id', 'ACC_tx_Account_Status__c': 'Account Status'})
            df = pd.merge(df, account_data[['AccountId', 'Account Legal Name', 'Top Parent Id', 'Account Status']], on='AccountId', how='left')

            df = df[['Id', 'ActivityDate', 'AccountId', 'Account Legal Name', 'Top Parent Id', 'Account Status', 'OwnerId', 'OwnerName__c']]

            account_name_mapping = account_data.set_index('AccountId')['Account Legal Name'].to_dict()
            df['Top Parent'] = df['Top Parent Id'].map(account_name_mapping)

            # Lógica para la columna 'Top' en el objeto 'event'
            df['TOP'] = df.apply(lambda row: row['Top Parent'] if pd.notnull(row['Top Parent']) else row['Account Legal Name'], axis=1)

            # Lógica para la columna 'Region' en el objeto 'event'
            self.asignar_region_a_evento(df)

            # Imprime el DataFrame con las nuevas columnas
            print(df[['Id', 'ActivityDate', 'AccountId', 'Account Legal Name', 'Top Parent Id', 'Account Status', 'OwnerId', 'OwnerName__c', 'Region']])

            df.to_excel(nombre_archivo, index=False, engine='openpyxl')

    def exportar_datos_multiple(self, objetos_sf):
        for i, objeto_sf in enumerate(objetos_sf, start=1):
            nombre_archivo = f'data/{objeto_sf}_data.xlsx'
            self.extraer_y_exportar_objeto_sf(objeto_sf, nombre_archivo)
            print(objeto_sf)
            print("="*50) 