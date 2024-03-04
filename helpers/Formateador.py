import pandas as pd

class ExcelReader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
    
    def read_excel(self, column_mapping=None):
        try:
            if column_mapping is None:
                self.df = pd.read_excel(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path, usecols=list(column_mapping.keys()))
                self.df.rename(columns=column_mapping, inplace=True)
            return self.df
        except FileNotFoundError:
            print("El archivo especificado no se encontró.")
        except Exception as e:
            print("Se produjo un error al leer el archivo:", e)
    
    def get_dataframe(self):
        return self.df

excel_reader = ExcelReader("data/historic_events.xlsx")

# Definimos un diccionario para mapear las columnas que queremos extraer con los nombres deseados
column_mapping = {
    'ACCOUNTID': 'AccountId',
    'ACTIVITYDATE': 'ActivityDate',
    'OWNERID': 'OwnerId',
    'OWNERNAME__C': 'OwnerName__c',
    'SUBJECT': 'Subject'
}

excel_reader.read_excel(column_mapping=column_mapping)

# Obtención del DataFrame con los datos extraídos
dataframe = excel_reader.get_dataframe()

print(dataframe)