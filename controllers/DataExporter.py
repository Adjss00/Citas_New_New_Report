import pandas as pd
import numpy as np
from controllers.date import semanas

class ExcelReader:
    def __init__(self, excel_info):
        self.file_path, self.sheet_name = excel_info
        self.df = None

    def read_sheet(self):
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
            print(f"Datos de la hoja '{self.sheet_name}' leídos con éxito.")
        except Exception as e:
            print(f"Error al leer la hoja '{self.sheet_name}': {e}")

    def display_data(self):
        if self.df is not None:
            # Filtrar los datos correspondientes
            dimasur_data = self.df[self.df['TOP'] == 'Impresiones Eloram']
            
            if not dimasur_data.empty:
                print("Datos para Impresiones Eloram:")
                print(dimasur_data)
            else:
                print("No hay datos para Impresiones Eloram.")
        else:
            print("No se ha leído ninguna hoja de Excel aún.")

    def calculate_date_difference(self):
        if self.df is not None:
            # Convierte la columna 'Date' a datetime
            self.df['ActivityDate'] = pd.to_datetime(self.df['ActivityDate'], format='%Y-%m-%d', dayfirst=True)

            # Ordena el DataFrame por 'TOP ID' y 'Date' en orden ascendente
            self.df = self.df.sort_values(by=['TOP', 'ActivityDate'])

            # Agrega la columna 'Last Date' que contiene la fecha de la fila anterior dentro de cada grupo 'TOP ID'
            self.df['Last Date'] = self.df.groupby('TOP')['ActivityDate'].shift(1)

            # Agrega la columna 'Owner Last Date' que contiene la fecha de la fila anterior dentro de cada grupo 'TOP ID'
            self.df['Owner Last Date'] = self.df.groupby('TOP')['OwnerName__c'].shift(1)

            # Agrega la columna 'Region Last Date' que contiene la fecha de la fila anterior dentro de cada grupo 'TOP ID'
            self.df['Region Last Date'] = self.df.groupby('TOP')['Region'].shift(1)

            # Calcula la diferencia entre fechas consecutivas dentro de cada grupo 'TOP ID'
            self.df['Date_Difference'] = (self.df['ActivityDate'] - self.df['Last Date']).dt.days

            # Llena las diferencias faltantes con 0 (para el primer valor de cada grupo)
            self.df['Date_Difference'] = self.df['Date_Difference'].fillna(0)

            # Obtener la fecha mínima dentro de cada grupo 'TOP ID'
            min_dates = self.df.groupby('TOP')['ActivityDate'].transform('min')

            # Agregar la columna 'Is_First' que indica si es la fecha más antigua del grupo
            self.df['Is_First'] = np.where(self.df['ActivityDate'] == min_dates, 1, 0)

            # Agrega la columna 'Is_Unic' que indica si 'TOP ID' es único
            self.df['Is_Unic'] = np.where(self.df['TOP'].duplicated(keep=False), 0, 1)

            # Agrega la columna 'Is_New_New' según las condiciones especificadas
            conditions = [
                (self.df['Date_Difference'] > 365),
                (self.df['Is_Unic'] == 1),
                (self.df['Is_First'] == 1),
                ((self.df['Account Status'] == 'Active') | 
                (self.df['Account Status'] == 'New Customer to EC') | 
                (self.df['Account Status'] == 'Dormant'))
            ]
            choices = ['Si', 'Si', 'Si', 'No']
            self.df['Is_New_New'] = np.select(conditions, choices, default='No')
            
            # Agregar columna 'is_customer' basada en 'Account Status'
            self.df['is_customer'] = np.where(
                (self.df['Account Status'] == 'Active') |
                (self.df['Account Status'] == 'New Customer to EC') |
                (self.df['Account Status'] == 'Dormant'),
                'Si',
                'No'
            )

            print("Diferencia de días, columnas 'Last Date', 'Is_Unic', 'Is_First', 'Is_New_New' y 'is_customer' calculadas con éxito.")
        else:
            print("No se ha leído ninguna hoja de Excel aún.")

    def export_to_xlsx(self, output_path='out/new_new_meets.xlsx'):
        if self.df is not None:
            if 'Is_New_New' not in self.df.columns:
                self.calculate_date_difference()
            
            # Añadir 'Semana' a las columnas
            column_order = ['AccountId', 'Account Legal Name', 'Top Parent Id', 'Top Parent', 'TOP', 'ActivityDate', 
                            'Last Date', 'OwnerId', 'OwnerName__c', 'Owner Last Date', 'Region', 'Region Last Date',
                            'Is_New_New', 'Date_Difference', 'Is_Unic', 'Is_First', 'Semana']  # Añade 'Semana' aquí

            # Obtener las columnas restantes en el orden original y agregarlas al final
            remaining_columns = [col for col in self.df.columns if col not in column_order]
            column_order.extend(remaining_columns)
            
            # Reorganizar el DataFrame con el nuevo orden de columnas
            self.df = self.df[column_order]
            
            # Exportar el DataFrame al archivo Excel
            self.df.to_excel(output_path, index=False)
            print(f"Datos exportados a '{output_path}' con éxito.")
        else:
            print("No hay datos para exportar.")

    def filter_and_export_swatt_data(self, output_path='out/swatt_data.xlsx'):
        if self.df is not None:
            swatt_data = self.df[(self.df['Region'] == 'SWATT LMM 1') | (self.df['Region'] == 'SWATT LMM 2')]
            if not swatt_data.empty:
                # Calcular las diferencias de fecha nuevamente antes de exportar
                self.df = swatt_data.copy()  # Utiliza una copia para evitar modificar el DataFrame original
                self.calculate_date_difference()
                # Exportar los datos actualizados al mismo archivo
                self.export_to_xlsx(output_path=output_path)
            else:
                print("No hay datos para exportar con la región SWATT.")
        else:
            print("No se ha leído ninguna hoja de Excel aún.")
    
    def assign_weeks(self):
        if self.df is not None:
            self.df['ActivityDate'] = pd.to_datetime(self.df['ActivityDate'], format='%Y-%m-%d', dayfirst=True)
            week_mapping = {}
            for week in semanas:
                inicio_semana = pd.to_datetime(week['Inicio'], format='%d/%m/%Y')
                fin_semana = pd.to_datetime(week['Fin'], format='%d/%m/%Y')
                week_mapping[(inicio_semana, fin_semana)] = week['semana']
            
            self.df['Semana'] = np.nan
            for key, value in week_mapping.items():
                mask = (self.df['ActivityDate'] >= key[0]) & (self.df['ActivityDate'] <= key[1])
                self.df.loc[mask, 'Semana'] = value
                print(f"Fechas en el rango {key} asignadas a la semana {value}")
            
            print("Columna 'Semana' asignada con éxito.")
        else:
            print("No se ha leído ninguna hoja de Excel aún.")

    def new_and_swatt(self, input_path="out/new_new_meets.xlsx", output_path="out/swatt_data.xlsx"):
        self.read_sheet()

        # Leer el primer excel
        new_new_meets_df = pd.read_excel(input_path)
        # print("first_df",new_new_meets_df)


        #Leer el segundo excel
        swatt_data_df = pd.read_excel(output_path)
        # print("second_df",swatt_data_df)

        # Obtener la lista de valores únicos de "Id" en swatt_data_df
        swatt_ids = swatt_data_df["Id"].tolist()
    
        # Sobreescribir la información en new_new_meets_df si el "Id" existe en swatt_data_df
        for id_value in swatt_ids:
            matching_rows = new_new_meets_df[new_new_meets_df["Id"] == id_value].index
            if not matching_rows.empty:
                new_rows = swatt_data_df[swatt_data_df["Id"] == id_value]
                for index, row in new_rows.iterrows():
                    # Validación de duplicados basada en las condiciones especificadas
                    print(new_new_meets_df.columns)
                    if ((new_new_meets_df.loc[matching_rows, 'ActivityDate'].values == row['Last Date']) & 
                        (new_new_meets_df.loc[matching_rows, 'OwnerName__c'].values == row['Owner Last Date'])).any():
                        # Si hay duplicados con las mismas fechas y nombres de propietario, eliminar la fila
                        new_new_meets_df.drop(index=matching_rows, inplace=True)
                    else:
                        # Si no hay duplicados, sobrescribir la información
                        new_new_meets_df.loc[matching_rows] = row.values
    
        

        # Exportar new_new_meets_df a un nuevo archivo Excel
        new_new_meets_df.to_excel("out/new_new_meets_updated.xlsx", index=False)

        print("Proceso completado. Archivo exportado con éxito.")



