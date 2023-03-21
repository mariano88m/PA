import pandas as pd
from geopy.geocoders import Bing
import os

class PersonalAntenas:
    def __init__(self, CruceID, CellID, Empresa,  CellDireccion, CellNum, CellLocalidad, CellProvincia, CellRadio, CellAzimuth, CellLatitud, CellLongitud):
        self.CruceID = CruceID
        self.CellID = CellID
        self.Empresa = Empresa
        self.CellDireccion = CellDireccion
        self.CellNum = CellNum
        self.CellLocalidad = CellLocalidad
        self.CellProvincia = CellProvincia
        self.CellRadio = CellRadio
        self.CellAzimuth = CellAzimuth
        self.CellLatitud = CellLatitud
        self.CellLongitud = CellLongitud

    def __str__(self):
        return f' {self.CellID} \t {self.CellDireccion} \t {self.CellNum}\t {self.CellLocalidad}\t {self.CellProvincia}\t {self.CellRadio} \t {self.CellAzimuth} \t {self.CellLatitud}\t {self.CellLongitud}'

geolocator = Bing(api_key='AmNF9w5RcCH2SPYsP0RPiC-sclj4vW1sqEsfLy05rMKpPmqrQsjUzDZYZrB-Kntn')

directory_path = 'C:\9794'

filenames = os.listdir(directory_path)

xls_filenames = [f for f in filenames if f.endswith('.xls')]

# create an empty list to store the dataframes
myList = ['Cell ID \t Celda Direccion \t Celda Num \t Celda Localidad \t Celda Provincia \t Radio Cobertura (en KM) \t Azimuth \t Latitud \t Longitud']

dfAux=[]
# iterate over the file paths and read each file as a dataframe
for path in xls_filenames:
    try:
        df = pd.read_excel(path, sheet_name='Datos Tecnicos')
        print("Read OK")

        df.dropna(inplace=True)
        #for i in range(2, len(df)):
        for i in range(2, len(df)):

            aux = str(df.iloc[i, 2]) + " " + str(df.iloc[i, 1]) + ", " + str(df.iloc[i, 3])
            location = geolocator.geocode(aux, timeout=None)
            p1 = PersonalAntenas(df.iloc[i, 0], df.iloc[i, 1], df.iloc[i, 2], df.iloc[i, 3], df.iloc[i, 4],
                                 df.iloc[i, 5], df.iloc[i, 6], location.latitude, location.longitude)
            myList.append(p1)

        print("OK")

    except:
        print("Something went wrong")

dfAux = pd.DataFrame(myList)
dfAux.to_csv('my_list.csv', index=False)