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
        return f' {self.CruceID} \t  {self.CellID} \t {self.Empresa} \t {self.CellDireccion} \t {self.CellNum} \t {self.CellLocalidad} \t {self.CellProvincia}\t {self.CellRadio} \t {self.CellAzimuth} \t {self.CellLatitud}\t {self.CellLongitud}'

geolocator = Bing(api_key='AmNF9w5RcCH2SPYsP0RPiC-sclj4vW1sqEsfLy05rMKpPmqrQsjUzDZYZrB-Kntn')

directory_path = 'C:\9794'

filenames = os.listdir(directory_path)

xls_filenames = [f for f in filenames if f.endswith('.xls')]

# create an empty list to store the dataframes
myList = ['Cruce \t Codigo \t Empresa \t Direccion \t Latitud \t Longitud \t Rango \t Azimuth \t Cobertura ']

dfAux=[]
# iterate over the file paths and read each file as a dataframe
for path in xls_filenames:
    try:
        df = pd.read_excel(path, sheet_name='Datos Tecnicos')
        df.dropna(inplace=True)
        for i in range(2, len(df)):
            try:

                aux = str(df.iloc[i, 1]) + " " + str(df.iloc[i, 2]) + ", " + str(df.iloc[i, 3])
                location = geolocator.geocode(aux, timeout=None)

                p1 = PersonalAntenas('1', df.iloc[i, 0], 'Personal', aux, location.latitude, location.longitude, " ",
                                     df.iloc[i, 6], df.iloc[i, 5],"","")
                myList.append(p1)
            except:
                pass
    except:
        print("Something went wrong")

dfAux = pd.DataFrame(myList)
dfAux.to_excel('my_list.xls', index=False)