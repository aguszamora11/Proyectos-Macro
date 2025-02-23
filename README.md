# Proyectos-Macro

# El proyecto levanta datos de la Macro local y algunas variables internacionales, los procesa y los exporta a una base de datos para ser levatnada por informes.

# -*- coding: utf-8 -*-
"""
Created on Wed Aug 28 11:49:36 2024

@author: AZ
"""

#Librerias

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options



import os
import time
import pyautogui
import xlwings as xw
import pandas as pd
import requests
import json
import datetime as dt
import openpyxl
import xlsxwriter
import xlrd
from pandas import DataFrame
import numpy as np
from numpy import matrix
from scipy import linalg
import matplotlib.pyplot as plt
import warnings
import io
import yfinance as yf
from datetime import date

# Marca de tiempo al inicio del programa
start_time = time.time()

# Si no funciona algún día, hay que actualizar el TOKEN !!!!!!!!!!!!!
# Si tampoco funciona puede que haya una fecha duplicada en alguna base.



# ___________________________________________________________ O _________________________________________________


#Elimina las bases de datos viejas

# Ruta de la carpeta donde están los archivos a eliminar
ruta_carpeta = 'D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos'

# Lista de nombres de archivos que deseas eliminar
archivos_a_eliminar = ['series.xlsm', 'Lefi.xlsx', 'ITCRMSerie.xlsx', "DOLAR MEP - Cotizaciones historicas.csv",
                       "DOLAR CCL - Cotizaciones historicas.csv", "din2_ser.txt", "depuva_2024.xls"
                       ]

# Eliminar archivos
for archivo in archivos_a_eliminar:
    ruta_archivo = os.path.join(ruta_carpeta, archivo)
    if os.path.exists(ruta_archivo):
        os.remove(ruta_archivo)
        print(f"Archivo eliminado: {archivo}")
    else:
        print(f"Archivo no encontrado: {archivo}")


# ___________________________________________________________ O _________________________________________________


# Realiza los querys a BCRA y Rava



# BCRA Series

# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configura las opciones de Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "download.filename": "Series.xlsm"  # Nombre del archivo descargado
})

# Crea el objeto Service para ChromeDriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el navegador Chrome con el Service y opciones configuradas
driver = webdriver.Chrome(service=service, options=chrome_options)

# URL de la página web que deseas abrir
url = "https://www.bcra.gob.ar/PublicacionesEstadisticas/Informe_monetario_diario.asp"
driver.get(url)

# Espera que la página cargue completamente
driver.implicitly_wait(5)  # Espera hasta 10 segundos

# Busca el enlace del icono "Series.xlsm" por el texto del enlace y haz clic
try:
    icono_series = driver.find_element(By.LINK_TEXT, "Series.xlsm")
    icono_series.click()
    print("   ")
    print("Clic en el icono 'Series.xlsm' exitoso.")
    
    # Espera 10 segundos para permitir la descarga del archivo
    time.sleep(10)
except Exception as e:
    print("   ")
    print(f"Error al intentar hacer clic en el icono 'Series.xlsm': {e}")

# Espera a que el archivo se descargue completamente
time.sleep(2)  # Ajusta el tiempo si es necesario

# Ruta del archivo descargado
downloaded_file = os.path.join(r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos", "Series.xlsm")

# Verifica si el archivo existe
if os.path.exists(downloaded_file):
    print("   ")
    print(f"El archivo se ha descargado como: {downloaded_file}")
else:
    print("   ")
    print("El archivo no se encontró en la ruta especificada.")

# Detén el programa aquí
driver.quit()
print("   ")
print("Descarga exitosa de Series.xlsm")



# ___________________________________________________________ O _________________________________________________


# BCRA Lefis


# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configuración de opciones de Chrome
options = Options()
options.add_argument("--start-maximized")

# Configura la carpeta de descarga y el nombre del archivo
download_path = r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos"
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Configura el servicio para usar el chromedriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el WebDriver
driver = webdriver.Chrome(service=service, options=options)

# Navega al sitio web
driver.get("https://www.bcra.gob.ar/PublicacionesEstadisticas/Operaciones-y-subastas.asp")

# Espera 10 segundos para que cargue la página
time.sleep(10)

# Haz clic en el enlace de "Serie histórica"
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Serie histórica")
link.click()

# Espera 10 segundos para que el archivo se descargue
time.sleep(10)

# Renombra el archivo descargado
downloaded_file = os.path.join(download_path, "Data_operaciones.xlsx")
new_file_name = os.path.join(download_path, "Lefi.xlsx")

# Renombrar el archivo descargado si existe
if os.path.exists(downloaded_file):
    os.rename(downloaded_file, new_file_name)

# Cierra el navegador
driver.quit()

print("   ")
print("Descarga exitosa de Lefi")

# ___________________________________________________________ O _________________________________________________

# BCRA Depósitos UVA

# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configuración de opciones de Chrome
options = Options()
options.add_argument("--start-maximized")

# Configura la carpeta de descarga
download_path = r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos"
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Configura el servicio para usar el chromedriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el WebDriver
driver = webdriver.Chrome(service=service, options=options)

# Navega al enlace de descarga directamente
driver.get("https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/depuva_2024.xls")

# Espera 10 segundos para que el archivo se descargue
time.sleep(5)

# Cierra el navegador
driver.quit()

print("   ")
print("Descarga exitosa de Depósitos UVA")

# ___________________________________________________________ O _________________________________________________

# BCRA ITCRM

# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configuración de opciones de Chrome
options = Options()
options.add_argument("--start-maximized")

# Configura la carpeta de descarga y el nombre del archivo
download_path = r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos"
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Configura el servicio para usar el chromedriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el WebDriver
driver = webdriver.Chrome(service=service, options=options)

# Navega al sitio web
driver.get("https://www.bcra.gob.ar/PublicacionesEstadisticas/Indices_tipo_cambio_multilateral.asp")

# Encuentra el enlace de descarga y hace clic en él
link = driver.find_element(By.LINK_TEXT, "Descargar la serie histórica")
link.click()

# Espera 10 segundos para que el archivo se descargue
time.sleep(10)

# Cierra el navegador
driver.quit()

print("  ")
print("Descarga exitosa de Índices de Tipo de Cambio Multilateral")


# ___________________________________________________________ O _________________________________________________


# Dolar Mep


# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configura las opciones de Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Crea el objeto Service para ChromeDriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el navegador Chrome con el Service y opciones configuradas
driver = webdriver.Chrome(service=service, options=chrome_options)

# URL de la página web que deseas abrir
url = "https://www.rava.com/perfil/DOLAR%20MEP"
driver.get(url)

time.sleep(3)  # Espera 2 segundos para que cargue la página

# Encuentra el botón usando XPATH
download_button = driver.find_element(By.XPATH, "//button[contains(text(),'Bajar en formato Excel')]")

# Haz clic en el botón usando JavaScript
driver.execute_script("arguments[0].click();", download_button)

# Espera para permitir la descarga del archivo (ajusta el tiempo si es necesario)
time.sleep(10)

# Cierra el navegador
driver.quit()

print("   ")
print("Descarga exitosa de Dólar Mep")


# ___________________________________________________________ O _________________________________________________



# Dolar CCL


# Configura el path al archivo chromedriver.exe
chrome_driver_path = r"C:\Users\outlet\anaconda3\chromedriver.exe"

# Configura las opciones de Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"D:\Archivos\Agus\Macro_Argentina\Bases_De_Datos",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Crea el objeto Service para ChromeDriver
service = Service(executable_path=chrome_driver_path)

# Inicializa el navegador Chrome con el Service y opciones configuradas
driver = webdriver.Chrome(service=service, options=chrome_options)

# URL de la página web que deseas abrir
url = "https://www.rava.com/perfil/DOLAR%20CCL"
driver.get(url)

time.sleep(3)  # Espera 3 segundos para que cargue la página

# Encuentra el botón usando XPATH
download_button = driver.find_element(By.XPATH, "//button[contains(text(),'Bajar en formato Excel')]")

# Haz clic en el botón usando JavaScript
driver.execute_script("arguments[0].click();", download_button)

# Espera para permitir la descarga del archivo (ajusta el tiempo si es necesario)
time.sleep(10)

# Cierra el navegador
driver.quit()

print("   ")
print("Descarga exitosa de Dólar CCL")

# ___________________________________________________________ O _________________________________________________



# Procesa el archivo Series

# Abre el archivo "Macros"
libro_macros = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\Macros.xlsm")

# Abre el archivo "series"
libro_series = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm")

# Activa el archivo "series"
libro_series.activate()

# Ejecuta la macro desde "Macros"
libro_macros.macro("Series")()



# Bloque: Elimina las filas que consolidan los meses. Limpia la base para que no queden duplicadas las fechas.



#lista de hojas
hojas = ["BASE MONETARIA", "RESERVAS", "DEPOSITOS", "PRESTAMOS"]

# Itera sobre cada hoja
for nombre_hoja in hojas:
    hoja = libro_series.sheets[nombre_hoja]

    # Lee los datos de la hoja en un DataFrame de pandas
    df = hoja.used_range.options(pd.DataFrame, header=1, index=False).value

    # Asegúrate de que la columna 'Fecha' esté en formato datetime
    df['Fecha'] = pd.to_datetime(df['Fecha'])

    # Paso 1: Identificar la diferencia entre las fechas sucesivas
    df['Diferencia'] = df['Fecha'].diff()

    # Paso 2: Encontrar el primer salto que sea mayor a 15 días, indicando el inicio de las fechas mensuales
    umbral = pd.Timedelta(days=15)
    fin_fechas_diarias = df[df['Diferencia'] > umbral].index[0]

    # Paso 3: Filtrar el DataFrame para conservar solo las fechas diarias,
    # eliminando también la fila anterior al primer salto mayor a 15 días
    df_diario = df.iloc[:fin_fechas_diarias - 1]

    # Opción: Eliminar la columna 'Diferencia' si ya no es necesaria
    df_diario = df_diario.drop(columns=['Diferencia'])

    # Escribir los datos filtrados de vuelta en la hoja de Excel
    hoja.clear_contents()
    hoja.range("A1").options(index=False).value = df_diario

print("   ")
print("Base Series limpia")

# Guarda y cierra el archivo
libro_series.save()
libro_series.close()






# ___________________________________________________________ O _________________________________________________


# Procesa el archivo Lefi

# Abre el archivo "Lefis"
libro_Lefi = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\Lefi.xlsx")

# Activa el archivo "Lefis"
libro_Lefi.activate()

# Ejecuta la macro desde "Macros"
libro_macros.macro("Lefi")()

# Graba "Lefis"
libro_Lefi.save()

# Cierra "Lefis"
libro_Lefi.close()



print("   ")
print("Base Lefi limpia")

# ___________________________________________________________ O _________________________________________________


# Procesa el archivo Depósitos UVA

# Abre el archivo "depuva_2024"
libro_depUVA = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\depuva_2024.xls")

# Activa el archivo "Depósitos UVA"
libro_depUVA.activate()

# Ejecuta la macro desde "Macros"
libro_macros.macro("Depositos_UVA")()

# Graba "Depósitos UVA"
libro_depUVA.save()

# Cierra "Depósitos UVA"
libro_depUVA.close()



print("   ")
print("Base Depósitos UVA Limpia")

# ___________________________________________________________ O _________________________________________________


# Procesa el archivo ITCRM

# Abre el archivo "ITCRM"
libro_ITCRM = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\ITCRMSerie.xlsx")

# Activa el archivo "ITCRM"
libro_ITCRM.activate()

# Ejecuta la macro desde "Macros"
libro_macros.macro("ITCRM")()

# Graba "ITCRM"
libro_ITCRM.save()

# Cierra "ITCRM"
libro_ITCRM.close()

libro_macros.close()

print("   ")
print("Base ITCRM limpia")

# ___________________________________________________________ O _________________________________________________


# Levanta Reservas Brutas, Encajes por depositos en USD expresados en pesos, depósitos del tesoro en el BCRA.


# URL del archivo de datos
url = "https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/din2_ser.txt"

# Realizar la solicitud GET para descargar el contenido, desactivando la verificación SSL
response = requests.get(url, verify=False)

# Verificar si la solicitud fue exitosa (código de estado 200)
if response.status_code == 200:
    # Especificar la ruta donde se guardará el archivo
    file_path = "D:/Archivos/Agus/Macro_Argentina/Bases_De_Datos/din2_ser.txt"
    
    # Guardar el contenido descargado en un archivo
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(response.text)
    
    print("   ")
    print("Desgarga exitosa de Reservas Brutas, Encajes y depósitos del tesoro en BCRA")
    
    # Volcar los datos en un DataFrame, separando por ";"
    data = io.StringIO(response.text)  # Utiliza StringIO para leer el contenido desde la memoria
    db16b = pd.read_csv(data, sep=';', header=None)  # Separador ajustado a ";"
    
    # Asignar nombres a las columnas
    db16b.columns = ['Código', 'Fecha', 'Valor']
    
    # Establecer la columna 'Código' como índice
    db16b.set_index('Código', inplace=True)
    db16b.index = pd.to_numeric(db16b.index, errors='coerce')
    
    # Convertir la columna 'Fecha' a tipo datetime
    db16b['Fecha'] = pd.to_datetime(db16b['Fecha'], format='%d/%m/%Y')

    # Definir la fecha de inicio para el filtro
    fecha_inicio = '2016-01-01'  

    # Filtrar el DataFrame desde la fecha de inicio
    db16b = db16b[db16b['Fecha'] >= fecha_inicio]
    
    # Crear la nueva columna 'Valor/1000'
    db16b['Valor'] = round(db16b['Valor'] / 1000, 0)
    
    # Reservas Brutas
    # Filtrar el DataFrame para mostrar solo el código 246
    db16 = db16b.loc[[246]]
    db16.rename(columns={'Valor': 'Reservas_Brutas'}, inplace=True)
    
    # Encajes del BCRA en pesos: Dividir por FX
    # Filtrar el DataFrame para mostrar solo el código 246
    db17 = db16b.loc[[258]]
    db17.rename(columns={'Valor': 'Encajes_en_USD'}, inplace=True)
    
    # Depositos del Tesoro en el BCRA.
    # Filtrar el DataFrame para mostrar solo el código 246
    db18 = db16b.loc[[8842]]
    db18.rename(columns={'Valor': 'Depositos_Tesoro_en_BCRA'}, inplace=True)

else:
    print("   ")
    print(f"Error al descargar el archivo. Código de estado: {response.status_code}")


# ___________________________________________________________ O _________________________________________________



# Levanta Precio del Gas Natural, Brent, US 3m, US 5y, US 10y, US 30y, S&P, DXY


def obtener_datos_diarios(tickers, fecha_inicio, fecha_fin):
    # Crear un DataFrame vacío para almacenar los datos
    precios_diarios = pd.DataFrame()

    # Diccionario con los nombres de las columnas que se utilizarán
    nombres_columnas = {
        "NG=F": "Natural_Gas",
        "BZ=F": "Brent",
        "^GSPC": "SyP",
        "DX-Y.NYB": "Dollar_Index",
        "^IRX": "US_3m",
        "^FVX": "US_5y",
        "^TNX": "US_10y",
        "^TYX": "US_30y"
    }

    # Iterar sobre cada ticker para descargar los datos
    for ticker, nombre in nombres_columnas.items():
        # Descargar los datos diarios de Yahoo Finance
        datos = yf.download(ticker, start=fecha_inicio3, end=fecha_fin3, interval="1d")
        
        if datos.empty:
            print(f"No se encontraron datos para {nombre} ({ticker}).")
            continue

        # Cambiar el nombre de la columna "Date" a "Fecha"
        datos['Fecha'] = datos.index.strftime('%Y-%m-%d')
        
        # Seleccionar solo la columna de cierre y renombrarla
        datos = datos[['Fecha', 'Close']].rename(columns={'Close': nombre})
        
        # Combinar con el DataFrame principal
        if precios_diarios.empty:
            precios_diarios = datos
        else:
            precios_diarios = pd.merge(precios_diarios, datos, on='Fecha', how='outer')
    
    return precios_diarios

# Rango de fechas
fecha_inicio3 = "2016-01-01"
fecha_fin3 = date.today()

# Obtener todos los datos diarios dentro del rango de fechas
db9 = obtener_datos_diarios(
    ["NG=F", "BZ=F", "^GSPC", "DX-Y.NYB", "^IRX", "^FVX", "^TNX", "^TYX"], 
    fecha_inicio3, fecha_fin3
)

db9["Fecha"] = pd.to_datetime(db9["Fecha"], errors='coerce')
db9.drop_duplicates(subset="Fecha", inplace=True)
db9.fillna(0, inplace=True)


db9['US_3m'] =  db9['US_3m'] / 100
db9['US_5y'] =  db9['US_5y'] / 100
db9['US_10y'] = db9['US_10y'] / 100
db9['US_30y'] = db9['US_30y'] / 100

# ___________________________________________________________ O _________________________________________________




# Indice

# db1 = Compra Divisas
# db2 = Base Monetaria
# db3 = Tasa política monetaria
# db4 = M2
# db5 = Prestamos Sector Privado ARS y USD
# db6 = Tipo de Cambio A3500
# db7 = Dolar MEP
# db8 = Dolar CCL
# db9 =  Gas Natural, Brent, US 3m, US 5y, US 10y, US 30y, S&P, DXY
# db10 = (Libre - se puede usar para Incluir variable)
# db11 = BADLAR
# db12 = TM20
# db13 = Circulante
# db14 = Lefis
# db15 = ITCRM
# db16 = Reservas Brutas
# db17 = Encajes de depósitos en USD en el BCRA expresados en pesos
# db18 = Depósitos del tesoro en el BCRA.
# db19 = UVA - Unidad de Valor Adquisitivo
# db20 = Depósitos UVA y Depositos UVA expresados en UVAs


# Levanta Compra de Reservas
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
db1 = pd.read_excel(filepath, sheet_name="RESERVAS")
db1 = db1[["Fecha", "Variaciones Diarias - Factores de explicación de las Reservas Internacionales - Compra de Divisas"]]
db1 = db1.rename(columns={"Variaciones Diarias - Factores de explicación de las Reservas Internacionales - Compra de Divisas": "Compra_Divisas"})


# Levanta Base Monetaria 
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
db2 = pd.read_excel(filepath, sheet_name="BASE MONETARIA")
db2 = db2[["Fecha", "Saldos Diarios - Base Monetaria - Total (12) = (8+9+10+11)"]]
db2 = db2.rename(columns={"Saldos Diarios - Base Monetaria - Total (12) = (8+9+10+11)": "Base_Monetaria"})


# Levanta Tasa BCRA
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
columnas = ["Fecha", "Tasas de Interés - Tasa de política monetaria (8) - TNA"]
db3 = pd.read_excel(filepath, sheet_name="INSTRUMENTOS DEL BCRA", usecols=columnas)
db3 = db3.rename(columns={"Tasas de Interés - Tasa de política monetaria (8) - TNA": "TASA_BCRA"})


# Levanta M2
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
db4 = pd.read_excel(filepath, sheet_name="DEPOSITOS")
db4 = db4.rename(columns={"Pasivos Sector Privado en Pesos - Total Depósitos": "Depositos_Sector_Privado_ARS"})
db4 = db4.rename(columns={"Depósitos en Dólares (expresados en Dólares) - Sector Privado": "Depositos_Sector_Privado_USD"})
db4 = db4.rename(columns={"M2(6) (expresado en Pesos)": "M2"})


# Levanta Prestamos Sector Privado en ARS
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
db5 = pd.read_excel(filepath, sheet_name="PRESTAMOS")
db5 = db5.rename(columns={"Préstamos al Sector Privado - Total Pesos": "Prestamos_Sector_Privado_ARS"})
db5 = db5.rename(columns={"Préstamos al Sector Privado - Total Dólares": "Prestamos_Sector_Privado_USD"})


# Levanta Tipo de Cambio Oficial
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
db6 = pd.read_excel(filepath, sheet_name="RESERVAS")
db6 = db6[["Fecha", "Tipo de cambio (1 u$s = … $)"]]
db6 = db6.rename(columns={"Tipo de cambio (1 u$s = … $)": "Tipo_De_Cambio_A3500"})


# Levanta Dolar MEP
ruta_csv_dolar_mep = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\DOLAR MEP - Cotizaciones historicas.csv"
db7 = pd.read_csv(ruta_csv_dolar_mep)
db7 = db7[["fecha", "cierre"]]
db7 = db7.rename(columns={"fecha": "Fecha"})
db7['Fecha'] = pd.to_datetime(db7['Fecha'])
db7 = db7.rename(columns={"cierre": "DOLAR_MEP"})


# Levanta Dolar CCL
ruta_csv_dolar_ccl = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\DOLAR CCL - Cotizaciones historicas.csv"
db8 = pd.read_csv(ruta_csv_dolar_ccl)
db8 = db8[["fecha", "cierre"]]
db8 = db8.rename(columns={"fecha": "Fecha"})
db8['Fecha'] = pd.to_datetime(db8['Fecha'])
db8 = db8.rename(columns={"cierre": "DOLAR_CCL"})


# Levanta Tasa BADLAR
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
columnas = ["Fecha", "BADLAR (1) - Pesos - Total - TNA"]
db11 = pd.read_excel(filepath, sheet_name="TASAS DE MERCADO", usecols=columnas)
db11 = db11.rename(columns={"BADLAR (1) - Pesos - Total - TNA": "TASA_BADLAR"})


# Levanta Tasa TM20
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
columnas = ["Fecha", "TM20 (3) - Pesos - Total - TNA"]
db12 = pd.read_excel(filepath, sheet_name="TASAS DE MERCADO", usecols=columnas)
db12 = db12.rename(columns={"TM20 (3) - Pesos - Total - TNA": "TASA_TM20"})


# Levanta Circulante
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\series.xlsm"
columnas = ["Fecha", "Circulante"]
db13 = pd.read_excel(filepath, sheet_name="BASE MONETARIA", usecols=columnas)


# Levanta Lefis
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\Lefi.xlsx"
columnas = ["Fecha", "Lefi"]
db14 = pd.read_excel(filepath, sheet_name="Info_Operaciones", usecols=columnas)
db14 = db14.dropna(subset=['Fecha'], how='any')


# Levanta ITCRM
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\ITCRMSerie.xlsx"
columnas = ["Fecha", "ITCRM"]
db15 = pd.read_excel(filepath, sheet_name="ITCRM y bilaterales", usecols=columnas)
db15 = db15.dropna(subset=["ITCRM"], how='any')


# Levanta UVA
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\depuva_2024.xls"
columnas = ["Fecha", "UVA"]
db19 = pd.read_excel(filepath, sheet_name="dep_dia_altas_exp_UVAs", usecols=columnas)
db19['Fecha'] = db19['Fecha'].astype(str).str.replace(r'(\d{4})(\d{2})(\d{2})', r'\1-\2-\3', regex=True)
db19['Fecha'] = pd.to_datetime(db19['Fecha'], errors='coerce')
db19['UVA'] = pd.to_numeric(db19['UVA'], errors='coerce').astype(float)
db19 = db19.dropna(subset=["UVA"], how='any')


# Levanta Depositos UVA
filepath = "D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\depuva_2024.xls"
columnas = ["Fecha", "Depositos_UVA"]
db20 = pd.read_excel(filepath, sheet_name="dep_dia_sal", usecols=columnas)
db20['Fecha'] = db20['Fecha'].astype(str).str.replace(r'(\d{4})(\d{2})(\d{2})', r'\1-\2-\3', regex=True)
db20['Fecha'] = pd.to_datetime(db20['Fecha'], errors='coerce')
db20 = db20.dropna(subset=["Depositos_UVA"], how='any')
db20['Depositos_UVA'] = pd.to_numeric(db20['Depositos_UVA'], errors='coerce').astype(float)
db20['Depositos_UVA'] = db20['Depositos_UVA'] / 1000
db20['Depositos_UVA_en_UVA'] = db20['Depositos_UVA'] / db19['UVA'] * 1000000


# Consolido la Base de datos.
db1.set_index("Fecha", inplace = True) #Pongo Fecha como index
db2.set_index("Fecha", inplace = True) #Pongo Fecha como index
db3.set_index("Fecha", inplace = True) #Pongo Fecha como index
db4.set_index("Fecha", inplace = True) #Pongo Fecha como index
db5.set_index("Fecha", inplace = True) #Pongo Fecha como index
db6.set_index("Fecha", inplace = True) #Pongo Fecha como index
db7.set_index("Fecha", inplace = True) #Pongo Fecha como index
db8.set_index("Fecha", inplace = True) #Pongo Fecha como index
db9.set_index("Fecha", inplace = True) #Pongo Fecha como index
#db10.set_index("Fecha", inplace = True) #Pongo Fecha como index
db11.set_index("Fecha",inplace = True) #Pongo Fecha como index
db12.set_index("Fecha",inplace = True) #Pongo Fecha como index
db13.set_index("Fecha",inplace = True) #Pongo Fecha como index
db14.set_index("Fecha",inplace = True) #Pongo Fecha como index
db15.set_index("Fecha",inplace = True) #Pongo Fecha como index
db16.set_index("Fecha",inplace = True) #Pongo Fecha como index
db17.set_index("Fecha",inplace = True) #Pongo Fecha como index
db18.set_index("Fecha",inplace = True) #Pongo Fecha como index
db19.set_index("Fecha",inplace = True) #Pongo Fecha como index
db20.set_index("Fecha",inplace = True) #Pongo Fecha como index

Macro_Argentina = pd.concat([db1,db2,db3,db4,db5,db6,db7,db8,db9,
                             db11,db12,db13,db14,db15,db16,db17,db18,db19,
                             db20], axis=1)

Macro_Argentina = Macro_Argentina[["Compra_Divisas", 
                                   "Base_Monetaria", 
                                   "Circulante", 
                                   "M2", 
                                   "Depositos_Sector_Privado_ARS", 
                                   "Prestamos_Sector_Privado_ARS", 
                                   "Tipo_De_Cambio_A3500", 
                                   "DOLAR_MEP",
                                   "DOLAR_CCL", 
                                   "TASA_BCRA", 
                                   "TASA_BADLAR", 
                                   "TASA_TM20", 
                                   "Depositos_Sector_Privado_USD", 
                                   "Prestamos_Sector_Privado_USD", 
                                   "Lefi", 
                                   "ITCRM",
                                   "Reservas_Brutas", 
                                   "Encajes_en_USD", 
                                   "Depositos_Tesoro_en_BCRA", 
                                   'UVA',
                                   "Depositos_UVA", 
                                   'Depositos_UVA_en_UVA',
                                   "Natural_Gas", 
                                   "Brent", 
                                   "SyP", 
                                   "Dollar_Index", 
                                   "US_3m", 
                                   "US_5y", 
                                   "US_10y", 
                                   "US_30y"
                                   ]]
# 
Macro_Argentina["Encajes_en_USD"] = Macro_Argentina["Encajes_en_USD"] / Macro_Argentina["Tipo_De_Cambio_A3500"]

# Definir la fecha de inicio para el filtro
fecha_inicio2 = '2016-01-01'  

Macro_Argentina = Macro_Argentina.reset_index()
# Filtrar el DataFrame desde la fecha de inicio
Macro_Argentina = Macro_Argentina[Macro_Argentina['Fecha'] >= fecha_inicio2]

# Construye la base de datos con todos los datos mensuales.
mensual = pd.DataFrame()
mensual["Compra_Divisas"] =                 round(Macro_Argentina.Compra_Divisas.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).sum())
mensual["Base_Monetaria"] =                 round(Macro_Argentina.Base_Monetaria.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean())    
mensual["Circulante"] =                     round(Macro_Argentina.Circulante.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean())    
mensual["M2"] =                             round(Macro_Argentina.M2.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last())
mensual["Depositos_Sector_Privado_ARS"] =   round(Macro_Argentina.Depositos_Sector_Privado_ARS.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last())
mensual["Prestamos_Sector_Privado_ARS"] =   round(Macro_Argentina.Prestamos_Sector_Privado_ARS.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last())
mensual["Tipo_De_Cambio_A3500"] =           round(Macro_Argentina.Tipo_De_Cambio_A3500.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 4)
mensual["DOLAR_MEP"] =                      round(Macro_Argentina.DOLAR_MEP.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["DOLAR_CCL"] =                      round(Macro_Argentina.DOLAR_CCL.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["TASA_BCRA"] =                      round(Macro_Argentina.TASA_BCRA.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean() / 100, 4)
mensual["TASA_BADLAR"] =                    round(Macro_Argentina.TASA_BADLAR.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean() / 100, 4)
mensual["TASA_TM20"] =                      round(Macro_Argentina.TASA_TM20.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean() / 100, 4)
mensual["Depositos_Sector_Privado_USD"] =   round(Macro_Argentina.Depositos_Sector_Privado_USD.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last())
mensual["Prestamos_Sector_Privado_USD"] =   round(Macro_Argentina.Prestamos_Sector_Privado_USD.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last())
mensual["Lefi"] =                           round(Macro_Argentina.Lefi.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["ITCRM"] =                          round(Macro_Argentina.ITCRM.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).mean(), 2)
mensual["Reservas_Brutas"] =                round(Macro_Argentina.Reservas_Brutas.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["Encajes_en_USD"] =                 round(Macro_Argentina.Encajes_en_USD.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["Depositos_Tesoro_en_BCRA"] =       round(Macro_Argentina.Depositos_Tesoro_en_BCRA.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["Depositos_UVA"] =                  round(Macro_Argentina.Depositos_UVA.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["Depositos_UVA_en_UVA"] =           round(Macro_Argentina.Depositos_UVA_en_UVA.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["UVA"] =                            round(Macro_Argentina.UVA.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["Natural_Gas"] =                    round(Macro_Argentina.Natural_Gas.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 3)
mensual["Brent"] =                          round(Macro_Argentina.Brent.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["SyP"] =                            round(Macro_Argentina.SyP.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 0)
mensual["Dollar_Index"] =                   round(Macro_Argentina.Dollar_Index.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 2)
mensual["US_3m"] =                          round(Macro_Argentina.US_3m.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 5)
mensual["US_5y"] =                          round(Macro_Argentina.US_5y.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 5)
mensual["US_10y"] =                         round(Macro_Argentina.US_10y.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 5)
mensual["US_30y"] =                         round(Macro_Argentina.US_30y.groupby([Macro_Argentina.Fecha.dt.year, Macro_Argentina.Fecha.dt.month]).last(), 5)



mensual.to_excel("D:\\Archivos\\Agus\\Macro_Argentina\\Macro Argentina.xlsx", index=True, header=True)

# Procesa el archivo Macro Argentina

# Abre el archivo "Macros"
libro_macros = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Bases_De_Datos\\Macros.xlsm")

# Abre el archivo "Macro Argentina"
libro_MacroArg = xw.Book("D:\\Archivos\\Agus\\Macro_Argentina\\Macro Argentina.xlsx")

# Activa el archivo "Macro Argentina"
libro_MacroArg.activate()

# Ejecuta la macro desde "Macros"
libro_macros.macro("Macro_Argentina")()

# Graba "Macro Argentina"
libro_MacroArg.save()

# Cierra "Macro Argentina"
libro_MacroArg.close()

libro_macros.close()

print("      ")
print("Base de Datos Actualizada")



# Calcular el tiempo de ejecución
end_time = time.time()
execution_time = end_time - start_time
minutes = int(execution_time // 60)
seconds = int(execution_time % 60)
print("   ")
print(f"El programa tardó {minutes} minutos y {seconds} segundos en ejecutarse.")
print("   ")


