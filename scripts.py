import time
import win32com.client
import sys
import subprocess
import time
import os
import pandas as pd
from datetime import datetime, timedelta
from google.auth import default
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from collections import OrderedDict
from datetime import datetime, timedelta
import xlwings as xw
import gspread
import numpy as np
import pytz
import pyodbc

#Datos ingreso SAP
sap_client = "400"
sap_user = "user"
sap_password = "password"
sap_language = "ES"

#Paths
path_yesterday_today_data = r'C:\Users\jacostae\Desktop\Daily_update\Data\Data_ayer_hoy.txt'
path_bitacora = r'Data\bitacora.txt'


def start_timer():
    """Inicia un temporizador y devuelve el tiempo de inicio."""
    return time.time()

def end_timer(start_time):
    """Calcula el tiempo transcurrido desde el inicio y devuelve el tiempo en segundos."""
    return time.time() - start_time

def current_data_Siclo(yesterday_date_str, current_date_str):

    # Data connection
    server = '10.111.111.43'
    database = 'databasename'
    username = 'username'
    password = 'password'

    sQuery = """
    SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
    SELECT *
    FROM (SELECT Remesa,
                 Negocio,
                 CONVERT(VARCHAR(16), [Fecha Despacho], 120) [Fecha Despacho],
                 CONVERT(VARCHAR(10), [Fecha Manifiesto], 120) [Fecha Manifiesto],
                 Destinatario,
                 Cliente,
                 Transportador,
                 Administrador,
                 Placa,
                 Remolque,
                 [Linea_codProducto],
                 [Linea_nombreProducto],
                 [Linea_UOM],
                 CONVERT(VARCHAR(20), [Linea_Cantidad]) [Linea_Cantidad],
                 CONVERT(VARCHAR(20), [Linea_Peso]) [Linea_Peso],
                 [Transporte Siclo],
                 [Transporte SAP],
                 [Cód. DANE],
                 Ciudad,
                 Departamento,
                 CONVERT(VARCHAR(16), [Báscula salida], 120) [Báscula salida],
                 CONVERT(VARCHAR(4), [Año]) [Año],
                 CONVERT(VARCHAR(2), Mes) Mes,
                 CONVERT(VARCHAR(2), [Día]) [Día],
                 Estado
          FROM view_DailyReport3
          WHERE (OperationalStatus >= 33
                 AND ReferenceDate BETWEEN 'par_DateIni'  AND 'par_DateFin' 
                 AND (CustomerName LIKE '%' OR '' = '')
                 AND shi2Id IS NULL)
                OR (OperationalStatus = -99
                    AND ReferenceDate BETWEEN 'par_DateIni'  AND 'par_DateFin' 
                    AND (CustomerName LIKE '%' OR '' = '')
                    AND shi3Id IS NOT NULL)
          UNION ALL
          SELECT '' Remesa,
                 '' Negocio,
                 '' [Fecha Despacho],
                 '' [Fecha Manifiesto],
                 '' Destinatario,
                 '' Cliente,
                 '' Transportador,
                 '' Administrador,
                 '' Placa,
                 '' Remolque,
                 '' [Linea_codProducto],
                 '' [Linea_nombreProducto],
                 '' [Linea_UOM],
                 '' [Linea_Cantidad],
                 '' [Linea_Peso],
                 '' [Transporte Siclo],
                 '' [Transporte SAP],
                 '' [Cód. DANE],
                 '' Ciudad,
                 '' Departamento,
                 '' [Báscula salida],
                 '' [Año],
                 '' Mes,
                 '' [Día],
                 '' Estado
          FROM dvNumbers
          WHERE Val < 2
          UNION ALL
          SELECT 'Reasignaciones realizadas anulando entrega original ' Remesa,
                 '' Negocio,
                 '' [Fecha Despacho],
                 '' [Fecha Manifiesto],
                 '' Destinatario,
                 '' Cliente,
                 '' Transportador,
                 '' Administrador,
                 '' Placa,
                 '' Remolque,
                 '' [Linea_codProducto],
                 '' [Linea_nombreProducto],
                 '' [Linea_UOM],
                 '' [Linea_Cantidad],
                 '' [Linea_Peso],
                 '' [Transporte Siclo],
                 '' [Transporte SAP],
                 '' [Cód. DANE],
                 '' Ciudad,
                 '' Departamento,
                 '' [Báscula salida],
                 '' [Año],
                 '' Mes,
                 '' [Día],
                 '' Estado
          FROM dvNumbers
          WHERE Val < 2
          UNION ALL
          SELECT 'Remesa' Remesa,
                 'Negocio' Negocio,
                 'Fecha Despacho' [Fecha Despacho],
                 'Fecha Manifiesto' [Fecha Manifiesto],
                 'Destinatario' Destinatario,
                 'Cliente' Cliente,
                 'Transportador' Transportador,
                 'Administrador' Administrador,
                 'Placa' Placa,
                 'Remolque' Remolque,
                 'Linea_codProducto' [Linea_codProducto],
                 'Linea_nombreProducto' [Linea_nombreProducto],
                 'Linea_UOM' [Linea_UOM],
                 'Linea_Cantidad' [Linea_Cantidad],
                 'Linea_Peso' [Linea_Peso],
                 'Transporte Siclo' [Transporte Siclo],
                 'Transporte SAP' [Transporte SAP],
                 'Cód. DANE' [Cód. DANE],
                 'Ciudad' Ciudad,
                 'Departamento' Departamento,
                 'Báscula salida' [Báscula salida],
                 'Año' [Año],
                 'Mes' Mes,
                 'Día' [Día],
                 'Estado' Estado
          FROM dvNumbers
          WHERE Val < 2
          UNION ALL
          SELECT Remesa,
                 Negocio,
                 CONVERT(VARCHAR(16), [Fecha Despacho], 120) [Fecha Despacho],
                 CONVERT(VARCHAR(10), [Fecha Manifiesto], 120) [Fecha Manifiesto],
                 Destinatario,
                 Cliente,
                 Transportador,
                 Administrador,
                 Placa,
                 Remolque,
                 [Linea_codProducto],
                 [Linea_nombreProducto],
                 [Linea_UOM],
                 CONVERT(VARCHAR(20), [Linea_Cantidad]) [Linea_Cantidad],
                 CONVERT(VARCHAR(20), [Linea_Peso]) [Linea_Peso],
                 [Transporte Siclo],
                 [Transporte SAP],
                 [Cód. DANE],
                 Ciudad,
                 Departamento,
                 CONVERT(VARCHAR(16), [Báscula salida], 120) [Báscula salida],
                 CONVERT(VARCHAR(4), [Año]) [Año],
                 CONVERT(VARCHAR(2), Mes) Mes,
                 CONVERT(VARCHAR(2), [Día]) [Día],
                 Estado
          FROM view_DailyReassignReport3
          WHERE OperationalStatus >= 33
                AND ReferenceDate BETWEEN 'par_DateIni' AND 'par_DateFin'
                AND (CustomerName LIKE '%' OR '' = '')
                AND shi2Id IS NOT NULL) A
    """

    # Execute function fill_parameters
    sQuery = fill_parameters(sQuery, yesterday_date_str, current_date_str)

    try:
        # Establecer la conexión
        with pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password) as connection:
            print("Login Siclo successful")

            # Crear un cursor
            with connection.cursor() as cursor:

                # Ejecutar la nueva consulta
                cursor.execute(sQuery)

                # Obtener los resultados
                rows = cursor.fetchall()

    except Exception as ex:
        print("Login Siclo failed:", str(ex))

    columnas=['Remesa', 'Negocio', 'Fecha Despacho', 'Fecha Manifiesto', 'Destinatario', 'Cliente', 'Transportador', 'Administrador', 'Placa', 'Remolque', 'Linea_codProducto', 'Linea_nombreProducto', 'Linea_UOM', 'Linea_Cantidad', 'Linea_Peso', 'Transporte Siclo', 'Transporte SAP', 'Cód. DANE', 'Ciudad', 'Departamento', 'Báscula salida', 'Año', 'Mes', 'Día', 'Estado']

    # Convertir cada fila en un diccionario
    datos = []
    for row in rows:
        fila = OrderedDict()
        for i, columna in enumerate(columnas):
            fila[columna] = row[i]
        datos.append(fila)

    # Crear el DataFrame a partir de la lista de diccionarios
    df_s = pd.DataFrame(datos)

    # Eliminar últimas 3 filas
    df_s= df_s.drop(df_s.index[-3:])

    #Filtrar columna 'Estado' para quitar reasignados
    df_filtrado = df_s[df_s['Estado'].eq('')]

    # Reemplazar valores None y NaN por cadenas vacías en todo el DataFrame
    df_filtrado = df_filtrado.replace({None: '', np.nan: ''})

    df_Siclo_dia = df_filtrado.sort_values(by='Fecha Despacho')

    # Convertir 'Remesa', 'Linea_Peso' y 'Linea_Cantidad' a entero
    df_Siclo_dia['Remesa'] = pd.to_numeric(df_Siclo_dia['Remesa'], errors='coerce').fillna(0).astype(int)
    df_Siclo_dia['Linea_Peso'] = pd.to_numeric(df_Siclo_dia['Linea_Peso'], errors='coerce').fillna(0).astype(int)
    df_Siclo_dia['Linea_Cantidad'] = pd.to_numeric(df_Siclo_dia['Linea_Cantidad'], errors='coerce').fillna(0).astype(int)
    return df_Siclo_dia


def login_and_download_data_SAP(path_yesterday_today_data, yesterday_date_SAP, today_date_SAP):
    """
    Login and download data of SAP
    
    """
    
    try:
        session, connection, application = open_SAPGUI()
        Msg, session, connection, application = enter_credentials(sap_client, sap_user, sap_password, sap_language, session, connection, application)
        pass
        download_data_SAP(yesterday_date_SAP, today_date_SAP, path_yesterday_today_data, session, application)
        
    except:
        disconnect_SAP(connection, application)
    
    finally:
        current_data_SAP(path_yesterday_today_data)
        
    return current_data_SAP(path_yesterday_today_data)

def open_SAPGUI():
    """
    Open to SAPGUI - (R/3 - Productivo)
    
    """
    
    subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
    time.sleep(1)  # Adjust sleep time as needed

    # Get SAP GUI scripting object
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not isinstance(SapGuiAuto, win32com.client.CDispatch):
        raise Exception("Unable to get SAP GUI scripting object")

    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection("R/3 - Productivo", True)
    session = connection.Children(0)
    return session, connection, application

def enter_credentials(sap_client, sap_user, sap_password, sap_language, session, connection, application):
    """
    Enter credentials in SAP
    
    """
    try:
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = sap_client
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = sap_user
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = sap_password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = sap_language
        session.findById("wnd[0]").sendVKey(0)
        Msg = session.findById("wnd[0]/sbar").Text

        if Msg == "Nombre o clave de acceso incorrectos (repita la entrada al sistema)":
            disconnect_SAP(connection, application)
            
    except:
        disconnect_SAP(connection, application)
    
    return Msg, session, connection, application

def disconnect_SAP(connection, application):
    """
    Disconnects from the SAP session and closes SAP GUI scripting objects.

    """

    if connection:
        connection.CloseSession('ses[0]')
        application.Quit()
        os.system("TASKKILL /F /IM saplogon.exe")
        
def download_data_SAP(yesterday_date_SAP,today_date_SAP, path_yesterday_today_data, session, application):
    """
    Download data of the transaction Y_CSD_80000073 in a file with format .txt
    
    Args:
    yesterday_date_SAP: Yesterday date.
    today_date_SAP: Today date.
    path_yesterday_today_data: 
    """

    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    print("Login SAP successful")
    #Acción para ingresar a la transacción y variante
    session.findById("wnd[0]/tbar[0]/okcd").Text = "Y_CSD_80000073"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (17)
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "JACOSTAE"
    session.findById("wnd[1]").sendVKey (8)
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSP$00015-LOW").Text = yesterday_date_SAP
    session.findById("wnd[0]/usr/ctxtSP$00015-HIGH").Text = today_date_SAP
    session.findById("wnd[0]").sendVKey (8)

    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[1]/usr/lbl[0,2]").SetFocus()
    session.findById("wnd[1]").sendVKey (2)
    session.findById("wnd[0]/tbar[1]/btn[9]").press()
    session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").Text = path_yesterday_today_data
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILETYPE").Text = "DAT"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    disconnect_SAP(connection, application)
    
def get_current_date_Colombia():
    """
    Function to get the current date in Colombia time zone
    """
    
    colombia_tz = pytz.timezone('America/Bogota')
    current_date = datetime.now(colombia_tz)
    
    return current_date


def calculate_dates_yesterday_today(current_date):
    """
    Function to calculate date values to download information from SAP and Siclo
    """

    current_date_ts = pd.Timestamp(current_date)
    start_month = current_date_ts.replace(day=1)
    end_month = current_date_ts + pd.offsets.MonthEnd(0)
    start_month_str = start_month.strftime('%Y%m%d')
    end_month_str = end_month.strftime('%Y%m%d')

    yesterday_date = current_date - timedelta(days=1)
    current_date_str = current_date.strftime('%Y%m%d')
    yesterday_date_str = yesterday_date.strftime('%Y%m%d')
    today_date_SAP = current_date.strftime("%d.%m.%Y")

    if current_date.strftime('%H:%M:%S') > '07:00:00':
        yesterday_date_str = current_date.strftime('%Y%m%d')
        yesterday_date_SAP = today_date_SAP
    else:
        yesterday_date_str = yesterday_date.strftime('%Y%m%d')
        yesterday_date_SAP = yesterday_date.strftime("%d.%m.%Y")
    
    return yesterday_date_str, current_date_str, today_date_SAP, yesterday_date_SAP, start_month_str, end_month_str 

def assign_time(row):
    actual_date = datetime.strptime(row['Fe.SMreal'], '%d.%m.%Y').date()
    current_date = get_current_date_Colombia()
    if actual_date == current_date.date():
        return get_current_date_Colombia().strftime('%H:%M:%S')
    elif actual_date < current_date.date():
        return '23:59:15'
    else:
        return '15:15:15'
    
def current_data_SAP(path_yesterday_today_data):
    """
    Read Latin-1 encoded file and add time
    
    """
    
    df_F_dia = pd.read_csv(path_yesterday_today_data, sep='\t', encoding='latin-1', index_col=None)
    df_F_dia = df_F_dia.rename(columns=lambda x: x.replace(" ", ""))
    df_F_dia['Hora'] = df_F_dia.apply(assign_time, axis=1)
    df_F_dia = df_F_dia.fillna('NA').astype(str)
    return df_F_dia

def current_data_Facturacion(path_yesterday_today_data, yesterday_date_SAP, today_date_SAP):
    try:
    
        df_F_dia = login_and_download_data_SAP(path_yesterday_today_data, yesterday_date_SAP, today_date_SAP)
    
    except:
        df_F_dia = current_data_SAP(path_yesterday_today_data)
        
    return df_F_dia


def join_data(df_Siclo_dia,BD_Siclo,df_F_dia,BD_SAP):
    """
    Join data from the current dataframes and the historical dataframes of the current month
    
    """
    # Concatenar los DataFrames información Despachos
    df_Siclo_U = pd.concat([BD_Siclo, df_Siclo_dia], ignore_index=True)

    # Convertir 'Entrega' a entero
    df_Siclo_U['Remesa'] = pd.to_numeric(df_Siclo_U['Remesa'], errors='coerce').fillna(0).astype(int)
    df_Siclo_U['Linea_Peso'] = pd.to_numeric(df_Siclo_U['Linea_Peso'], errors='coerce').fillna(0).astype(int)

    # Concatenar los DataFrames información Facturación
    df_Facturacion_U = pd.concat([BD_SAP, df_F_dia], ignore_index=True)

    # Convertir 'Entrega' a entero
    df_Facturacion_U['Entrega'] = pd.to_numeric(df_Facturacion_U['Entrega'], errors='coerce').fillna(0).astype(int)

    # Eliminar duplicados datos Facturación
    df_Facturacion_final = df_Facturacion_U.drop_duplicates(subset=["Entrega",'Textobrevedematerial', "NTGEW_OK"], keep='last')

    # Restablecer el índice
    df_Facturacion_final = df_Facturacion_final.reset_index(drop=True)

    # Convertir la columna a tipo numérico
    df_Facturacion_final['NTGEW_OK'] = pd.to_numeric(df_Facturacion_final['NTGEW_OK'], errors='coerce')

    # Rellenar NaN con 'NA' y convertir todas las columnas a tipo str
    df_Facturacion_final = df_Facturacion_final.fillna('NA').astype(str)

    df_Facturacion_final = df_Facturacion_U.drop_duplicates(subset=["Entrega", "Textobrevedematerial", "Pos."], keep='first')
    
    df_Facturacion_final = sort_dataframe_by_date_and_time(df_Facturacion_final)

    # Eliminar duplicados datos Despachos
    df_Siclo_final = df_Siclo_U.drop_duplicates(subset=["Remesa", "Linea_nombreProducto","Linea_Peso"], keep='first')

    # Restablecer el índice
    df_Siclo_final = df_Siclo_final.reset_index(drop=True)

    # Reemplazar valores None y NaN por cadenas vacías en todo el DataFrame
    df_Siclo_final = df_Siclo_final.replace({None: '', np.nan: ''})

    # Rellenar NaN con 'NA' y convertir todas las columnas a tipo str
    df_Siclo_final['Linea_Cantidad'] = df_Siclo_final['Linea_Cantidad'].astype(str)

    # Crear una nueva columna en los dataframes 
    df_Siclo_final.loc[:, 'Concatenada'] = df_Siclo_final['Remesa'].astype(str) + df_Siclo_final['Linea_codProducto'].astype(str)+df_Siclo_final['Linea_Cantidad'].str.split('.').str[0]

    # Crear una nueva columna en dataframeF
    df_Facturacion_final.loc[:, 'Concatenada'] = df_Facturacion_final['Entrega'].astype(str) + df_Facturacion_final['Material'].astype(str)+df_Facturacion_final['Cantidadentrega'].str.split(',').str[0]
    
    print ("Dataframes unidos exitosamente")
    
    return df_Siclo_final, df_Facturacion_final


def match_Siclo_SAP(df_Siclo_final, df_Facturacion_final):
    """
    Match Siclo and SAP data
    """
    # Combinar ambos dataframes utilizando la columna concatenada como clave de combinación
    df_cierre = pd.merge(df_Siclo_final, df_Facturacion_final[['Concatenada', 'Entrega', 'Fe.SMreal', 'NTGEW_OK', 'Hora']], on='Concatenada', how='left')

    registros_no_cierre = df_Facturacion_final[~df_Facturacion_final['Concatenada'].isin(df_cierre['Concatenada'])]

    # Eliminar la columna 'Concatenada' en los dataframes
    df_cierre = df_cierre.drop(columns=['Concatenada'])
    df_Siclo_final=df_Siclo_final.drop(columns=['Concatenada'])
    df_Facturacion_final=df_Facturacion_final.drop(columns=['Concatenada'])
    registros_no_cierre=registros_no_cierre.drop(columns=['Concatenada'])

    df_cierre.loc[:, 'Concatenada_']=df_cierre['Entrega'].astype(str) + df_cierre['Linea_codProducto'].astype(str)+df_cierre['Linea_Peso'].astype(str)

    df_cierre = df_cierre.drop_duplicates(subset=['Concatenada_'])

    df_cierre = df_cierre.drop(columns=['Concatenada_'])

    # Rellenar NaN con 'NA' y convertir todas las columnas a tipo str
    df_cierre = df_cierre.fillna('NA').astype(str)
    return df_cierre, registros_no_cierre, df_Siclo_final, df_Facturacion_final

def discarded_logs(registros_no_cierre):
    """
    Data discarded from main match
    """
    # Filtrar el DataFrame
    registros_filtrados = registros_no_cierre[registros_no_cierre['PsEx'] == 'ANO0']

    # Convertir la columna 'Entrega' a texto
    registros_filtrados.loc[:,'Entrega'] = registros_filtrados['Entrega'].astype(str)

    #columnas_deseadas = ['CLIENTE', 'Entrega', 'Incot', 'Fe.SMreal', 'Textobrevedematerial', 'Cantidadentrega', 'UMV_OK']
    df_no_registros = registros_filtrados.loc[:, ['CLIENTE', 'Entrega', 'Incot', 'Fe.SMreal', 'Textobrevedematerial', 'Cantidadentrega', 'UMV_OK']]

    # Filtrar registros donde 'Entrega' inicia con '3'
    df_inicia3 = df_no_registros[df_no_registros['Entrega'].str.startswith('3')].copy()

    # Filtrar registros donde 'Entrega' inicia con '8'
    df_inicia8 = df_no_registros[df_no_registros['Entrega'].str.startswith('8')].copy()

    # Crear una nueva columna en los dataframes 
    df_inicia3['Concatenada'] = df_inicia3['CLIENTE'].astype(str) + df_inicia3['Textobrevedematerial'].astype(str) + df_inicia3['Cantidadentrega'].astype(str)

    # Crear una nueva columna en dataframeF
    df_inicia8['Concatenada'] = df_inicia8['CLIENTE'].astype(str) + df_inicia8['Textobrevedematerial'].astype(str) + df_inicia8['Cantidadentrega'].astype(str)

    # Asignar valores a una nueva columna según la frecuencia de cada valor
    df_inicia8['Cant_Reg'] = 1 + df_inicia8.groupby('Concatenada')['Concatenada'].cumcount()
    df_inicia3['Cant_Reg'] = 1 + df_inicia3.groupby('Concatenada')['Concatenada'].cumcount()

    # Crear una nueva columna en los dataframes 
    df_inicia3['Concatenada_N'] = df_inicia3['Concatenada'] + df_inicia3['Cant_Reg'].astype(str)
    df_inicia3['Concatenada_B'] = df_inicia3['CLIENTE'].astype(str) + df_inicia3['Textobrevedematerial'].astype(str) + df_inicia3['Cantidadentrega'].str.split(',').str[0] + df_inicia3['Cant_Reg'].astype(str)
    df_inicia3['Concatenada_C'] = df_inicia3['Textobrevedematerial'].astype(str) + df_inicia3['Cantidadentrega'].astype(str) + df_inicia3['Cant_Reg'].astype(str)

    # Crear una nueva columna en dataframeF
    df_inicia8['Concatenada_N'] = df_inicia8['Concatenada'] + df_inicia8['Cant_Reg'].astype(str)
    df_inicia8['Concatenada_B'] = df_inicia8['CLIENTE'].astype(str) + df_inicia8['Textobrevedematerial'].astype(str) + df_inicia8['Cantidadentrega'].str.split(',').str[0] + df_inicia8['Cant_Reg'].astype(str)
    df_inicia8['Concatenada_C'] = df_inicia8['Textobrevedematerial'].astype(str) + df_inicia8['Cantidadentrega'].astype(str) + df_inicia8['Cant_Reg'].astype(str)

    # Combinar ambos dataframes utilizando la columna concatenada como clave de combinación
    df_cruce_ = pd.merge(df_inicia3,
                         df_inicia8[['Concatenada_N','Entrega', 'Fe.SMreal', 'UMV_OK']],
                         on='Concatenada_N', how='left')

    df_cruce_B = pd.merge(df_inicia3,
                         df_inicia8[['Concatenada_B','Entrega', 'Fe.SMreal', 'UMV_OK']],
                         on='Concatenada_B', how='left')

    df_cruce_C = pd.merge(df_inicia3,
                         df_inicia8[['Concatenada_C','Entrega', 'Fe.SMreal', 'UMV_OK']],
                         on='Concatenada_C', how='left')

    orden_columnas = ['Fe.SMreal_x', 'CLIENTE', 'Textobrevedematerial', 'Entrega_x', 'UMV_OK_x',
                    'Incot', 'Cantidadentrega', 'Concatenada', 'Cant_Reg', 'Concatenada_N','Concatenada_B',
                    'Concatenada_C','Fe.SMreal_y', 'Entrega_y', 'UMV_OK_y']

    # Reindexar el DataFrame con el nuevo orden de columnas
    df_cruce_ = df_cruce_.reindex(columns=orden_columnas)
    df_cruce_B = df_cruce_B.reindex(columns=orden_columnas)
    df_cruce_C = df_cruce_C.reindex(columns=orden_columnas)

    # Eliminar filas con valores NaN en la columna 'Entrega_y'
    df_cruce_ = df_cruce_.dropna(subset=['Entrega_y'])
    df_cruce_B = df_cruce_B.dropna(subset=['Entrega_y'])
    #df_cruce_C = df_cruce_C.dropna(subset=['Entrega_y'])

    # Concatenar los DataFrames
    df_cruce_1 = pd.concat([df_cruce_, df_cruce_B], ignore_index=True)

    # Eliminar duplicados
    df_cruce_1 = df_cruce_1.drop_duplicates(subset=["Concatenada_N"], keep='first')

    # Concatenar los DataFrames
    df_cruce_final = pd.concat([df_cruce_1, df_cruce_C], ignore_index=True)

    # Eliminar duplicados
    df_cruce_final = df_cruce_final.drop_duplicates(subset=["Concatenada_N"], keep='first')

    # Restablecer el índice
    df_cruce_final = df_cruce_final.reset_index(drop=True)

    # Definir el nuevo nombre de las columnas
    columnas_renombradas = {
        'Fe.SMreal_x': 'Fecha1',
        'CLIENTE': 'Cliente',
        'Textobrevedematerial': 'Producto',
        'Entrega_x': 'Entrega1',
        'UMV_OK_x': 'Peso1',
        'Incot': 'Incoterm',
        'Cantidadentrega': 'Cantidad entregada',
        'Fe.SMreal_y': 'Fecha2',
        'Entrega_y': 'Entrega2',
        'UMV_OK_y': 'Peso2'
    }

    # Cambiar el nombre de las columnas
    df_cruce_final = df_cruce_final.rename(columns=columnas_renombradas)

    # Seleccionar las columnas deseadas
    df_cruce_final = df_cruce_final[['Fecha1', 'Cliente', 'Producto', 'Entrega1', 'Peso1', 'Fecha2', 'Entrega2', 'Peso2']]

    # Filtrar las filas en df_inicia8 donde la columna 'Entrega' no está presente en df_cruce_final['Entrega2']
    registros_no_cruce = df_inicia8[~df_inicia8['Entrega'].isin(df_cruce_final['Entrega2'])]

    registros_no_cruce = registros_no_cruce.copy()
    registros_no_cruce=registros_no_cruce[['CLIENTE','Textobrevedematerial','Cantidadentrega','Fe.SMreal','Entrega','UMV_OK']]

    # Renombrar las columnas Fecha1 y Fecha2
    registros_no_cruce.rename(columns={'CLIENTE':'Cliente','Textobrevedematerial':'Producto','Cantidadentrega':'Peso1','Fe.SMreal': 'Fecha2', 'Entrega': 'Entrega2', 'UMV_OK': 'Peso2'}, inplace=True)

    # Unir los DataFrames uno encima del otro
    df_concatenado = pd.concat([df_cruce_final, registros_no_cruce])

    # Reemplazar valores None y NaN por cadenas vacías en todo el DataFrame
    df_concatenado = df_concatenado.replace({None: '', np.nan: ''})

    df_concatenado = df_concatenado.reset_index(drop=True)
    
    return df_concatenado

def load_credentials(key, scopes):
    """
    Load Google API credentials.
    """
    return service_account.Credentials.from_service_account_file(key, scopes=scopes)

def authorize_google_sheets(creds):
    """
    Authorize access to Google Sheets using gspread.
    """
    return gspread.authorize(creds)

def open_spreadsheet(gc, url):
    """
    Open Google Sheets spreadsheet by URL.
    """
    return gc.open_by_url(url)

def get_worksheet(spreadsheet, sheet_name):
    """
    Get specific worksheet from the spreadsheet.
    """
    return spreadsheet.worksheet(sheet_name)

def get_column_values(worksheet, column):
    """
    Get all values from a specific column.
    """
    return worksheet.col_values(column)

def find_last_row(column_values):
    """
    Find the last non-empty row in a column.
    """
    return len(column_values)

def find_month_change_indices(column_values):
    """
    Find indices where the month changes.
    """
    return [i for i, val in enumerate(column_values) if i > 0 and column_values[i] != column_values[i - 1]]

def process_despacho_data(worksheet, start_row, end_row, titles):
    """
    Process 'Despacho' worksheet data.
    """
    data = worksheet.get(f'A{start_row}:Y{end_row}')
    df = pd.DataFrame(data, columns=titles)
    df['Estado'] = ''
    return df

def process_facturacion_data(worksheet, start_row, end_row, titles):
    """
    Process 'Facturacion' worksheet data.
    """
    data = worksheet.get(f'A{start_row}:BL{end_row}')
    df = pd.DataFrame(data, columns=titles)
    df = df.fillna('NA').astype(str)
    df = df.rename(columns=lambda x: x.replace(" ", ""))
    return df

def data_spreadsheet(n):
    """
    Download data from the 'Data_Transcem' spreadsheet.
    """
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    KEY = 'credentials.json'
    URL = 'https://docs.google.com/spreadsheets/d/1l4QNZ1P3HweFLViHHGwrkFCF9pS8gECDAUKP9xBlcB0'
    
    creds = load_credentials(KEY, SCOPES)
    gc = authorize_google_sheets(creds)
    spreadsheet = open_spreadsheet(gc, URL)

    worksheet_D = get_worksheet(spreadsheet, 'Despacho')
    worksheet_F = get_worksheet(spreadsheet, "Facturación")
    #worksheet_C = get_worksheet(spreadsheet, "Cierre")
    #worksheet_H = get_worksheet(spreadsheet, "Hora")
    #worksheet_R = get_worksheet(spreadsheet, "Registros_NoCruce")

    column_values_D = get_column_values(worksheet_D, 24)
    column_values_F = get_column_values(worksheet_F, 16)
    #meses = [datetime.strptime(fecha, '%d.%m.%Y').month for fecha in column_values_F[1:]]

    last_row_D = find_last_row(column_values_D)
    last_row_F = find_last_row(column_values_F)

    cambio_mes_indices_D = find_month_change_indices(column_values_D)
    cambio_mes_indices_F = find_month_change_indices(column_values_F)

    posicion_ultimo_mes_D = cambio_mes_indices_D[-n] + 1
    posicion_ultimo_mes_F = cambio_mes_indices_F[-n] + 1

    titulos_D = [
        'Remesa', 'Negocio', 'Fecha Despacho', 'Fecha Manifiesto', 'Destinatario', 'Cliente', 
        'Transportador', 'Administrador', 'Placa', 'Remolque', 'Linea_codProducto', 'Linea_nombreProducto',
        'Linea_UOM', 'Linea_Cantidad', 'Linea_Peso', 'Transporte Siclo', 'Transporte SAP', 'Cód. DANE', 
        'Ciudad', 'Departamento', 'Báscula salida', 'Año', 'Mes', 'Día'
    ]

    titulos_F = [
        'Se', 'PsEx', 'Solic.', 'CLIENTE', 'Destinat.', 'ECOBRA', 'Referencia', 'Fe.carga',
        'Identif.externadenotaentreg', 'Entrega', 'ClEnt', 'Tp.DC', 'Fe.prev.SM', 'Incot', 'CE',
        'Fe.SMreal', 'GuiaManual', 'Textobrevedematerial', 'Gr.1', 'Alm.', 'TPos', 'GrpPortM',
        'Div.', 'Cantidadentrega', 'UM', 'Material', 'Pos.', 'GVen', 'UM.1', 'Doc.Ventas',
        'Fabricante', 'GrM3', 'GrM', 'ZTOBRA', 'PRVOBRA', 'RPago', 'RespPago', 'Referencia1',
        'Grupodeclientes1', 'STATUSSM', 'PedidoCompra', 'ClavedeaccesoaSRIEC', 'UMV_OK', 'UM.2',
        'NTGEW_OK', 'Un', 'Z640', 'Mon.', 'Z627', 'Mon..1', 'Z645', 'Mon..2', 'Z641', 'Mon..3',
        'OrgVt', 'SectorMaterial', 'SectorMaterial.1', 'Z672', 'Mon..4', 'Z673', 'Mon..5', 'Z675',
        'Mon..6', 'Hora'
    ]

    BD_Siclo = process_despacho_data(worksheet_D, posicion_ultimo_mes_D, last_row_D, titulos_D)
    BD_SAP = process_facturacion_data(worksheet_F, posicion_ultimo_mes_F, last_row_F, titulos_F)

    return BD_Siclo, BD_SAP, posicion_ultimo_mes_D, posicion_ultimo_mes_F, worksheet_D, worksheet_F
    
def prepare_data(df_Siclo_final, posicion_ultimo_mes, df_Facturacion_final, posicion_ultimo_mes_F, df_cierre):
    """
    Prepare the data for update spreadsheet
    
    """
    # Obtener los datos
    data = df_Siclo_final.values.tolist()

    # Definir el rango de celdas a actualizar
    start_row = posicion_ultimo_mes
    start_col = 1  # Columna A
    end_row = start_row + len(data) - 1
    end_col = start_col + len(df_Siclo_final.columns) - 1
    range_str = f'A{start_row}:Y{end_row}'
    
    # Obtener los datos Facturación
    data_F = df_Facturacion_final.values.tolist()

    # Definir el rango de celdas a actualizar
    start_row_F = posicion_ultimo_mes_F
    start_col_F = 1  # Columna A
    end_row_F = start_row_F + len(data_F) - 1
    end_col_F = start_col_F + len(df_Facturacion_final.columns) - 1
    range_str_F = f'A{start_row_F}:BL{end_row_F}'
    
    # Obtener los datos de Cierre
    data_C = df_cierre.values.tolist()

    # Definir el rango de celdas a actualizar
    start_row_C = posicion_ultimo_mes
    start_col_C = 1  # Columna A
    end_row_C = start_row_C + len(data_C) - 1
    end_col_C = start_col_C + len(df_cierre.columns) - 1
    range_str_C = f'A{start_row_C}:AC{end_row_C}'
    
    return data, range_str, data_F, range_str_F, data_C, range_str_C

def eliminate_logs_spreadsheet(posicion_ultimo_mes, posicion_ultimo_mes_F, worksheet_D, worksheet_F, worksheet_C, worksheet_R, last_row, last_row_F):
    """
    Elimnate the data of spreadsheet of current month
    
    """
    worksheet_R.clear()
    
    #Borrar datos de Despachos en la Hoja de Calculo
    worksheet_D.batch_clear([f'A{posicion_ultimo_mes+1}:Y{last_row}'])

    #Borrar datos de Despachos en la Hoja de Calculo
    worksheet_C.batch_clear([f'A{posicion_ultimo_mes+1}:Y{last_row}'])

    #Borrar datos de Facturación en la Hoja de Calculo
    worksheet_F.batch_clear([f'A{posicion_ultimo_mes_F+2}:BL{last_row_F}'])
    
def clear_worksheet_data(worksheet, start_col, end_col, start_row, end_row):
    """
    Clears data in a specified range in the worksheet.
    """
    worksheet.batch_clear([f'{start_col}{start_row}:{end_col}{end_row}'])

def update_worksheet_data(worksheet, data, range_str):
    """
    Updates data in a specified range in the worksheet.
    """
    worksheet.update(data, range_str)

def update_data_spreadsheet(worksheet_D, worksheet_F, df_concatenado, current_date, data, range_str, data_F, range_str_F, data_C, range_str_C):
    """
    Update data of the spreadsheet 'Data_Transcem'
    
    """
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    KEY = 'credentials.json'
    URL = 'https://docs.google.com/spreadsheets/d/1l4QNZ1P3HweFLViHHGwrkFCF9pS8gECDAUKP9xBlcB0'
    
    creds = load_credentials(KEY, SCOPES)
    gc = authorize_google_sheets(creds)
    spreadsheet = open_spreadsheet(gc, URL)

    worksheet_C = get_worksheet(spreadsheet, "Cierre")
    worksheet_H = get_worksheet(spreadsheet, "Hora")
    worksheet_R = get_worksheet(spreadsheet, "Registros_NoCruce")
    
    worksheet_R.update([df_concatenado.columns.values.tolist()] + df_concatenado.values.tolist())

    # Register current date and time in the 'Hora' worksheet
    worksheet_H.update([[current_date.strftime('%Y-%m-%d %H:%M:%S')]], 'A1')

    # Update data in worksheets
    worksheet_D.update(data,range_str)
    worksheet_F.update(data_F,range_str_F)
    worksheet_C.update(data_C,range_str_C)

    
def fill_parameters(sQuery, date_ini, date_fin):
    sQuery = sQuery.replace('par_DateIni', date_ini)
    sQuery = sQuery.replace('par_DateFin', date_fin)
    return sQuery

def save_time_log(current_date, tiempo, path_bitacora):
    with open(path_bitacora, 'a') as archivo:
        cadena_datos = current_date.strftime('%Y-%m-%d %H:%M:%S') + " | " +str(tiempo)
        archivo.write(cadena_datos + '\n')
        
def sort_dataframe_by_date_and_time(df, date_column='Fe.SMreal', time_column='Hora'):
    """
    Sorts a DataFrame by date and time columns from smallest to largest.

    Args:
    df (pd.DataFrame): El DataFrame que contiene las columnas de fecha y hora.
    date_column (str): El nombre de la columna de fecha en el formato "DD.MM.AAAA".
    time_column (str): El nombre de la columna de hora en el formato "hh:mm:ss".

    Returns:
    pd.DataFrame: El DataFrame ordenado por fecha y hora con el formato original de la fecha.
    """
    
    # Trabajar con una copia del DataFrame original
    df = df.copy()
    
    # Convertir la columna de fecha al formato datetime, ignorando errores
    df[date_column] = pd.to_datetime(df[date_column], format='%d.%m.%Y', errors='coerce')
    
    # Convertir la columna de hora al formato datetime.time, ignorando errores
    df[time_column] = pd.to_datetime(df[time_column], format='%H:%M:%S', errors='coerce').dt.time
    
    # Eliminar filas con valores NaT en la columna de fecha
    df = df.dropna(subset=[date_column])
    
    # Ordenar el DataFrame por fecha y hora
    df = df.sort_values(by=[date_column, time_column])
    
    # Convertir la columna de fecha de vuelta al formato original "DD.MM.AAAA"
    df[date_column] = df[date_column].dt.strftime('%d.%m.%Y')
    
    # Convertir la columna de hora al formato de cadena
    df[time_column] = df[time_column].astype(str)
    
    return df