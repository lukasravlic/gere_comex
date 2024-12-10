# %%
import pandas as pd
import numpy as np
import os
import datetime
import win32com.client
import time
import getpass
from datetime import timedelta

usuario = getpass.getuser()

inicio = time.time()


# %%
import os
from datetime import datetime

# # Define the directory path
# carpeta_fechas = "C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Actualización diaria fechas DT/Actualización Diaria fechas Dts R3"

# # Get the list of files in the directory
# ruta_fechas = os.listdir(carpeta_fechas)

# # Extract dates from file names (assuming dates are in YYYY-MM-DD format)
# hoy = (datetime.today()).strftime('%d-%m-%Y') 
# ayer = (datetime.today()- timedelta(days=1)).strftime('%d-%m-%Y')
# columnas = ['Nro. DT','Vía (Texto)']
# max_days_back = 30

# # Function to try reading the file for a given date
# def try_read_file(date_str):
#     ruta_arch = f'{date_str} Actualizacion fechas diaria  Dts OEM.xlsx'
#     ruta = os.path.join(carpeta_fechas, ruta_arch)
#     if os.path.exists(ruta):
#         df_fechas = pd.read_excel(ruta, sheet_name='Data', dtype = {'Nro. DT':'str'})
#         print(f"File found: {ruta}")
#         return df_fechas
#     return None

# # Try to find a file for today and up to max_days_back days ago
# for days_back in range(max_days_back + 1):
#     date_to_try = (datetime.today() - timedelta(days=days_back)).strftime('%d-%m-%Y')
#     df_fechas = try_read_file(date_to_try)
#     if df_fechas is not None:
#         break
# else:
#     print("No file found within the given date range.")

# # Continue with your further processing
# if df_fechas is not None:
#     # Do something with df_fechas
#     pass
# else:
#     # Handle the case where no file was found
#     pass



df_fechas = pd.read_excel("C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Actualización diaria fechas DT/Actualización Diaria fechas Dts R3/04-12-2024 Actualizacion fechas diaria  Dts OEM (002).xlsx", sheet_name='Data', dtype = {'Nro. DT':'str'})

# %%
filtro = df_fechas[df_fechas['Vía (Texto)'].isin(['Maritimo','Terrestre'])]['Nro. DT']

# %%
filtro.to_clipboard(index=False, header=False)



# %%
import win32com.client

try:
    # Initialize SAP GUI scripting
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
except Exception as e:
    print(f"Error obtaining SAP GUI session: {str(e)}")
    exit(1)

try:
    # Maximize the window
    session.findById("wnd[0]").maximize()

    # Execute transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Navigate to specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_DT").select()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_DT/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1005/btn%_SO_TKNUM_%_APP_%-VALU_PUSH").press()

    # Handle pop-up windows
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)

    # Load variant
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&LOAD")

    # Select variant row
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(2, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export to Excel
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]").close()

    # Handle export pop-up windows
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\lravlic\\Codigos\\automatizacion_gere_comex"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "dts.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("dts.XLSX")
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close SAP windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

except Exception as e:
    print(f"Error during SAP GUI interaction: {str(e)}")











# %%
import xlwings as xw
try:
    book = xw.Book("C:/Users/lravlic/Codigos/automatizacion_gere_comex/dts.XLSX")
    book.close()
except Exception as e:
    print(e)

# %%
import win32com.client

try:
    # Initialize SAP GUI scripting
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    if not application:
        raise Exception("Error obtaining SAP GUI application")

    connection = application.Children(0)
    if not connection:
        raise Exception("Error obtaining SAP GUI connection")

    session = connection.Children(0)
    if not session:
        raise Exception("Error obtaining SAP GUI session")

    # Connect to WScript if available
    try:
        WScript = win32com.client.Dispatch("WScript")
        WScript.ConnectObject(session, "on")
        WScript.ConnectObject(application, "on")
    except Exception as e:
        print(f"Error connecting to WScript: {str(e)}")

    # Maximize the window
    session.findById("wnd[0]").maximize()

    # Execute transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmm_seguim_comex_cl"
    session.findById("wnd[0]").sendVKey(0)

    # Press buttons
    session.findById("wnd[0]/usr/btn%_P_TKNUM_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)

    # Export to Excel
    session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\lravlic\\Codigos\\automatizacion_gere_comex"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "comex.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close SAP windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

except Exception as e:
    print(f"Error during SAP GUI interaction: {str(e)}")




