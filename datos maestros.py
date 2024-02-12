import sys
import win32com.client as win32
import subprocess
import time
import schedule
import openpyxl
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

user = ""
passwd = ""
sap_erp = ["1. Grupo Nutresa_ERP_PRD", 300, "ES"]

# Resto del código de la clase SapGui ...

def subir_a_drive(archivo_local, carpeta_drive):
    gauth = GoogleAuth()
    # Intentar cargar las credenciales desde un archivo existente
    gauth.LoadCredentialsFile("mycreds.txt")
    if gauth.credentials is None:
        # Realizar la autenticación manual
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Actualizar las credenciales si han expirado
        gauth.Refresh()
    else:
        # Autorizar las credenciales
        gauth.Authorize()

    # Guardar las credenciales para futuras ejecuciones
    gauth.SaveCredentialsFile("mycreds.txt")

    drive = GoogleDrive(gauth)

    # Crear un archivo en Google Drive y cargar el archivo local
    file_drive = drive.CreateFile({"title": archivo_local, "parents": [{"kind": "drive#fileLink", "id": carpeta_drive}]})
    file_drive.Upload()
    print(f"Archivo '{archivo_local}' subido a Google Drive en la carpeta con ID {carpeta_drive}")

def ejecutar_script():
    sap = SapGui()
    try:
        sap.session.findById("wnd[0]").maximize
        sap.session.findById("wnd[0]/tbar[0]/okcd").text = "ZPP_POL_1308"
        sap.session.findById("wnd[0]").sendVKey(0)
        sap.session.findById("wnd[0]/tbar[1]/btn[17]").press
        sap.session.findById("wnd[1]/usr/txtENAME-LOW").text = "ACJEVELANDIA"
        sap.session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
        sap.session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 12
        sap.session.findById("wnd[1]/tbar[0]/btn[8]").press
        sap.session.findById("wnd[0]/tbar[1]/btn[8]").press
        sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(5, "MEINH")
        sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "5"
        sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").contextMenu
        sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
        sap.session.findById("wnd[1]/tbar[0]/btn[0]").press
        sap.session.findById("wnd[1]/tbar[0]/btn[0]").press

        # Esperar un tiempo suficiente para que la descarga se complete (ajusta según sea necesario)
        time.sleep(5)

        # Obtener datos de la celda activa
        active_cell_data = sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").getCellValue(5, "MEINH")

        # Crear un archivo Excel y almacenar los datos
        archivo_excel = "datos_descargados.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet['A1'] = "Datos Descargados"
        sheet['A2'] = "Celda Activa:"
        sheet['B2'] = active_cell_data

        workbook.save(archivo_excel)
        print(f"Datos descargados y guardados en {archivo_excel}")

        # Subir el archivo a Google Drive
        carpeta_drive_id = "1X_ZsHhHANqNfGiTuYpMUo9zAQ82FIUzz"  # ID de la carpeta en Google Drive
        subir_a_drive(archivo_excel, carpeta_drive_id)
    finally:
        sap.cerrarSAP()

# Horario de ejecución (cada día a las 10:00 AM)
hora_ejecucion = "10:00"
schedule.every().day.at(hora_ejecucion).do(ejecutar_script)

while True:
    schedule.run_pending()
    time.sleep(1)