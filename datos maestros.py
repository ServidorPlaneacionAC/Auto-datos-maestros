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

class SapGui():
    def __init__(self):
        try:
            import win32com.client as win32

            # Interfaz básica con SAP
            self.SapGuiAuto = win32.GetObject("SAPGUI")
            self.application = self.SapGuiAuto.GetScriptingEngine
            self.connection = self.application.Children(0)
            self.session = self.connection.Children(0)

            if not self.__run_SAP():
                self.__arrancar_SAP()

        except:
            self.__arrancar_SAP()

    def __arrancar_SAP(self):
        """Función para arrancar de forma limpia una ventana de saplogon"""

        # Hacer inicialización limpia de SAP
        cmd("C:\\Windows\\System32\\TASKKILL /IM saplogon.exe /F")
        self.path = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(3)

        self.SapGuiAuto = win32.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32.client.CDispatch:
            return
        self.application = self.SapGuiAuto.GetScriptingEngine
        self.connection = self.application.OpenConnection(sap_erp[0], True)
        time.sleep(3)

        self.session = self.connection.Children(0)
        time.sleep(3)

        try:
            if not self.__run_SAP():
                # Logueo con usuario y contraseña
                self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = sap_erp[1]      # Mandante
                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = Acwagavilan     # Usuario
                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = Marzo2024-      # Contraseña
                self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = sap_erp[2]      # Idioma
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(3)
        except:
            pass

        self.__limpiar_msje()

    def __run_SAP(self):
        """Función para detectar si hay una sesión activa de SAP"""

        # Validar si la sesion de SAP está activo
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n1"
        self.session.findById("wnd[0]").sendVKey(0)
        mensaje = self.session.findById("wnd[0]/sbar/pane[0]").Text
        time.sleep(3)

        if mensaje == "La transacción 1 no existe":
            return True

        else:
            return False

    def __limpiar_msje(self):
        """Función para limpiar de mensajes emergentes y esporádicos el inicio de saplogon"""

        # Eliminar mensajes de comienzo de sesión
        try:
            self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select
            self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press
            self.session.findById("wnd[1]").sendVKey(0)
            self.session.findById("wnd[1]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[1]").sendVKey(0)

        except:
            pass

    def cerrarSAP(self):
        """Función para cerrar completamente todos los procesos y ventanas de SAP"""

        try:
            self.session.findById("wnd[0]").close()
            time.sleep(1.0)
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # Boton SI
            SapGui.terminate()
        except:
            cmd("C:\\Windows\\System32\\TASKKILL /IM saplogon.exe /F")
        print("Todos los procesos en SAP se han cerrado exitosamente")

    def subir_a_drive(self, archivo_local, carpeta_drive):
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
        sap.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
        sap.session.findById("wnd[1]/tbar[0]/btn[0]").press
        sap.session.findById("wnd[1]/tbar[0]/btn[0]").press

        # Guardar el archivo localmente con openpyxl
        nombre_archivo_local = "nombre_del_archivo.xlsx"
        libro_excel = openpyxl.Workbook()
        libro_excel.save(nombre_archivo_local)

        # Subir el archivo a Google Drive
        carpeta_drive = "ID_de_la_carpeta_en_Google_Drive"
        sap.subir_a_drive(nombre_archivo_local, carpeta_drive)

    finally:
        sap.cerrarSAP()

# Horario de ejecución (cada día a las 10:00 AM)
hora_ejecucion = "10:00"
schedule.every().day.at(hora_ejecucion).do(ejecutar_script)

while True:
    schedule.run_pending()
    time.sleep(1)
