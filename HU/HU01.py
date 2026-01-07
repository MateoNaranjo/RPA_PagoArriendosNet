# ===============================
# HU01: Nombre HU
# Autor: Santiago Pinzon - Desarrollador RPA
# Descripcion: Descripcion de la HU 
# Ultima modificacion: 2/1/2026
# Propiedad de Colsubsidio
# Cambios: Si aplica
# ===============================
from funciones.ControlHU import control_hu
from funciones.EscribirLog import WriteLog
from HU.pagoArriendos import ConexionSAP
from config.settings import SAP_CONFIG
from config.init_config import in_config
from HU.ME2L import TransaccionME2L
from funciones.Excel import Excel
from repositorios.excel import Excel as ExcelRepo
from pywinauto import Desktop

def HU01_Prueba():
    """
    Docstring for HU01_Prueba
    """
    # =========================
    # CONFIGURACION DEL PROCESO
    # =========================
    task_name = "HU01_Prueba"

    try:
        # === Inicio HU01 ===
        control_hu(task_name, 0)
        # GestionTicketInsumo(estado, id, maquina, observaciones)
        # WriteLog(mensaje="Inicio HU01", estado="INFO", task_name=task_name)
        
        # ============================= Inicio acciones =============================
        sap = ConexionSAP(SAP_CONFIG.get('SAP_USUARIO'),
                        SAP_CONFIG.get('SAP_PASSWORD'),
                        in_config('SAP_CLIENTE'),
                        in_config('SAP_IDIOMA'),
                        in_config('SAP_PATH'),
                        in_config('SAP_SISTEMA')
                    )
        sap.iniciar_sesion_sap()
        sap.abrir_transaccion("ME2L")
        me2l = TransaccionME2L(sap)
        oc = me2l.buscar_oc_activa('900346807')
        ruta_archivo = in_config("PathTemp")

        me2l.exportar_tabla(ruta_archivo)

        print(oc)

        ruta_excel = in_config("PathInsumos")+"\BaseMedicamentos.xlsx"

        print(ruta_excel)

        
        #if ruta_excel:
        #    Excel.ejecutar_bulk(ruta_excel)

        TablaBase = ExcelRepo.obtener_valores()

        for registro in TablaBase:
            print(registro['NIT'])
        # ============================= Finalizacion HU =============================

        control_hu(task_name, 100)
    
    except Exception as e:
        print(f"Error en ejecucion: ({e}) ")
        # WriteLog()
        control_hu(task_name, 99)
        raise
        # GestionTicketInsumo(id, observaciones, estado, maquina)

    finally:
        # WriteLog()
        log = "Finalizacion HU"
        print(log)
        