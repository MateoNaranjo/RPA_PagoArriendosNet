# ===============================
# HU01: Cadena
# Autor: Santiago Pinzon - Desarrollador RPA
# Descripcion: Descripcion de la HU 
# Ultima modificacion: 2/1/2026
# Propiedad de Colsubsidio
# Cambios: Si aplica
# ===============================
from funciones.Cadena import login_colsubsidio, realizar_consulta, descargar_xml_final
from funciones.ControlHU import control_hu
from repositorios.excel import ExcelRepo

def HU01_Cadena():
    """
    Docstring for HU01_Prueba
    """
    # =========================
    # CONFIGURACION DEL PROCESO
    # =========================
    task_name = "HU01_Cadena"

    try:
        # === Inicio HU01 ===
        control_hu(task_name, 0)
        # GestionTicketInsumo(estado, id, maquina, observaciones)
        # WriteLog(mensaje="Inicio HU01", estado="INFO", task_name=task_name)
        # ============================= Inicio acciones =============================

        for i in range(3):
            try:
                TablaBase = ExcelRepo.obtener_valores("basemedicamentoslimpio")

                for registro in TablaBase[29:]:
                    sesion = login_colsubsidio()
                    realizar_consulta(sesion, oc="4001249504")
                    descargar_xml_final(sesion)
                break
            except Exception as e:
                raise

        # ============================= Finalizacion HU =============================
        control_hu(task_name, 100)
    
    except Exception as e:
        print(f"Error en ejecucion: ({e}) ")
        # WriteLog()
        control_hu(task_name, 99)
        raise

    finally:
        # WriteLog()
        log = "Finalizacion HU"
        print(log)