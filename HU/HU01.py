# ===============================
# HU01: Nombre HU
# Autor: Santiago Pinzon - Desarrollador RPA
# Descripcion: Descripcion de la HU 
# Ultima modificacion: 2/1/2026
# Propiedad de Colsubsidio
# Cambios: Si aplica
# ===============================
from funciones.ControlHU import control_hu
from config.init_config import in_config
from funciones.EscribirLog import WriteLog

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
        
        print(f'Ejecutado, valor de prueba in_config: {in_config("Prueba")}')

        # ============================= Finalizacion HU =============================

        control_hu(task_name, 100)
    
    except Exception as e:
        print(f"Error en ejecucion: ({e}) ")
        # WriteLog()
        control_hu(task_name, 99)
        # GestionTicketInsumo(id, observaciones, estado, maquina)

    finally:
        # WriteLog()
        log = "Finalizacion HU"
        print(log)
        