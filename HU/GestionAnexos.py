import time
import logging
import threading
import pyautogui
from pywinauto import Desktop

class GestionAnexos:
    def __init__(self, sap_conexion):
        self.sesion = sap_conexion.sesion
        self.logger = logging.getLogger("main_proceso_masivo")

    def _interaccion_ventana_windows(self, ruta_archivo):
        """
        Maneja la ventana externa 'Import file' inyectando la ruta y dando Enter.
        """
        self.logger.info("Hilo secundario: Vigilando ventana 'Import file'...")
        inicio = time.time()
        timeout = 20
        
        while (time.time() - inicio) < timeout:
            try:
                desktop = Desktop(backend="win32")
                ventana = desktop.window(title_re="(?i)Import.*file.*")
                
                if ventana.exists():
                    ventana.set_focus()
                    time.sleep(1)
                    
                    # Inyectar ruta directamente
                    edit_box = ventana.child_window(class_name="Edit")
                    edit_box.set_edit_text(ruta_archivo)
                    time.sleep(1)
                    
                    # Forzar cierre con Enter
                    edit_box.type_keys('{ENTER}')
                    
                    # Respaldo si no cierra
                    time.sleep(1)
                    if ventana.exists():
                        pyautogui.press('enter')
                    
                    self.logger.info("Hilo secundario: Ventana de Windows procesada.")
                    return
            except:
                pass
            time.sleep(0.5)

    def cargar_archivo_gos(self, num_oc, ruta_archivo):
        try:
            # 1. Navegar y cargar OC en ME22N
            self.sesion.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.sesion.findById("wnd[0]/tbar[0]/okcd").text = "ME22N"
            self.sesion.findById("wnd[0]").sendVKey(0)
            
            self.sesion.findById("wnd[0]/tbar[1]/btn[17]").press()
            self.sesion.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = num_oc
            self.sesion.findById("wnd[1]").sendVKey(0)
            time.sleep(1)

            # 2. LANZAR HILO PARA VENTANA DE WINDOWS
            hilo_externo = threading.Thread(target=self._interaccion_ventana_windows, args=(ruta_archivo,))
            hilo_externo.daemon = True
            hilo_externo.start()

            # 3. DISPARAR ACCIÓN (Bloqueo de SAP)
            self.sesion.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
            self.sesion.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_PCATTA_CREA")

            # 4. ESPERAR DESBLOQUEO Y MANEJAR POPUP DE SEGURIDAD (Imagen 2)
            hilo_externo.join(timeout=15)
            time.sleep(2) # Tiempo para que aparezca la ventana de 'Allow'

            # --- NUEVO AJUSTE: MANEJO DE 'ALLOW' EN SAP ---
            try:
                # Si hay una ventana secundaria abierta (wnd[1])
                if self.sesion.Children.Count > 1:
                    self.logger.info("Detectada ventana de seguridad de SAP. Intentando presionar 'Allow'...")
                    
                    # Intentamos los dos IDs más comunes para el botón 'Allow'
                    try:
                        # Opción A (Estandar SPOP)
                        self.sesion.findById("wnd[1]/usr/btnSPOP-VAROCB1").press()
                    except:
                        # Opción B (Estandar Button)
                        self.sesion.findById("wnd[1]/usr/btnBUTTON_1").press()
                    
                    self.logger.info("Clic en 'Allow' realizado exitosamente.")
                    time.sleep(1)
            except Exception as e:
                self.logger.warning(f"No se pudo interactuar con el popup de seguridad: {e}")

            # 5. GUARDAR CAMBIOS
            # Importante: Esto confirma el anexo en la OC
            self.sesion.findById("wnd[0]/tbar[0]/btn[11]").press()
            
            self.logger.info(f"OC {num_oc}: Anexo cargado y pedido grabado.")
            return True

        except Exception as e:
            self.logger.error(f"Error en proceso masivo para OC {num_oc}: {e}")
            return False