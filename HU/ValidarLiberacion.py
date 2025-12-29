import time

class ValidarLiberacion:
    def __init__(self, sap_conexion):
        self.sap = sap_conexion
        # Usamos sap_conexion.sesion porque así se llama en tu clase ConexionSAP
        self.sesion = sap_conexion.sesion 
        self.logger = sap_conexion.logger if hasattr(sap_conexion, 'logger') else None

    def verificar_estado(self, num_oc):
        try:
            # 1. Abrir la transacción ME23N (Visualizar Pedido)
            self.sesion.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
            self.sesion.findById("wnd[0]").sendVKey(0)
            
            # 2. Seleccionar la OC específica (Botón 'Other Purchase Order' - Shift+F5)
            # Intentamos presionar el botón de selección de documento
            self.sesion.findById("wnd[0]/tbar[1]/btn[17]").press()
            
            # Escribir el número de la OC en el campo de entrada
            self.sesion.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = num_oc
            self.sesion.findById("wnd[1]").sendVKey(0) 
            time.sleep(2)

            # 3. Asegurar que la cabecera (Header) esté expandida
            # Si el botón de expansión está visible, lo presionamos
            btn_expandir = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0020/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/btnDYN_1105-BUTTON"
            try:
                # Si el texto del tooltip indica que está contraído, expandimos
                if "Expand" in self.sesion.findById(btn_expandir).tooltip:
                    self.sesion.findById(btn_expandir).press()
            except:
                pass # Si no lo encuentra, asumimos que ya está expandido

            # 4. Localizar y hacer clic en la pestaña 'Release Strategy' (Estrategia de Liberación)
            # Nota: El ID 'tabpTABHDT9' es el estándar para esta pestaña
            id_pestaña = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0020/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/tabsTS_HEADER/tabpTABHDT9"
            
            try:
                self.sesion.findById(id_pestaña).select()
            except Exception:
                if self.logger: self.logger.warning(f"No se visualiza pestaña de liberación para OC {num_oc}. Puede que no la requiera.")
                return True # Si no existe la pestaña, SAP no exige liberación para esta OC

            # 5. Leer el texto del estado de liberación
            # Este campo indica si hay bloqueos o si ya está 'Release Completed'
            id_texto_estado = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0020/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/tabsTS_HEADER/tabpTABHDT9/ssubSUB_TABSTRIP:SAPLMEGUI:1107/txtMEPO1217-RLSST"
            estado_texto = self.sesion.findById(id_texto_estado).text
            
            if self.logger: self.logger.info(f"Estado de liberación detectado: {estado_texto}")

            # Lógica de validación: Si el texto indica bloqueo o falta de firmas
            # (Depende de cómo esté configurado tu SAP, usualmente 'X' o 'Blocked' es mala señal)
            if "X" in estado_texto or "Bloqueado" in estado_texto.lower():
                if self.logger: self.logger.error(f"La OC {num_oc} está pendiente de liberación.")
                return False
            
            if self.logger: self.logger.info(f"¡OC {num_oc} liberada correctamente!")
            return True

        except Exception as e:
            if self.logger: self.logger.error(f"Error técnico validando liberación: {str(e)}")
            return False