import pandas as pd
import logging
import time
from HU.pagoArriendos import ConexionSAP
from HU.ValidarLiberacion import ValidarLiberacion
from HU.GestionAnexos import GestionAnexos # La nueva clase para ME22N
from config.settings import SAP_CONFIG, RUTAS

def main_proceso_masivo():
    # 1. Configuración de Logs
    logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')
    logger = logging.getLogger(__name__)

    # 2. Leer el Excel
    ruta_excel = RUTAS['PATH_INSUMO'] + "/Listado_OC.xlsx"
    try:
        df = pd.read_excel(ruta_excel)
        logger.info(f"Procesando {len(df)} órdenes de compra.")
    except Exception as e:
        logger.error(f"Error al leer Excel: {e}")
        return

    # 3. Conexión a SAP (Una sola vez)
    sap = ConexionSAP(SAP_CONFIG.get('SAP_USUARIO'),
                      SAP_CONFIG.get('SAP_PASSWORD'),
                      SAP_CONFIG.get('SAP_CLIENTE'),
                      SAP_CONFIG.get('SAP_IDIOMA'),
                      SAP_CONFIG.get('SAP_PATH'),
                      SAP_CONFIG.get('SAP_SISTEMA'))
    sap.iniciar_sesion_sap()
    
    # Instanciar las dos herramientas
    cargador = GestionAnexos(sap)
    validador = ValidarLiberacion(sap)
    
    resultados = []

    # 4. LOOP de procesamiento
    for index, fila in df.iterrows():
        oc_actual = str(fila['OC']).split('.')[0]
        # Si tienes una columna con la ruta del PDF, la usamos. Si no, simulamos una ruta.
        ruta_pdf = str(fila.get('RUTA_ANEXO', '')) 
        
        logger.info(f"--- INICIO OC {oc_actual} ({index + 1}/{len(df)}) ---")
        
        try:
            # PASO A: CARGAR ANEXO (ME22N)
            # Solo intentamos cargar si existe una ruta de archivo válida
            if ruta_pdf != 'nan' and ruta_pdf != '':
                logger.info(f"Paso 1: Cargando anexo en ME22N para OC {oc_actual}...")
                cargador.cargar_archivo_gos(oc_actual, ruta_pdf)
            else:
                logger.info(f"Paso 1: Saltando carga de anexos (No hay archivo definido).")

            # PASO B: VALIDAR LIBERACIÓN (ME23N)
            logger.info(f"Paso 2: Validando estado de liberación en ME23N...")
            esta_liberada = validador.verificar_estado(oc_actual)
            
            if esta_liberada:
                logger.info(f"RESULTADO: OC {oc_actual} LIBERADA.")
                resultados.append("Cargado y Liberado")
            else:
                logger.warning(f"RESULTADO: OC {oc_actual} BLOQUEADA.")
                resultados.append("Pendiente de firmas")

        except Exception as e:
            logger.error(f"Error en loop para OC {oc_actual}: {e}")
            resultados.append("Error en proceso")

    # 5. Cierre y Reporte
    df['Resultado_Final'] = resultados
    df.to_excel(RUTAS['PATH_RESULTADO'] + "/Reporte_Completo_Arriendos.xlsx", index=False)
    logger.info("Proceso terminado exitosamente.")

if __name__ == "__main__":
    main_proceso_masivo()