from HU.pagoArriendos import ConexionSAP
from HU.LeerXML import LectorFacturaXML
from HU.ME2L import TransaccionME2L
from HU.MIGO import TransaccionMIGO
from config.settings import SAP_CONFIG, RUTAS

def main():
    # 1. Conexión
    sap = ConexionSAP(SAP_CONFIG.get('SAP_USUARIO'),
                           SAP_CONFIG.get('SAP_PASSWORD'),
                           SAP_CONFIG.get('SAP_CLIENTE'),
                           SAP_CONFIG.get('SAP_IDIOMA'),
                           SAP_CONFIG.get('SAP_PATH'),
                           SAP_CONFIG.get('SAP_SISTEMA'))
    sap.iniciar_sesion_sap()
    
    # 2. Leer XML (Suponiendo que está en tu carpeta de Insumos)
    xml_path = RUTAS['PATH_INSUMO'] + "/ad090063145002525021701C7.xml"
    datos = LectorFacturaXML(xml_path).obtener_datos()
    
    # 3. Buscar OC
    me2l = TransaccionME2L(sap)
    oc = me2l.buscar_oc_activa(datos['nit'])
    
    # 4. Registrar MIGO
    if oc:
        migo = TransaccionMIGO(sap)
        migo.contabilizar_entrada(oc, datos['factura'])

if __name__ == "__main__":
    main()