# ===============================
# HU01: Nombre HU
# Autor: Santiago Pinzon - Desarrollador RPA
# Descripcion: Descripcion de la HU 
# Ultima modificacion: 2/1/2026
# Propiedad de Colsubsidio
# Cambios: Si aplica
# ===============================
import os
import time
from funciones.ControlHU import control_hu
from funciones.EscribirLog import WriteLog
from HU.pagoArriendos import ConexionSAP
from config.settings import SAP_CONFIG
from config.init_config import in_config
from HU.ME2L import TransaccionME2L
from funciones.Excel import Excel
from repositorios.excel import Excel as ExcelRepo
from funciones.ME80FN import ME80FN
import pandas as pd


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
        ruta_excel = in_config("PathInsumos")+"\BaseMedicamentos.xlsx"

        try:
            TablaBase = ExcelRepo.obtener_valores("BaseMedicamentos")
        except:
            
            columnas_medicamentos = {
                "cod_fin": "cod_fin",
                "nit": "nit",
                "orden_2025": "orden_2025",
                "mts2_segun_contrato": "mts2",
                "iva": "iva",
                "tipo": "tipo",
                "enero": "enero",
                "febrero": "febrero",
                "marzo": "marzo",
                "abril": "abril",
                "mayo": "mayo",
                "junio": "junio",
                "julio": "julio",
                "agosto": "agosto",
                "septiembre": "septiembre",
                "actubre": "octubre",
                "noviembre": "noviembre",
                "diciembre": "diciembre",
                "observacion_de_pagos": "observaciones",
                "no_de_contrato": "numero_contrato",
                "nombre_facturador": "nombre_facturador"
            }
            
            orden_final = [
            "cod_fin", "nit", "orden_2025", "mts2", "iva", "tipo",
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
            "observaciones", "numero_contrato", "nombre_facturador"
            ]

            COLUMNAS_BASE_MEDICAMENTOS = {
                "CodFin": "VARCHAR(100)",
                "NIT": "VARCHAR(100)",
                "Orden2025": "VARCHAR(MAX)",
                "MTS2": "VARCHAR(100)",
                "IVA": "VARCHAR(100)",
                "Tipo": "VARCHAR(100)",
                "Enero": "VARCHAR(100)",
                "Febrero": "VARCHAR(100)",
                "Marzo": "VARCHAR(100)",
                "Abril": "VARCHAR(100)",
                "Mayo": "VARCHAR(100)",
                "Junio": "VARCHAR(100)",
                "Julio": "VARCHAR(100)",
                "Agosto": "VARCHAR(100)",
                "Septiembre": "VARCHAR(100)",
                "Octubre": "VARCHAR(100)",
                "Noviembre": "VARCHAR(100)",
                "Diciembre": "VARCHAR(100)",
                "Observaciones": "VARCHAR(300)",
                "NumeroContrato": "VARCHAR(300)",
                "NombreFacturador": "VARCHAR(MAX)"
            }

            if ruta_excel:
                Excel.ejecutar_bulk(
                    ruta_excel,
                    "BaseMedicamentos",
                    COLUMNAS_BASE_MEDICAMENTOS,
                    orden_final,
                    columnas_medicamentos,
                    3
                    )

                TablaBase = ExcelRepo.obtener_valores("BaseMedicamentos")

        for registro in TablaBase[6:]:
            sap.abrir_transaccion("ME2L")
            me2l = TransaccionME2L(sap)
            me2l.BuscarPorOC(registro["Orden2025"])

            ruta_archivo = in_config("PathTemp")+"\ComprasNit.xlsx"
            if os.path.exists(ruta_archivo):
                os.remove(ruta_archivo)
            me2l.exportar_tabla(ruta_archivo)
            time.sleep(3)
            os.system("taskkill /f /im excel.exe")
            time.sleep(5)

            """
            Asignamos las columnas y datos para la ejecucion del bulk para la tabla de la transaccion ME2L
            """

            # ==============================
            # CONFIGURACION BULK ME2L
            # ==============================

            df_me2l = pd.read_excel(ruta_archivo, header=None)
            df_me2l = df_me2l.dropna(how="all").reset_index(drop=True)
            df_me2l = df_me2l.rename(columns={
                    0: "proveedor",
                    2: "grupo_de_compras",
                    5: "oc",
                    11: "posicion",
                    12: "material",
                    13: "texto_breve",
                    16: "valor_neto"
                })
            df_me2l.to_excel(ruta_archivo, index=False)

            tabla_me2l = "TablaME2L"
            orden_me2l = [
                "proveedor", "grupo_de_compras", "oc", "posicion", "material",
                "texto_breve", "valor_neto"
            ]
            columnas_me2l= {
                "Proveedor": "VARCHAR(100)",
                "GrupoCompras": "VARCHAR(100)",
                "OC":"VARCHAR(100)",
                "Posicion":"INT",
                "Material": "VARCHAR(100)",
                "TextoBreve":"VARCHAR(100)",
                "ValorNeto":"VARCHAR(100)"
            }

            columnas_map_me2l= {
                "proveedor": "proveedor",
                "grupo_de_compras": "grupo_de_compras",
                "oc": "oc",
                "posicion": "posicion",
                "material": "material",
                "texto_breve": "texto_breve",
                "valor_neto": "valor_neto"
            }

            Excel.ejecutar_bulk(
                ruta_archivo,
                tabla_me2l,
                columnas_me2l, 
                orden_me2l, 
                columnas_map_me2l,
                0
                )


            ruta_cabecera= in_config("PathTemp")+"\cabecera.xlsx"
            ruta_repartos= in_config("PathTemp")+"\Reparto.xlsx"
            ruta_tabla_final=in_config("PathTemp")+"\ValoresAComparar.xlsx"

            sap.MenuPrincipal()
            sap.abrir_transaccion("ME80FN")
            me80fn = ME80FN(sap)
            me80fn.ingresar_oc(registro["Orden2025"])
            time.sleep(2)
            me80fn.exportar_tabla(ruta_cabecera, "cabecera")
            time.sleep(3)
            os.system("taskkill /f /im excel.exe")
            time.sleep(5)
            me80fn.entrar_repartos()
            me80fn.exportar_tabla(ruta_repartos, "repartos")
            time.sleep(3)
            os.system("taskkill /f /im excel.exe")
            time.sleep(5)
            # Operaciones con los excel
            try:
                cabecera = pd.read_excel(ruta_cabecera, header=None)
                reparto = pd.read_excel(ruta_repartos, header=None)

                cabecera= cabecera.dropna(how="all").reset_index(drop=True)
                reparto= reparto.dropna(how="all").reset_index(drop=True)

                cabecera = cabecera.rename(columns={
                    0: "OC",
                    1: "posicion",
                    2: "material",
                    3: "descripcion",
                    10: "proveedor",
                    12: "grupo_de_Compras",
                    14: "valor_neto"
                })

                reparto = reparto.rename(columns={
                    4: "posicion",
                    9: "fecha_entrega"
                })

                cabecera["posicion"] = cabecera["posicion"].astype(str)
                reparto["posicion"] = reparto["posicion"].astype(str)

                cabecera["fecha_entrega"] = cabecera["posicion"].map(
                    reparto.set_index("posicion")["fecha_entrega"]
                )

                cabecera = cabecera[
                        [
                            "OC",
                            "posicion",
                            "material",
                            "descripcion",
                            "proveedor",
                            "grupo_de_Compras",
                            "valor_neto",
                            "fecha_entrega"
                        ]
                    ]

                cabecera.to_excel(ruta_tabla_final,index=False)
                print("Cruce completado")

                # Se ejecuta el bulk en la base de datos
                tabla_me80fn = "TablaME80FN"
                orden_me80fn = [
                    "oc", "posicion", "material", "descripcion", "proveedor",
                    "grupo_de_compras", "valor_neto", "fecha_entrega"
                ]
                columnas_me80fn= {
                    "OC": "VARCHAR(100)",
                    "Posicion": "FLOAT",
                    "Material": "VARCHAR(100)",
                    "Descripcion":"VARCHAR(100)",
                    "Proveedor":"VARCHAR(100)",
                    "GrupoCompras":"VARCHAR(100)",
                    "ValorNeto":"VARCHAR(100)",
                    "FechaEntrega": "VARCHAR(100)"
                }
                columnas_mapeadas= {
                    "oc": "oc",
                    "posicion": "posicion",
                    "material": "material",
                    "descripcion": "descripcion",
                    "proveedor": "proveedor",
                    "grupo_de_compras": "grupo_de_compras",
                    "valor_neto": "valor_neto",
                    "fecha_entrega": "fecha_entrega"
                }

                Excel.ejecutar_bulk(
                    ruta_tabla_final,
                    tabla_me80fn,
                    columnas_me80fn, 
                    orden_me80fn, 
                    columnas_mapeadas,
                    0
                    )
                
            except Exception as e:
                print("Error en combinar columnas", e)

            

            DatosME80FN = ExcelRepo.obtener_datos_por_posicion(tabla_me80fn)
            DatosME2L = ExcelRepo.obtener_datos_por_posicion(tabla_me2l)

            for d, datome2l in zip(DatosME80FN, DatosME2L):
                print("ValorNeto ME80FN:", d["ValorNeto"])
                print("ValorNeto ME2L:", datome2l["ValorNeto"])

                if d["ValorNeto"] == datome2l["ValorNeto"]:
                    print("Coinciden")
                else:
                    print("No coinciden")

            # Finaliza proceso de operaciones
            if os.path.exists(ruta_cabecera) and os.path.exists(ruta_repartos) and os.path.exists(ruta_archivo) and os.path.exists(ruta_tabla_final) :              
                os.remove(ruta_cabecera)
                os.remove(ruta_repartos)
                os.remove(ruta_archivo)
                os.remove(ruta_tabla_final)
                print("Archivos temporales eliminados")

            print("Finalizacion del proceso el registro con oc:", registro["Orden2025"])

            for i in range(2):
                sap.MenuPrincipal()

        # ============================= Finalizacion HU =============================

        control_hu(task_name, 100)

        Estado = 100
        return Estado
    
    except Exception as e:
        print(f"Error en ejecucion: ({e}) ")
        # WriteLog()
        # GestionTicketInsumo(id, observaciones, estado, maquina)
        control_hu(task_name, 99)
        # Estado = 99
        # return Estado
        raise

    finally:
        # WriteLog()
        log = "Finalizacion HU"
        print(log)
        