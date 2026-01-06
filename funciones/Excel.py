from config.init_config import in_config
from repositorios.excel import Excel as ExcelRepo
import pandas as pd
import csv
import os
import unicodedata
import re
import warnings

class Excel:

    def normalize_column(col) -> str:
        col = str(col)
        col = col.strip().lower()
        col = unicodedata.normalize("NFKD", col)
        col = col.encode("ascii", "ignore").decode("utf-8")
        col = col.replace(" ", "_").replace(".", "")
        return col

    def limpiar_texto(valor):
        if pd.isna(valor):
            return ""
        valor = str(valor)
        valor = re.sub(r"[\x00-\x1F\x7F]", "", valor)
        valor = valor.replace("\u00A0", " ")
        valor = valor.replace("\n", " ").replace("\r", " ")
        return valor.strip()

    COLUMN_MAP = {
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
        "no_de_contratro": "numero_contrato",
        "no_de_contrato": "numero_contrato",
        "nombre_facturador": "nombre_facturador"
    }

    def excel_a_csv(ruta_excel: str) -> str:
        warnings.filterwarnings(
            "ignore",
            category=UserWarning,
            module="openpyxl"
        )

        # leer excel
        df = pd.read_excel(
            ruta_excel,
            header=3,
            dtype=str,
            engine="openpyxl"
        )

        # normalizar headers
        df.columns = [Excel.normalize_column(c) for c in df.columns]

        # filtrar solo columnas necesarias (las que existan)
        columnas_presentes = {}
        for col in df.columns:
            if col in Excel.COLUMN_MAP:
                columnas_presentes[col] = Excel.COLUMN_MAP[col]

        if not columnas_presentes:
            raise ValueError("No se encontró ninguna columna esperada en el Excel")

        df = df[list(columnas_presentes.keys())]
        df = df.rename(columns=columnas_presentes)

        # asegurar orden final
        orden_final = [
            "cod_fin", "nit", "orden_2025", "mts2", "iva", "tipo",
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
            "observaciones", "numero_contrato", "nombre_facturador"
        ]

        df = df.reindex(columns=orden_final)

        # limpiar contenido
        df = df.map(Excel.limpiar_texto)

        # exportar CSV limpio
        nombre_base= os.path.splitext(os.path.basename(ruta_excel))[0]
        carpeta_temp= in_config("PathTemp")
        ruta_csv = os.path.join(carpeta_temp, f"{nombre_base}.csv")
    
        df.to_csv(
            ruta_csv,
            sep=";",
            index=False,
            encoding="utf-8-sig"
        )

        print(f"CSV generado correctamente en: {ruta_csv}")

        return ruta_csv

    def sanitize_text(value: str) -> str:
        if value is None:
            return "NULL"

        value = str(value)

        # Normalizar Unicode
        value = unicodedata.normalize("NFKC", value)

        # Remover BOM
        value = value.replace("\ufeff", "")

        # Reemplazar espacios invisibles por espacio normal
        value = value.replace("\u00A0", " ")
        value = value.replace("\u2009", " ")
        value = value.replace("\u202F", " ")

        # Eliminar caracteres de control (incluye CR y LF)
        value = re.sub(r"[\x00-\x1F\x7F]", " ", value)

        # Quitar comillas
        value = value.replace('"', "").strip()

        # Si queda vacío → NULL
        return value if value else "NULL"

    def convertirTxt(csv_path: str) -> bool:
        try:
            if not os.path.exists(csv_path):
                return True

            txt_path = os.path.splitext(csv_path)[0] + ".txt"

            if os.path.exists(txt_path):
                os.remove(txt_path)

            with open(csv_path, "r", encoding="latin1", newline="") as csv_file, \
                open(txt_path, "w", encoding="utf-8", newline="\n") as txt_file:

                reader = csv.reader(csv_file)

                for raw_row in reader:
                    # Sanitizar cada campo
                    cleaned_row = [Excel.sanitize_text(field) for field in raw_row]

                    # Si la fila está vacía → ignorar
                    if all(f == "NULL" for f in cleaned_row):
                        continue

                    # Garantizar que no existan saltos de línea dentro de los campos
                    cleaned_row = [f.replace("\r", " ").replace("\n", " ") for f in cleaned_row]

                    # Forzar línea terminada únicamente en \n
                    line = ";".join(cleaned_row)

                    txt_file.write(line + "\n")

            return False

        except Exception as e:
            print("Error durante limpieza:", e)
            return True

    @staticmethod
    def ejecutar_bulk(ruta_excel: str):
        nombre_base= os.path.splitext(os.path.basename(ruta_excel))[0]
        carpeta_temp= in_config("PathTemp")
        ruta_txt = os.path.join(carpeta_temp, f"{nombre_base}.txt")

        if os.path.exists(ruta_txt):
            os.remove(ruta_txt)
        
        try:
            ruta_csv = Excel.excel_a_csv(ruta_excel)
            Excel.convertirTxt(ruta_csv)
            ExcelRepo.ejecutar_bulk(ruta_txt)
        
        except Exception as e:
            print(f"Error: {e}")
            raise
        
        finally:
            if os.path.exists(ruta_txt):
                os.remove(ruta_txt)

            if os.path.exists(ruta_csv):
                os.remove(ruta_csv)