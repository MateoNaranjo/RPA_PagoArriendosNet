from config.database import Database

class ExcelRepo:

    # -----------------------------
    # CREAR COLUMNAS DINÁMICAS
    # -----------------------------
    @staticmethod
    def _construir_columnas(columnas: list[str]) -> str:
        return ",\n".join(
            f"{col} VARCHAR(MAX) NULL"
            for col in columnas
        )

    # -----------------------------
    # CREAR TABLA TEMPORAL
    # -----------------------------
    @staticmethod
    def crear_tabla_temp(tabla: str, columnas: list[str]) -> bool:

        tabla_temp = f"{tabla}_temp"
        if not columnas:
            raise ValueError("La lista de columnas está vacía")

        columnas_sql = ",\n".join(
            f"[{col}] NVARCHAR(MAX)"
            for col in columnas
        )

        query = f"""
        IF OBJECT_ID('PagoArriendos.{tabla_temp}', 'U') IS NOT NULL
            DROP TABLE PagoArriendos.{tabla_temp};

        CREATE TABLE PagoArriendos.{tabla_temp} (
            {columnas_sql}
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                cursor.close()
                print("Tabla temporal creada con exito")
            return True
        except Exception as e:
            print(f"Error creando tabla temporal {tabla_temp}: {e}")
            return False

    # -----------------------------
    # CREAR TABLA FINAL
    # -----------------------------
    @staticmethod
    def crear_tabla_final(tabla: str, columnas: list[str]) -> bool:

        columnas_sql = ExcelRepo._construir_columnas(columnas)

        query = f"""
        IF OBJECT_ID('PagoArriendos.{tabla}', 'U') IS NOT NULL
            DROP TABLE PagoArriendos.{tabla};

        CREATE TABLE PagoArriendos.{tabla} (
            {columnas_sql},
            EstadoRegistro VARCHAR(50) NOT NULL DEFAULT 'Pendiente'
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                print("Tabla final creada con exito")
                cursor.close()
            return True
        except Exception as e:
            print(f"Error creando tabla final {tabla}: {e}")
            return False

    # -----------------------------
    # BULK A TEMP + TRANSFERENCIA
    # -----------------------------
    @staticmethod
    def ejecutar_bulk_dinamico(
        ruta_txt: str,
        tabla: str,
        columnas: list[str]
    ):

        tabla_temp = f"{tabla}_temp"
        columnas_sql = ", ".join(columnas)

        bulk_query = f"""
        BULK INSERT PagoArriendos.{tabla_temp}
        FROM '{ruta_txt}'
        WITH (
            FIRSTROW = 2,
            FIELDTERMINATOR = ';',
            ROWTERMINATOR = '0x0a',
            CODEPAGE = '65001',
            TABLOCK
        );
        """

        insert_query = f"""
        INSERT INTO PagoArriendos.{tabla} ({columnas_sql}, EstadoRegistro)
        SELECT {columnas_sql}, 'Pendiente'
        FROM PagoArriendos.{tabla_temp};
        """

        drop_query = f"""
        DROP TABLE PagoArriendos.{tabla_temp};
        """

        try:
            with Database.get_connection() as conn:
                conn.autocommit = True
                cursor = conn.cursor()

                if not ExcelRepo.crear_tabla_temp(tabla, columnas):
                    return

                ExcelRepo.crear_tabla_final(tabla, columnas)

                cursor.execute(bulk_query)
                cursor.execute(insert_query)
                cursor.execute(drop_query)

                cursor.close()

            print(f"Carga BULK ejecutada correctamente en {tabla}")

        except Exception as e:
            print(f"Error durante BULK dinámico en {tabla}: {e}")

   
    # -----------------------------
    # OBTENER SOLO PENDIENTES
    # -----------------------------
    @staticmethod
    def obtener_valores(tabla: str):

        query = f"""
        SELECT TOP 30 *
        FROM PagoArriendos.{tabla}
        WHERE EstadoRegistro = 'Pendiente'
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)

            columnas = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            cursor.close()

            return [dict(zip(columnas, fila)) for fila in rows]
        

