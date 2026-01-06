from config.database import Database

class Excel:

    @staticmethod
    def ejecutar_bulk(ruta_txt: str):
        query=f"""
            BULK INSERT PagoArriendos.BaseMedicamentos
            FROM '{ruta_txt}'
            WITH (
                FIRSTROW = 2,
                FIELDTERMINATOR = ';',
                ROWTERMINATOR = '0x0a',
                CODEPAGE = '65001',        
                TABLOCK
            )
            """
        with Database.get_connection() as conn:
            conn.autocommit = True,
            cursor = conn.cursor()
            try:
                cursor.execute(query)
                print("Consulta ejecutada correctamente")
                cursor.commit()
                cursor.close()

            except Exception as e:
                print(f"Error al realizar el BULK: {e}")
                cursor.close()

    @staticmethod
    def obtener_valores():
        query="""
        SELECT TOP 5 * FROM PagoArriendos.BaseMedicamentos
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            column = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

            return [dict(zip(column, fila)) for fila in rows]
            