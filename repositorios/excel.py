from config.database import Database

class Excel:

    @staticmethod
    def CrearTablaBM():
        query = """
        DROP TABLE IF EXISTS PagoArriendos.BaseMedicamentos

        CREATE TABLE PagoArriendos.BaseMedicamentos (
            CodFin              VARCHAR(100) NULL,
            NIT                 VARCHAR(100) NULL,
            Orden2025           VARCHAR(MAX) NULL,
            MTS2                VARCHAR(100) NULL,
            IVA                 VARCHAR(100) NULL,
            Tipo				VARCHAR(100) NULL,
            Enero               VARCHAR(100) NULL,
            Febrero             VARCHAR(100) NULL,
            Marzo               VARCHAR(100) NULL,
            Abril               VARCHAR(100) NULL,
            Mayo                VARCHAR(100) NULL,
            Junio               VARCHAR(100) NULL,
            Julio               VARCHAR(100) NULL,
            Agosto              VARCHAR(100) NULL,
            Septiembre          VARCHAR(100) NULL,
            Octubre             VARCHAR(100) NULL,
            Noviembre           VARCHAR(100) NULL,
            Diciembre            VARCHAR(100) NULL,
            Observaciones       VARCHAR(300) NULL,
            NumeroContrato      VARCHAR(300) NULL,
            NombreFacturador    VARCHAR(MAX) NULL
        );
        """
        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                cursor.close()

                return True
        except Exception as e:
            print(f"Error al crear la tabla: {e}")
            
            return False
        
    @staticmethod
    def ejecutar_bulk(ruta_txt: str):
        
        # Excel.CrearTablaBM()

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
        SELECT TOP 10 * FROM PagoArriendos.BaseMedicamentos
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            column = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

            return [dict(zip(column, fila)) for fila in rows]