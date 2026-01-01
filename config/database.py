import os
import pyodbc
import logging

logger = logging.getLogger(__name__)

class Database:
    """Gestión básica de conexión a SQL Server"""

    @staticmethod
    def get_connection():
        """
        Abre conexión bajo demanda.
        El cierre se maneja con 'with'.
        """
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={os.getenv('DB_SERVER')};"
                f"DATABASE={os.getenv('DB_NAME')};"
                f"UID={os.getenv('DB_USER')};"
                f"PWD={os.getenv('DB_PASSWORD')};"
                "TrustServerCertificate=yes;"
            )
            return conn

        except Exception:
            logger.error("Error conectando a SQL Server", exc_info=True)
            raise
