import os
import warnings
import mysql.connector
import pandas as pd
from dotenv import load_dotenv

load_dotenv()
warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")


class MySQLConnector:

    def __init__(self, host=None, port=None, user=None, password=None, database=None):
        self.config = {
            "host": host or os.getenv("DB_HOST", "localhost"),
            "port": int(port or os.getenv("DB_PORT", 3306)),
            "user": user or os.getenv("DB_USER", "root"),
            "password": password or os.getenv("DB_PASSWORD", ""),
            "database": database or os.getenv("DB_NAME", "mineria_datos"),
        }

    def _get_connection(self):
        return mysql.connector.connect(**self.config)

    def extract_data(self, query: str) -> pd.DataFrame:
        conn = self._get_connection()
        try:
            df = pd.read_sql(query, conn)
        finally:
            conn.close()
        return df
