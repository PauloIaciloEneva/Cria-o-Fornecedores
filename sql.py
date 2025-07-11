import os
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy.engine import URL
from sqlalchemy import create_engine


load_dotenv('.env', override= True)

driver = os.getenv('DRIVER')
server = os.getenv('SERVER')
db = os.getenv('DATABASE')
user = os.getenv('USERNAME')
password = os.getenv('PASSWORD')


class Conection():
    def __init__(self):

        connection_string = f"DRIVER={os.getenv('DRIVER')};SERVER={os.getenv('SERVER')};DATABASE={os.getenv('DATABASE')};UID={os.getenv('USERNAME')};PWD={os.getenv('PASSWORD')};Authentication=ActiveDirectoryPassword;"
        connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
        self.engine = create_engine(connection_url, echo=False, pool_size=10, max_overflow=20, pool_timeout=30, pool_recycle=1800)
        
        self.db_connection = self.engine.connect()
        
    def fetch_data(self, query):
        return pd.read_sql(query, con=self.db_connection)
    
    
    #Filtrar data ZZ_P4IDATE
    
    def LFA1(self):
        query_LFA1 = """
        SELECT 
        LIFNR, 
        NAME1,
        STCD1, 
        STCD2
        FROM [SAP_ECC_TGT].[LFA1]
        WHERE STCD2 IS NOT NULL
        """
        return self.fetch_data(query_LFA1)