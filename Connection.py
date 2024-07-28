import pyodbc
from dotenv import load_dotenv
import os
# Load environment variables from the .env file (if present)
load_dotenv()
def connection():
    con = f'DRIVER={os.getenv('SQLSERVER_DIVER')};SERVER={os.getenv('SQLSERVER_HOST')};DATABASE={os.getenv('SQLSERVER_BD')};UID={os.getenv('SQLSERVER_USER')};PWD={os.getenv('SQLSERVER_PASSWORD')}'
    try:
        connection = pyodbc.connect(con)
        return connection
    except Exception as ex:
        print("Error durante la conexión: {}".format(ex))