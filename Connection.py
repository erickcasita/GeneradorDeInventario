import pyodbc

def connection():
    try:
        connection = pyodbc.connect('DRIVER={SQL Server};SERVER=DBMSATUXTLA;DATABASE=BDSIVE;UID=comersat;PWD=Soporte%%2022')
        return connection
    except Exception as ex:
        print("Error durante la conexión: {}".format(ex))