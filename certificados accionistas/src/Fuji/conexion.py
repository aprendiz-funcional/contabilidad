import pyodbc, os

from dotenv import load_dotenv

class Conexion:
    #Metos constructor
    def __init__(self):
        # Cargar las variables de entorno desde el archivo .env
        load_dotenv()
        # Obtener las credenciales de las variables de entorno
        self.server = os.getenv('SERVER')
        self.database = os.getenv('DATABASE')
    #Metodo para conectar a la BD
    def conexion(self):
        # Construir cadena de conexión
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};Trusted_Connection=yes;'
        try:
            conn = pyodbc.connect(connection_string)
            print("Conexión exitosa a la base de datos.")
            return conn
        except pyodbc.Error as e:
            print("Error al conectar a la base de datos:", e)
            return None