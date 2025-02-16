import pyodbc

from src.Fuji.conexion import Conexion

class GetData:
    #Metodo contructor
    def __init__(self):
        pass
    #Metodo para consultar la BD
    def get_datos_id(self,id_unico):
        conexion_ = Conexion()
        conn = conexion_.conexion()
        #Se verifica si la conexion fue exitosa
        if conn:
            try:
                cursor = conn.cursor()
                # Llamar al procedimiento almacenado con el parámetro
                cursor.execute("{CALL sp_GetData(?)}", id_unico)
                # Obtener todos los registros
                rows = cursor.fetchall()
                # Obtener los nombres de las columnas
                columns = [column[0] for column in cursor.description]
                # Procesar cada registro sin imprimir datos sensibles
                for row in rows:
                    data = {column: getattr(row, column) for column in columns}
                    # Usa los datos de manera segura en tu aplicación
                    return data
            except pyodbc.Error as e:
                print("Error al obtener datos:", e)
            except Exception as e:
                print("Error al obtener datos:", e)
            finally:
                cursor.close()
                conn.close()
        else:
            print("No se pudo establecer conexión a la base de datos.")

