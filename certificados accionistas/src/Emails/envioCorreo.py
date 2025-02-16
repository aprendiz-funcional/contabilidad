import os

from email.mime.base import MIMEBase
from email import encoders

from src.Emails.correo import Correo
from dotenv import load_dotenv

class EnvioCorreo(Correo):
    #Metodo constructor
    def __init__(self):
        super().__init__()
        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path_doc = os.path.join(self.path_, 'Documentos')
        load_dotenv()
        self.to_email = os.getenv('to_email')
    #Metodo para obtener el nombre de los archivos para adjuntar
    def __obtener_nombre_archivos(self):
        try:
            # Obtener todos los nombres de los archivos con extensión .xlsx
            nom_archivos_excel = [f for f in os.listdir(self.path_doc) if f.endswith(('.xlsx'))]
            print(nom_archivos_excel)
            return nom_archivos_excel
        except Exception as e:
            print(e)
    #Metodo para adjuntar los archivos
    def __adjuntar_archivos(self,msg, lista_archivos):
        try:
            if lista_archivos:
                for archivo in lista_archivos:
                    #Ruta del archivo
                    file_path =  os.path.join(self.path_doc, archivo)
                    print(file_path)
                    # Verificar si el archivo existe
                    if not os.path.isfile(file_path):
                        print(f"El archivo no existe: {file_path}")
                        continue
                    file_name = os.path.basename(file_path)
                    print(file_name)
                    with open(file_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={file_name}",
                        )
                        msg.attach(part)
                print(f"{len(lista_archivos)} archivos adjuntados correctamente.")
            else:
                print("No hay archivos para adjuntar")
        except Exception as e:
            print(f"Error al adjuntar los archivos: {e}")
    #Metodo para enviar el correo al proceso de comercial
    def enviar_correo(self, correo):
        # Crear el cuerpo del correo
        body = f"""
        <html>
        <body>
            <p>Cordial saludo,<br>
            <br>
            <b>Los vendedores han sido actualizados con éxito. 
            En el archivo adjunto se pueden ver los detalles de los 
            vendedores actualizados.</b>
            <br><br>
            Correo de la persona que solicitó la actualización:<br><div"><b>{correo}</b></div>
            <br>
            Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
            </p>
        </body>
        </html>
        """
        lista_archivos = ['Procesados.xlsx']#self.__obtener_nombre_archivos()
        subject = 'Proceso de actualizacion de vendedores'

        print(correo)
        to_email_final = self.to_email.replace("{correo_vendedor}", correo)
        to_email_final = to_email_final.split(",")
        msg = super().crear_mensaje(subject, body, to_email_final)
        #Adjuntar el archivo al mensaje
        self.__adjuntar_archivos(msg, lista_archivos)
        #Llamar al método de la clase padre para enviar el correo 
        super().enviar(msg, to_email_final) 
    