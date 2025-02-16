import os

from src.Emails.correo import Correo
from dotenv import load_dotenv

class ErrorCorreo(Correo):
    #Metodo constructor
    def __init__(self):
        super().__init__()
        load_dotenv()
        # lista de Correos del proceso comercial a los que se le enviara confirmacion
        self.to_email = os.getenv('to_email')
        # correos del proceso de automatizacion para errores
        self.to_email_b_entrada = os.getenv('to_email_b_entrada').split(',')
    #Metodo para enviar correos de error
    def enviar_error(self,mensaje, correo):
        # Contenido del correo (puede incluir HTML)
        body = f"""
        <html>
        <body>
            <p>Cordial saludo,<br>
            <br>
            <p><b><u>
            Ocurrió un error en el proceso de la actualizacion de  vendedores Indirectos.</u></b></P>
            <br>
            Detalles del error:<br><div">{mensaje}</div>
            <br>
            Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
            </p>
        </body>
        </html>
        """
        print(correo)
        to_email_final = self.to_email.replace("{correo_vendedor}", correo)
        to_email_final = to_email_final.split(",")
        subject = 'Error en el proceso de Actualizacion vendedores Indirectos'
        msg = super().crear_mensaje(subject, body, to_email_final)
        super().enviar(msg, to_email_final)  # Llamar al método de la clase padre
    #Metodo para enviar correos de error
    def enviar_error_b_entrada(self,mensaje):
        # Contenido del correo (puede incluir HTML)
        body = f"""
        <html>
        <body>
            <p>Cordial saludo,<br>
            <br>
            <p><b><u>
            Ocurrió un error en el proceso de la actualizacion de  vendedores Indiretos.</u></b></P><br>
            <br>
            Detalles del error:<br><div">{mensaje}</div>
            <br>
            Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
            </p>
        </body>
        </html>
        """
        subject = 'Error en el proceso de Actualizacion vendedores Indirectos'
        msg = super().crear_mensaje(subject, body, self.to_email_b_entrada)
        super().enviar(msg, self.to_email_b_entrada)  # Llamar al método de la clase padre