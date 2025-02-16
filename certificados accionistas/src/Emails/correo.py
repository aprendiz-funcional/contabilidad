import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from src.Fuji.get_data import GetData

class Correo:
    #Metodo constructor
    def __init__(self): 
        Data = GetData()
        data = Data.get_datos_id('4')
        self.SMTP_SERVER = data['server_smtp']
        self.SMTP_PORT = data['port_smtp']
        self.SMTP_USERNAME = data['user_smtp']
        self.SMTP_PASSWORD = data['pass_smtp']
        self.from_email = self.SMTP_USERNAME     
    #Metodo para crear el mensaje
    def crear_mensaje(self, subject, body, to_email):
        # Crear el mensaje
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = ', '.join(to_email)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        return msg
    #Metodo para enviar el mensaje
    def enviar(self, msg, to_email):
        # Iniciar sesión en el servidor SMTP y enviar correo
        try:
            server = smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT)
            server.starttls()
            server.login(self.SMTP_USERNAME, self.SMTP_PASSWORD)
            server.sendmail(self.from_email, to_email, msg.as_string())
            server.quit()
            print("Correo enviado correctamente")
        except Exception as e:
            print("Error al enviar el correo:", str(e))

