import os

from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv

from src.Fuji.get_data import GetData
from src.Scrapping.validaciones import Validaciones

class DescargaCorreo:
    #Metodo contructor
    def __init__(self):
        load_dotenv()
        self.CLIENT_ID = None
        self.CLIENT_SECRET = None
        self.TENANT_ID = None
        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_name_template = "plantillaParaActualizacion.xlsx"
        self.download_path = os.path.join(self.path_, 'Documentos')
        self.path_email = os.path.join(self.path_, 'Emails')
        self.db_siga_excel = os.path.join(self.download_path, self.file_name_template)
        self.mensajes = []
    #Metodo para la conexion al correo
    def __conexion_BD(self):
        # definir credenciales de acceso a la cuenta de correo este
        Data = GetData()
        data = Data.get_datos_id('4')
        self.CLIENT_ID = data['client_id']
        self.CLIENT_SECRET = data['secret_id']
        self.TENANT_ID = data['tenant_id']
        
        credentials = (self.CLIENT_ID, self.CLIENT_SECRET)
        token_path = os.path.join(self.path_email, 'o365_token.txt')
        token_backend = FileSystemTokenBackend(token_path=token_path)
        account = Account(credentials, tenant_id=self.TENANT_ID, token_backend=token_backend)

        # Iniciar sesión en la cuenta de correo
        if not account.is_authenticated:  # No almacenamos el token o el token ha expirado
            if account.authenticate(scopes=['basic', 'message_all']):
                print('Autenticado correctamente y token almacenado.')
            else:
                print('Error de autenticación')
        else:
            print('Autenticación exitosa utilizando el token almacenado.')
        return account
    #Medotodo para descargar el adjunto
    def descargarAdjunto(self):
        try:
            datos, mensaje = ['',''], None     
            # Conectar a la cuenta
            account = self.__conexion_BD()
            if not account.is_authenticated:
                return False, "No se pudo autenticar la cuenta."
            # Obtener todos los mensajes no leídos de la carpeta 'Actualizacion Vendedores'
            messages = list(
                account.mailbox()
                    .get_folder(folder_name='Actualizacion Vendedores')
                    .get_messages(
                            query="isRead eq false and subject eq 'ACTUALIZACION VENDEDORES'",
                            download_attachments=True
                    )
            )
            actualizacion = Validaciones()
            if messages:
                # Procesar cada correo no leído
                for message in messages:
                    adjunto_valido_en_correo = False
                    
                    # Extraer el remitente y la fecha, y guardarlos en una tupla
                    datos = (message.sender.address, message.received)
                    print(f"Procesando correo de: {message.sender} en fecha: {message.received}")
                    # Verificar si el correo tiene adjuntos
                    if message.attachments:
                        cont = 0
                        # Verificar si el adjunto es el correcto
                        for attachment in message.attachments:
                            if attachment.name == "plantillaParaActualizacion.xlsx":
                                # Guardar el adjunto en la ruta especificada
                                attachment.save(self.download_path)
                                mensaje = f"Adjunto guardado: {attachment.name}"
                                print(mensaje)
                                adjunto_valido_en_correo = True
                                # Ejecutar la actualización
                                actualizacion.ejecutar_actualizacion(datos, True, mensaje)
                                break
                            else:
                                mensaje = (
                                    f"El nombre del archivo adjunto '{attachment.name}' es INCORRECTO, "
                                    "debe estar nombrado como 'plantillaParaActualizacion.xlsx'."
                                )
                                print(mensaje)
                                self.mensajes.append(mensaje)
                        if cont > 0:
                            mensaje = "<br>ERROR. Hubo un problema con los archivos Adjuntos.<br><br><b>Por favor revisar y volver a enviar.<b><br>"
                            self.mensajes.append(mensaje)
                            actualizacion.ejecutar_actualizacion(datos, False, self.mensajes)
                        if not adjunto_valido_en_correo:
                            mensaje = (
                                f"No se encontró un adjunto válido en el correo de {message.sender} "
                                f"con fecha {message.date}\n"
                            )
                            print(mensaje)
                            actualizacion.ejecutar_actualizacion(datos, False, mensaje)
                    else:
                        mensaje = (
                            f"El correo de {message.sender} con fecha {message.received} NO tiene archivos adjuntos."
                        )
                        print(mensaje)
                        actualizacion.ejecutar_actualizacion(datos, False, mensaje)
                    
                    # Marcar el correo como leído para evitar reprocesarlo
                    message.mark_as_read()
            else:
                mensaje = "En la bandeja de entrada NO hay correos nuevos."
                print(mensaje)
                actualizacion.ejecutar_actualizacion(datos, None, mensaje)
        except Exception as e:
            mensaje = f"ERROR INESPERADO en la descarga de correo: {e}"
            print(mensaje)
            actualizacion.ejecutar_actualizacion(datos, None, mensaje)

