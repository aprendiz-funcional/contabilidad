import os

from O365 import Account, FileSystemTokenBackend
#from dotenv import load_dotenv

from src.Fuji.get_data import GetData

class DescargaCorreo:
    #Metodo contructor
    def __init__(self):
        self.mensajes = []
        self.CLIENT_ID = None
        self.CLIENT_SECRET = None
        self.TENANT_ID = None

        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_name_template = "plantillaParaActivacion.xlsx"
        self.download_path = os.path.join(self.path_, 'Documentos')
        self.path_email = os.path.join(self.path_, 'Emails')
        self.db_siga_excel = os.path.join(self.download_path, self.file_name_template)
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
            sw = False
            mensaje = ""
            account = self.__conexion_BD()
            # obtener los mensajes de la bandeja
            if account.is_authenticated:
                # Obtener mensajes de la bandeja de entrada con un límite de 5 mensajes no leídos
                messages = list(account.mailbox().get_folder(folder_name='Ingreso Vendedores').get_messages(
                    query="isRead eq false and subject eq 'INGRESO VENDEDORES'",download_attachments=True))
            # Verificar si hay correos no leídos
            if messages:
                # Obtener solo el último mensaje
                last_message = messages[-1]
                # Obtener los adjuntos del mensaje actual
                attachments = last_message.attachments
                #Se verifica si hay adjuntos
                print(f"El correo tiene {len(attachments)} adjunto(s).")
                if attachments:
                    cont = 0
                    # Iterar a través de cada adjunto
                    for attachment in attachments:
                        match attachment.name:
                            case "plantillaParaActivacion.xlsx":
                                # Guardar el adjunto en la ruta especificada
                                attachment.save(self.download_path)
                                # Imprimir el nombre del archivo guardado
                                mensaje = f"Adjunto guardado Correctamente: <b>{attachment.name}.</b><br>"
                                self.mensajes.append(mensaje)
                                print(mensaje)
                                cont +=1
                            case "imagenes.zip":
                                # Guardar el adjunto en la ruta especificada
                                attachment.save(self.download_path)
                                # Imprimir el nombre del archivo guardado
                                mensaje = f"Adjunto guardado Correctamente: <b>{attachment.name}.</b><br>"
                                self.mensajes.append(mensaje)
                                print(mensaje)
                                cont +=1
                            case _:
                                mensaje = f"""
                                    <p>El nombre del archivo adjunto: <u><i>'{attachment.name}'</i></u> es INCORRECTO.<br>
                                    Los nombres correctos son:<br>
                                    <b><i>--plantillaParaActivacion.xlsx Ó imágenes.zip--</i></b><br><br>
                                    </p>
                                    """
                                print(mensaje)
                                self.mensajes.append(mensaje)
                    if cont == 2:
                        self.mensajes = "Archivos adjuntos descargados correctamente"
                        sw = True
                    else:
                        mensaje = "<br>ERROR. Hubo un problema con los archivos Adjuntos.<br><br><b>Por favor revisar y volver a enviar.<b><br>"
                        self.mensajes.append(mensaje)
                        sw = False
                    last_message.mark_as_read()
                else:
                    # Imprimir que el mensaje no tiene adjuntos
                    self.mensajes = "El correo enviado NO tiene archivos adjuntos"
                    print(self.mensajes)
                    last_message.mark_as_read()
                    sw = False
                    # Marcar el mensaje como leído         
            else:
                self.mensajes = "En la bandeja de entrada NO hay correos nuevos"
                print(self.mensajes)
                sw = False
            return sw, self.mensajes
        except Exception as e:
            mensaje = f"ERROR INESPERATADO en la descarga de correo {e}"
            print(mensaje)
            return False, mensaje
