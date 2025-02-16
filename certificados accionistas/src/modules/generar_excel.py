import os, openpyxl
from src.modules.generar_pdf import GenerarCertificado


class GenerarCertificadoExcel():
    #Constructor de la clase
    def __init__(self):
        __path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path_documentos = os.path.join(__path_, "Documentos")
        self.path_plantillas = os.path.join(self.path_documentos, "Plantillas")       
        self.path_plantilla_excel = os.path.join(self.path_plantillas, "plantilla.xlsx")
        self.path_plantilla_excel2 = os.path.join(self.path_plantillas, "plantilla2.xlsx")
        self.mensajes = []
        self.path_procesados = os.path.join(__path_, "Procesados")
        self.path_certificados = os.path.join(self.path_procesados, "Certificados")

        self.certificado = GenerarCertificado()
        self.path_plantilla_general = os.path.join(self.path_documentos, "PLANTILLA ACCIONISTAS GENERAL.xlsx")
    def main(self, df_datos):
        try:
            df_datos.rename(columns={
                "COMPROBANTE": "cantidad de acciones",
                "SECUENCIA": "porcentaje de participación",
                "FUENTE": "cuotas año anterior"
            }, inplace=True)
            # 📌 Verificar que el DataFrame no esté vacío
            if df_datos.empty:
                return "No hay datos válidos para procesar"

            # 📌 Iterar sobre cada accionista en el DataFrame
            for _, df_dato in df_datos.iterrows():
                identificacion = df_dato['TERCERO']
                nombre = df_dato['Nombre Tercero']
                nombre_certificado = f"{identificacion}_{nombre}"
                
                # 📌 Generar certificado
                sw = self.__generar_certificado_excel_socio(df_dato, nombre_certificado)
                
                if sw:
                    self.certificado.generarCertificado(nombre_certificado)
                else:
                    mensaje = f"⚠️ ERROR: al generar certificado para {nombre_certificado}"
                    self.mensajes.append(mensaje)
                    print(mensaje)
            return self.mensajes
        except Exception as e:
            mensaje = f"❌ ERROR en la generación de certificados: {e}"
            print(mensaje)
            return mensaje
    # Metodo para generar el certificado de un socio
    def __generar_certificado_excel_socio1(self, wb_socio, nombre_cerificado):
        try:
            # Cargar el archivo de Excel y seleccionar la hoja "plantilla"
            wb = openpyxl.load_workbook(self.path_plantilla_excel)
            ws = wb["plantilla"]
            
            path_excel = os.path.join(self.path_certificados,"Certificados_excel", f"{nombre_cerificado}.xlsx")

            # Mapear los valores en las celdas correspondientes
            ws["B8"] = wb_socio['identificación accionista']
            ws["B9"] = wb_socio['Nombre accionista']
            ws["B10"] = wb_socio['cantidad de acciones']
            ws["B13"] = wb_socio['porcentaje de participación']
            ws["B15"] = wb_socio['cuotas año anterior']
            ws["B17"] = wb_socio['Suma Enero-Junio']
            ws["B19"] = wb_socio['Suma Julio-Septiembre']
            ws["B21"] = wb_socio['Suma Octubre-Diciembre']
            ws["B23"] = wb_socio['retefuente']  # Descomentarlo si es necesario
            ws["B24"] = wb_socio['rete ICA']

            # Guardar los cambios en el mismo archivo
            wb.save(path_excel)
            wb.close()

            print("Certificado generado correctamente.")
            return True
        except Exception as e:
            print(f"ERROR al generar certificado de socio {wb_socio['identificación accionista']}: {str(e)}")
            return False
    def __generar_certificado_excel_socio(self, socio_row, nombre_certificado):
        try:
            # 📌 Cargar la plantilla de Excel
            wb = openpyxl.load_workbook(self.path_plantilla_excel)
            ws = wb.active
            #ws = wb["plantilla"]
            
            # 📌 Definir la ruta donde se guardará el certificado
            path_excel = os.path.join(self.path_certificados, "Certificados_excel", f"{nombre_certificado}.xlsx")

            # 📌 Asignar valores a las celdas de la plantilla
            ws["B8"] = socio_row.get('TERCERO', 'N/A')
            ws["B9"] = socio_row.get('Nombre Tercero', 'N/A')
            ws["B10"] = socio_row.get('cantidad de acciones', 0)
            ws["B13"] = socio_row.get('porcentaje de participación', 0)
            ws["B15"] = socio_row.get('cuotas año anterior', 0)
            ws["B17"] = socio_row.get('Suma Enero-Junio', 0)
            ws["B19"] = socio_row.get('Suma Julio-Septiembre', 0)
            ws["B21"] = socio_row.get('Suma Octubre-Diciembre', 0)
            ws["B23"] = socio_row.get('retefuente', 0)  
            ws["B24"] = socio_row.get('rete ICA', 0)

            # 📌 Guardar el archivo en la ruta especificada
            wb.save(path_excel)
            wb.close()

            print(f"✅ Certificado generado correctamente: {nombre_certificado}")
            return True

        except Exception as e:
            print(f"❌ ERROR al generar certificado para {socio_row.get('identificación accionista', 'N/A')}: {str(e)}")
            return False
