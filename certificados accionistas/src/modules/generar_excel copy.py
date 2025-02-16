import os, openpyxl
import pandas as pd
from src.modules.generar_pdf import GenerarCertificado

from src.modules.leer_excel import LeerExcel

class GenerarCertificadoExcel(LeerExcel):
    #Constructor de la clase
    def __init__(self):
        self.generarCertificado = GenerarCertificado()
        __path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path_documentos = os.path.join(__path_, "Documentos")
        self.path_plantillas = os.path.join(self.path_documentos, "Plantillas")       
        self.path_plantilla_excel = os.path.join(self.path_plantillas, "plantilla.xlsx")
        self.path_plantilla_excel2 = os.path.join(self.path_plantillas, "plantilla2.xlsx")

        self.path_procesados = os.path.join(__path_, "Procesados")
        self.path_certificados = os.path.join(self.path_procesados, "Certificados")

        self.path_plantilla_general = os.path.join(self.path_documentos, "PLANTILLA ACCIONISTAS GENERAL.xlsx")
    #Metodo principal para generar los certificados
    def main(self):
        wb_socios, mensaje = super().leer_datos_excel(self.path_plantilla_general, 'BASE CERTIFICADOS')
        if wb_socios is None:
            return mensaje
        print(wb_socios)
        
        wb_socios = self.__eliminar_filas_excel(wb_socios)
        print(wb_socios)
        if not wb_socios:
            return "No hay datos válidos para procesar"
        
        for wb_socio in wb_socios:
            identificacion = wb_socio['identificación accionista']
            nombre = wb_socio['Nombre accionista']
            nombre_certificado = f"{identificacion}_{nombre}"
            sw = self.__generar_certificado_excel_socio(wb_socio, nombre_certificado)
            
            if not sw:
                return "Error al generar certificado"
        
        return "Certificados generados correctamente"
    #Metodo para leer los datos del excel
    def __eliminar_filas_excel(self, datos):
        print("Encabezados:", [key for key in datos[0].keys()])
        return [fila for fila in datos if str(fila.get('identificación accionista', '')).strip() not in ('', 'N/A')]
    # Metodo para generar el certificado de un socio
    def __generar_certificado_excel_socio(self, wb_socio, nombre_cerificado):
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
            ws["B17"] = wb_socio['cuotas 6 meses']
            ws["B19"] = wb_socio['cuotas 3 meses']
            ws["B21"] = wb_socio['cuaota proximo año']
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

    def __eliminar_filas_excel1(self, df_socios):
        # Eliminar las primeras 4 filas y usar la fila 5 como encabezado
        df_socios = df_socios.iloc[3:].reset_index(drop=True)
        df_socios.columns = df_socios.iloc[0]  # Asignar la fila 5 como encabezado
        df_socios = df_socios.iloc[1:]  # Eliminar la fila duplicada en los datos
        return df_socios
    
    def __generar_certificado_excel_socio1(self,df_socio):
        try:
            identificacion = df_socio[' identificación accionista'] #b8
            nombre = df_socio['Nombre accionista'] #b9
            cantidad_acciones = df_socio[' cantidad de acciones'] #b10   
            porcentaje_participacion = df_socio['porcentaje de participación'] #b13
            retefuente = df_socio['retefuente'] #b23
            rete_ica = df_socio['rete ICA '] #b24
            cuotas_annio_anterior = df_socio[' cuotas año anterior'] #b15
            cuotas_6_meses = df_socio['cuotas 6 meses'] #b17
            cuota_3_meses = df_socio['cuotas 3 meses'] #b19
            cuotas_proximo_annio = df_socio['cuaota proximo año'] #b21
            

            df_plantilla, mensaje = self.leer_datos_excel(self.path_plantilla_excel,"plantilla")



            print(df_socio)
            return True
            #self.certificado.generarCertificado(df_socio)
        except Exception as e:
            print(f"ERROR al generar certificado de socio {df_socio['IDENTIFICACION']}<br>")
            return False
        
    def __generar_certificado_excel_socio1(self, df_socio):
        try:
            # Leer la plantilla
            df_plantilla, mensaje = self.leer_datos_excel(self.path_plantilla_excel, "plantilla")
            
            # Mapear los valores de df_socio a df_plantilla en las celdas correspondientes
            df_plantilla.at['B8', 'Valor'] = df_socio['identificación accionista']
            df_plantilla.at['B9', 'Valor'] = df_socio['Nombre accionista']
            df_plantilla.at['B10', 'Valor'] = df_socio['cantidad de acciones']
            df_plantilla.at['B13', 'Valor'] = df_socio['porcentaje de participación']
            df_plantilla.at['B15', 'Valor'] = df_socio['cuotas año anterior']
            df_plantilla.at['B17', 'Valor'] = df_socio['cuotas 6 meses']
            df_plantilla.at['B19', 'Valor'] = df_socio['cuotas 3']
            df_plantilla.at['B21', 'Valor'] = df_socio['cuaota proximo año']
            #df_plantilla.at['B23', 'Valor'] = df_socio['retefuente']
            df_plantilla.at['B24', 'Valor'] = df_socio['rete ICA ']
            
            # Guardar los cambios en el mismo archivo Excel
            with pd.ExcelWriter(self.path_plantilla_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_plantilla.to_excel(writer, sheet_name="plantilla", index=False)
            
            print("Certificado generado correctamente.")
            return True
        except Exception as e:
            print(f"ERROR al generar certificado de socio {df_socio['identificación accionista']}: {str(e)}")
            return False

    
    def __generar_certificado_excel_socio1(self, df_socio, nombre_cerificado):
        try:
            # Cargar el archivo de Excel y seleccionar la hoja "plantilla"
            wb = openpyxl.load_workbook(self.path_plantilla_excel)
            ws = wb["plantilla"]
            
            path_excel = os.path.join(self.path_certificados,"Certificados_excel", f"{nombre_cerificado}.xlsx")

            # Mapear los valores en las celdas correspondientes
            ws["B8"] = df_socio['identificación accionista']
            ws["B9"] = df_socio['Nombre accionista']
            ws["B10"] = df_socio['cantidad de acciones']
            ws["B13"] = df_socio['porcentaje de participación']
            ws["B15"] = df_socio['cuotas año anterior']
            ws["B17"] = df_socio['cuotas 6 meses']
            ws["B19"] = df_socio['cuotas 3 meses']
            ws["B21"] = df_socio['cuaota proximo año']
            ws["B23"] = df_socio['retefuente']  # Descomentarlo si es necesario
            ws["B24"] = df_socio['rete ICA']

            # Guardar los cambios en el mismo archivo
            wb.save(path_excel)
            wb.close()

            print("Certificado generado correctamente.")
            return True
        except Exception as e:
            print(f"ERROR al generar certificado de socio {df_socio['identificación accionista']}: {str(e)}")
            return False

    def __leer_datos_excel(self, ruta, nomhoja):
        try:
            #file_path = os.path.join(self.path_, ruta)
            # Validar existencia del archivo
            if not os.path.exists(ruta):
                mensaje = f"El archivo '{ruta}' no existe"
                print(mensaje)
                return None, mensaje
            df_socios = pd.DataFrame()
            #Se lee los datos del excel
            with open(ruta, mode='rb') as fp:
                df_socios = pd.read_excel(fp, sheet_name=nomhoja, engine='openpyxl', dtype=str)
            mensaje = f"Datos de excel leidos con exito en el paso"
            print(mensaje)
            #print(df_socios)
            return df_socios, mensaje
        except Exception as e:
            mensaje = f"ERROR al leer los datos de excel<br>"
            print(f"{mensaje} {e}")
            return None, mensaje
        
    def __generar(self):
        df_socios, mensaje = self.leer_datos_excel(self.path_plantilla_general, 'BASE CERTIFICADOS')
        if df_socios is None:
            return mensaje
        df_socios = self.__eliminar_filas_excel(df_socios)
        
        if df_socios is None:
            return mensaje
        #df_socios.to_excel('socios2.xlsx', index=False)
        #print(df_socios)
        for _, df_socio in df_socios.iterrows():
            #print(type(df_socio))
            #print(df_socio)
            identificacion = df_socio['identificación accionista']
            nombre = df_socio['Nombre accionista']
            nombre_certificado = f"{identificacion}_{nombre}"
            sw = self.__generar_certificado_excel_socio(df_socio, nombre_certificado)
            if not sw:
                return "Error al generar certificado"
            
            self.generarCertificado.generarCertificado(nombre_certificado)
            
            #self.certificado.generarCertificado(df_socio)
    

