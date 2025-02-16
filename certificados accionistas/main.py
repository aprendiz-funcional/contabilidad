import time

from src.modules.modelar_informacion import ModelarInformacion
from src.modules.generar_excel import GenerarCertificadoExcel

hora_inico = time.time()
print(f"Hora Inicio: {hora_inico}")
df_datos = ModelarInformacion().main()

mensaje = GenerarCertificadoExcel().main(df_datos)


tiempo_total = time.time() - hora_inico
horas, resto = divmod(tiempo_total, 3600)
minutos, segundos = divmod(resto, 60)

print(f"✅ Tiempo total: {int(horas)} horas, {int(minutos)} minutos y {segundos:.2f} segundos")


