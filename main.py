import os
import pdfplumber
import pandas as pd
import re

# Carpeta donde están los PDFs
carpeta_facturas = "facturas"

# Lista para almacenar los datos extraídos
datos = []

# Función para extraer datos del PDF
def extraer_datos(ruta_pdf):
    num_autorizacion = "No encontrado"
    num_factura = "No encontrado"
    total_retenido = 0.0

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables()
            if tablas:
                for tabla in tablas:
                    for fila in tabla:
                        if fila and fila[0]:  # Asegurar que la fila y su primer elemento no sean None
                            # Buscar el Número de Factura
                            if "FACTURA" in str(fila[0]).upper() and num_factura == "No encontrado":
                                num_factura = re.sub(r"\s+", "", str(fila[1]))  # Limpiar espacios y saltos de línea

                        # Sumar total retenido (última columna)
                        try:
                            if fila and fila[-1]:  # Asegurar que la fila y su última columna no sean None
                                valor_columna = fila[-1].strip()
                                valores = re.findall(r"\d+\.\d+", valor_columna)
                                if valores:
                                    total_retenido += float(valores[-1])
                        except:
                            continue

            # Buscar el Número de Autorización en el texto (una sola vez)
            if num_autorizacion == "No encontrado":
                texto = page.extract_text()
                if texto:
                    match = re.search(r"NÚMERO DE AUTORIZACIÓN[:\s]*([\d]+)", texto)
                    if match:
                        num_autorizacion = match.group(1)

            # Si ya encontramos todo, salimos del bucle
            if num_autorizacion != "No encontrado" and num_factura != "No encontrado":
                break

    return num_autorizacion, num_factura, round(total_retenido, 2)

# Recorrer la carpeta de facturas
for archivo in os.listdir(carpeta_facturas):
    if archivo.lower().endswith('.pdf'):
        ruta_completa = os.path.join(carpeta_facturas, archivo)
        num_aut, num_fact, total_retenido = extraer_datos(ruta_completa)
        datos.append([archivo, num_aut, num_fact, total_retenido])  

# Crear un DataFrame de Pandas
df = pd.DataFrame(datos, columns=["Archivo", "Número de Autorización", "Número de Factura", "Total Retenido"])

# Guardar en un archivo Excel
nombre_excel = "facturas_procesadas.xlsx"
df.to_excel(nombre_excel, index=False)

print(f"\n✅ Proceso finalizado. Archivo Excel generado: {nombre_excel}")
