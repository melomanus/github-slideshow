# ------------------------------------ #
# Configuración y Procesamiento de Datos SIAF
# Autor: Alonso Sarmiento
# Fecha: 24/12/2024
# Descripción:
# Este script procesa archivos de datos del SIAF generados dinámicamente,
# elimina columnas innecesarias, combina múltiples bases y exporta
# el resultado en un archivo delimitado por '|' para análisis posterior.
# ------------------------------------ #

import os
import pandas as pd
from datetime import datetime

# ------------------------------------ #
# 1. Configuración de Directorios y Variables
# ------------------------------------ #

# Ruta principal del proyecto
main = "F:/Herramienta_2024"

# Subdirectorios para organización
Input = os.path.join(main, "1. Input")  # Directorio de entrada
Mean = os.path.join(main, "2. Mean")   # Directorio de resultados intermedios
Output = os.path.join(main, "3. Output") # Directorio de resultados finales
Script = os.path.join(main, "4. Script") # Directorio de scripts
Cross = os.path.join(main, "5. Cross")  # Directorio de cruces

# Fecha dinámica en formato DDMMYYYY (Ejemplo: 24122024)
Fecha_SIAF = datetime.now().strftime("%d%m%Y")

# Directorio donde se encuentran los archivos de "Gastos Gobierno"
Sistema = os.path.join(main, "Gastos gobierno")

# ------------------------------------ #
# 2. Instalación y Carga de Paquetes
# ------------------------------------ #

# Librerías necesarias están instaladas en el entorno Spyder.
# ------------------------------------ #
# 3. Variables Dinámicas para el Proceso
# ------------------------------------ #

anio_actual = datetime.now().strftime("%Y")  # Año actual dinámico
bases = ["GN", "GR", "GL"]  # Tipos de bases a procesar
archivos_temporales = {}  # Diccionario para almacenar datos temporalmente

# ------------------------------------ #
# 4. Procesamiento de Archivos
# ------------------------------------ #

for i, base in enumerate(bases, start=1):  # Iterar sobre cada base (GN, GR, GL)
    # Construir el nombre del archivo de entrada
    archivo_entrada = os.path.join(
        Sistema,
        f"{i}.PIAPIMDevGirxMetaEsp_{anio_actual}_{base}_{Fecha_SIAF}.xlsx"
    )

    # Verificar si el archivo existe antes de procesarlo
    if os.path.exists(archivo_entrada):
        # Leer los datos del archivo Excel
        datos = pd.read_excel(
            archivo_entrada,
            engine='openpyxl'  # Usar el motor de lectura adecuado
        )

        # Eliminar columnas innecesarias
        columnas_a_eliminar = ["DES_UE_GRUPO", "GRUPO_BS", "IND_AHORRO", "CLASE_GTO"]
        datos = datos.drop(columns=[col for col in columnas_a_eliminar if col in datos.columns])

        # Almacenar los datos procesados en el diccionario temporal
        archivos_temporales[base] = datos
    else:
        # Generar una advertencia si el archivo no existe
        print(f"Advertencia: Archivo no encontrado: {archivo_entrada}")

# ------------------------------------ #
# 5. Combinación de Bases y Exportación
# ------------------------------------ #

# Verificar si hay datos en el diccionario temporal
if archivos_temporales:
    # Combinar todas las bases en una sola tabla
    datos_combinados = pd.concat(archivos_temporales.values(), ignore_index=True)

    # Nombre del archivo de salida basado en el año actual
    archivo_salida = f"SIAF_{anio_actual}.txt"

    # Exportar los datos combinados en un archivo delimitado por '|'
    datos_combinados.to_csv(
        os.path.join(Mean, archivo_salida),
        sep='|',
        index=False,
        encoding='utf-8'
    )

    # Mensaje informando éxito en la exportación
    print("Archivo combinado y exportado exitosamente.")
else:
    # Advertencia si no hay datos para combinar
    print("Advertencia: No se encontraron datos para combinar.")

#######################################
datos_inicio = datos_combinados.copy()
datos_combinados = datos_inicio.copy()
#######################################


# ------------------------------------ #
# 6. Modificación de la Columna ETIQUETA_LINEA en datos_combinados
# ------------------------------------ #

# Descripción:
# Esta sección crea y modifica la columna `ETIQUETA_LINEA` en el dataframe `datos_combinados`,
# asignando valores predeterminados y realizando cambios condicionales basados
# en el contenido de la columna `ACT_OBRA_ACCINV`.

# Asignar valor por defecto "1. GASTO RECURRENTE" a toda la columna
datos_combinados["ETIQUETA_LINEA"] = "1. GASTO RECURRENTE"

# Modificar la columna ETIQUETA_LINEA para registros que cumplan la condición
# Si ACT_OBRA_ACCINV contiene el valor "5006373", se asigna "3. GASTO REACTIVACIÓN"
datos_combinados.loc[datos_combinados["ACT_OBRA_ACCINV"].str.contains("5006373", na=False), "ETIQUETA_LINEA"] = "3. GASTO REACTIVACIÓN"

# Mensaje de éxito
print("Modificación de la columna ETIQUETA_LINEA en datos_combinados completada.")

# ------------------------------------ #
# 7. Corrección de Variables
# ------------------------------------ #

# Descripción:
# Esta sección realiza la corrección de nombres de columnas en el dataframe `datos_combinados`,
# renombrando y eliminando columnas según sea necesario.

# Eliminar columnas innecesarias
datos_combinados = datos_combinados.drop(columns=["ANO_EJE"])

# Renombrar columnas según las especificaciones
datos_combinados = datos_combinados.rename(columns={
    "NIVEL_GOBIERNO": "NIVEL_GOB",
    "CATEGORIA": "CATEGORIA_GTO",
    "PROGRAMA_PPTAL": "PROGRAMA_PPTAL"
})

# Eliminar filas donde `NIVEL_GOB` y `PLIEGO` estén vacías
datos_combinados = datos_combinados.dropna(subset=["NIVEL_GOB", "PLIEGO"], how="all")

# Mensaje de éxito
print("Corrección de variables completada.")

# ------------------------------------ #
# 8. Limpieza Adicional y Creación de Variables
# ------------------------------------ #

# Crear la columna TIPO_TRANSACCION con el valor "2. GASTOS PRESUPUESTARIOS"
datos_combinados["TIPO_TRANSACCION"] = "2. GASTOS PRESUPUESTARIOS"

# Crear columnas temporales I y F con valores vacíos
datos_combinados["I"] = ""
datos_combinados["F"] = ""

# Reordenar las columnas: colocar I al inicio y F al final
cols = ["I"] + [col for col in datos_combinados if col not in ["I", "F"]] + ["F"]
datos_combinados = datos_combinados[cols]

# Recortar espacios en blanco de las columnas I y F
datos_combinados["I"] = datos_combinados["I"].str.strip()
datos_combinados["F"] = datos_combinados["F"].str.strip()

# Eliminar las columnas temporales I y F
datos_combinados = datos_combinados.drop(columns=["I", "F"])

# Corregir valores en la columna EJECUTORA
# Eliminar el carácter "\" si está presente en la columna EJECUTORA
datos_combinados["EJECUTORA"] = datos_combinados["EJECUTORA"].str.replace(r"\\", "", regex=True)

# Corregir valores en la columna ACT_OBRA_ACCINV
# Eliminar el carácter "\" si está presente en la columna ACT_OBRA_ACCINV
datos_combinados["ACT_OBRA_ACCINV"] = datos_combinados["ACT_OBRA_ACCINV"].str.replace(r"\\", "", regex=True)

# ------------------------------------ #
# Finalización de la Limpieza
# ------------------------------------ #

# Mensaje informando que la limpieza adicional ha sido completada
print("Limpieza adicional y correcciones en datos_combinados completadas.")

# ------------------------------------ #
# 9. Manipulación de Variables y Codificación
# ------------------------------------ #

# Lista de columnas a procesar
m = ["FUNCION", "ACT_OBRA_ACCINV", "DIVISION_FUNCIONAL", "GRUPO_FUNCIONAL", 
     "PRODUCTO_PROYECTO", "FINALIDAD", "PROGRAMA_PPTAL"]

# Cambiar el tipo de las columnas a cadenas de texto largas
datos_combinados[m] = datos_combinados[m].astype(str)

# Crear códigos para las columnas en la lista `m`
for col in m:
    # Crear nueva columna COD_<col> basada en los primeros caracteres antes del "."
    datos_combinados[f"COD_{col}"] = datos_combinados[col].str.split(".").str[0].str.replace(".", "", regex=True)
    
    # Eliminar la columna original
    datos_combinados = datos_combinados.drop(columns=[col])

# Respaldar el dataframe
backup = datos_combinados.copy()
datos_combinados = backup.copy()

# ------------------------------------ #
# 10. Integración con Estructura Maestra
# ------------------------------------ #

# Ruta al archivo maestro
# Define la ruta al archivo maestro que contiene las estructuras necesarias.
estructura_maestro = os.path.join(Cross, "Estructura_SIAF_Maestro2024.xlsx")

# Iterar sobre las columnas en `m`
for col in m:
    # Leer la hoja correspondiente del archivo maestro
    # Se espera que cada hoja contenga datos relacionados con la columna en proceso.
    hoja_maestro = pd.read_excel(estructura_maestro, sheet_name=col, dtype=str)

    # Eliminar columna ESTRUCTURA si existe
    # Esto es necesario para evitar redundancias en los datos combinados.
    if "ESTRUCTURA" in hoja_maestro.columns:
        hoja_maestro = hoja_maestro.drop(columns=["ESTRUCTURA"])

    # Identificar el nombre correcto de la columna clave en el archivo maestro
    # Busca automáticamente una columna que contenga el prefijo "COD_" para utilizarla como clave.
    clave_y = [c for c in hoja_maestro.columns if c.startswith("COD_")]

    # Validar que la columna clave exista en la hoja maestro
    # Si no se encuentra una clave válida, detiene la ejecución con un mensaje de error.
    if len(clave_y) == 0:
        raise ValueError(f"No se encontró una columna clave para la hoja: {col}")

    # Realizar el merge con la estructura maestra
    # Combina los datos utilizando la columna generada ("COD_<col>") y la clave del maestro.
    datos_combinados = pd.merge(
        datos_combinados,
        hoja_maestro,
        left_on=f"COD_{col}",
        right_on=clave_y[0],
        how="left"
    )

    # Eliminar la columna de código después del merge
    # Esto ayuda a mantener el dataframe limpio y evita columnas redundantes.
    datos_combinados = datos_combinados.drop(columns=[f"COD_{col}"])


# ------------------------------------ #
# 11. Reordenamiento de Columnas
# ------------------------------------ #

# Ordenar las columnas por FUENTE, precediendo a SEC_EJEC
datos_combinados = datos_combinados[["FUENTE", "SEC_EJEC"] + [col for col in datos_combinados if col not in ["FUENTE", "SEC_EJEC"]]]

# Iterar sobre las columnas originales y reordenarlas precediendo a FUENTE
for col in m:
    if col in datos_combinados.columns:
        cols = [col] + ["FUENTE"] + [c for c in datos_combinados if c not in [col, "FUENTE"]]
        datos_combinados = datos_combinados[cols]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Integración y reordenamiento completados.")

# ------------------------------------ #
# 12. Unión con Lista Histórica de UE
# ------------------------------------ #

# Convertir la columna SEC_EJEC a numérica (equivalente a `destring` en Stata)
datos_combinados["SEC_EJEC"] = pd.to_numeric(datos_combinados["SEC_EJEC"], errors="coerce")

# Eliminar columnas innecesarias, utilizando los nombres correctos
columnas_a_eliminar = ["NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA"]
datos_combinados = datos_combinados.drop(columns=[col for col in columnas_a_eliminar if col in datos_combinados.columns])

# Ruta al archivo maestro de Lista Histórica
lista_ue_hist = os.path.join(Cross, "Lista_UE_Hist_2014-2024.xlsx")

# Realizar un merge (equivalente a `merge m:1 SEC_EJEC`)
lista_hist_datos = pd.read_excel(lista_ue_hist, dtype=str)

# Convertir SEC_EJEC a numérica en el archivo maestro
lista_hist_datos["SEC_EJEC"] = pd.to_numeric(lista_hist_datos["SEC_EJEC"], errors="coerce")

# Realizar la unión
datos_combinados = pd.merge(
    datos_combinados,
    lista_hist_datos,
    on="SEC_EJEC",
    how="left"
)

# Eliminar filas que no encontraron coincidencia (_merge == 2 en Stata)
datos_combinados = datos_combinados.dropna(subset=["NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA"])

# Reordenar las columnas para que NIVEL_GOB, SECTOR, PLIEGO y EJECUTORA precedan a SEC_EJEC
cols = ["NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA", "SEC_EJEC"] + [c for c in datos_combinados if c not in ["NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA", "SEC_EJEC"]]
datos_combinados = datos_combinados[cols]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Unión con lista histórica de UE completada y columnas reordenadas.")

# ------------------------------------ #
# 13. Creación de la Columna CLASIFICADOR_GASTO
# ------------------------------------ #

# Crear la columna CLASIFICADOR_GASTO combinando las partes numéricas de cada columna
datos_combinados["CLASIFICADOR_GASTO"] = datos_combinados.apply(
    lambda row: ".".join([
        row["CATEGORIA_GTO"].split(".")[0] if pd.notna(row["CATEGORIA_GTO"]) else "",
        row["TIPO_TRANSACCION"].split(".")[0] if pd.notna(row["TIPO_TRANSACCION"]) else "",
        row["GENERICA"].split(".")[0] if pd.notna(row["GENERICA"]) else "",
        row["SUBGENERICA"].split(".")[0] if pd.notna(row["SUBGENERICA"]) else "",
        row["SUBGENERICA_DET"].split(".")[0] if pd.notna(row["SUBGENERICA_DET"]) else "",
        row["ESPECIFICA"].split(".")[0] if pd.notna(row["ESPECIFICA"]) else "",
        row["ESPECIFICA_DET"].split(".")[0] if pd.notna(row["ESPECIFICA_DET"]) else ""
    ]), axis=1
)

# Limpiar posibles espacios o valores incompletos
datos_combinados["CLASIFICADOR_GASTO"] = datos_combinados["CLASIFICADOR_GASTO"].str.strip()
datos_combinados = datos_combinados[(datos_combinados["CLASIFICADOR_GASTO"].notna()) & (datos_combinados["CLASIFICADOR_GASTO"] != "")]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #
print("Columna CLASIFICADOR_GASTO creada correctamente.")

# ------------------------------------ #
# 14. Reordenamiento de Columnas
# ------------------------------------ #

# Reordenar la columna CLASIFICADOR_GASTO para que preceda a ESPECIFICA_DET
cols = ["CLASIFICADOR_GASTO"] + [c for c in datos_combinados if c != "CLASIFICADOR_GASTO"]
datos_combinados = datos_combinados[cols]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Columna CLASIFICADOR_GASTO creada y columnas reordenadas.")

#######################################
backup1_1 = datos_combinados.copy()  # Crear un respaldo del dataframe original
datos_combinados = backup1_1.copy()  # Restaurar el dataframe en caso de problemas
#######################################


# ------------------------------------ #
# 15. Integración con Estructura Maestra para CLASIFICADOR_GASTO
# ------------------------------------ #


# Eliminar las columnas conflictivas en datos_combinados
datos_combinados = datos_combinados.drop(columns=[
    "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", 
    "SUBGENERICA", "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET"
], errors="ignore")

# Cargar la hoja "CLASIFICADOR_GASTO" del archivo maestro
estructura_clasificador = pd.read_excel(
    os.path.join(Cross, "Estructura_SIAF_Maestro2024.xlsx"),
    sheet_name="CLASIFICADOR_GASTO",
    dtype=str
)

# Verificar y limpiar datos del clasificador
estructura_clasificador = estructura_clasificador.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

# Realizar un merge (similar a `merge m:1 CLASIFICADOR_GASTO` en Stata)
datos_combinados = datos_combinados.merge(
    estructura_clasificador,
    on="CLASIFICADOR_GASTO",
    how="left"
)

# Verificar si hay columnas con sufijo `.y`
columnas_maestro = [col for col in datos_combinados.columns if col.endswith(".y")]

# Si hay columnas con sufijo `.y`, filtrar filas sin coincidencias
if columnas_maestro:
    datos_combinados = datos_combinados[datos_combinados[columnas_maestro].notna().any(axis=1)]

# Si hay columnas `.y`, renombrarlas y eliminar las `.x`
columnas_x = [col for col in datos_combinados.columns if col.endswith(".x")]
columnas_y = [col for col in datos_combinados.columns if col.endswith(".y")]

if columnas_y:
    nuevos_nombres = [col.replace(".y", "") for col in columnas_y]
    datos_combinados.rename(columns=dict(zip(columnas_y, nuevos_nombres)), inplace=True)
    datos_combinados.drop(columns=columnas_x, inplace=True)


# ------------------------------------ #
# 16. Reordenamiento de Columnas
# ------------------------------------ #

# Reordenar las columnas para que las eliminadas precedan a `CLASIFICADOR_GASTO`
columnas_reordenadas = [
    "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", "SUBGENERICA",
    "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET", "CLASIFICADOR_GASTO"
]
columnas_existentes = [col for col in columnas_reordenadas if col in datos_combinados.columns]
datos_combinados = datos_combinados[columnas_existentes + [col for col in datos_combinados.columns if col not in columnas_existentes]]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Integración y reordenamiento de CLASIFICADOR_GASTO completados.")

# ------------------------------------ #
# 17. Filtrado y Modificación de Datos
# ------------------------------------ #

# Mantener solo las filas donde CATEGORIA_GTO contiene "5."
datos_combinados = datos_combinados[datos_combinados["CATEGORIA_GTO"].str.startswith("5.", na=False)]

# Eliminar filas donde GENERICA sea "0. RESERVA DE CONTINGENCIA"
datos_combinados = datos_combinados[datos_combinados["GENERICA"] != "0. RESERVA DE CONTINGENCIA"]

# ------------------------------------ #
# 18. Cálculo de Columnas y Renombramiento
# ------------------------------------ #

# Renombrar columnas
datos_combinados = datos_combinados.rename(columns={
    "COMPROMISO_ANUAL": "COMPROMISO_ANUAL_2024",
    "CERTIFICADO": "CERTIFICADO_ANUAL_2024"
})

# ------------------------------------ #
# 19. Reordenamiento de Columnas
# ------------------------------------ #

# Especificar el nuevo orden de columnas
nuevo_orden = [
    "NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA", "SEC_EJEC", 
    "FUENTE", "RUBRO", "PROGRAMA_PPTAL", "FUNCION", "ACT_OBRA_ACCINV", 
    "PRODUCTO_PROYECTO", "FINALIDAD", "CATEGORIA_GTO", "TIPO_TRANSACCION", 
    "GENERICA", "SUBGENERICA", "SUBGENERICA_DET", "ESPECIFICA", 
    "ESPECIFICA_DET", "CLASIFICADOR_GASTO", "PIA", "PIM", 
    "COMPROMISO_ANUAL_2024", "CERTIFICADO_ANUAL_2024"
] + [col for col in datos_combinados.columns if col.startswith("DEV")]

# Reordenar las columnas
datos_combinados = datos_combinados[nuevo_orden]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Filtrado, cálculos y reordenamiento de columnas completados.")

# ------------------------------------ #
# 20. Renombrar Columnas
# ------------------------------------ #

# Meses en español
meses = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

# Renombrar columnas DEV01-DEV12 a formato DEV_<MES>_2024
for n, mes in enumerate(meses, start=1):
    col_old1 = f"DEV{str(n).zfill(2)}"  # Formato DEV01, DEV02...
    col_old2 = f"DEV{n}"                # Formato DEV1, DEV2...
    col_new = f"DEV_{mes}_2024"

    # Renombrar si existen las columnas
    if col_old1 in datos_combinados.columns:
        datos_combinados = datos_combinados.rename(columns={col_old1: col_new})
    if col_old2 in datos_combinados.columns:
        datos_combinados = datos_combinados.rename(columns={col_old2: col_new})

# Renombrar columnas PIA y PIM a PIA_2024 y PIM_2024
datos_combinados = datos_combinados.rename(columns={
    "PIA": "PIA_2024",
    "PIM": "PIM_2024"
})

###########################################
data_intermedia = datos_combinados.copy()
datos_combinados = data_intermedia.copy()
###########################################


# ------------------------------------ #
# 21. Consolidación de Datos y Limpieza de Duplicados
# ------------------------------------ #

import pandas as pd

# Verifica las columnas que cumplen con los patrones
columnas_relevantes = [col for col in datos_combinados.columns if col.startswith(("PIA_", "PIM_", "COMPROMISO", "CERTIFICADO", "DEV_"))]

if not columnas_relevantes:
    raise ValueError("No se encontraron columnas que coincidan con los patrones especificados.")

# Agrupa y resume
datos_combinados = datos_combinados.groupby([
    "NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA", "SEC_EJEC", "FUENTE", "RUBRO", 
    "PROGRAMA_PPTAL", "FUNCION", "ACT_OBRA_ACCINV", "PRODUCTO_PROYECTO", "FINALIDAD", 
    "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", "SUBGENERICA", 
    "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET", "CLASIFICADOR_GASTO"
])[columnas_relevantes].sum().reset_index()

# Crear el nombre de la columna DEV_ANUAL dinámicamente
columna_dev_anual = f"DEV_ANUAL_{anio_actual}"

# Calcular DEV_ANUAL y SALDO basados en el año dinámico
datos_combinados[columna_dev_anual] = datos_combinados.filter(like="DEV_").sum(axis=1, skipna=True)
datos_combinados["SALDO"] = datos_combinados[f"PIM_{anio_actual}"] - datos_combinados[columna_dev_anual]

###############################################################################################################
#CREAMOS UN BACKUP
###############################################################################################################

backup2 = datos_combinados.copy()
datos_combinados = backup2.copy()
################################################################################################################

# ------------------------------------ #
# 22. Creación de la Columna Partida_Historica
# ------------------------------------ #

import pandas as pd
import numpy as np
import re

# Crear las condiciones para Partida_Historica en datos_combinados
def asignar_partida_historica(row):
    """
    Asigna un valor a la columna 'Partida_Historica' basado en condiciones específicas
    de las columnas 'PLIEGO' y 'CLASIFICADOR_GASTO'.

    Args:
        row (Series): Una fila del DataFrame `datos_combinados`.

    Returns:
        str: El valor correspondiente a 'Partida_Historica' o NaN si no cumple ninguna condición.
    """
    # Condición específica para "6.0 CPMP"
    if row["PLIEGO"] in ["026. M. DE DEFENSA", "007. M. DEL INTERIOR"] and pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.5\.2\.1\.1\.99(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "6.0 CPMP"
    # Condición específica para "7.0 BONOS DE RECONOCIMIENTO"
    elif row["PLIEGO"] == "095. OFICINA DE NORMALIZACION PREVISIONAL-ONP" and pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.8\.1\.2\.2\.99(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "7.0 BONOS DE RECONOCIMIENTO"
    # Condición específica para "5.0 INDEMNIZACIONES/COMPENSACIONES"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.5\.5\.2(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "5.0 INDEMNIZACIONES/COMPENSACIONES"
    # Condición específica para "4.0 NEGOCIACIONES COLECTIVAS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.5\.6(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "4.0 NEGOCIACIONES COLECTIVAS"
    # Condición específica para "3.1 OTROS TIPOS DE CONTRATO DE PERSONAL"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and (re.match(r"^5\.2\.3\.2\.7\.5(\.\d+)*$", row["CLASIFICADOR_GASTO"]) or re.match(r"^5\.2\.3\.2\.7\.12(\.\d+)*$", row["CLASIFICADOR_GASTO"])) and not re.match(r"^5\.2\.3\.2\.7\.5\.9(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "3.1 OTROS TIPOS DE CONTRATO DE PERSONAL"
    # Condición específica para "3.0 LOCACIÓN DE SERVICIOS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.3\.2\.9\.1\.1(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "3.0 LOCACIÓN DE SERVICIOS"
    # Condición específica para "2.1 JUDICIALES PENSIONISTAS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.2\.1\.2(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "2.1 JUDICIALES PENSIONISTAS"
    # Condición específica para "2.0 PENSIONISTAS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.2(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "2.0 PENSIONISTAS"
    # Condición específica para "1.2 JUDICIALES ACTIVOS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.1\.5(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "1.2 JUDICIALES ACTIVOS"
    # Condición específica para "1.1 CAS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and (re.match(r"^5\.2\.1\.1\.13(\.\d+)*$", row["CLASIFICADOR_GASTO"]) or re.match(r"^5\.2\.1\.1\.9\.1\.4(\.\d+)*$", row["CLASIFICADOR_GASTO"]) or re.match(r"^5\.2\.1\.3\.1\.1\.15(\.\d+)*$", row["CLASIFICADOR_GASTO"])):
        return "1.1 CAS"
    # Condición general para "1.0 ACTIVOS"
    elif pd.notna(row["CLASIFICADOR_GASTO"]) and re.match(r"^5\.2\.1(\.\d+)*$", row["CLASIFICADOR_GASTO"]):
        return "1.0 ACTIVOS"
    else:
        return np.nan

# Asigna valores a la nueva columna 'Partida_Historica' basándose en condiciones personalizadas
# Esto permite categorizar los registros para un análisis más estructurado y segmentado.
datos_combinados["Partida_Historica"] = datos_combinados.apply(asignar_partida_historica, axis=1)

# Mensaje de finalización
print("Columna Partida_Historica creada correctamente.")

# Backup inicial
backup4 = datos_combinados.copy()
datos_combinados = backup4.copy()

# ------------------------------------ #
# 23. Estructura y Partidas de Gasto
# ------------------------------------ #

# Convertir SEC_EJEC a numérico
datos_combinados["SEC_EJEC"] = pd.to_numeric(datos_combinados["SEC_EJEC"], errors="coerce")

# Cargar la lista histórica y realizar el merge (equivalente a m:1 en Stata)
lista_ue_hist = pd.read_excel(os.path.join(Cross, "Lista_UE_Hist_2014-2024.xlsx"), dtype=str)
lista_ue_hist["SEC_EJEC"] = pd.to_numeric(lista_ue_hist["SEC_EJEC"], errors="coerce")  # Convertir SEC_EJEC en el archivo histórico

# Realizar la unión muchos-a-uno
datos_combinados = datos_combinados.merge(lista_ue_hist, on="SEC_EJEC", how="left")

# Depuración: eliminar columnas duplicadas (mantener `.x` y eliminar `.y`)
datos_combinados = datos_combinados[[col for col in datos_combinados.columns if not col.endswith(".y")]]

# Renombrar columnas con sufijo `.x` para quitar el sufijo
datos_combinados.columns = [col.replace(".x", "") for col in datos_combinados.columns]

# Eliminar filas que solo están en lista_ue_hist (equivalente a `_merge == 2` en Stata)
datos_combinados = datos_combinados[datos_combinados["SEC_EJEC"].notna()]



# ------------------------------------ #
# 24. Procesamiento de Variables Clave
# ------------------------------------ #

# Vector con las columnas a procesar
m = ["FUNCION", "ACT_OBRA_ACCINV", "FINALIDAD", "PROGRAMA_PPTAL", "RUBRO"]

for i in m:
    # Crear una nueva columna COD_`i' extrayendo texto hasta el primer punto
    datos_combinados[f"COD_{i}"] = datos_combinados[i].str.split(".").str[0]

    # Eliminar el punto si existe en COD_`i'
    datos_combinados[f"COD_{i}"] = datos_combinados[f"COD_{i}"].str.replace(".", "", regex=True)

    # Convertir la columna COD_`i' a numérica
    datos_combinados[f"COD_{i}"] = pd.to_numeric(datos_combinados[f"COD_{i}"], errors="coerce")

    # Eliminar la columna original `i'
    datos_combinados = datos_combinados.drop(columns=[i])

# Eliminar la columna FUENTE si existe
if "FUENTE" in datos_combinados.columns:
    datos_combinados = datos_combinados.drop(columns=["FUENTE"])

# Backup inicial
backup5 = datos_combinados.copy()
datos_combinados = backup5.copy()

import pandas as pd
import os

# ------------------------------------ #
# 25. Procesamiento de Hojas desde Excel
# ------------------------------------ #

# Ruta al archivo Excel
ruta_excel = os.path.join(Cross, "Estructura_SIAF_Maestro2024.xlsx")

# Vector con las hojas que deseas procesar
m = ["FUNCION", "ACT_OBRA_ACCINV", "FINALIDAD", "PROGRAMA_PPTAL", "RUBRO"]

# Procesamiento de cada hoja
for i in m:
    # Importar la hoja del archivo Excel
    df = pd.read_excel(ruta_excel, sheet_name=i, dtype=str)

    # Eliminar la columna ESTRUCTURA si existe
    if "ESTRUCTURA" in df.columns:
        df = df.drop(columns=["ESTRUCTURA"])

    # Agregar al dataframe existente datos_combinados con una clave específica
    datos_combinados = datos_combinados.merge(
        df,
        left_on=f"COD_{i}",
        right_on=f"COD_{i}",
        how="left"
    )

# ------------------------------------ #
# 26. Procesar CLASIFICADOR_GASTO
# ------------------------------------ #

# Cargar la estructura maestra para CLASIFICADOR_GASTO desde Excel
estructura_clasificador = pd.read_excel(
    os.path.join(Cross, "Estructura_SIAF_Maestro2024.xlsx"),
    sheet_name="CLASIFICADOR_GASTO",
    dtype=str
)

# Guardar temporalmente la estructura (equivalente a tempfile)
archivo_temporal = "CLASIFICADOR_GASTO_temp.pkl"
estructura_clasificador.to_pickle(archivo_temporal)

# Restaurar el estado anterior del dataframe
datos_combinados = backup5.copy()

# Eliminar las columnas relacionadas con el clasificador
datos_combinados = datos_combinados.drop(columns=[
    "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", "SUBGENERICA", 
    "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET"
], errors="ignore")

# Realizar el merge con la estructura de CLASIFICADOR_GASTO
estructura_clasificador = pd.read_pickle(archivo_temporal)

datos_combinados = datos_combinados.merge(
    estructura_clasificador,
    on="CLASIFICADOR_GASTO",
    how="left"
)

# Eliminar filas donde no hubo coincidencias
datos_combinados = datos_combinados[datos_combinados["SECTOR"].notna()]

# Eliminar columnas relacionadas con el merge
datos_combinados = datos_combinados[[col for col in datos_combinados.columns if not col.startswith("_merge")]]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Procesamiento de CLASIFICADOR_GASTO completado.")

# ------------------------------------ #
# 27. Generar ETIQUETA_ART_9_13
# ------------------------------------ #

# Crear la columna ETIQUETA_ART_9_13 con el valor predeterminado
datos_combinados["ETIQUETA_ART_9_13"] = "FINALIDADES NO RESTRINGIDAS"

# Actualizar ETIQUETA_ART_9_13 para las finalidades restringidas
datos_combinados.loc[
    datos_combinados["FINALIDAD"].str.contains("0267928|0267929|0334300", na=False),
    "ETIQUETA_ART_9_13"
] = "FINALIDADES RESTRINGIDAS"

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Generación de ETIQUETA_ART_9_13 completada.")

# ------------------------------------ #
# 28. Generar Variable PARTIDA_GASTO
# ------------------------------------ #

# Crear la columna PARTIDA_GASTO e inicializarla como cadena vacía
datos_combinados["PARTIDA_GASTO"] = ""

# GG1: Actualizar PARTIDA_GASTO para casos específicos de PERSONAL Y OBLIGACIONES SOCIALES
datos_combinados.loc[
    datos_combinados["GENERICA"].str.contains("PERSONAL Y OBLIGACIONES SOCIALES", na=False),
    "PARTIDA_GASTO"
] = "2.1. PERSONAL Y OBLIGACIONES SOCIALES"
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"].str.contains("5.2.1.1.13.1.2", na=False),
    "PARTIDA_GASTO"
] = "2.1.1.13.1.2 CONTRATO ADMINISTRATIVO DE SERVICIOS"
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"].str.contains("5.2.1.1.13.1.1", na=False),
    "PARTIDA_GASTO"
] = "2.1.1.13.1.1 CONTRATO ADMINISTRATIVO DE SERVICIOS"

# GG2: Actualizar PARTIDA_GASTO para PENSIONES Y OTRAS PRESTACIONES SOCIALES
datos_combinados.loc[
    (datos_combinados["GENERICA"].str.contains("PENSIONES Y OTRAS PRESTACIONES SOCIALES", na=False)) &
    (datos_combinados["SUBGENERICA_DET"].str.contains("PENSIONES", na=False)),
    "PARTIDA_GASTO"
] = "2.2. PENSIONES Y OTRAS PRESTACIONES SOCIALES"
datos_combinados.loc[
    (datos_combinados["GENERICA"].str.contains("PENSIONES Y OTRAS PRESTACIONES SOCIALES", na=False)) &
    (datos_combinados["SUBGENERICA_DET"].str.contains("PRESTACIONES DE SALUD Y OTROS BENEFICIOS", na=False)),
    "PARTIDA_GASTO"
] = "2.2. PENSIONES Y OTRAS PRESTACIONES SOCIALES"
datos_combinados.loc[
    (datos_combinados["GENERICA"].str.contains("PENSIONES Y OTRAS PRESTACIONES SOCIALES", na=False)) &
    (datos_combinados["SUBGENERICA_DET"].str.contains("ASISTENCIA SOCIAL EN PENSIONES E INDEMNIZACIONES", na=False)),
    "PARTIDA_GASTO"
] = "2.2. PENSIONES Y OTRAS PRESTACIONES SOCIALES"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PAGO DE SENTENCIAS JUDICIALES EN CALIDAD DE COSA JUZGADA Y LAUDOS ARBITRALES DEFINITIVOS POR ADEUDOS", na=False),
    "PARTIDA_GASTO"
] = "2.2. PENSIONES Y OTRAS PRESTACIONES SOCIALES"

# GG3: Actualizar PARTIDA_GASTO para casos específicos de prácticas y locaciones
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"].str.contains("5.2.3.2.7.5", na=False),
    "PARTIDA_GASTO"
] = "2.3.2.7.5. PRACTICANTES, SECIGRISTAS Y SIMILARES"
datos_combinados.loc[
    datos_combinados["ESPECIFICA_DET"].str.contains("FONDO DE APOYO GERENCIAL|PERSONAL ALTAMENTE CALIFICADO", na=False),
    "PARTIDA_GASTO"
] = "2.3.2.7.12. OTRAS MODALIDADES DE CONTRATACION DE PERSONAS NATURALES"
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"] == "5.2.3.2.9.1.1",
    "PARTIDA_GASTO"
] = "2.3.2.9.1.1. LOCACIÓN DE SERVICIOS RELACIONADAS AL ROL DE LA ENTIDAD"

# GG4 y GG5: Actualizar PARTIDA_GASTO para sentencias, indemnizaciones y transferencias
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PAGO DE SENTENCIAS JUDICIALES EN CALIDAD DE COSA JUZGADA POR ADEUDOS", na=False),
    "PARTIDA_GASTO"
] = "2.1.5.1. SENTENCIAS/LAUDOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("INDEMNIZACIONES Y COMPENSACIONES", na=False),
    "PARTIDA_GASTO"
] = "2.5.5.2. INDEMNIZACIÓN/COMPENSACIÓN"
datos_combinados.loc[
    (datos_combinados["PLIEGO"].str.contains("M. DEL INTERIOR|M. DE DEFENSA", na=False)) &
    (datos_combinados["FUNCION"].str.contains("24", na=False)) &
    (datos_combinados["ESPECIFICA_DET"].str.contains("A OTRAS ORGANIZACIONES", na=False)),
    "PARTIDA_GASTO"
] = "2.5. TRANSFERENCIAS A LA CAJA DE PENSIONES MILITAR POLICIAL"

# GG5 Negociaciones Colectivas
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"] == "5.2.5.6.1.1.1",
    "PARTIDA_GASTO"
] = "2.5.6.1.1.1. NEGOCIACIÓN COLECTIVA - NIVEL DESCENTRALIZADO POR ENTIDAD PÚBLICA"
datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"] == "5.2.5.6.1.1.2",
    "PARTIDA_GASTO"
] = "2.5.6.1.1.2. NEGOCIACIÓN COLECTIVA - NIVEL DESCENTRALIZADO EN EL ÁMBITO SECTORIAL"

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Variable PARTIDA_GASTO generada y actualizada correctamente.")


backup6 = datos_combinados.copy()
datos_combinados = backup6.copy()

# ------------------------------------ #
# 29. Filtrar Datos Basados en PARTIDA_GASTO
# ------------------------------------ #

# Eliminar filas donde PARTIDA_GASTO esté vacío
datos_combinados = datos_combinados[datos_combinados["PARTIDA_GASTO"] != ""]

# Eliminar filas donde PARTIDA_GASTO contenga "2.8. BONOS"
datos_combinados = datos_combinados[~datos_combinados["PARTIDA_GASTO"].str.contains("2.8. BONOS", na=False)]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Filtrado de PARTIDA_GASTO completado.")


# ------------------------------------ #
# 30. Generar Variable SUBPARTIDA_GASTO
# ------------------------------------ #

# Crear la columna SUBPARTIDA_GASTO e inicializarla como cadena vacía
datos_combinados["SUBPARTIDA_GASTO"] = ""

# Actualizar SUBPARTIDA_GASTO para las condiciones especificadas
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL ADMINISTRATIVO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.1. PERSONAL ADMINISTRATIVO"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL DEL MAGISTERIO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.2. PERSONAL DEL MAGISTERIO"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL DE LA SALUD", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.3. PERSONAL DE LA SALUD"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL JUDICIAL", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.4. PERSONAL JUDICIAL"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("DOCENTES UNIVERSITARIOS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.5. DOCENTES UNIVERSITARIOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL DIPLOMATICO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.6. PERSONAL DIPLOMATICO"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL MILITAR Y POLICIAL", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.7. PERSONAL MILITAR Y POLICIAL"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL OBRERO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.8. PERSONAL OBRERO"
datos_combinados.loc[
    (datos_combinados["SUBGENERICA_DET"].str.contains("GASTOS VARIABLES Y OCASIONALES", na=False)) &
    (datos_combinados["SUBGENERICA"].str.contains("RETRIBUCIONES Y COMPLEMENTOS EN EFECTIVO", na=False)),
    "SUBPARTIDA_GASTO"
] = "2.1.1.9. GASTOS VARIABLES Y OCASIONALES"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("DIETAS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.10. DIETAS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PERSONAL PENITENCIARIO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.11. PERSONAL PENITENCIARIO"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("DOCENTES DE INSTITUTOS DE EDUCACION SUPERIOR", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.12. DOCENTES DE INSTITUTOS DE EDUCACION SUPERIOR"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("CONTRATO ADMINISTRATIVO DE SERVICIOS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.1.13. CONTRATO ADMINISTRATIVO DE SERVICIOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("RETRIBUCIONES EN BIENES O SERVICIOS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.2.1. RETRIBUCIONES EN BIENES O SERVICIOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("OBLIGACIONES DEL EMPLEADOR", na=False),
    "SUBPARTIDA_GASTO"
] = "2.1.3.1. OBLIGACIONES DEL EMPLEADOR"
datos_combinados.loc[
    (datos_combinados["SUBGENERICA_DET"].str.contains("GASTOS VARIABLES Y OCASIONALES", na=False)) &
    (datos_combinados["SUBGENERICA"].str.contains("RETRIBUCIONES Y COMPLEMENTOS EN EFECTIVO VARIABLES", na=False)),
    "SUBPARTIDA_GASTO"
] = "2.1.4.1. GASTOS VARIABLES Y OCASIONALES"
datos_combinados.loc[
    (datos_combinados["SUBGENERICA_DET"].str.contains("PAGO DE SENTENCIAS JUDICIALES EN CALIDAD DE COSA JUZGADA POR ADEUDOS", na=False)) &
    (datos_combinados["SUBGENERICA"].str.contains("PAGO DE SENTENCIAS JUDICIALES EN CALIDAD DE COSA JUZGADA Y LAUDOS ARBITRALES DEFINITIVOS POR ADEUDOS", na=False)),
    "SUBPARTIDA_GASTO"
] = "2.1.5.1. PAGO DE SENTENCIAS JUDICIALES EN CALIDAD DE COSA JUZGADA POR ADEUDOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PENSIONES", na=False),
    "SUBPARTIDA_GASTO"
] = "2.2.1.1. PENSIONES"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("PRESTACIONES DE SALUD Y OTROS BENEFICIOS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.2.2.1. PRESTACIONES DE SALUD Y OTROS BENEFICIOS"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("ASISTENCIA SOCIAL EN PENSIONES E INDEMNIZACIONES", na=False),
    "SUBPARTIDA_GASTO"
] = "2.2.2.2. ASISTENCIA SOCIAL EN PENSIONES E INDEMNIZACIONES"
datos_combinados.loc[
    datos_combinados["SUBGENERICA_DET"].str.contains("ENTREGA DE BIENES Y SERVICIOS", na=False),
    "SUBPARTIDA_GASTO"
] = "2.2.2.3. ENTREGA DE BIENES Y SERVICIOS"
datos_combinados.loc[
    datos_combinados["ESPECIFICA"].str.contains("PRACTICANTES, SECIGRISTAS Y SIMILARES", na=False),
    "SUBPARTIDA_GASTO"
] = "2.3.2.7.5. PRACTICANTES, SECIGRISTAS Y SIMILARES"
datos_combinados.loc[
    datos_combinados["ESPECIFICA_DET"].str.contains("FONDO DE APOYO GERENCIAL", na=False),
    "SUBPARTIDA_GASTO"
] = "2.3.2.7.12.1. FONDO DE APOYO GERENCIAL"
datos_combinados.loc[
    datos_combinados["ESPECIFICA_DET"].str.contains("PERSONAL ALTAMENTE CALIFICADO", na=False),
    "SUBPARTIDA_GASTO"
] = "2.3.2.7.12.2. PERSONAL ALTAMENTE CALIFICADO"
datos_combinados.loc[
    datos_combinados["ESPECIFICA_DET"].str.contains("LOCACIÓN DE SERVICIOS REALIZADOS POR PERSONA NATURAL", na=False),
    "SUBPARTIDA_GASTO"
] = "2.3.2.9.1.1. LOCACIÓN DE SERVICIOS REALIZADOS POR PERSONA NATURAL"
datos_combinados.loc[
    (datos_combinados["ESPECIFICA"].str.contains("AL SECTOR PRIVADO", na=False)) &
    (datos_combinados["SUBGENERICA"].str.contains("PAGO DE SENTENCIAS JUDICIALES, LAUDOS ARBITRALES Y SIMILARES", na=False)),
    "SUBPARTIDA_GASTO"
] = "2.5.5.1.3. SENTENCIAS/LAUDOS - SECTOR PRIVADO"
datos_combinados.loc[
    datos_combinados["SUBPARTIDA_GASTO"] == "",
    "SUBPARTIDA_GASTO"
] = datos_combinados["PARTIDA_GASTO"]

# ------------------------------------ #
# Ordenar y Tabular
# ------------------------------------ #

# Reordenar columnas: SUBPARTIDA_GASTO precede a PARTIDA_GASTO
columnas = ["SUBPARTIDA_GASTO", "PARTIDA_GASTO"] + [col for col in datos_combinados.columns if col not in ["SUBPARTIDA_GASTO", "PARTIDA_GASTO"]]
datos_combinados = datos_combinados[columnas]

# Tabular SUBPARTIDA_GASTO para verificar
conteo_subpartida = datos_combinados["SUBPARTIDA_GASTO"].value_counts()
print(conteo_subpartida)

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Variable SUBPARTIDA_GASTO generada y actualizada correctamente.")


# ------------------------------------ #
# 31. Eliminar Filas Basadas en SUBPARTIDA_GASTO
# ------------------------------------ #

# Eliminar filas donde SUBPARTIDA_GASTO contiene "2.2.2.3. ENTREGA DE BIENES Y SERVICIOS"
datos_combinados = datos_combinados[~datos_combinados["SUBPARTIDA_GASTO"].str.contains("2.2.2.3. ENTREGA DE BIENES Y SERVICIOS", na=False)]

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Filtrado de SUBPARTIDA_GASTO completado.")

# ------------------------------------ #
# 32. Generar Variable VALIDAR
# ------------------------------------ #

# Crear la columna 'validar' e inicializar en 0
datos_combinados["validar"] = 0

# Actualizar 'validar' a 1 si 'CLASIFICADOR_GASTO' coincide con alguno de los valores especificados
clasificadores_validar = [
    "5.2.1.1.1.2.3", "5.2.1.1.1.2.4", "5.2.1.1.3.1.5", "5.2.1.1.3.1.6",
    "5.2.1.1.3.3.9", "5.2.1.1.3.3.1", "5.2.1.1.9.2.1", "5.2.1.4.1.1.3",
    "5.2.1.4.1.1.4", "5.2.1.4.1.1.6", "5.2.1.1.9.3.4", "5.2.1.1.9.3.5",
    "5.2.1.1.9.3.6", "5.2.1.1.9.3.7", "5.2.1.1.9.3.10", "5.2.1.1.9.3.11",
    "5.2.1.1.9.3.12", "5.2.1.1.9.3.13", "5.2.1.1.9.3.99", "5.2.1.1.10.1.3",
    "5.2.1.2.1.1.1", "5.2.1.2.1.1.99", "5.2.1.2.1.2.1", "5.2.1.2.1.2.2",
    "5.2.1.2.1.2.3", "5.2.1.2.1.2.4", "5.2.1.2.1.2.5", "5.2.1.2.1.2.99",
    "5.2.2.2.1.2.4", "5.2.1.1.2.2.2", "5.2.2.1.1.2.2", "5.2.2.1.1.2.6",
    "5.2.2.2.1.1.1", "5.2.2.2.1.2.1", "5.2.2.2.1.2.2", "5.2.2.2.1.2.3",
    "5.2.2.2.1.2.99", "5.2.2.2.2.1.3", "5.2.2.2.2.1.99", "5.2.2.2.3.1.1",
    "5.2.2.2.3.1.99", "5.2.2.2.3.2.1", "5.2.2.2.3.2.2", "5.2.2.2.3.2.99",
    "5.2.2.2.3.3.1", "5.2.2.2.3.3.99", "5.2.2.2.3.4.1", "5.2.2.2.3.4.2",
    "5.2.2.2.3.4.3", "5.2.2.2.3.99.99", "5.2.2.2.1.3.1", "5.2.2.2.1.3.2",
    "5.2.2.2.1.3.3", "5.2.1.4.1.1.6", "5.2.3.2.8.1.6", "5.2.3.2.8.1.7",
    "5.2.1.1.9.3.10", "5.2.3.2.8.1.10", "5.2.1.1.9.3.13", "5.2.3.2.9.1.1",
    "5.2.3.2.7.5.9", "5.2.3.2.7.5.97", "5.2.3.2.7.2.7"
]

datos_combinados.loc[datos_combinados["CLASIFICADOR_GASTO"].isin(clasificadores_validar), "validar"] = 1

# ------------------------------------ #
# Tabular la columna 'validar' para verificar
# ------------------------------------ #

conteo_validar = datos_combinados["validar"].value_counts()
print(conteo_validar)

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Generación de la variable 'validar' completada.")

# ------------------------------------ #
# 33. Generar Variable ETIQUETA_APM_DES
# ------------------------------------ #

# Crear la columna ETIQUETA_APM_DES con un valor inicial basado en la columna 'validar'
datos_combinados["ETIQUETA_APM_DES"] = datos_combinados["validar"].apply(lambda x: "LB: Conceptos AIRHSP" if x == 0 else "LB: Conceptos NO AIRHSP")

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Generación de la variable 'ETIQUETA_APM_DES' completada.")


# ------------------------------------ #
# 34. Generar Variable CONCEPTO
# ------------------------------------ #

# Crear la columna CONCEPTO con valores iniciales basados en PARTIDA_GASTO y CLASIFICADOR_GASTO
datos_combinados["CONCEPTO"] = None

datos_combinados.loc[
    datos_combinados["PARTIDA_GASTO"].isin(["2.5.5.1. SENTENCIAS/LAUDOS", "2.5.5.2. INDEMNIZACIÓN/COMPENSACIÓN"]),
    "CONCEPTO"
] = "SENTENCIAS JUDICIALES"

datos_combinados.loc[
    datos_combinados["PARTIDA_GASTO"] == "2.5. TRANSFERENCIAS A LA CAJA DE PENSIONES MILITAR POLICIAL",
    "CONCEPTO"
] = "TRANSFERENCIAS A LA CAJA DE PENSIONES MILITAR POLICIAL"

datos_combinados.loc[
    datos_combinados["CLASIFICADOR_GASTO"].isin(["5.2.1.1.9.3.98", "5.2.2.1.1.2.98", "5.2.3.2.7.5.99", "5.2.3.2.8.1.99"]),
    "CONCEPTO"
] = "RESTRINGIDA"

datos_combinados.loc[
    (datos_combinados["CLASIFICADOR_GASTO"] == "5.2.1.3.1.1.3") &
    (datos_combinados["FINALIDAD"].str.contains("0238543. PAGO DE CUOTAS DEL REPRO-AFP|0239152. PAGO DE CUOTAS DEL SINCERAMIENTO DE DEUDAS POR APORTACIONES A LA ONP", na=False)),
    "CONCEPTO"
] = "REPRO/SIDEA"

datos_combinados.loc[
    (datos_combinados["CLASIFICADOR_GASTO"] == "5.2.1.3.1.1.5") &
    (datos_combinados["FINALIDAD"].str.contains("0238544. PAGO DE CUOTAS DEL SINCERAMIENTO DE DEUDAS POR APORTACIONES AL ESSALUD", na=False)),
    "CONCEPTO"
] = "REPRO/SIDEA"

datos_combinados.loc[
    (datos_combinados["ETIQUETA_APM_DES"] == "LB: Conceptos NO AIRHSP") & (datos_combinados["CONCEPTO"].isna()),
    "CONCEPTO"
] = "NO AIRHSP"

datos_combinados.loc[
    datos_combinados["CONCEPTO"].isna(),
    "CONCEPTO"
] = "AIRHSP"

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Generación de la variable 'CONCEPTO' completada.")

# ------------------------------------ #
# 35. Ordenar las Columnas del DataFrame
# ------------------------------------ #

# Especificar el orden deseado de las columnas
columnas_orden = [
    "ETIQUETA_APM_DES", "ETIQUETA_ART_9_13", "NIVEL_GOB", "SECTOR", "PLIEGO", "EJECUTORA",
    "SEC_EJEC", "FINALIDAD", "FUENTE", "RUBRO", "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", 
    "SUBGENERICA", "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET", "PARTIDA_GASTO", 
    "SUBPARTIDA_GASTO", "CLASIFICADOR_GASTO", "CONCEPTO"
] + \
    [col for col in datos_combinados.columns if col.startswith("PIA_")] + \
    [col for col in datos_combinados.columns if col.startswith("PIM_")] + \
    [col for col in datos_combinados.columns if col.startswith("COMP")] + \
    [col for col in datos_combinados.columns if col.startswith("CERT")] + \
    [col for col in datos_combinados.columns if col.startswith("DEV")]

# Reordenar las columnas del DataFrame
datos_combinados = datos_combinados[columnas_orden]
# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Ordenamiento de columnas completado.")

# ------------------------------------ #
# 36. Eliminar Duplicados y Colapsar Datos
# ------------------------------------ #

# Especificar las columnas a colapsar
columnas_sumar = [
    col for col in datos_combinados.columns if col.startswith(("PIA_", "PIM_", "COMP", "CERT", "DEV", "SALDO"))
]

# Especificar las columnas de agrupación como todas las que no están en columnas_sumar
columnas_agrupacion = [
    col for col in datos_combinados.columns if col not in columnas_sumar
]

# Colapsar datos mediante la suma
datos_combinados = datos_combinados.groupby(columnas_agrupacion, as_index=False)[columnas_sumar].sum()

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Colapso de datos y eliminación de duplicados completado.")


# ------------------------------------ #
# 37. Guardar Datos en Formato CSV
# ------------------------------------ #

# Generar el nombre dinámico del archivo
archivo_salida = os.path.join(Mean, f"1. SIAF_{Fecha_SIAF}.csv")

# Guardar los datos en formato CSV con separador "|", codificación UTF-8
datos_combinados.to_csv(
    archivo_salida,
    sep="|",
    index=False,
    encoding="utf-8",
    quoting=3  # Evita comillas en los campos
)

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print(f"Archivo guardado exitosamente en: {archivo_salida}")

backup7 = datos_combinados.copy()
datos_combinados = backup7.copy()

# ------------------------------------ #
# 38. Procesar y Colapsar Datos
# ------------------------------------ #

# Convertir 'SEC_EJEC' a numérico
datos_combinados["SEC_EJEC"] = pd.to_numeric(datos_combinados["SEC_EJEC"], errors="coerce")

# Eliminar columnas 'PARTIDA_GASTO' y 'SUBPARTIDA_GASTO'
datos_combinados = datos_combinados.drop(columns=["PARTIDA_GASTO", "SUBPARTIDA_GASTO"], errors="ignore")

# Especificar las columnas a colapsar
columnas_sumar = [
    col for col in datos_combinados.columns if col.startswith(("PIA_", "PIM_", "COMPROMISO_", "DEV_", "CERTIFICADO_"))
]

# Agrupar por las columnas deseadas
columnas_agrupacion = [
    "ETIQUETA_APM_DES", "ETIQUETA_ART_9_13", "NIVEL_GOB", "SECTOR", 
    "PLIEGO", "EJECUTORA", "SEC_EJEC", "FINALIDAD", "FUENTE", "RUBRO", 
    "CATEGORIA_GTO", "TIPO_TRANSACCION", "GENERICA", "SUBGENERICA", 
    "SUBGENERICA_DET", "ESPECIFICA", "ESPECIFICA_DET", "CLASIFICADOR_GASTO", 
    "CONCEPTO"
]

# Colapsar datos mediante la suma
datos_combinados = datos_combinados.groupby(columnas_agrupacion, as_index=False)[columnas_sumar].sum()

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print("Conversión, eliminación y colapso de datos completados.")

# ------------------------------------ #
# 39. Guardar Datos en Formato CSV
# ------------------------------------ #

# Generar el nombre del archivo de salida
archivo_salida = os.path.join(Mean, f"1. SIAF_{Fecha_SIAF}_sinPGvs2.csv")

# Guardar los datos en formato CSV con separador "|", codificación UTF-8
datos_combinados.to_csv(
    archivo_salida,
    sep="|",
    index=False,
    encoding="utf-8",
    quoting=3  # Evita comillas en los campos
)

# ------------------------------------ #
# Mensaje Final
# ------------------------------------ #

print(f"Archivo guardado exitosamente en: {archivo_salida}")


