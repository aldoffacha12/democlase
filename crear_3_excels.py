

import pandas as pd


# EJERCICIO 1 - TRANSPORTE

# Contexto: Empresa Municipal de Transporte Urbano (EMTU).
# Objetivo/target: predecir la columna "retraso_mayor_10" (1 = tuvo retraso >10 min, 0 = no).
# Algoritmo sugerido: RandomForestClassifier (clasificación porque la respuesta es 0/1).

# Lista de filas crudas. Cada sublista = 1 viaje, en el mismo orden que "columns" de abajo.
# Nota: los datos traen errores a propósito (duplicados, nulos con None,
# categorías mal escritas tipo "L1"/"Linea 1"/"L-2", y un outlier de 180 pasajeros).
# Eso es justamente lo que el estudiante debe detectar y limpiar.
datos_transporte = [
    ["V001", "2026-05-01", "L1", "Mañana", 48, 16.2, "No", 12, 0],
    ["V002", "2026-05-01", "L2", "Tarde", 61, 17.1, "Sí", 7, 1],
    ["V003", "2026-05-01", "Linea 1", "mañana", None, 15.5, "no", 10, 0],   # nulo en pasajeros + "Linea 1" inconsistente
    ["V004", "2026-05-02", "L-2", "Tarde", 70, 18.0, "SI", 5, 1],           # "L-2" inconsistente + "SI" debería ser "Sí"
    ["V005", "2026-05-02", "L3", "Noche", 55, 14.8, "No", 30, 0],
    ["V006", "2026-05-02", "L3", "noche", 53, None, "NO", 28, 1],           # nulo en temperatura + "noche" en minúscula
    ["V007", "2026-05-03", "L1", "Mañana", 180, 16.9, "no", 11, 0],         # OUTLIER: 180 pasajeros (los demás ~45-70)
    ["V008", "2026-05-03", "L2", "Tarde", 64, 19.1, "Sí", None, 1],         # nulo en mantenimiento_dias
    ["V009", "2026-05-03", "Línea 3", "Noche", 51, 13.9, "No", 22, 0],      # "Línea 3" con tilde y texto = inconsistente
    ["V010", "2026-05-04", "L1", "Mañana", 47, 17.2, "No", 9, 0],
    ["V011", "2026-05-04", "L2", "Tarde", None, 20.1, "Si", 6, 1],          # nulo en pasajeros
    ["V012", "2026-05-04", "L3", "Noche", 58, 14.2, "No", 25, 1],
    ["V012", "2026-05-04", "L3", "Noche", 58, 14.2, "No", 25, 1],           # FILA DUPLICADA de la anterior
]

# pd.DataFrame convierte la lista en una tabla con encabezados.
# columns = nombres de las columnas (variables del dataset).
df_transporte = pd.DataFrame(datos_transporte, columns=[
    "viaje_id", "fecha", "linea", "turno", "pasajeros",
    "temperatura", "lluvia", "mantenimiento_dias",
    "retraso_mayor_10"
])

# HOJA "Diccionario": documenta qué significa cada columna y su tipo
# (Texto/Fecha/Categórica/Numérica/Binaria). Ayuda a entender los datos antes de limpiar.
dic_transporte = pd.DataFrame([
    ["viaje_id", "Identificador único del viaje", "Texto"],
    ["fecha", "Fecha del viaje", "Fecha"],
    ["linea", "Línea de transporte", "Categórica"],
    ["turno", "Turno del viaje", "Categórica"],
    ["pasajeros", "Cantidad de pasajeros", "Numérica"],
    ["temperatura", "Temperatura registrada", "Numérica"],
    ["lluvia", "Existencia de lluvia", "Categórica"],
    ["mantenimiento_dias", "Días desde mantenimiento", "Numérica"],
    ["retraso_mayor_10", "Objetivo: retraso > 10 minutos", "Binaria"],
], columns=["Variable", "Descripción", "Tipo"])

# HOJA "Referencia": lista los valores VÁLIDOS de cada categoría.
# Sirve como "policía" para estandarizar: todo lo que no aparezca aquí
# (ej. "L-2", "Linea 1", "SI", "noche") hay que corregirlo o mapearlo a estos valores.
ref_transporte = pd.DataFrame([
    ["linea", "L1", "Línea 1"],
    ["linea", "L2", "Línea 2"],
    ["linea", "L3", "Línea 3"],
    ["turno", "Mañana", "Mañana"],
    ["turno", "Tarde", "Tarde"],
    ["turno", "Noche", "Noche"],
    ["lluvia", "Sí", "Sí"],
    ["lluvia", "No", "No"],
], columns=["Variable", "Valor", "Valor_estandar"])

# HOJA "Instrucciones": enunciado del ejercicio
# (problema a resolver, empresa, algoritmo a usar, tipos de limpieza y columna objetivo).
inst_transporte = pd.DataFrame([
    ["Problema", "Predecir si un viaje tendrá retraso mayor a 10 minutos."],
    ["Empresa", "Empresa Municipal de Transporte Urbano (EMTU)"],
    ["Algoritmo", "RandomForestClassifier"],
    ["Limpieza", "Duplicados, nulos, categorías inconsistentes y valores extremos."],
    ["Objetivo", "retraso_mayor_10"],
], columns=["Elemento", "Descripción"])


# ESCRITURA DEL ARCHIVO EXCEL
# pd.ExcelWriter abre un archivo .xlsx para escritura; engine="openpyxl" = usar openpyxl como motor.
# "with" = al salir del bloque, el archivo se cierra y se GUZA automáticamente.
# Cada to_excel() = una hoja (sheet). index=False evita que pandas agregue una columna extra de índice 0,1,2...
with pd.ExcelWriter("01_Transporte_EMTU.xlsx", engine="openpyxl") as writer:
    df_transporte.to_excel(writer, sheet_name="Datos_Crudos", index=False)
    dic_transporte.to_excel(writer, sheet_name="Diccionario", index=False)
    inst_transporte.to_excel(writer, sheet_name="Instrucciones", index=False)
    ref_transporte.to_excel(writer, sheet_name="Referencia", index=False)



# EJERCICIO 2 - MANTENIMIENTO INDUSTRIAL

# Contexto: Industrias Andinas de Manufactura S.A.
# Objetivo: predecir "falla_24h" (1 = la máquina fallará en las próximas 24 horas).
# Mismos problemas de calidad que el ejercicio 1:
#   - Duplicados (M012 aparece dos veces)
#   - Nulos (None en temperatura, vibracion, presion)
#   - Categorías inconsistentes ("torno" vs "Torno", "FRESADORA" vs "Fresadora", "BUENA" vs "Buena")
#   - Outliers (M007: 95° de temperatura, 9.2 de vibración, 8900 horas)
datos_maquinas = [
    ["M001", "Torno", 72.5, 2.1, 8.2, 4500, "Buena", 10, 0],
    ["M002", "Fresadora", 81.2, 4.8, 9.1, 5200, "Media", 25, 1],
    ["M003", "torno", None, 2.4, 8.5, 4700, "Buena", 12, 0],        # nulo en temperatura + "torno" minúscula
    ["M004", "FRESADORA", 85.4, 5.1, 9.4, 6100, "Baja", 40, 1],     # "FRESADORA" en mayúsculas
    ["M005", "Prensa", 68.2, 1.8, 7.9, 3200, "Buena", 8, 0],
    ["M006", "prensa", 70.1, None, 8.1, 3500, "Media", 15, 0],      # nulo en vibración
    ["M007", "Torno", 95.0, 9.2, 12.5, 8900, "Baja", 65, 1],        # OUTLIER en varias columnas
    ["M008", "Fresadora", 79.4, 4.2, None, 5100, "Media", 20, 1],   # nulo en presión
    ["M009", "Prensa", 67.8, 1.5, 7.8, 3000, "BUENA", 7, 0],        # "BUENA" en mayúsculas
    ["M010", "Torno", 73.2, 2.2, 8.3, 4600, "Buena", 11, 0],
    ["M011", "Fresadora", 88.0, 6.5, 9.8, 7000, "Baja", 50, 1],
    ["M012", "Prensa", 69.5, 1.7, 8.0, 3400, "Media", 18, 0],
    ["M012", "Prensa", 69.5, 1.7, 8.0, 3400, "Media", 18, 0],       # FILA DUPLICADA
]

# Tabla de datos con los nombres de columnas del dataset.
df_maquinas = pd.DataFrame(datos_maquinas, columns=[
    "maquina_id", "tipo_maquina", "temperatura", "vibracion",
    "presion", "horas_operacion", "lubricacion",
    "mantenimiento_dias", "falla_24h"
])

# Diccionario de variables: qué significa y de qué tipo es cada columna.
dic_maquinas = pd.DataFrame([
    ["maquina_id", "Identificador de la máquina", "Texto"],
    ["tipo_maquina", "Tipo de máquina", "Categórica"],
    ["temperatura", "Temperatura de operación", "Numérica"],
    ["vibracion", "Nivel de vibración", "Numérica"],
    ["presion", "Presión registrada", "Numérica"],
    ["horas_operacion", "Horas acumuladas", "Numérica"],
    ["lubricacion", "Estado de lubricación", "Categórica"],
    ["mantenimiento_dias", "Días desde mantenimiento", "Numérica"],
    ["falla_24h", "Falla en las próximas 24 horas", "Binaria"],
], columns=["Variable", "Descripción", "Tipo"])

# Valores válidos de las categorías (aquí solo 2 columnas: Variable + Valor_valido).
ref_maquinas = pd.DataFrame([
    ["tipo_maquina", "Torno"],
    ["tipo_maquina", "Fresadora"],
    ["tipo_maquina", "Prensa"],
    ["lubricacion", "Buena"],
    ["lubricacion", "Media"],
    ["lubricacion", "Baja"],
], columns=["Variable", "Valor_valido"])

# Enunciado del ejercicio 2.
inst_maquinas = pd.DataFrame([
    ["Problema", "Predecir si una máquina tendrá una falla en las próximas 24 horas."],
    ["Empresa", "Industrias Andinas de Manufactura S.A."],
    ["Algoritmo", "RandomForestClassifier"],
    ["Limpieza", "Duplicados, valores faltantes, categorías inconsistentes y valores extremos."],
    ["Objetivo", "falla_24h"],
], columns=["Elemento", "Descripción"])


# Genera el segundo archivo Excel con sus 4 hojas (mismo proceso que el ejercicio 1).
with pd.ExcelWriter(
    "02_Mantenimiento_Industrial.xlsx",
    engine="openpyxl"
) as writer:
    df_maquinas.to_excel(writer, sheet_name="Datos_Crudos", index=False)
    dic_maquinas.to_excel(writer, sheet_name="Diccionario", index=False)
    inst_maquinas.to_excel(writer, sheet_name="Instrucciones", index=False)
    ref_maquinas.to_excel(writer, sheet_name="Referencia", index=False)



# EJERCICIO 3 - CONSUMO DE ENERGÍA

# Contexto: Administración del Complejo Administrativo Central.
# Objetivo: predecir "consumo_kwh" (kWh consumidos).
# OJO: aquí el objetivo es NUMÉRICO continuo, no 0/1, por eso el algoritmo
# sugerido es RandomForestRegressor (REGRESIÓN) en vez de Classifier.
# Problemas de calidad:
#   - Duplicados (E012 aparece dos veces)
#   - Nulos (None en temperatura, humedad, equipos_activos)
#   - Categorías inconsistentes ("martes"/"Martes", "Miercoles"/"Miércoles", "OFICINAS"/"Oficinas")
#   - Outlier EXTREMO: E013 con temperatura 500 y consumo 9999 kWh
datos_energia = [
    ["E001", "2026-06-01", 8, "Lunes", 18.2, 55, 42, 15, "Oficinas", 82.5],
    ["E002", "2026-06-01", 10, "Lunes", 21.4, 52, 75, 28, "Oficinas", 118.3],
    ["E003", "2026-06-01", 12, "Lunes", None, 50, 90, 35, "Servidores", 156.8],  # nulo en temperatura
    ["E004", "2026-06-01", 14, "Lunes", 24.8, 48, 82, 31, "Oficinas", 141.2],
    ["E005", "2026-06-02", 9, "Martes", 19.5, None, 60, 22, "Laboratorio", 105.4],  # nulo en humedad
    ["E006", "2026-06-02", 11, "martes", 22.1, 49, 78, None, "Laboratorio", 132.6],  # "martes" minúscula + nulo
    ["E007", "2026-06-02", 13, "Martes", 23.5, 47, 85, 30, "Servidores", 160.1],
    ["E008", "2026-06-02", 15, "Martes", 25.2, 45, 80, 29, "OFICINAS", 149.7],  # "OFICINAS" en mayúsculas
    ["E009", "2026-06-03", 8, "Miércoles", 17.8, 57, 40, 12, "Oficinas", 78.4],
    ["E010", "2026-06-03", 10, "Miercoles", 20.5, 54, 70, 20, "Laboratorio", 111.9],  # "Miercoles" sin tilde
    ["E011", "2026-06-03", 12, "Miércoles", 22.7, 51, 88, 33, "Servidores", 170.5],
    ["E012", "2026-06-03", 14, "Miércoles", 23.9, 49, 82, 28, "Oficinas", 145.3],
    ["E013", "2026-06-04", 12, "Jueves", 500, 50, 85, 35, "Servidores", 9999],  # OUTLIER EXTREMO (500° y 9999 kWh)
    ["E012", "2026-06-03", 14, "Miércoles", 23.9, 49, 82, 28, "Oficinas", 145.3],  # FILA DUPLICADA
]

# Tabla con los nombres de columnas del dataset de energía.
df_energia = pd.DataFrame(datos_energia, columns=[
    "registro_id", "fecha", "hora", "dia_semana", "temperatura",
    "humedad", "personas", "equipos_activos", "area",
    "consumo_kwh"
])

# Diccionario de variables del ejercicio 3.
# Nota: "consumo_kwh" se marca como "Numérica / objetivo" porque es la columna a predecir (y).
dic_energia = pd.DataFrame([
    ["registro_id", "Identificador del registro", "Texto"],
    ["fecha", "Fecha de medición", "Fecha"],
    ["hora", "Hora de medición", "Numérica"],
    ["dia_semana", "Día de la semana", "Categórica"],
    ["temperatura", "Temperatura exterior", "Numérica"],
    ["humedad", "Humedad relativa", "Numérica"],
    ["personas", "Personas presentes", "Numérica"],
    ["equipos_activos", "Equipos encendidos", "Numérica"],
    ["area", "Área del edificio", "Categórica"],
    ["consumo_kwh", "Consumo eléctrico", "Numérica / objetivo"],
], columns=["Variable", "Descripción", "Tipo"])

# Valores válidos de día_semana y area (para detectar "martes", "Miercoles", "OFICINAS", etc.).
ref_energia = pd.DataFrame([
    ["dia_semana", "Lunes"],
    ["dia_semana", "Martes"],
    ["dia_semana", "Miércoles"],
    ["dia_semana", "Jueves"],
    ["dia_semana", "Viernes"],
    ["area", "Oficinas"],
    ["area", "Laboratorio"],
    ["area", "Servidores"],
], columns=["Variable", "Valor_valido"])

# Enunciado del ejercicio 3 (nota: usa RandomForestRegressor y "Institución" en vez de "Empresa").
inst_energia = pd.DataFrame([
    ["Problema", "Predecir el consumo eléctrico del edificio en kWh."],
    ["Institución", "Administración del Complejo Administrativo Central"],
    ["Algoritmo", "RandomForestRegressor"],
    ["Limpieza", "Duplicados, nulos, categorías inconsistentes y valores extremos."],
    ["Objetivo", "consumo_kwh"],
], columns=["Elemento", "Descripción"])


# Genera el tercer archivo Excel con sus 4 hojas.
with pd.ExcelWriter(
    "03_Consumo_Energia.xlsx",
    engine="openpyxl"
) as writer:
    df_energia.to_excel(writer, sheet_name="Datos_Crudos", index=False)
    dic_energia.to_excel(writer, sheet_name="Diccionario", index=False)
    inst_energia.to_excel(writer, sheet_name="Instrucciones", index=False)
    ref_energia.to_excel(writer, sheet_name="Referencia", index=False)


# Mensaje final de confirmación en consola.
print("Los 3 archivos Excel fueron creados correctamente.")
