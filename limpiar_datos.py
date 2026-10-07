"""
Script de LIMPIEZA DE DATOS para los 3 ejercicios.

Lee los archivos creados por 'crear_3_excels.py' (hoja Datos_Crudos),
corrige los errores de cada tabla y genera versiones limpias
en archivos nuevos (*_limpio.xlsx) para poder entrenar el modelo.

Limpiezas aplicadas:
  1. Duplicados              -> se eliminan las filas repetidas.
  2. Categorías inconsistentes -> se estandarizan a los valores válidos
                                 de la hoja "Referencia".
  3. Nulos                    -> se rellenan con la MEDIANA de la columna.
  4. Outliers (valores extremos) -> se recortan (winsorización) usando IQR:
                                   límites = Q1 - 1.5*IQR  y  Q3 + 1.5*IQR.

Cada archivo limpio tiene 5 hojas:
  Datos_Limpios      -> los datos ya corregidos
  Reporte_Limpieza   -> registro de TODO lo que se hizo (para justificar)
  Diccionario        -> copiada del archivo original
  Instrucciones      -> copiada del archivo original
  Referencia         -> copiada del archivo original

pandas: manipula las tablas. openpyxl: escribe los .xlsx.
"""

import pandas as pd

# ----------------------------------------------------------------------------
# Funciones auxiliares (reusadas en los 3 ejercicios)
# ----------------------------------------------------------------------------

def estandarizar(df, columna, mapa, reporte, PASO="Categorías"):
    """Convierte todas las variantes (mayúsculas, minúsculas, sin tilde, alias)
    a un único valor válido usando un mapa. Los cambios quedan en el reporte."""
    antes = df[columna].copy()
    normalizado = df[columna].str.strip().str.lower()
    corregido = normalizado.map(mapa)
    sin_mapear = corregido.isna() & antes.notna()
    if sin_mapear.any():
        raise ValueError(f"Columna '{columna}': sin mapeo -> {antes[sin_mapear].unique()}")
    df[columna] = corregido
    cambios = int((antes != df[columna]).sum())
    reporte.append([PASO, f"'{columna}' estandarizada a valores válidos", cambios])


def imputar_mediana(df, columnas, reporte):
    """Rellena los valores faltantes (NaN) con la mediana de cada columna."""
    for c in columnas:
        faltan = int(df[c].isna().sum())
        if faltan > 0:
            mediana = df[c].median()
            df[c] = df[c].fillna(mediana)
            reporte.append(["Nulos", f"'{c}' rellenos con la mediana ({mediana:.1f})", faltan])


def recortar_outliers(df, columnas, reporte):
    """Detecta outliers con el método IQR y los recorta al límite más cercano
    (no borra filas: suaviza el valor extremo). Deja constancia en el reporte."""
    for c in columnas:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
        q1, q3 = df[c].quantile(0.25), df[c].quantile(0.75)
        iqr = q3 - q1
        limite_inf = q1 - 1.5 * iqr
        limite_sup = q3 + 1.5 * iqr
        fuera = (df[c] < limite_inf) | (df[c] > limite_sup)
        n = int(fuera.sum())
        if n > 0:
            antes = df.loc[fuera, c].tolist()
            df.loc[fuera, c] = df[c].clip(lower=limite_inf, upper=limite_sup)[fuera]
            despues = df.loc[fuera, c].tolist()
            reporte.append(
                ["Outliers",
                 f"'{c}' recortados a [{limite_inf:.1f}, {limite_sup:.1f}] "
                 f"(antes={antes}, después={despues})",
                 n]
            )


# ============================================================
# EJERCICIO 1 - TRANSPORTE
# ============================================================
# Empresa Municipal de Transporte Urbano (EMTU).
# Objetivo: predecir "retraso_mayor_10" (1 = retraso >10 min, 0 = no).

# 1. Cargar los datos CRUDOS (con los errores) desde el archivo original.
df_transporte = pd.read_excel("01_Transporte_EMTU.xlsx", sheet_name="Datos_Crudos")
n_originales = len(df_transporte)

# Lista que después se convierte en la hoja "Reporte_Limpieza".
reporte_transporte = []

# 2. Eliminar duplicados (V012 estaba repetido).
n_duplicados = int(df_transporte.duplicated().sum())
df_transporte = df_transporte.drop_duplicates()
reporte_transporte.append(["Duplicados", "filas duplicadas eliminadas", n_duplicados])

# 3. Estandarizar categorías usando los valores válidos de la hoja "Referencia".
#    Antes: "L1","Linea 1","L-2","Línea 3"... -> Ahora: L1, L2, L3.
estandarizar(df_transporte, "linea", {
    "l1": "L1", "linea 1": "L1", "línea 1": "L1", "linea 1": "L1",
    "l2": "L2", "l-2": "L2", "linea 2": "L2", "línea 2": "L2",
    "l3": "L3", "linea 3": "L3", "línea 3": "L3",
}, reporte_transporte)
#    turno: "mañana" -> "Mañana", "noche" -> "Noche".
estandarizar(df_transporte, "turno", {
    "mañana": "Mañana", "manana": "Mañana", "tarde": "Tarde", "noche": "Noche",
}, reporte_transporte)
#    lluvia: "si"/"SI" -> "Sí", "no"/"NO" -> "No".
estandarizar(df_transporte, "lluvia", {
    "sí": "Sí", "si": "Sí", "no": "No",
}, reporte_transporte)

# 4. Convertir la fecha de texto a tipo fecha (datetime).
df_transporte["fecha"] = pd.to_datetime(df_transporte["fecha"])
reporte_transporte.append(["Tipos", "'fecha' convertida a datetime", len(df_transporte)])

# 5. Rellenar nulos con la mediana (pasajeros, temperatura, mantenimiento_dias).
imputar_mediana(df_transporte, ["pasajeros", "temperatura", "mantenimiento_dias"],
                reporte_transporte)

# 6. Recortar outliers: V007 tenía 180 pasajeros (todos van de ~45 a ~70).
recortar_outliers(df_transporte,
                  ["pasajeros", "temperatura", "mantenimiento_dias"],
                  reporte_transporte)

# 7. Asegurar que la columna objetivo quede como 0/1 entero.
df_transporte["retraso_mayor_10"] = df_transporte["retraso_mayor_10"].astype(int)

reporte_transporte.append(
    ["Filas", "total: originales -> limpias",
     f"{n_originales} -> {len(df_transporte)}"]
)

# Convertir la lista en tabla (así se escribe como hoja de Excel).
reporte_transporte = pd.DataFrame(
    reporte_transporte, columns=["Paso", "Acción aplicada", "Cantidad"]
)


# 8. Escribir el archivo LIMPIO (5 hojas, igual que el patrón del script original).
with pd.ExcelWriter("01_Transporte_EMTU_limpio.xlsx", engine="openpyxl") as writer:
    df_transporte.to_excel(writer, sheet_name="Datos_Limpios", index=False)
    reporte_transporte.to_excel(writer, sheet_name="Reporte_Limpieza", index=False)
    pd.read_excel("01_Transporte_EMTU.xlsx", "Diccionario").to_excel(
        writer, sheet_name="Diccionario", index=False)
    pd.read_excel("01_Transporte_EMTU.xlsx", "Instrucciones").to_excel(
        writer, sheet_name="Instrucciones", index=False)
    pd.read_excel("01_Transporte_EMTU.xlsx", "Referencia").to_excel(
        writer, sheet_name="Referencia", index=False)

print("01_Transporte_EMTU_limpio.xlsx creado.")


# ============================================================
# EJERCICIO 2 - MANTENIMIENTO INDUSTRIAL
# ============================================================
# Industrias Andinas de Manufactura S.A.
# Objetivo: predecir "falla_24h" (1 = la máquina fallará en 24 horas).

# 1. Cargar los datos crudos.
df_maquinas = pd.read_excel("02_Mantenimiento_Industrial.xlsx", sheet_name="Datos_Crudos")
n_originales = len(df_maquinas)

reporte_maquinas = []

# 2. Eliminar duplicados (M012 estaba repetido).
n_duplicados = int(df_maquinas.duplicated().sum())
df_maquinas = df_maquinas.drop_duplicates()
reporte_maquinas.append(["Duplicados", "filas duplicadas eliminadas", n_duplicados])

# 3. Estandarizar categorías: "torno"/"FRESADORA" -> "Torno"/"Fresadora",
#    y lubricación "BUENA" -> "Buena".
estandarizar(df_maquinas, "tipo_maquina", {
    "torno": "Torno", "fresadora": "Fresadora", "prensa": "Prensa",
}, reporte_maquinas)
estandarizar(df_maquinas, "lubricacion", {
    "buena": "Buena", "media": "Media", "baja": "Baja",
}, reporte_maquinas)

# 4. Rellenar nulos con la mediana (temperatura, vibracion, presion).
imputar_mediana(df_maquinas, ["temperatura", "vibracion", "presion"],
                reporte_maquinas)

# 5. Recortar outliers: M007 tenía 95° / vibración 9.2 / 8900 horas.
recortar_outliers(
    df_maquinas,
    ["temperatura", "vibracion", "presion", "horas_operacion", "mantenimiento_dias"],
    reporte_maquinas
)

# 6. Columna objetivo como 0/1 entero.
df_maquinas["falla_24h"] = df_maquinas["falla_24h"].astype(int)

reporte_maquinas.append(
    ["Filas", "total: originales -> limpias", f"{n_originales} -> {len(df_maquinas)}"]
)

reporte_maquinas = pd.DataFrame(
    reporte_maquinas, columns=["Paso", "Acción aplicada", "Cantidad"]
)


# 7. Escribir el archivo LIMPIO.
with pd.ExcelWriter("02_Mantenimiento_Industrial_limpio.xlsx",
                    engine="openpyxl") as writer:
    df_maquinas.to_excel(writer, sheet_name="Datos_Limpios", index=False)
    reporte_maquinas.to_excel(writer, sheet_name="Reporte_Limpieza", index=False)
    pd.read_excel("02_Mantenimiento_Industrial.xlsx", "Diccionario").to_excel(
        writer, sheet_name="Diccionario", index=False)
    pd.read_excel("02_Mantenimiento_Industrial.xlsx", "Instrucciones").to_excel(
        writer, sheet_name="Instrucciones", index=False)
    pd.read_excel("02_Mantenimiento_Industrial.xlsx", "Referencia").to_excel(
        writer, sheet_name="Referencia", index=False)

print("02_Mantenimiento_Industrial_limpio.xlsx creado.")


# ============================================================
# EJERCICIO 3 - CONSUMO DE ENERGÍA
# ============================================================
# Administración del Complejo Administrativo Central.
# Objetivo: predecir "consumo_kwh" (numérico -> este sí es regresión).

# 1. Cargar los datos crudos.
df_energia = pd.read_excel("03_Consumo_Energia.xlsx", sheet_name="Datos_Crudos")
n_originales = len(df_energia)

reporte_energia = []

# 2. Eliminar duplicados (E012 estaba repetido).
n_duplicados = int(df_energia.duplicated().sum())
df_energia = df_energia.drop_duplicates()
reporte_energia.append(["Duplicados", "filas duplicadas eliminadas", n_duplicados])

# 3. Estandarizar categorías: "martes" -> "Martes", "Miercoles" -> "Miércoles",
#    "OFICINAS" -> "Oficinas".
estandarizar(df_energia, "dia_semana", {
    "lunes": "Lunes", "martes": "Martes",
    "miércoles": "Miércoles", "miercoles": "Miércoles",
    "jueves": "Jueves", "viernes": "Viernes",
}, reporte_energia)
estandarizar(df_energia, "area", {
    "oficinas": "Oficinas", "laboratorio": "Laboratorio", "servidores": "Servidores",
}, reporte_energia)

# 4. Convertir la fecha a tipo fecha (datetime).
df_energia["fecha"] = pd.to_datetime(df_energia["fecha"])
reporte_energia.append(["Tipos", "'fecha' convertida a datetime", len(df_energia)])

# 5. Rellenar nulos con la mediana (temperatura, humedad, equipos_activos).
imputar_mediana(df_energia, ["temperatura", "humedad", "equipos_activos"],
                reporte_energia)

# 6. Recortar outliers: E013 tenía temperatura 500 y consumo ¡9999 kWh!
recortar_outliers(
    df_energia,
    ["temperatura", "humedad", "personas", "equipos_activos", "consumo_kwh"],
    reporte_energia
)

# 7. La hora queda como entero.
df_energia["hora"] = df_energia["hora"].astype(int)

reporte_energia.append(
    ["Filas", "total: originales -> limpias", f"{n_originales} -> {len(df_energia)}"]
)

reporte_energia = pd.DataFrame(
    reporte_energia, columns=["Paso", "Acción aplicada", "Cantidad"]
)


# 8. Escribir el archivo LIMPIO.
with pd.ExcelWriter("03_Consumo_Energia_limpio.xlsx", engine="openpyxl") as writer:
    df_energia.to_excel(writer, sheet_name="Datos_Limpios", index=False)
    reporte_energia.to_excel(writer, sheet_name="Reporte_Limpieza", index=False)
    pd.read_excel("03_Consumo_Energia.xlsx", "Diccionario").to_excel(
        writer, sheet_name="Diccionario", index=False)
    pd.read_excel("03_Consumo_Energia.xlsx", "Instrucciones").to_excel(
        writer, sheet_name="Instrucciones", index=False)
    pd.read_excel("03_Consumo_Energia.xlsx", "Referencia").to_excel(
        writer, sheet_name="Referencia", index=False)

print("03_Consumo_Energia_limpio.xlsx creado.")


# ============================================================
# RESUMEN FINAL
# ============================================================
print("\nLimpieza completada: los 3 archivos LIMPIOS fueron creados.")