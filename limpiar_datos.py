"""
Limpieza de datos: lee los 3 archivos "sucios" (Datos_Crudos) y
genera versiones limpias en archivos nuevos (*_limpio.xlsx).

Limpiezas aplicadas:
  1. Duplicados    -> se eliminan filas repetidas.
  2. Nulos          -> se imputan con la MEDIANA de la columna.
  3. Categorías     -> se estandarizan (minúsculas/mayúsculas, tildes, alias)
                       usando la hoja "Referencia" como guía.
  4. Outliers       -> se recortan (winsorización) con el método IQR:
                       límites = Q1 - 1.5*IQR  y  Q3 + 1.5*IQR.

Cada archivo limpio contiene las hojas:
  Datos_Limpios   -> datos ya corregidos
  Reporte_Limpieza -> registro de todo lo que se hizo
  Diccionario / Referencia / Instrucciones -> copiadas del original
"""

import pandas as pd

ORIGINALES = [
    "01_Transporte_EMTU.xlsx",
    "02_Mantenimiento_Industrial.xlsx",
    "03_Consumo_Energia.xlsx",
]
LIMPIOS = [
    "01_Transporte_EMTU_limpio.xlsx",
    "02_Mantenimiento_Industrial_limpio.xlsx",
    "03_Consumo_Energia_limpio.xlsx",
]


def leer(archivo):
    return pd.read_excel(archivo, sheet_name="Datos_Crudos")


def reportar(logs):
    return pd.DataFrame(logs, columns=["Paso", "Acción aplicada", "Cantidad"])


def estandarizar(df, col, mapeo, logs):
    """Estandariza una columna categórica usando un mapeo de alias (lower-case key)."""
    original = df[col]
    normalizado = original.str.strip().str.lower()
    corregido = normalizado.map(mapeo)
    no_mapeadas = corregido.isna() & original.notna()
    if no_mapeadas.any():
        raros = original[no_mapeadas].unique()
        raise ValueError(f"Columna '{col}': valores sin mapeo: {raros}")
    df[col] = corregido
    n_cambios = int((original != df[col]).sum())
    logs.append(["Categorías", f"'{col}' estandarizada a valores válidos", n_cambios])


def imputar_mediana(df, cols, logs):
    for c in cols:
        n = int(df[c].isna().sum())
        if n > 0:
            med = df[c].median()
            df[c] = df[c].fillna(med)
            logs.append(["Nulos", f"'{c}' imputados con la mediana ({med:.1f})", n])


def recortar_outliers_iqr(df, cols, logs):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
        q1 = df[c].quantile(0.25)
        q3 = df[c].quantile(0.75)
        iqr = q3 - q1
        lim_inf = q1 - 1.5 * iqr
        lim_sup = q3 + 1.5 * iqr
        fuera = (df[c] < lim_inf) | (df[c] > lim_sup)
        n = int(fuera.sum())
        if n > 0:
            antes = df.loc[fuera, c].tolist()
            df.loc[fuera, c] = df[c].clip(lower=lim_inf, upper=lim_sup)[fuera]
            despues = df.loc[fuera, c].tolist()
            logs.append(
                ["Outliers",
                 f"'{c}' recortados a [{lim_inf:.1f}, {lim_sup:.1f}] "
                 f"(antes={antes}, después={despues})",
                 n]
            )


def limpiar_transporte():
    df = leer(ORIGINALES[0])
    n_orig = len(df)
    logs = []

    n_dup = int(df.duplicated().sum())
    if n_dup:
        df = df.drop_duplicates()
    logs.append(["Duplicados", "filas duplicadas eliminadas", n_dup])

    estandarizar(df, "linea", {
        "l1": "L1", "linea 1": "L1", "línea 1": "L1",
        "l2": "L2", "l-2": "L2", "linea 2": "L2", "línea 2": "L2",
        "l3": "L3", "linea 3": "L3", "línea 3": "L3",
    }, logs)
    estandarizar(df, "turno", {
        "mañana": "Mañana", "manana": "Mañana",
        "tarde": "Tarde", "noche": "Noche",
    }, logs)
    estandarizar(df, "lluvia", {
        "sí": "Sí", "si": "Sí", "no": "No",
    }, logs)

    df["fecha"] = pd.to_datetime(df["fecha"])

    imputar_mediana(df, ["pasajeros", "temperatura", "mantenimiento_dias"], logs)
    recortar_outliers_iqr(
        df, ["pasajeros", "temperatura", "mantenimiento_dias"], logs
    )
    df["retraso_mayor_10"] = df["retraso_mayor_10"].astype(int)

    logs.append(["Filas", "total: originales -> limpias", f"{n_orig} -> {len(df)}"])
    logs.append(["Tipos", "'fecha' convertida a fecha (datetime)", len(df)])
    return df, reportar(logs)


def limpiar_maquinas():
    df = leer(ORIGINALES[1])
    n_orig = len(df)
    logs = []

    n_dup = int(df.duplicated().sum())
    if n_dup:
        df = df.drop_duplicates()
    logs.append(["Duplicados", "filas duplicadas eliminadas", n_dup])

    estandarizar(df, "tipo_maquina", {
        "torno": "Torno", "fresadora": "Fresadora", "prensa": "Prensa",
    }, logs)
    estandarizar(df, "lubricacion", {
        "buena": "Buena", "media": "Media", "baja": "Baja",
    }, logs)

    imputar_mediana(df, ["temperatura", "vibracion", "presion"], logs)
    recortar_outliers_iqr(
        df,
        ["temperatura", "vibracion", "presion", "horas_operacion", "mantenimiento_dias"],
        logs,
    )
    df["falla_24h"] = df["falla_24h"].astype(int)

    logs.append(["Filas", "total: originales -> limpias", f"{n_orig} -> {len(df)}"])
    return df, reportar(logs)


def limpiar_energia():
    df = leer(ORIGINALES[2])
    n_orig = len(df)
    logs = []

    n_dup = int(df.duplicated().sum())
    if n_dup:
        df = df.drop_duplicates()
    logs.append(["Duplicados", "filas duplicadas eliminadas", n_dup])

    estandarizar(df, "dia_semana", {
        "lunes": "Lunes", "martes": "Martes",
        "miércoles": "Miércoles", "miercoles": "Miércoles",
        "jueves": "Jueves", "viernes": "Viernes",
    }, logs)
    estandarizar(df, "area", {
        "oficinas": "Oficinas", "laboratorio": "Laboratorio", "servidores": "Servidores",
    }, logs)

    df["fecha"] = pd.to_datetime(df["fecha"])

    imputar_mediana(df, ["temperatura", "humedad", "equipos_activos"], logs)
    recortar_outliers_iqr(
        df,
        ["temperatura", "humedad", "personas", "equipos_activos", "consumo_kwh"],
        logs,
    )
    df["hora"] = df["hora"].astype(int)

    logs.append(["Filas", "total: originales -> limpias", f"{n_orig} -> {len(df)}"])
    logs.append(["Tipos", "'fecha' convertida a fecha (datetime)", len(df)])
    return df, reportar(logs)


def guardar(archivo_origen, archivo_limpio, df_limpio, reporte):
    """Crea el archivo limpio: Datos_Limpios + Reporte_Limpieza + hojas copiadas."""
    with pd.ExcelWriter(archivo_limpio, engine="openpyxl") as writer:
        df_limpio.to_excel(writer, sheet_name="Datos_Limpios", index=False)
        reporte.to_excel(writer, sheet_name="Reporte_Limpieza", index=False)
        for hoja in ["Diccionario", "Instrucciones", "Referencia"]:
            pd.read_excel(archivo_origen, sheet_name=hoja).to_excel(
                writer, sheet_name=hoja, index=False
            )
    print(f"  -> {archivo_limpio} creado.")


def main():
    resultados = [
        limpiar_transporte(),
        limpiar_maquinas(),
        limpiar_energia(),
    ]
    for origen, limpio, (df, reporte) in zip(ORIGINALES, LIMPIOS, resultados):
        print(f"\n=== {origen} ===")
        print(reporte.to_string(index=False))
        print(f"Datos limpios: {len(df)} filas x {len(df.columns)} columnas")
        guardar(origen, limpio, df, reporte)

    print("\nLos 3 archivos LIMPIOS fueron creados.")


if __name__ == "__main__":
    main()