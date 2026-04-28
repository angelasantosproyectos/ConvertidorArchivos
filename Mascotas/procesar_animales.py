"""
procesar_animales.py
====================
Script para procesar el archivo Excel de perfiles de animales de compañía
y generar un CSV final con transformaciones y enriquecimiento de datos.

Uso:
    python procesar_animales.py

Dependencias:
    pip install pandas openpyxl
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date
import pandas as pd


# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

# Separador del CSV de salida (cambiar a "," si se prefiere)
CSV_SEPARATOR = ";"

# Encoding de salida
CSV_ENCODING = "utf-8"

# Columnas de fecha requeridas en el archivo de entrada
COL_FECHA_NAC = "Fec_nacimiento"
COL_FECHA_EFECTO = "Fec_efecto_spto"

# Columnas de raza en el archivo de entrada
COL_RAZA1 = "Raza_1"
COL_RAZA2 = "Raza_2"

# Columna que se usa como base para D_s_a_Asis_Vet_Premium
# Si en el futuro se cambia la lógica, modificar aquí.
COL_ASIS_VET = "D_s_a_Asis_Vet"


# ---------------------------------------------------------------------------
# SELECCIÓN DE ARCHIVOS (tkinter)
# ---------------------------------------------------------------------------

def seleccionar_archivo_entrada() -> str:
    """Abre un diálogo para seleccionar el Excel de entrada."""
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel de entrada",
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
    )
    root.destroy()
    if not ruta:
        print("❌ No se seleccionó ningún archivo de entrada. Saliendo.")
        sys.exit(0)
    return ruta


def seleccionar_archivo_razas() -> str:
    """Abre un diálogo para seleccionar el Excel de correspondencias de razas."""
    root = tk.Tk()
    root.withdraw()
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo de correspondencias de razas (RazaCod)",
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
    )
    root.destroy()
    if not ruta:
        print("❌ No se seleccionó el archivo de razas. Saliendo.")
        sys.exit(0)
    return ruta


def seleccionar_archivo_salida() -> str:
    """Abre un diálogo para elegir nombre y ubicación del CSV de salida."""
    root = tk.Tk()
    root.withdraw()
    ruta = filedialog.asksaveasfilename(
        title="Guarda el CSV de salida",
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
    )
    root.destroy()
    if not ruta:
        print("❌ No se seleccionó destino de salida. Saliendo.")
        sys.exit(0)
    return ruta


# ---------------------------------------------------------------------------
# LECTURA DE DATOS
# ---------------------------------------------------------------------------

def leer_excel_entrada(ruta: str) -> pd.DataFrame:
    """Lee el archivo Excel principal con pandas."""
    try:
        df = pd.read_excel(ruta)
        print(f"✅ Archivo de entrada cargado: {df.shape[0]} filas × {df.shape[1]} columnas")
        return df
    except FileNotFoundError:
        print(f"❌ Archivo no encontrado: {ruta}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error al leer el archivo de entrada: {e}")
        sys.exit(1)


def leer_excel_razas(ruta: str) -> dict:
    """
    Lee el archivo de correspondencias de razas y devuelve dos diccionarios:
        - raza1_map: {nombre_raza -> código_int}  (basado en columnas Raza_1 / Raza_1.1)
        - raza2_map: {nombre_raza -> código_int}  (basado en columnas Raza_2 / Raza_2.1)
    """
    try:
        df_razas = pd.read_excel(ruta)
    except FileNotFoundError:
        print(f"❌ Archivo de razas no encontrado: {ruta}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error al leer el archivo de razas: {e}")
        sys.exit(1)

    # --- Mapa Raza 1 ---
    # Columnas esperadas: 'Raza_1' (nombre) y 'Raza_1.1' (código)
    if "Raza_1" not in df_razas.columns or "Raza_1.1" not in df_razas.columns:
        print("⚠️  Columnas de Raza_1 no encontradas en el archivo de razas.")
        raza1_map = {}
    else:
        df_r1 = df_razas[["Raza_1", "Raza_1.1"]].dropna(subset=["Raza_1", "Raza_1.1"])
        df_r1 = df_r1.drop_duplicates(subset=["Raza_1"])
        raza1_map = dict(zip(df_r1["Raza_1"].astype(str).str.strip(), df_r1["Raza_1.1"].astype(int)))

    # --- Mapa Raza 2 ---
    # En el archivo de razas: 'Raza_2' contiene el CÓDIGO (float) y 'Raza_2.1' el NOMBRE (str).
    # El mapeo que necesitamos es nombre → código, por tanto: Raza_2.1 → Raza_2.
    if "Raza_2" not in df_razas.columns or "Raza_2.1" not in df_razas.columns:
        print("⚠️  Columnas de Raza_2 no encontradas en el archivo de razas.")
        raza2_map = {}
    else:
        df_r2 = df_razas[["Raza_2.1", "Raza_2"]].dropna(subset=["Raza_2.1", "Raza_2"])
        df_r2 = df_r2.drop_duplicates(subset=["Raza_2.1"])
        raza2_map = dict(zip(df_r2["Raza_2.1"].astype(str).str.strip(), df_r2["Raza_2"].astype(int)))

    print(f"✅ Mapa de razas cargado: {len(raza1_map)} entradas Raza1, {len(raza2_map)} entradas Raza2")
    return raza1_map, raza2_map


# ---------------------------------------------------------------------------
# VALIDACIONES
# ---------------------------------------------------------------------------

def validar_columnas(df: pd.DataFrame, columnas_requeridas: list) -> None:
    """Verifica que las columnas necesarias existan en el DataFrame."""
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        print(f"❌ Columnas requeridas no encontradas en el archivo: {faltantes}")
        sys.exit(1)
    print(f"✅ Todas las columnas requeridas están presentes.")


# ---------------------------------------------------------------------------
# TRANSFORMACIONES
# ---------------------------------------------------------------------------

def calcular_edad_anios(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la edad en años completos desde Fec_nacimiento hasta hoy.
    Tiene en cuenta día y mes para no redondear al alza antes del cumpleaños.
    Resultado en columna nueva: NUM_EDAD_ANIO
    """
    hoy = date.today()

    def edad_exacta(fec_nac):
        if pd.isna(fec_nac):
            return None
        try:
            nac = fec_nac.date() if hasattr(fec_nac, "date") else pd.Timestamp(fec_nac).date()
            anios = hoy.year - nac.year
            # Restar 1 si aún no ha llegado el cumpleaños este año
            if (hoy.month, hoy.day) < (nac.month, nac.day):
                anios -= 1
            return anios
        except Exception:
            return None

    df["NUM_EDAD_ANIO"] = df[COL_FECHA_NAC].apply(edad_exacta)
    return df


def calcular_asis_vet_premium(df: pd.DataFrame) -> pd.DataFrame:
    """
    PLACEHOLDER — D_s_a_Asis_Vet_Premium.
    Por ahora se copia el valor de D_s_a_Asis_Vet como punto de partida.
    Cuando se defina la lógica exacta, modificar esta función.
    """
    if COL_ASIS_VET in df.columns:
        df["D_s_a_Asis_Vet_Premium"] = df[COL_ASIS_VET]
    else:
        # Si la columna origen no existe, se crea vacía documentada
        df["D_s_a_Asis_Vet_Premium"] = None
        print(f"⚠️  Columna '{COL_ASIS_VET}' no encontrada; D_s_a_Asis_Vet_Premium queda vacía.")
    return df


def limpiar_regimen(df: pd.DataFrame) -> pd.DataFrame:
    """
    Regimen: extrae solo la letra inicial antes del guión.
      'J-Tomador jurídico' → 'J'
      'F-Tomador físico'   → 'F'
      Cualquier otro valor → '0'
    """
    if "Regimen" not in df.columns:
        return df
    df["Regimen"] = (
        df["Regimen"]
        .astype(str)
        .str.strip()
        .str.extract(r"^([JF])", expand=False)
        .fillna("0")
    )
    return df


def limpiar_modalidad(df: pd.DataFrame) -> pd.DataFrame:
    """
    Modalidad: mantiene solo el código numérico inicial (antes del guión).
      '200132-Selección' → '200132'
      Sin coincidencia    → '0'
    """
    if "Modalidad" not in df.columns:
        return df
    df["Modalidad"] = (
        df["Modalidad"]
        .astype(str)
        .str.strip()
        .str.extract(r"^(\d+)", expand=False)
        .fillna("0")
    )
    return df


def limpiar_forma_pago(df: pd.DataFrame) -> pd.DataFrame:
    """
    Forma_Pago: mantiene solo el dígito inicial (1, 2 o 3).
      '1-Anual'       → '1'
      '3-Trimestral'  → '3'
      Sin coincidencia → '0'
    """
    if "Forma_Pago" not in df.columns:
        return df
    df["Forma_Pago"] = (
        df["Forma_Pago"]
        .astype(str)
        .str.strip()
        .str.extract(r"^([123])", expand=False)
        .fillna("0")
    )
    return df


def limpiar_especie_y_cod_clase(df: pd.DataFrame) -> pd.DataFrame:
    """
    Especie: el código puede ser de 1 o 2 dígitos antes del guión.
      '1-Perro'         → Especie='1', COD_CLASE_OTROS='0'
      '2-Gatos'         → Especie='2', COD_CLASE_OTROS='0'
      '31-Aves Exóticas'→ Especie='3', COD_CLASE_OTROS='1'
      '32-Aves Rapaces' → Especie='3', COD_CLASE_OTROS='2'

    Regla general para códigos de 2 dígitos (XY-...):
      Especie        → primer dígito  (X)
      COD_CLASE_OTROS→ segundo dígito (Y)

    Sin coincidencia → ambos quedan '0'.
    """
    if "Especie" not in df.columns:
        df["COD_CLASE_OTROS"] = "0"
        return df

    raw = df["Especie"].astype(str).str.strip()

    # Extraer la parte numérica antes del guión (1 o 2 dígitos)
    codigo = raw.str.extract(r"^(\d+)", expand=False).fillna("")

    # Especie = primer dígito del código
    df["Especie"] = codigo.str[:1].replace("", "0")

    # COD_CLASE_OTROS = segundo dígito si existe, si no '0'
    df["COD_CLASE_OTROS"] = codigo.str[1:2].replace("", "0").fillna("0")
    df["COD_CLASE_OTROS"] = df["COD_CLASE_OTROS"].where(
        df["COD_CLASE_OTROS"] != "", "0"
    )

    return df


def limpiar_provincia(df: pd.DataFrame) -> pd.DataFrame:
    """
    Provincia: mantiene solo el código numérico antes de los dos puntos.
      '50: Zaragoza' → '50'
      Sin coincidencia → '0'
    """
    if "Provincia" not in df.columns:
        return df
    df["Provincia"] = (
        df["Provincia"]
        .astype(str)
        .str.strip()
        .str.extract(r"^(\d+)", expand=False)
        .fillna("0")
    )
    return df


def limpiar_booleano_sn(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """
    Para columnas tipo S/N (Perro_Peligroso, Dec_Salud):
      'S-Sí', 'S' → 'S'
      'N-No', 'N' → 'N'
      Cualquier otro → '0'
    """
    if columna not in df.columns:
        return df
    raw = df[columna].astype(str).str.strip()
    df[columna] = raw.str.extract(r"^([SN])", expand=False).fillna("0")
    return df


def limpiar_sexo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sexo: mantiene solo M o H.
      'M-Macho'  → 'M'
      'H-Hembra' → 'H'
      Otro       → '0'
    """
    if "Sexo" not in df.columns:
        return df
    df["Sexo"] = (
        df["Sexo"]
        .astype(str)
        .str.strip()
        .str.extract(r"^([MH])", expand=False)
        .fillna("0")
    )
    return df


def limpiar_uso_animal(df: pd.DataFrame) -> pd.DataFrame:
    """
    Uso_animal: mantiene solo F o A.
      'F-Familiar' → 'F'
      'A-...'      → 'A'
      Otro         → '0'
    """
    if "Uso_animal" not in df.columns:
        return df
    df["Uso_animal"] = (
        df["Uso_animal"]
        .astype(str)
        .str.strip()
        .str.extract(r"^([FA])", expand=False)
        .fillna("0")
    )
    return df


def limpiar_clase(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clase: mantiene solo el dígito inicial (1 o 2).
      '1-Pura Raza'      → '1'
      '2-Mestizo o Cruce' → '2'
      Otro               → '0'
    """
    if "Clase" not in df.columns:
        return df
    df["Clase"] = (
        df["Clase"]
        .astype(str)
        .str.strip()
        .str.extract(r"^([12])", expand=False)
        .fillna("0")
    )
    return df


def formatear_fecha_efecto(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte Fec_efecto_spto a formato ISO 8601 con zona horaria fija +0100.
    Ejemplo: 2025-05-14T00:00:00.000+0100
    """
    df["FEC_EFECTO_SPTO_FORMATEADA"] = pd.to_datetime(
        df[COL_FECHA_EFECTO], errors="coerce"
    ).dt.strftime("%Y-%m-%dT00:00:00.000+0100")
    return df


def mapear_razas(df: pd.DataFrame, raza1_map: dict, raza2_map: dict) -> pd.DataFrame:
    """
    Añade columnas RazaCod1 y RazaCod2 con los códigos numéricos
    correspondientes a los nombres de Raza_1 y Raza_2.
    Las razas sin correspondencia quedan como None.
    """
    if COL_RAZA1 in df.columns:
        df["RazaCod1"] = df[COL_RAZA1].astype(str).str.strip().map(raza1_map)
    else:
        df["RazaCod1"] = None
        print(f"⚠️  Columna '{COL_RAZA1}' no encontrada; RazaCod1 queda vacía.")

    if COL_RAZA2 in df.columns:
        df["RazaCod2"] = df[COL_RAZA2].astype(str).str.strip().map(raza2_map)
    else:
        df["RazaCod2"] = None
        print(f"⚠️  Columna '{COL_RAZA2}' no encontrada; RazaCod2 queda vacía.")

    return df


def normalizar_decimales(df: pd.DataFrame) -> pd.DataFrame:
    """
    En columnas de tipo object que contengan números con coma como separador decimal,
    reemplaza la coma por punto y convierte a float.
    Columnas ya numéricas se dejan intactas.
    """
    for col in df.columns:
        if df[col].dtype == object:
            # Intentar convertir sólo si la columna parece numérica con comas
            muestra = df[col].dropna().astype(str).head(20)
            if muestra.str.match(r"^-?\d+,\d+$").any():
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace(",", ".", regex=False)
                    .pipe(pd.to_numeric, errors="coerce")
                )
    return df


# ---------------------------------------------------------------------------
# EXPORTACIÓN
# ---------------------------------------------------------------------------

def exportar_csv(df: pd.DataFrame, ruta_salida: str) -> None:
    """Exporta el DataFrame a CSV con separador ';' y encoding UTF-8."""
    try:
        df.to_csv(ruta_salida, sep=CSV_SEPARATOR, encoding=CSV_ENCODING, index=False)
        print(f"✅ CSV generado correctamente en: {ruta_salida}")
    except Exception as e:
        print(f"❌ Error al guardar el CSV: {e}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# RESUMEN FINAL
# ---------------------------------------------------------------------------

def mostrar_resumen(df: pd.DataFrame, columnas_nuevas: list) -> None:
    """Muestra un resumen de las transformaciones realizadas."""
    print("\n" + "=" * 55)
    print("  RESUMEN DEL PROCESAMIENTO")
    print("=" * 55)
    print(f"  Filas procesadas   : {df.shape[0]:,}")
    print(f"  Columnas totales   : {df.shape[1]}")
    print(f"  Columnas añadidas  : {len(columnas_nuevas)}")
    for col in columnas_nuevas:
        no_nulos = df[col].notna().sum()
        print(f"    · {col:<35} {no_nulos:>6} valores no nulos")
    print("=" * 55 + "\n")


# ---------------------------------------------------------------------------
# FLUJO PRINCIPAL
# ---------------------------------------------------------------------------

def main():
    print("\n🐾 Procesador de perfiles — Animales de Compañía\n")

    # 1. Selección de archivos
    print("→ Selecciona el archivo Excel de ENTRADA...")
    ruta_entrada = seleccionar_archivo_entrada()

    print("→ Selecciona el archivo de correspondencias de RAZAS...")
    ruta_razas = seleccionar_archivo_razas()

    print("→ Selecciona la ubicación y nombre del CSV de SALIDA...")
    ruta_salida = seleccionar_archivo_salida()

    # 2. Lectura
    df = leer_excel_entrada(ruta_entrada)
    raza1_map, raza2_map = leer_excel_razas(ruta_razas)

    # 3. Validación de columnas mínimas requeridas
    columnas_requeridas = [COL_FECHA_NAC, COL_FECHA_EFECTO]
    validar_columnas(df, columnas_requeridas)

    # 4. Transformaciones
    columnas_nuevas = [
        "NUM_EDAD_ANIO",
        "D_s_a_Asis_Vet_Premium",
        "COD_CLASE_OTROS",
        "FEC_EFECTO_SPTO_FORMATEADA",
        "RazaCod1",
        "RazaCod2",
    ]

    print("\n→ Aplicando transformaciones...")

    # — Limpieza de columnas existentes —
    df = limpiar_regimen(df)
    df = limpiar_modalidad(df)
    df = limpiar_forma_pago(df)
    df = limpiar_especie_y_cod_clase(df)   # también genera COD_CLASE_OTROS
    df = limpiar_provincia(df)
    df = limpiar_booleano_sn(df, "Perro_Peligroso")
    df = limpiar_booleano_sn(df, "Dec_Salud")
    df = limpiar_sexo(df)
    df = limpiar_uso_animal(df)
    df = limpiar_clase(df)

    # — Nuevas columnas calculadas —
    df = calcular_edad_anios(df)
    df = calcular_asis_vet_premium(df)
    df = formatear_fecha_efecto(df)
    df = mapear_razas(df, raza1_map, raza2_map)
    df = normalizar_decimales(df)

    # — Rellenar campos vacíos con 0 —
    # Las columnas de texto reciben '0'; las numéricas reciben 0
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("0")
        else:
            df[col] = df[col].fillna(0)

    # 5. Exportación
    exportar_csv(df, ruta_salida)

    # 6. Resumen
    mostrar_resumen(df, columnas_nuevas)


if __name__ == "__main__":
    main()