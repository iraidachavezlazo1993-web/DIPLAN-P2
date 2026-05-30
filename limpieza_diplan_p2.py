# -*- coding: utf-8 -*-
"""
================================================================================
 DIPLAN-P2 | Pipeline de limpieza, normalizacion y consolidacion de bases
================================================================================

Autor      : Equipo DIPLAN
Descripcion:
    Lee todas las bases (libros de Excel) entregadas, las limpia segun reglas
    de negocio, las normaliza y las organiza en 11 GRUPOS definidos en el
    insumo "Ordenamiento_ultimo.xlsx" (hoja "Grupos" / "Resumen").

Reglas de limpieza aplicadas a todas las bases (salvo insumos):
    1.  Fechas          -> formato DD/MM/YYYY (sin hora).
    2.  Montos/Devengado-> numerico puro (sin letras, sin simbolos S/).
    3.  cod_local       -> un solo codigo por celda, 6 digitos (ceros izq.).
                           Si falta, se imputa cruzando con "vinculaciones"
                           (via cui o cod_mod).
    4.  cod_modular     -> 7 digitos; SE PERMITE apilado (varios por celda)
                           porque asi corresponde a la realidad modular.
    5.  cui             -> 7 digitos (ceros izq.).
    6.  Campos de texto -> normalizados (trim, espacios colapsados, sin saltos
                           de linea, vacios/placeholders -> nulo).
    7.  Campos COMENTARIO-> se conservan completos (solo trim), porque suelen
                           contener mayor informacion.

Insumos (NO se limpian, solo se usan):
    - Ordenamiento_ultimo.xlsx           (define los 11 grupos)
    - df_vinculaciones_updated_*.xlsx    (cruce cui/cod_mod -> cod_local)

NOTA: El Grupo 2 (PI) NO se procesa: el cliente ya lo tiene limpio.

Salidas (carpeta ./output):
    - bases_limpias/Grupo_XX_<NOMBRE>/<df>.xlsx  y  <df>.parquet
    - base_final/base_final_armonizada.xlsx / .parquet
    - base_final/base_final_union_completa.xlsx / .parquet
    - diccionario_datos.xlsx
    - reporte_limpieza.xlsx
================================================================================
"""

import os
import re
import glob
import unicodedata
import warnings
from datetime import datetime, date

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ------------------------------------------------------------------ #
#  RUTAS
# ------------------------------------------------------------------ #
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
OUT_DIR    = os.path.join(BASE_DIR, "output")
LIMPIAS_DIR = os.path.join(OUT_DIR, "bases_limpias")
FINAL_DIR  = os.path.join(OUT_DIR, "base_final")
for d in (OUT_DIR, LIMPIAS_DIR, FINAL_DIR):
    os.makedirs(d, exist_ok=True)

ARCHIVO_VINCULACIONES = "df_vinculaciones_updated_20260527.xlsx"
ARCHIVO_ORDENAMIENTO  = "Ordenamiento_ultimo.xlsx"
ARCHIVO_UBIGEO        = "ubigeo_UGEL.xlsx"   # ubigeo enriquecido con DRE/UGEL
ARCHIVO_PADRON        = "Copia de Padron_web.csv"   # padron web nacional (insumo)
ARCHIVO_PRONIED       = "ubigeo_pronied.xlsx"        # ubigeo PRONIED (cod_local)
ARCHIVO_UBIGEO_CUI    = "ubigeo_cui.csv"             # ubigeo por CUI (insumo)

# ------------------------------------------------------------------ #
#  NOMBRES DE LOS 11 GRUPOS (insumo Ordenamiento, hoja "Grupos")
# ------------------------------------------------------------------ #
GRUPOS = {
    1:  "MOBILIARIO Y EQUIPO",
    2:  "PI",
    3:  "MODULOS",
    4:  "ACONDICIONAMIENTO",
    5:  "ASISTENCIAS-SIMILAR",
    6:  "MANTENIMIENTO",
    7:  "SANEAMIENTO",
    8:  "PROGRAMAS PRESUPUESTALES",
    9:  "OTROS",
    10: "INSPECCIONES",
    11: "DEMOLICIONES",
}
GRUPO_OMITIR = None   # el Grupo 2 ahora se incorpora con su base final ya limpia

# ------------------------------------------------------------------ #
#  FUENTE (unidad/insumo que reporta principalmente la base)
#  Para Anexo 1 y 2 la fuente se toma por fila del campo "Remitente".
# ------------------------------------------------------------------ #
FUENTE_LIBRO = {
    "ANIN.xlsx": "ANIN",
    "FONCODES.xlsx": "FONCODES",
    "PEIP.xlsx": "PEIP",
    "PEIP_solomant.xlsx": "PEIP",
    "UGM.xlsx": "UGM",
    "UGME.xlsx": "UGME",
    "UGRD.xlsx": "UGRD",
    "UGSC.xlsx": "UGSC",
    "2026.03.30 Zonales_UZ.xlsx": "Unidades Zonales (UZ)",
    "DIGEGED.xlsx": "DIGEGED",
    "DRELM.xlsx": "DRELM",
    "UE118.xlsx": "UE118",
    "UGEO.xlsx": "UGEO",
    "df_pi_listados_final_v7.xlsx": "PI (Banco de Inversiones)",
}

# ------------------------------------------------------------------ #
#  CONFIGURACION DE BASES A PROCESAR
#  (derivada del insumo Ordenamiento_ultimo.xlsx, hoja "Resumen")
#  header = None  -> autodeteccion de la fila de encabezado
#  anio   = valor a forzar cuando la base no tiene campo fecha
# ------------------------------------------------------------------ #
ANEXO1 = ("Anexo 01 - Solicitud de información de intervenciones - inversiones "
          "- Consolidado 300426.xlsx")
ANEXO2 = ("Anexo 02 - Solicitud de información de intervenciones - no inversiones "
          "- Consolidado 300426_depurada.xlsx")

CONFIG = [
    # archivo, hoja, df_name, nro, grupo, anio
    ("ANIN.xlsx",                "MANTENIMIENTO",                "df_anin_mantenimiento",      1,  6,  None),
    ("FONCODES.xlsx",            "REPORTE_MANT_ACOND_2025",      "df_foncode_2025",            3,  6,  2025),
    ("FONCODES.xlsx",            "REPORTE_MANT_ACOND_2026",      "df_foncode_2026",            4,  6,  2026),
    ("PEIP.xlsx",                "CONTINGENCIA",                 "df_peip_contingencia",       7,  3,  None),
    ("PEIP_solomant.xlsx",       "MANTENIMIENTO",                "df_peip_mantenimiento",      8,  6,  None),
    ("UGM.xlsx",                 "ACONDICIONAMIENTO",            "df_ugm_acondicionamiento",   12, 4,  None),
    ("UGM.xlsx",                 "MANTENIMIENTO 2025",           "df_ugm_mantenimiento_2025",  13, 6,  2025),
    ("UGM.xlsx",                 "MANTENIMIENTO 2026",           "df_ugm_mantenimiento_2026",  14, 6,  2026),
    ("UGM.xlsx",                 "ACCESIBILIDAD 2026-2",         "df_ugm_accesibilidad_2026",  46, 4,  2026),
    ("UGME.xlsx",                "SISTEMAS MODULARES",           "df_ugme_modulares",          24, 3,  None),
    ("UGME.xlsx",                "MOBILIARIO Y EQUIPAMIENTO",    "df_ugme_mobiliario",         25, 1,  None),
    ("UGME.xlsx",                "PLAN DE CONSERVACIÓN MODULAR", "df_ugme_conservacion",       26, 3,  None),
    ("UGRD.xlsx",                "PIRCC",                        "df_ugrd_pircc_pircc",        63, 3,  None),
    ("UGRD.xlsx",                "MBR",                          "df_ugrd_mbr",                28, 3,  None),
    ("UGRD.xlsx",                "ME",                           "df_ugrd_me",                 29, 3,  None),
    ("UGSC.xlsx",                "ASITEC-SIAT",                  "df_ugsc_asitec",             30, 5,  None),
    ("UGSC.xlsx",                "SEGUIMIENTO DE PI FINANCIADOS ","df_ugsc_seguimiento",       31, 10, None),
    ("2026.03.30 Zonales_UZ.xlsx","INSPECCIONES",                "df_uz_inspecciones",         32, 10, None),
    ("2026.03.30 Zonales_UZ.xlsx","ASESORAMIENTO",               "df_uz_asesoramiento",        33, 5,  None),
    ("DIGEGED.xlsx",             "MTO",                          "df_digeged_mto",             35, 6,  None),
    ("DIGEGED.xlsx",             "SFL",                          "df_digeged_sfl",             36, 7,  None),
    ("DIGEGED.xlsx",             "PROGRAMA",                     "df_digeged_programa",        37, 8,  None),
    ("DIGEGED.xlsx",             "OTROS25",                      "df_digeged_otros25",         38, 9,  None),
    ("DRELM.xlsx",               "Mobiliario",                   "df_drelm_mobiliario",        40, 1,  2026),
    ("DRELM.xlsx",               "MANTENIMIENTO",                "df_drelm_mantenimiento",     42, 6,  None),
    ("DRELM.xlsx",               "MODULO",                       "df_drelm_modulo",            43, 3,  2026),
    ("DRELM.xlsx",               "SFL",                          "df_drelm_sfl",               44, 7,  2026),
    ("DRELM.xlsx",               "SERVICIOS_OTROS",              "df_drelm_servicios",         45, 9,  2026),
    (ANEXO2,                     "A",                            "df_anexo2_a",                71, 3,  None),
    (ANEXO2,                     "B",                            "df_anexo2_b",                72, 1,  None),
    (ANEXO2,                     "C",                            "df_anexo2_c",                73, 6,  None),
    (ANEXO2,                     "D",                            "df_anexo2_d",                74, 11, None),
    (ANEXO2,                     "E",                            "df_anexo2_e",                75, 4,  None),
    # Grupo 2 (PI): base FINAL ya consolidada y limpia, se incorpora tal cual
    ("df_pi_listados_final_v7.xlsx", "con_cod_local",            "df_pi_listados_final_v7",    77, 2,  None),
]

# ================================================================== #
#  UTILIDADES DE NORMALIZACION
# ================================================================== #
NULL_TOKENS = {"", "-", "--", "·", ".", "n/a", "na", "s/d", "sd", "sin dato",
               "sin datos", "none", "null", "nan", "no aplica", "no aplica.",
               "(en blanco)", "(vacio)", "#n/a", "#¡valor!", "#ref!"}


def quitar_acentos(texto: str) -> str:
    """Elimina tildes/diacriticos para comparaciones robustas."""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def norm_key(texto) -> str:
    """Normaliza un nombre de columna a una clave comparable (sin acentos,
    minuscula, espacios colapsados)."""
    if texto is None:
        return ""
    t = quitar_acentos(str(texto)).lower()
    t = t.replace("\n", " ").replace("\r", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return t


def limpiar_nombre_columna(texto) -> str:
    """Limpia el nombre visible de una columna: sin saltos de linea, sin
    espacios dobles, recortado."""
    if texto is None:
        return ""
    t = str(texto).replace("\n", " ").replace("\r", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return t


def normalizar_texto(valor):
    """Normaliza un valor de texto disperso: trim + colapsa espacios + quita
    saltos de linea. Convierte placeholders de vacio a NaN."""
    if valor is None:
        return np.nan
    if isinstance(valor, float) and np.isnan(valor):
        return np.nan
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return valor
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return valor
    t = str(valor).replace("\n", " ").replace("\r", " ")
    t = re.sub(r"\s+", " ", t).strip()
    if norm_key(t) in NULL_TOKENS:
        return np.nan
    return t


def normalizar_comentario(valor):
    """Para campos COMENTARIO: solo recorta espacios extremos y colapsa
    espacios internos, pero CONSERVA el contenido (no descarta como nulo)."""
    if valor is None:
        return np.nan
    if isinstance(valor, float) and np.isnan(valor):
        return np.nan
    t = str(valor).replace("\r", " ").strip()
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n\s*", " | ", t).strip()   # apila multilinea en una sola
    if t == "" or norm_key(t) in {"-", "·", "."}:
        return np.nan
    return t


# ---------- Fechas -> DD/MM/YYYY -------------------------------------------- #
_PATRONES_FECHA = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y",
    "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y", "%d.%m.%Y",
]


def limpiar_fecha(valor):
    """Devuelve la fecha en formato DD/MM/YYYY (texto). Si no se puede
    interpretar, devuelve NaN (o el texto original si parece relevante)."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        try:
            return pd.Timestamp(valor).strftime("%d/%m/%Y")
        except Exception:
            return np.nan
    s = str(valor).strip()
    if norm_key(s) in NULL_TOKENS:
        return np.nan
    # numero serial de Excel
    if re.fullmatch(r"\d{4,6}(\.0+)?", s):
        try:
            base = datetime(1899, 12, 30)
            return (base + pd.to_timedelta(float(s), unit="D")).strftime("%d/%m/%Y")
        except Exception:
            pass
    s2 = s.split(" ")[0] if re.match(r"\d{4}-\d{2}-\d{2}", s) else s
    for fmt in _PATRONES_FECHA:
        try:
            return datetime.strptime(s2, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return np.nan


# ---------- Montos / Devengado -> numerico ---------------------------------- #
def limpiar_monto(valor):
    """Convierte un monto a numero puro. Maneja 'S/ 1,500.00', '1.234,56',
    textos con letras, etc. Si no hay numero -> NaN."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)
    s = str(valor).strip()
    if norm_key(s) in NULL_TOKENS:
        return np.nan
    s = s.replace("S/", "").replace("s/", "").replace("$", "")
    s = re.sub(r"[^\d,.\-]", "", s)            # deja digitos , . -
    if s in ("", "-", ".", ","):
        return np.nan
    # decide separador decimal
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):        # formato 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:                                  # formato 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        # coma como decimal si hay <=2 digitos despues
        if re.search(r",\d{1,2}$", s):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return round(float(s), 2)
    except Exception:
        return np.nan


# ---------- Codigos --------------------------------------------------------- #
def _solo_digitos(s):
    return re.sub(r"\D", "", str(s))


def limpiar_cod_local(valor):
    """Un unico codigo local, 6 digitos con ceros a la izquierda.
    Si la celda trae varios apilados, toma el primero valido."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    s = str(valor)
    if norm_key(s) in NULL_TOKENS:
        return np.nan
    # separa posibles multiples codigos
    candidatos = re.split(r"[\s,;/|\n]+|\]\[|\]|\[", s)
    for c in candidatos:
        d = _solo_digitos(c)
        if d and 1 <= len(d) <= 7:
            return d.zfill(6)
    d = _solo_digitos(s)
    return d.zfill(6) if d else np.nan


def limpiar_cui(valor):
    """CUI a 7 digitos con ceros a la izquierda (un solo codigo)."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    s = str(valor)
    if norm_key(s) in NULL_TOKENS:
        return np.nan
    d = _solo_digitos(s.split(".")[0])
    return d.zfill(7) if d else np.nan


def limpiar_cod_modular(valor):
    """Codigo(s) modular(es). SE PERMITE apilado. Cada codigo a 7 digitos.
    Devuelve los codigos separados por ' | '."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return np.nan
    s = str(valor)
    if norm_key(s) in NULL_TOKENS:
        return np.nan
    # extrae todas las secuencias de digitos de >=5 caracteres
    crudos = re.findall(r"\d{5,8}", s)
    if not crudos:
        # quiza vengan como '[A2 - 0257048]' -> ya cubierto; intenta cualquiera
        crudos = re.findall(r"\d+", s)
    codigos = []
    for c in crudos:
        c7 = c.zfill(7)
        if c7 not in codigos:
            codigos.append(c7)
    return " | ".join(codigos) if codigos else np.nan


# ================================================================== #
#  CANONIZACION DE NOMBRES DE VARIABLES (minuscula, snake_case)
# ================================================================== #
# Sinonimos -> nombre canonico (clave = norm_key del nombre original)
SINONIMOS_CANON = {
    "n°": "n", "nº": "n", "no": "n", "item": "item",
    "cui": "cui", "cui (iri)": "cui",
    "codigo unico de inversiones (cui)": "cui",
    "codigo unico de inversion": "cui", "codigo unico de inversiones": "cui",
    "codigo de inversion": "cui", "codigo de inversiones": "cui",
    "codigo local": "cod_local", "codigo del local educativo": "cod_local",
    "codigo de local educativo": "cod_local", "codigo de local": "cod_local",
    "cod_local": "cod_local", "code local": "cod_local",
    "codigo modular": "cod_mod", "codigos modulares": "cod_mod",
    "codigos modulares intervenidos": "cod_mod", "cod. modular": "cod_mod",
    "cod modular": "cod_mod", "cod_modular": "cod_mod", "cod_mod": "cod_mod",
    "codigo institucional": "cod_institucional",
    "nombre de la i.e.": "nombre_ie", "nombre de la ie": "nombre_ie",
    "nombre del iri": "nombre_ie", "nombre del peip": "nombre_ie",
    "nombre de institucion educativa": "nombre_ie",
    "nombre del servicio educativo": "nombre_ie", "local escolar": "nombre_ie",
    "region": "departamento", "departamento": "departamento",
    "nombre departamento": "departamento",
    "provincia": "provincia", "nombre provincia": "provincia",
    "distrito": "distrito", "nombre distrito": "distrito",
    "monto de la inversion (s/)": "monto", "monto de inversion (s/)": "monto",
    "monto contractual (s/)": "monto_contractual",
    "monto devengado": "devengado",
    "avance fisico (%)": "avance_fisico", "avance (%)": "avance",
    "fecha de inicio (o estimada)": "fecha_inicio",
    "fecha de culminacion (o estimada)": "fecha_culminacion",
    "comentarios": "comentario", "comentario": "comentario",
    "fuentes_concat": "fuentes_concat",
}


def canonizar_nombre(nombre) -> str:
    """Convierte un nombre de columna a su forma canonica: minuscula,
    sin acentos, snake_case; mapea sinonimos conocidos (p.ej. 'Codigo Unico
    de Inversiones' -> 'cui')."""
    k = norm_key(nombre)
    if k in SINONIMOS_CANON:
        return SINONIMOS_CANON[k]
    if "comentario" in k:
        return "comentario"
    if "observ" in k:
        return "observacion"
    s = quitar_acentos(str(nombre)).lower()
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    s = re.sub(r"_+", "_", s)
    if len(s) > 60:                       # recorta nombres descriptivos largos
        s = s[:60].rstrip("_")
    return s or "col"


def canonizar_columnas(df):
    """Renombra todas las columnas a su forma canonica (minuscula/snake_case),
    resolviendo colisiones con sufijos. Devuelve (df, mapping original->canonico)."""
    nuevos, usados, mapping = [], set(), {}
    for c in df.columns:
        base = canonizar_nombre(c)
        nombre, i = base, 1
        while nombre in usados:           # garantiza unicidad real
            i += 1
            nombre = f"{base}_{i}"
        usados.add(nombre)
        nuevos.append(nombre)
        mapping[c] = nombre
    df = df.copy()
    df.columns = nuevos
    return df, mapping


# ================================================================== #
#  CLASIFICACION DE COLUMNAS POR NOMBRE
# ================================================================== #
def es_col_fecha(nombre: str) -> bool:
    n = norm_key(nombre)
    return ("fecha" in n) or n in {"periodo inicial", "periodo final / actual"}


def es_col_monto(nombre: str) -> bool:
    n = norm_key(nombre)
    claves = ["monto", "devengado", "dev. acumulado", "dev acumulado",
              "dev_acum", "dev acum",
              "costo actualizado", "costo", "pim", "presupuesto",
              "inversion (s/", "monto de inversion", "monto contractual",
              "monto incurrido", "monto asignado", "monto transferido",
              "monto total"]
    return any(k in n for k in claves)


def es_col_comentario(nombre: str) -> bool:
    n = norm_key(nombre)
    return any(k in n for k in ["comentario", "observ", "obs", "riesgo",
                                "medidas de accion", "alertas", "detalle",
                                "titular / ocurrencia"])


def es_col_cod_local(nombre: str) -> bool:
    n = norm_key(nombre)
    return n in {"codigo local", "codigo del local educativo",
                 "codigo de local educativo", "code local", "cod_local",
                 "codigo de local"} or n == "codigo local "


def es_col_cui(nombre: str) -> bool:
    n = norm_key(nombre)
    return (n == "cui" or n.startswith("cui ") or "codigo unico de inversion" in n
            or n in {"cui (iri)", "codigo unico de inversiones (cui)"})


def es_col_cod_modular(nombre: str) -> bool:
    n = norm_key(nombre)
    return any(k in n for k in ["codigo modular", "codigos modulares",
                                "cod. modular", "cod modular", "cod_modular",
                                "cod_mod", "cod mod", "codigo institucional"])


# ================================================================== #
#  LECTURA Y DETECCION DE ENCABEZADO
# ================================================================== #
_TOKENS_HDR_EXACTO = {"n", "no", "nro", "cui", "item", "code", "id"}
_TOKENS_HDR_INICIO = ("codigo", "cod", "cui", "item", "relacion", "code")


def _es_celda_encabezado(valor) -> bool:
    """True si el texto luce como inicio de fila de encabezado (N°, Código,
    CUI, Item, ...). Robusto a 'N°', 'Nº', 'CODIGO LOCAL', etc."""
    limpio = re.sub(r"[^a-z0-9]", "", quitar_acentos(str(valor)).lower())
    if limpio in _TOKENS_HDR_EXACTO:
        return True
    return any(limpio.startswith(t) for t in _TOKENS_HDR_INICIO)


def detectar_fila_encabezado(df_raw: pd.DataFrame, max_scan: int = 12) -> int:
    """Detecta la fila de encabezado: primera fila (de arriba hacia abajo) cuya
    primera celda no vacia luce como encabezado y tiene >=3 celdas; si ninguna
    coincide, la fila con mas celdas no nulas entre las primeras filas."""
    mejor_idx, mejor_n = 0, -1
    for i in range(min(max_scan, len(df_raw))):
        fila = df_raw.iloc[i].tolist()
        no_nulos = [c for c in fila if c is not None and str(c).strip() != ""]
        n = len(no_nulos)
        if no_nulos and n >= 3 and _es_celda_encabezado(no_nulos[0]):
            return i
        if n > mejor_n:
            mejor_idx, mejor_n = i, n
    return mejor_idx


def cargar_hoja(archivo: str, hoja: str) -> pd.DataFrame:
    """Carga una hoja, detecta encabezado, asigna nombres de columnas limpios y
    desecha filas/columnas totalmente vacias."""
    raw = pd.read_excel(archivo, sheet_name=hoja, header=None, dtype=object)
    hdr = detectar_fila_encabezado(raw)
    encabezados = raw.iloc[hdr].tolist()
    datos = raw.iloc[hdr + 1:].reset_index(drop=True)

    # nombres de columnas: limpios y unicos
    cols, vistos = [], {}
    for j, c in enumerate(encabezados):
        nombre = limpiar_nombre_columna(c)
        if nombre == "":
            nombre = f"columna_{j+1}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre} ({vistos[nombre]})"
        else:
            vistos[nombre] = 0
        cols.append(nombre)
    datos.columns = cols

    # descarta filas totalmente vacias
    datos = datos.dropna(axis=0, how="all")
    # elimina columnas autogeneradas (sin nombre) que esten 100% vacias
    auto_vacias = [c for c in datos.columns
                   if re.fullmatch(r"columna_\d+", str(c)) and datos[c].isna().all()]
    datos = datos.drop(columns=auto_vacias)
    datos = datos.reset_index(drop=True)
    return datos


# ================================================================== #
#  CRUCE CON VINCULACIONES (imputacion de cod_local)
# ================================================================== #
def construir_indices_vinculaciones():
    """Devuelve dicts cui->cod_local, cod_mod->cod_local y cod_local->[cod_mod]."""
    print("  · Cargando insumo de vinculaciones (puede tardar)...")
    v = pd.read_excel(ARCHIVO_VINCULACIONES, sheet_name=0, dtype=object)
    v.columns = [limpiar_nombre_columna(c) for c in v.columns]
    cui_a_local, mod_a_local, local2mods = {}, {}, {}
    for _, r in v.iterrows():
        cl = limpiar_cod_local(r.get("cod_local"))
        if cl is np.nan or cl is None:
            continue
        cui = _solo_digitos(str(r.get("cui", "")))
        mod = _solo_digitos(str(r.get("cod_mod", "")))
        if cui and cui not in cui_a_local:
            cui_a_local[cui.lstrip("0") or "0"] = cl
        if mod:
            mod_a_local.setdefault(mod.lstrip("0") or "0", cl)
            local2mods.setdefault(cl, set()).add(mod.zfill(7))
    print(f"    -> {len(cui_a_local):,} CUI y {len(mod_a_local):,} cod_mod indexados")
    return cui_a_local, mod_a_local, local2mods


def imputar_cod_local(df, cui_a_local, mod_a_local, inst_a_local=None):
    """Si el df no tiene cod_local (o tiene nulos), intenta imputarlo cruzando
    por cui, cod_modular o codigo institucional. Devuelve (df, n_imputados)."""
    inst_a_local = inst_a_local or {}
    col_local = next((c for c in df.columns if es_col_cod_local(c)), None)
    col_cui   = next((c for c in df.columns if es_col_cui(c)), None)
    col_mod   = next((c for c in df.columns if es_col_cod_modular(c)), None)
    col_inst  = next((c for c in df.columns
                      if "codigo institucional" in norm_key(c)
                      or norm_key(c) in {"codinst", "cod_inst"}), None)

    if col_local is None:
        df["cod_local"] = np.nan
        col_local = "cod_local"

    def buscar(row):
        actual = row[col_local]
        if isinstance(actual, str) and actual.strip() and actual.strip().lower() != "nan":
            return actual
        if col_cui is not None and pd.notna(row.get(col_cui)):
            k = _solo_digitos(str(row[col_cui])).lstrip("0") or "0"
            if k in cui_a_local:
                return cui_a_local[k]
        if col_mod is not None and pd.notna(row.get(col_mod)):
            for c in re.findall(r"\d{5,8}", str(row[col_mod])):
                k = c.lstrip("0") or "0"
                if k in mod_a_local:
                    return mod_a_local[k]
        if col_inst is not None and pd.notna(row.get(col_inst)):
            k = _solo_digitos(str(row[col_inst])).lstrip("0") or "0"
            if k in inst_a_local:
                return inst_a_local[k]
        return actual

    antes = df[col_local].notna().sum()
    df[col_local] = df.apply(buscar, axis=1)
    despues = df[col_local].notna().sum()
    return df, max(0, despues - antes), col_local


# ================================================================== #
#  CRUCE CON UBIGEO (departamento / provincia / distrito)
# ================================================================== #
def construir_indice_ubigeo():
    """cod_local (6 dig) -> {departamento, provincia, distrito,
    centro_poblado, ubigeo} a partir del insumo ubigeo.xlsx."""
    print("  · Cargando insumo de ubigeo...")
    u = cargar_hoja(ARCHIVO_UBIGEO, 0)
    idx = {}
    c_local = next((c for c in u.columns if es_col_cod_local(c)), u.columns[0])
    c_reg  = buscar_col(u, "region", "departamento")
    c_prov = buscar_col(u, "provincia")
    c_dist = buscar_col(u, "distrito")
    c_cp   = buscar_col(u, "centro poblado")
    c_ubi  = buscar_col(u, "ubigeo")
    c_dre  = buscar_col(u, "dre")
    c_ugel = buscar_col(u, "ugel")
    c_nom  = buscar_col(u, "nro. y/o nombre", "nombre de los", "nombre de la i", "nombre")
    for _, r in u.iterrows():
        cl = limpiar_cod_local(r.get(c_local))
        if not isinstance(cl, str):
            continue
        idx.setdefault(cl, {
            "departamento":  normalizar_texto(r.get(c_reg))  if c_reg  else np.nan,
            "provincia":     normalizar_texto(r.get(c_prov)) if c_prov else np.nan,
            "distrito":      normalizar_texto(r.get(c_dist)) if c_dist else np.nan,
            "centro_poblado":normalizar_texto(r.get(c_cp))   if c_cp   else np.nan,
            "ubigeo":        (lambda v: re.sub(r"\D", "", str(v)).zfill(6)
                              if pd.notna(v) and re.sub(r"\D", "", str(v)) else np.nan)(r.get(c_ubi)) if c_ubi else np.nan,
            "dre":           normalizar_texto(r.get(c_dre))  if c_dre  else np.nan,
            "ugel":          normalizar_texto(r.get(c_ugel)) if c_ugel else np.nan,
            "nombre_ie":     normalizar_texto(r.get(c_nom))  if c_nom  else np.nan,
            "area":          np.nan,
        })
    print(f"    -> {len(idx):,} locales georreferenciados (con DRE/UGEL)")
    return idx


def _norm_area(v):
    """Normaliza el ambito a binario 'Urbana'/'Rural' (admite 'Urbano',
    'Rural 1/2/3', etc.)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    s = norm_key(v)
    if s.startswith("urban"):
        return "Urbana"
    if s.startswith("rural"):
        return "Rural"
    return normalizar_texto(v)


def construir_indice_pronied():
    """ubigeo PRONIED: cod_local -> geo (con DRE/UGEL, nombre IE, ruralidad) y
    mod_local (codigos modulares apilados -> cod_local). Insumo adicional."""
    if not os.path.exists(ARCHIVO_PRONIED):
        print("  · (ubigeo_pronied no encontrado; se omite)")
        return {}, {}
    print("  · Cargando ubigeo PRONIED...")
    u = cargar_hoja(ARCHIVO_PRONIED, 0)
    c_local = next((c for c in u.columns if es_col_cod_local(c)), u.columns[0])
    c_mod  = buscar_col(u, "modular")
    c_dpto = buscar_col(u, "departamento")
    c_prov = buscar_col(u, "provincia")
    c_dist = buscar_col(u, "distrito")
    c_ubi  = buscar_col(u, "ubigeo")
    c_cp   = buscar_col(u, "centro poblado")
    c_ugel = buscar_col(u, "dre / ugel", "ugel", "dre/ugel")
    c_nom  = buscar_col(u, "numero y nombre", "nombre")
    c_area = buscar_col(u, "gradiente", "ruralidad", "area")
    geo, mod2local, local2mods = {}, {}, {}
    for _, r in u.iterrows():
        cl = limpiar_cod_local(r.get(c_local))
        if not isinstance(cl, str):
            continue
        geo.setdefault(cl, {
            "departamento":   normalizar_texto(r.get(c_dpto)) if c_dpto else np.nan,
            "provincia":      normalizar_texto(r.get(c_prov)) if c_prov else np.nan,
            "distrito":       normalizar_texto(r.get(c_dist)) if c_dist else np.nan,
            "centro_poblado": normalizar_texto(r.get(c_cp))   if c_cp   else np.nan,
            "ubigeo": (lambda v: re.sub(r"\D", "", str(v)).zfill(6)
                       if pd.notna(v) and re.sub(r"\D", "", str(v)) else np.nan)(r.get(c_ubi)) if c_ubi else np.nan,
            "dre":            np.nan,
            "ugel":           normalizar_texto(r.get(c_ugel)) if c_ugel else np.nan,
            "nombre_ie":      normalizar_texto(r.get(c_nom))  if c_nom  else np.nan,
            "area":           _norm_area(r.get(c_area))       if c_area else np.nan,
        })
        if c_mod and pd.notna(r.get(c_mod)):
            for m in re.findall(r"\d{5,8}", str(r.get(c_mod))):
                mod2local.setdefault(m.lstrip("0") or "0", cl)
                local2mods.setdefault(cl, set()).add(m.zfill(7))
    print(f"    -> PRONIED: {len(geo):,} locales, {len(mod2local):,} cod_mod")
    return geo, mod2local, local2mods


def cargar_ubigeo_cui():
    """ubigeo por CUI: codigo_unico (CUI) -> {departamento, provincia, distrito,
    ubigeo}. Permite enriquecer geografia por CUI cuando falla por cod_local."""
    if not os.path.exists(ARCHIVO_UBIGEO_CUI):
        print("  · (ubigeo_cui no encontrado; se omite)")
        return {}
    print("  · Cargando ubigeo por CUI...")
    p = pd.read_csv(ARCHIVO_UBIGEO_CUI, sep=";", dtype=str, encoding="latin-1",
                    on_bad_lines="skip")
    p.columns = [limpiar_nombre_columna(c) for c in p.columns]
    colf = lambda nm: p[nm].tolist() if nm in p.columns else [None] * len(p)
    L_cui = colf("CODIGO_UNICO"); L_d = colf("DEPARTAMENTO")
    L_p = colf("PROVINCIA"); L_di = colf("DISTRITO"); L_u = colf("UBIGEO")
    cui2geo = {}
    for i in range(len(p)):
        k = re.sub(r"\D", "", str(L_cui[i])) if L_cui[i] is not None else ""
        if not k:
            continue
        k = k.lstrip("0") or "0"
        if k in cui2geo:
            continue
        ub = re.sub(r"\D", "", str(L_u[i])) if L_u[i] is not None else ""
        cui2geo[k] = {
            "departamento": normalizar_texto(L_d[i]),
            "provincia":    normalizar_texto(L_p[i]),
            "distrito":     normalizar_texto(L_di[i]),
            "ubigeo":       ub.zfill(6) if ub else np.nan,
        }
    print(f"    -> ubigeo_cui: {len(cui2geo):,} CUI con ubicacion")
    return cui2geo


def cargar_padron():
    """Lee el padron web nacional. Devuelve:
      - mod2local   : cod_mod (sin ceros) -> cod_local (6 dig)
      - inst2local  : codinst (sin ceros) -> cod_local (6 dig)
      - geo_padron  : cod_local (6 dig) -> {departamento, provincia, distrito,
                       centro_poblado, ubigeo, dre, ugel}
    Sirve como fuente adicional de imputacion y respaldo geografico."""
    if not os.path.exists(ARCHIVO_PADRON):
        print("  · (padron web no encontrado; se omite)")
        return {}, {}, {}
    print("  · Cargando padron web nacional...")
    p = pd.read_csv(ARCHIVO_PADRON, sep=";", dtype=str, encoding="latin-1",
                    on_bad_lines="skip")
    p.columns = [limpiar_nombre_columna(c) for c in p.columns]
    col = lambda nm: p[nm].tolist() if nm in p.columns else [None] * len(p)
    L_local = col("CODLOCAL"); L_mod = col("COD_MOD"); L_inst = col("CODINST")
    L_dpto = col("D_DPTO"); L_prov = col("D_PROV"); L_dist = col("D_DIST")
    L_geo = col("CODGEO"); L_reg = col("D_REGION"); L_ugel = col("D_DREUGEL")
    L_nom = col("CEN_EDU"); L_area = col("DAREACENSO")
    mod2local, inst2local, geo = {}, {}, {}
    for i in range(len(p)):
        cl = limpiar_cod_local(L_local[i])
        if not isinstance(cl, str):
            continue
        mod = re.sub(r"\D", "", str(L_mod[i])) if L_mod[i] is not None else ""
        if mod:
            mod2local.setdefault(mod.lstrip("0") or "0", cl)
        inst = re.sub(r"\D", "", str(L_inst[i])) if L_inst[i] is not None else ""
        if inst:
            inst2local.setdefault(inst.lstrip("0") or "0", cl)
        if cl not in geo:
            ub = re.sub(r"\D", "", str(L_geo[i])) if L_geo[i] is not None else ""
            geo[cl] = {
                "departamento":   normalizar_texto(L_dpto[i]),
                "provincia":      normalizar_texto(L_prov[i]),
                "distrito":       normalizar_texto(L_dist[i]),
                "centro_poblado": np.nan,
                "ubigeo":         ub.zfill(6) if ub else np.nan,
                "dre":            normalizar_texto(L_reg[i]),
                "ugel":           normalizar_texto(L_ugel[i]),
                "nombre_ie":      normalizar_texto(L_nom[i]),
                "area":           _norm_area(L_area[i]),
            }
    print(f"    -> padron: {len(mod2local):,} cod_mod, {len(inst2local):,} codinst, "
          f"{len(geo):,} locales con ubicacion")
    return mod2local, inst2local, geo


def agregar_ubigeo(df, col_local, ubigeo_idx, cui2geo=None):
    """Agrega columnas canonicas (departamento/provincia/distrito/centro_poblado/
    ubigeo/dre/ugel/nombre_ie/area) a partir del cod_local. Para los que quedan
    sin ubicacion, intenta completar geografia por CUI. Devuelve (df, n_geo)."""
    campos = ["departamento", "provincia", "distrito", "centro_poblado",
              "ubigeo", "dre", "ugel", "nombre_ie", "area"]
    vacio = {k: np.nan for k in campos}
    locales = (df[col_local] if col_local in df.columns
               else pd.Series([np.nan] * len(df))).tolist()
    datos = [dict(ubigeo_idx.get(cl, vacio)) if isinstance(cl, str) else dict(vacio)
             for cl in locales]

    # fallback por CUI para los registros sin departamento
    if cui2geo:
        col_cui = next((c for c in df.columns if es_col_cui(c)), None)
        if col_cui is not None:
            cuis = df[col_cui].tolist()
            for i, d in enumerate(datos):
                if pd.isna(d.get("departamento")) and pd.notna(cuis[i]):
                    k = _solo_digitos(str(cuis[i])).lstrip("0") or "0"
                    g = cui2geo.get(k)
                    if g:
                        for kk in ("departamento", "provincia", "distrito", "ubigeo"):
                            if pd.isna(d.get(kk)) and pd.notna(g.get(kk)):
                                d[kk] = g[kk]

    for k in campos:
        nombre = k
        if nombre in df.columns:          # evita colision con columna original
            nombre = f"{k}_ubigeo"
        df[nombre] = [d[k] for d in datos]
    n_geo = int(sum(1 for d in datos if pd.notna(d["departamento"])))
    return df, n_geo


def enriquecer_por_nombre(df, dept_idx, ugel_idx):
    """Para registros que quedaron SIN departamento (no tienen cod_local),
    intenta asignar departamento/dre/ugel cruzando por NOMBRE de UGEL o de
    Unidad Zonal contra el padron. Devuelve (df, n_asignados)."""
    if "departamento" not in df.columns:
        return df, 0
    mask = df["departamento"].isna()
    if not mask.any():
        return df, 0

    # columnas candidatas que pueden contener nombres de UGEL/Unidad Zonal/entidad
    cands = [c for c in df.columns if any(t in norm_key(c) for t in
             ["ugel", "unidad zonal", "nombre de entidad", "nombre de la unidad",
              "dre", "region", "departamento"])]
    if not cands:
        return df, 0

    # asegura dtype object para poder asignar texto (evita columnas all-NaN float64)
    for c in ("departamento", "dre", "ugel"):
        if c in df.columns:
            df[c] = df[c].astype(object)

    # claves ordenadas (mas largas primero, para evitar coincidencias parciales)
    ugel_keys = sorted(ugel_idx.keys(), key=len, reverse=True)
    dept_keys = sorted(dept_idx.keys(), key=len, reverse=True)

    n = 0
    for i in df.index[mask]:
        celdas = [norm_key(df.at[i, c]) for c in cands if pd.notna(df.at[i, c])]
        if not celdas:
            continue
        texto = " ".join(celdas)
        info = None
        # 1) coincidencia exacta de una celda con un nombre de UGEL
        for cel in celdas:
            if cel in ugel_idx:
                info = ugel_idx[cel]; break
        # 2) departamento como palabra dentro del texto (p.ej. "UNIDAD ZONAL AREQUIPA")
        if info is None:
            for dk in dept_keys:
                if dk and re.search(rf"\b{re.escape(dk)}\b", texto):
                    info = dept_idx[dk]; break
        if info is None:
            continue
        if pd.notna(info.get("departamento")):
            df.at[i, "departamento"] = info["departamento"]
        if "dre" in df.columns and pd.notna(info.get("dre")):
            df.at[i, "dre"] = info["dre"]
        if "ugel" in df.columns and info.get("ugel") and pd.isna(df.at[i, "ugel"]):
            df.at[i, "ugel"] = info["ugel"]
        n += 1
    return df, n


# ================================================================== #
#  LIMPIEZA DE UN DATAFRAME
# ================================================================== #
def limpiar_df(df, cui_a_local, mod_a_local, ubigeo_idx, inst_a_local=None,
               dept_idx=None, ugel_idx=None, cui2geo=None, anio=None):
    """Aplica todas las reglas de limpieza columna por columna."""
    reporte_cols = []
    for col in list(df.columns):
        if es_col_comentario(col):
            df[col] = df[col].map(normalizar_comentario)
            tipo = "comentario"
        elif es_col_fecha(col):
            df[col] = df[col].map(limpiar_fecha)
            tipo = "fecha (DD/MM/YYYY)"
        elif es_col_cod_local(col):
            df[col] = df[col].map(limpiar_cod_local)
            tipo = "cod_local (6 dig)"
        elif es_col_cui(col):
            df[col] = df[col].map(limpiar_cui)
            tipo = "cui (7 dig)"
        elif es_col_cod_modular(col):
            df[col] = df[col].map(limpiar_cod_modular)
            tipo = "cod_modular (7 dig, apilable)"
        elif es_col_monto(col):
            df[col] = df[col].map(limpiar_monto)
            tipo = "monto/devengado (numerico)"
        else:
            df[col] = df[col].map(normalizar_texto)
            tipo = "texto normalizado"
        reporte_cols.append((col, tipo))

    # imputa cod_local (vinculaciones + padron via cui/cod_mod/codinst)
    df, n_imp, col_local = imputar_cod_local(df, cui_a_local, mod_a_local, inst_a_local)

    # enriquece con ubigeo (departamento/provincia/distrito) desde cod_local
    cols_previas = set(df.columns)
    df, n_geo = agregar_ubigeo(df, col_local, ubigeo_idx, cui2geo)
    for c in df.columns:
        if c not in cols_previas:
            reporte_cols.append((c, "ubigeo (derivado de cod_local)"))

    # para registros sin cod_local: asigna departamento/dre/ugel por nombre
    if dept_idx or ugel_idx:
        df, _ = enriquecer_por_nombre(df, dept_idx or {}, ugel_idx or {})

    # anio: si la base no tiene fecha, agrega columna anio
    tiene_fecha = any(es_col_fecha(c) for c in df.columns)
    if anio is not None and not tiene_fecha and "anio" not in df.columns:
        df["anio"] = anio
        reporte_cols.append(("anio", "anio (forzado segun Ordenamiento)"))

    # elimina filas completamente vacias tras limpiar
    df = df.dropna(axis=0, how="all").reset_index(drop=True)
    return df, reporte_cols, n_imp, col_local, n_geo


# ================================================================== #
#  HARMONIZACION (base final armonizada)
# ================================================================== #
def primera_col(df, predicado):
    for c in df.columns:
        if predicado(c):
            return c
    return None


def buscar_col(df, *patrones):
    for c in df.columns:
        n = norm_key(c)
        if any(p in n for p in patrones):
            return c
    return None


def etiqueta_fuente(libro):
    """Etiqueta legible de la unidad/insumo que reporta la base."""
    if libro in FUENTE_LIBRO:
        return FUENTE_LIBRO[libro]
    n = norm_key(libro)
    if n.startswith("anexo 01") or "anexo 1" in n:
        return "Anexo 1"
    if n.startswith("anexo 02") or "anexo 2" in n:
        return "Anexo 2"
    return libro


def asignar_fuente(df, libro):
    """Devuelve la serie 'fuente' por fila.
      - Grupo 2 (PI): toma el campo 'fuentes_concat' de la propia base.
      - Anexos 1 y 2: toma el campo 'Remitente' (quien reporta).
      - Resto: la unidad/insumo que reporta."""
    etiqueta = etiqueta_fuente(libro)
    if libro == "df_pi_listados_final_v7.xlsx":
        c_fc = buscar_col(df, "fuentes_concat", "fuentes concat", "fuente")
        if c_fc is not None:
            return df[c_fc].astype(object).where(df[c_fc].notna(), etiqueta)
    if etiqueta in ("Anexo 1", "Anexo 2"):
        c_rem = buscar_col(df, "remitente")
        if c_rem is not None:
            base = f"{etiqueta} - " + df[c_rem].astype(object).where(
                df[c_rem].notna(), "Sin remitente").astype(str)
            return base
    return pd.Series([etiqueta] * len(df), index=df.index)


def _col_fecha_inicio(df):
    kws = ["inicio", "viabilidad", "aprobacion"]
    for c in df.columns:                       # preferir columnas tipo fecha
        if es_col_fecha(c) and any(k in norm_key(c) for k in kws):
            return c
    for c in df.columns:                       # si no, por nombre
        if any(k in norm_key(c) for k in kws):
            return c
    return None


def _col_fecha_fin(df):
    kws = ["culmin", "entrega", "recepcion", "cierre", "termino",
           "caducidad", "estimad"]
    def match(n):
        return any(k in n for k in kws) or bool(re.search(r"\bfin\b", n))
    for c in df.columns:                       # preferir columnas tipo fecha
        if es_col_fecha(c) and match(norm_key(c)):
            return c
    for c in df.columns:                       # si no, por nombre (cierre/termino)
        if match(norm_key(c)):
            return c
    return None


def armonizar(df, df_name, grupo, libro, hoja, col_local):
    """Extrae campos estandar de un df limpio para la base final armonizada."""
    n = len(df)
    out = pd.DataFrame(index=range(n))
    out["grupo"]         = grupo
    out["grupo_nombre"]  = GRUPOS[grupo]
    out["df_name"]       = df_name
    out["fuente"]        = df["fuente"].values if "fuente" in df.columns else etiqueta_fuente(libro)
    out["libro"]         = libro
    out["hoja"]          = hoja

    c_local = col_local if col_local in df.columns else primera_col(df, es_col_cod_local)
    c_mod   = primera_col(df, es_col_cod_modular)
    c_cui   = primera_col(df, es_col_cui)
    c_nom   = ("nombre_ie" if "nombre_ie" in df.columns else
               buscar_col(df, "nombre de la i", "nombre del", "nombre de instituc",
                          "local escolar", "nombre ie", "locales educativos",
                          "nombre del servicio"))
    # prioriza columnas canonicas derivadas de ubigeo
    c_dep   = "departamento" if "departamento" in df.columns else buscar_col(df, "region", "departamento", "nombre departamento", "dre/gre")
    c_prov  = "provincia"    if "provincia"    in df.columns else buscar_col(df, "provincia", "nombre provincia")
    c_dist  = "distrito"     if "distrito"     in df.columns else buscar_col(df, "distrito", "nombre distrito")
    # detalle del activo / intervencion: busca por palabras clave en la base
    c_activo = buscar_col(df, "activo", "bien", "descripcion", "componente",
                          "tipo de mantenimiento", "tipo de intervencion",
                          "tipo de inversion", "tipo de sistema", "sistema modular",
                          "actividad", "tipologia", "producto", "grupo", "tema",
                          "titulo del asesoramiento", "alcance del plan",
                          "denominacion")
    c_est   = buscar_col(df, "estado", "fase de la obra", "etapa de la intervencion",
                         "fase del proceso", "situacion")
    c_monto = buscar_col(df, "monto de la inversion", "monto de inversion",
                         "monto contractual", "monto total", "monto de la intervencion",
                         "costo actualizado", "monto", "pim")
    c_dev   = buscar_col(df, "devengado", "dev. acumulado", "dev acumulado", "dev_acum")
    c_av    = buscar_col(df, "avance fisico", "avance físico", "avance")
    c_fini  = _col_fecha_inicio(df)
    c_ffin  = _col_fecha_fin(df)
    c_fecha = primera_col(df, es_col_fecha)
    c_com   = primera_col(df, es_col_comentario)

    def col(c):
        return df[c].values if c in df.columns else np.nan

    def fechacol(c):
        # formatea a DD/MM/YYYY (idempotente) incluso columnas no marcadas como fecha
        return df[c].map(limpiar_fecha).values if c and c in df.columns else np.nan

    # tipo_intervencion = nombre del grupo (para G2 = tipo de inversion);
    # tipo_activo = detalle del activo intervenido (para G2 = intervencion detectada)
    if grupo == 2:
        c_tinv = buscar_col(df, "tipo_inversion", "tipo de inversion")
        c_tdet = buscar_col(df, "tipo_intervencion_detectada", "intervencion detectada")
        if c_tinv:                          # tipo de inversion; si falta -> nombre grupo
            s = df[c_tinv].astype(object)
            tipo_intervencion = s.where(s.notna(), GRUPOS[grupo]).values
        else:
            tipo_intervencion = GRUPOS[grupo]
        tipo_activo = (df[c_tdet].values if c_tdet else col(c_activo))
    else:
        tipo_intervencion = GRUPOS[grupo]
        tipo_activo = col(c_activo)

    out["cod_local"]        = col(c_local)
    out["cod_modular"]      = col(c_mod)
    out["cui"]              = col(c_cui)
    out["nombre_ie"]        = col(c_nom)
    out["departamento"]     = col(c_dep)
    out["provincia"]        = col(c_prov)
    out["distrito"]         = col(c_dist)
    out["centro_poblado"]   = col("centro_poblado") if "centro_poblado" in df.columns else np.nan
    out["ubigeo"]           = col("ubigeo") if "ubigeo" in df.columns else np.nan
    out["dre"]              = col("dre") if "dre" in df.columns else np.nan
    out["ugel"]             = col("ugel") if "ugel" in df.columns else np.nan
    out["area"]             = col("area") if "area" in df.columns else np.nan
    out["tipo_intervencion"]= tipo_intervencion
    out["tipo_activo"]      = tipo_activo
    out["estado"]           = col(c_est)
    out["monto"]            = col(c_monto)
    out["devengado"]        = col(c_dev)
    out["avance_fisico"]    = col(c_av)
    out["fecha_inicio"]     = fechacol(c_fini)
    out["fecha_fin"]        = fechacol(c_ffin) if c_ffin else (fechacol(c_fecha) if c_fini is None else np.nan)
    out["anio"]             = col("anio") if "anio" in df.columns else np.nan
    out["comentario"]       = col(c_com)
    # deriva anio desde las fechas cuando falte
    if out["anio"].isna().all():
        for fc in ("fecha_fin", "fecha_inicio"):
            anios = out[fc].astype(str).str.extract(r"(\d{4})")[0]
            out["anio"] = out["anio"].fillna(anios) if "anio" in out else anios
    return out


# ================================================================== #
#  GUARDADO
# ================================================================== #
def nombre_carpeta_grupo(g):
    base = re.sub(r"[^A-Za-z0-9]+", "_", quitar_acentos(GRUPOS[g])).strip("_")
    return os.path.join(LIMPIAS_DIR, f"Grupo_{g:02d}_{base}")


def guardar_parquet(df, ruta):
    """Guarda en parquet preservando tipos numericos; castea solo las columnas
    de tipo objeto (texto/codigos/mixtas) a 'string' para evitar errores de
    tipos mixtos en pyarrow."""
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype("string")
    out.to_parquet(ruta, index=False)


def guardar_df(df, carpeta, df_name):
    os.makedirs(carpeta, exist_ok=True)
    df.to_excel(os.path.join(carpeta, f"{df_name}.xlsx"), index=False)
    guardar_parquet(df, os.path.join(carpeta, f"{df_name}.parquet"))


# ================================================================== #
#  PROCESO PRINCIPAL
# ================================================================== #
def main():
    print("=" * 70)
    print(" DIPLAN-P2 | Pipeline de limpieza y consolidacion")
    print("=" * 70)

    cui_a_local, mod_a_local, local2mods = construir_indices_vinculaciones()
    ubigeo_idx = construir_indice_ubigeo()
    mod_padron, inst_padron, geo_padron = cargar_padron()
    geo_pronied, mod_pronied, local2mods_p = construir_indice_pronied()
    cui2geo = cargar_ubigeo_cui()

    # combina cod_local -> [cod_mod] (vinculaciones + PRONIED) para el Grupo 2
    for cl, mods in local2mods_p.items():
        local2mods.setdefault(cl, set()).update(mods)

    def fusionar_geo(destino, origen):
        """Agrega 'origen' a 'destino' sin sobrescribir; completa campos NaN."""
        for cl, info in origen.items():
            if cl not in destino:
                destino[cl] = dict(info)
            else:
                actual = destino[cl]
                for campo, valor in info.items():
                    if pd.isna(actual.get(campo)) and pd.notna(valor):
                        actual[campo] = valor

    # fusiona TODAS las fuentes de geo (sin sobrescribir lo ya conocido)
    for k, v in mod_padron.items():
        mod_a_local.setdefault(k, v)
    for k, v in mod_pronied.items():
        mod_a_local.setdefault(k, v)
    fusionar_geo(ubigeo_idx, geo_padron)
    fusionar_geo(ubigeo_idx, geo_pronied)

    # indices por NOMBRE (para registros sin cod_local): UGEL y departamento
    dept_idx, ugel_idx = {}, {}
    for info in ubigeo_idx.values():
        dep = info.get("departamento")
        if isinstance(dep, str) and norm_key(dep):
            dept_idx.setdefault(norm_key(dep),
                                {"departamento": dep, "dre": info.get("dre")})
        ug = info.get("ugel")
        if isinstance(ug, str) and norm_key(ug):
            ugel_idx.setdefault(norm_key(ug), {"departamento": dep,
                                "dre": info.get("dre"), "ugel": ug})
    print(f"  · Indices combinados -> cod_mod: {len(mod_a_local):,} | "
          f"locales con ubicacion: {len(ubigeo_idx):,} | "
          f"UGEL: {len(ugel_idx):,} | departamentos: {len(dept_idx):,}")

    filas_dicc        = []   # diccionario de datos
    filas_reporte     = []   # reporte por base
    armonizadas       = []   # para base final armonizada
    union_completa    = []   # para base final union completa
    consolidado_grupo = {}   # para una base consolidada por cada grupo

    for archivo, hoja, df_name, nro, grupo, anio in CONFIG:
        if grupo == GRUPO_OMITIR:
            continue
        if not os.path.exists(archivo):
            print(f"  [FALTA] {archivo} :: {hoja}  -> archivo no encontrado")
            filas_reporte.append({"df_name": df_name, "libro": archivo, "hoja": hoja,
                                  "grupo": grupo, "estado": "ARCHIVO FALTANTE",
                                  "filas": 0, "columnas": 0, "cod_local_imputados": 0})
            continue

        print(f"  · [{nro:>2}] G{grupo:<2} {df_name:<28} <- {hoja}")
        try:
            df = cargar_hoja(archivo, hoja)
        except Exception as e:
            print(f"      ERROR al cargar: {e}")
            filas_reporte.append({"df_name": df_name, "libro": archivo, "hoja": hoja,
                                  "grupo": grupo, "estado": f"ERROR: {e}",
                                  "filas": 0, "columnas": 0, "cod_local_imputados": 0})
            continue

        df, rep_cols, n_imp, col_local, n_geo = limpiar_df(
            df, cui_a_local, mod_a_local, ubigeo_idx, inst_padron,
            dept_idx, ugel_idx, cui2geo, anio)

        # agrega campo 'fuente' (unidad/remitente que reporta) a cada base
        df["fuente"] = asignar_fuente(df, archivo).values
        rep_cols.append(("fuente", "fuente (origen del reporte)"))

        # SOLO Grupo 2 (PI): completa cod_mod con los modulos del local (padron)
        if df_name == "df_pi_listados_final_v7":
            cmod = next((c for c in df.columns if es_col_cod_modular(c)), None)
            if cmod is None:
                cmod = "cod_mod"; df[cmod] = np.nan
            def _mods(cl):
                if isinstance(cl, str) and cl in local2mods and local2mods[cl]:
                    return " | ".join(sorted(local2mods[cl]))
                return np.nan
            llenos_prev = df[cmod].notna()
            nuevos_mod = df[col_local].map(_mods)
            df[cmod] = df[cmod].where(llenos_prev, nuevos_mod)
            rep_cols.append((cmod, "cod_modular (completado desde padron)"))

        # armoniza ANTES de canonizar (usa los nombres originales para detectar)
        armonizadas.append(armonizar(df, df_name, grupo, archivo, hoja, col_local))

        # reporte (antes de renombrar)
        col_local_lleno = int(df[col_local].notna().sum()) if col_local in df.columns else 0
        filas_reporte.append({"df_name": df_name, "libro": archivo, "hoja": hoja,
                              "grupo": grupo, "estado": "OK",
                              "filas": len(df), "columnas": df.shape[1],
                              "cod_local_no_nulos": col_local_lleno,
                              "cod_local_imputados": n_imp,
                              "filas_con_ubigeo": n_geo})

        # CANONIZA nombres de columnas (minuscula / snake_case)
        df_canon, mapping = canonizar_columnas(df)

        # guarda base limpia (ya canonizada)
        guardar_df(df_canon, nombre_carpeta_grupo(grupo), df_name)

        # diccionario de datos (variable canonica + nombre original)
        tipos = dict(rep_cols)
        for orig, canon in mapping.items():
            serie = df[orig]
            ej = serie.dropna()
            filas_dicc.append({
                "grupo": grupo, "grupo_nombre": GRUPOS[grupo], "df_name": df_name,
                "libro": archivo, "hoja": hoja,
                "variable": canon, "nombre_original": orig,
                "tipo_limpieza": tipos.get(orig, "texto normalizado"),
                "no_nulos": int(serie.notna().sum()), "total_filas": len(df),
                "ejemplo": str(ej.iloc[0])[:60] if len(ej) else "",
            })

        # base consolidada por grupo (conserva TODOS los campos) + union total
        uc = df_canon.copy()
        uc.insert(0, "grupo", grupo)
        uc.insert(1, "grupo_nombre", GRUPOS[grupo])
        uc.insert(2, "df_name", df_name)
        uc.insert(3, "libro", archivo)
        uc.insert(4, "hoja", hoja)
        consolidado_grupo.setdefault(grupo, []).append(uc)
        union_completa.append(uc)

    # -------- BASE FINAL ARMONIZADA -------- #
    print("\n  · Generando BASE FINAL ARMONIZADA...")
    base_arm = pd.concat(armonizadas, ignore_index=True)
    base_arm.to_excel(os.path.join(FINAL_DIR, "base_final_armonizada.xlsx"), index=False)
    guardar_parquet(base_arm, os.path.join(FINAL_DIR, "base_final_armonizada.parquet"))
    print(f"    -> {len(base_arm):,} filas x {base_arm.shape[1]} columnas")

    # -------- BASE FINAL UNION COMPLETA -------- #
    print("  · Generando BASE FINAL UNION COMPLETA...")
    base_uni = pd.concat(union_completa, ignore_index=True, sort=False)
    guardar_parquet(base_uni, os.path.join(FINAL_DIR, "base_final_union_completa.parquet"))
    # Excel solo si es manejable (openpyxl es lento con tablas muy grandes)
    celdas = len(base_uni) * base_uni.shape[1]
    if len(base_uni) < 1_048_576 and celdas <= 3_000_000:
        base_uni.to_excel(
            os.path.join(FINAL_DIR, "base_final_union_completa.xlsx"), index=False)
    else:
        print(f"    (tabla muy grande: {len(base_uni):,}x{base_uni.shape[1]} = "
              f"{celdas:,} celdas; se entrega solo en Parquet)")
    print(f"    -> {len(base_uni):,} filas x {base_uni.shape[1]} columnas")

    # -------- BASES CONSOLIDADAS POR GRUPO (todos los campos) -------- #
    print("  · Generando BASES CONSOLIDADAS POR GRUPO...")
    cons_dir = os.path.join(FINAL_DIR, "consolidado_por_grupo")
    os.makedirs(cons_dir, exist_ok=True)
    for g in sorted(consolidado_grupo):
        cg = pd.concat(consolidado_grupo[g], ignore_index=True, sort=False)
        base = re.sub(r"[^A-Za-z0-9]+", "_", quitar_acentos(GRUPOS[g])).strip("_")
        nombre = f"Grupo_{g:02d}_{base}_consolidado"
        ruta = os.path.join(cons_dir, nombre)
        guardar_parquet(cg, ruta + ".parquet")
        celdas = len(cg) * cg.shape[1]
        if len(cg) < 1_048_576 and celdas <= 4_000_000:
            cg.to_excel(ruta + ".xlsx", index=False)
            ex = "xlsx+parquet"
        else:
            ex = "solo parquet (muy grande)"
        print(f"    - Grupo {g:02d} {GRUPOS[g]:<24}: {len(cg):>7,} filas x "
              f"{cg.shape[1]:>3} cols [{ex}]")

    # -------- DICCIONARIO DE DATOS -------- #
    print("  · Generando DICCIONARIO DE DATOS...")
    dicc = pd.DataFrame(filas_dicc)
    dicc.to_excel(os.path.join(OUT_DIR, "diccionario_datos.xlsx"), index=False)

    # -------- REPORTE DE LIMPIEZA -------- #
    print("  · Generando REPORTE DE LIMPIEZA...")
    rep = pd.DataFrame(filas_reporte)
    with pd.ExcelWriter(os.path.join(OUT_DIR, "reporte_limpieza.xlsx")) as xw:
        rep.to_excel(xw, sheet_name="bases", index=False)
        resumen_grupos = (rep[rep["estado"] == "OK"]
                          .groupby(["grupo"])
                          .agg(bases=("df_name", "count"), filas=("filas", "sum"))
                          .reset_index())
        resumen_grupos["grupo_nombre"] = resumen_grupos["grupo"].map(GRUPOS)
        resumen_grupos.to_excel(xw, sheet_name="resumen_grupos", index=False)

    print("\n" + "=" * 70)
    print(" PROCESO COMPLETADO. Salidas en:", OUT_DIR)
    print("=" * 70)


if __name__ == "__main__":
    main()
