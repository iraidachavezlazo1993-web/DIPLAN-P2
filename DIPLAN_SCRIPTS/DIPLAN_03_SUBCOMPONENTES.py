# -*- coding: utf-8 -*-
"""
================================================================================
 DIPLAN-P2 | Subcomponentes de cierre de brecha sobre la BASE CONSOLIDADA TOTAL
================================================================================
Aplica los 17 subcomponentes (Fichas_Indicadores_Cierre_Brecha_PNIE) a TODOS los
cod_local de la base consolidada completa (los 11 grupos), validando el
cumplimiento con las fuentes oficiales:

  - Rep_Activos_F7_02SET2024.csv  (CUI -> ACTIVO intervenido en F7)
  - HR 079179-2026 - Inversiones.csv.gz  (CUI -> ESTADO / FASE / MTO_F9_CIERRE)
  - vinculaciones / base consolidada (llaves cui, cod_local, cod_mod)

Salidas (output/base_final/):
  1) subcomponentes_total_matriz.(xlsx|parquet)
        1 fila por cod_local (todos los grupos); 1 columna 1/0 por subcomponente;
        incluye llaves (cui, cod_mod), grupo(s) y total de subcomponentes.
  2) subcomponentes_total_evidencia.(xlsx|parquet)
        1 fila por (cod_local, subcomponente) con cumple (1/0) + la EVIDENCIA:
        activo que valido, estado/fase/F9 de la inversion, y las llaves, para que
        el revisor pueda cotejar por que se cumplio el criterio.
  3) subcomponentes_total_diccionario.xlsx

Regla de cumplimiento (segun decision): cumple=1 si el local presenta el ACTIVO
asociado al subcomponente (Rep_Activos_F7). El ESTADO/FASE de la inversion se
adjunta en la base de evidencia para cotejo (no condiciona el 1/0).
================================================================================
"""
import os
import re
import gzip
import csv
import unicodedata
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR  = os.path.dirname(BASE_DIR)            # raiz del proyecto (insumos)
os.chdir(ROOT_DIR)
FINAL_DIR = os.path.join(ROOT_DIR, "DIPLAN_OUTPUT", "base_final")
CONS_PATH = os.path.join(FINAL_DIR, "base_final_armonizada.parquet")
ACTIVOS_ZIP = "Rep_Activos_F7_02SET2024.zip"
ACTIVOS_CSV = "Rep_Activos_F7_02SET2024.csv"
INVERS_GZ = "HR 079179-2026 - Inversiones.csv.gz"


def na(t):
    t = unicodedata.normalize("NFKD", str(t))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"\s+", " ", t).strip()


def _dig(x):
    if x is None:
        return ""
    return re.sub(r"\D", "", str(x).split(".")[0])   # ignora decimales de floats


# --- 17 subcomponentes: (codigo, grupo_pnie, nombre, claves de ACTIVO) -------
SUBCOMPONENTES = [
    ("ct_st_d", "G1", "Demolicion de locales para sustitucion total",
     ["demolicion", "edificacion", "obras exteriores"]),
    ("ct_sp_d", "G1", "Demolicion de edificaciones para sustitucion parcial",
     ["demolicion", "edificacion"]),
    ("ct_ri", "G1", "Reforzamiento incremental",
     ["reforzamiento", "muro de contencion", "estructura"]),
    ("ct_ic", "G1", "Intervencion contingente",
     ["modulo", "prefabricad", "provisional", "contingen"]),
    ("ct_cp", "G1", "Cercos perimetricos",
     ["cerco perimetrico", "cerco"]),
    ("ct_aad", "G2", "Acceso a agua y desague",
     ["agua", "desague", "saneamiento", "ubs", "unidades basicas",
      "reservorio", "captacion", "linea de conduccion", "red de distribucion",
      "pozo", "tanque", "letrina", "biodigestor"]),
    ("ct_cad", "G2", "Mejora de la calidad del servicio de agua y saneamiento",
     ["servicios higienicos", "vestidor", "cisterna", "tanque elevado",
      "planta de tratamiento", "agua potable", "ptar"]),
    ("ct_ce", "G3", "Mejora de la calidad del servicio de energia electrica",
     ["electric", "energia", "subestacion", "panel solar", "red electrica",
      "tablero", "instalacion electrica"]),
    ("ct_me", "G3", "Mobiliario y equipamiento de la infraestructura existente",
     ["mobiliario", "equipamiento", "equipo", "mueble", "carpeta",
      "pizarra", "computo"]),
    ("ct_ene", "G3", "Mantenimiento correctivo de elementos no estructurales",
     ["piso", "pavimento", "vereda", "cobertura", "pintura", "carpinteria",
      "acabado", "cuneta", "sardinel", "techo"]),
    ("ct_sp_r", "G4", "Reposicion de edificaciones por demoler (sustitucion parcial)",
     ["edificacion", "aula", "ambiente", "pabellon", "taller", "laboratorio"]),
    ("ct_rc", "G4", "Reforzamiento convencional",
     ["reforzamiento", "estructura", "muro de contencion"]),
    ("ct_ic_r", "G4", "Reposicion posterior a intervencion contingente",
     ["edificacion", "aula", "modulo"]),
    ("ct_ae", "G4", "Acceso al servicio de energia electrica",
     ["electric", "energia", "panel solar", "subestacion", "red electrica"]),
    ("ct_acc", "G4", "Accesibilidad para personas con discapacidad",
     ["rampa", "accesibilidad", "ascensor", "elevador", "baranda", "pasamano"]),
    ("ct_st", "G5", "Sustitucion total integral del local",
     ["edificacion", "aula", "ambiente", "pabellon", "infraestructura",
      "usos multiples", "taller", "laboratorio"]),
    ("b_safil", "ET", "Saneamiento fisico legal",
     ["saneamiento fisico legal", "fisico legal", "intangible",
      "otras acciones de intangibles", "titulacion"]),
]


def cargar_activos_por_cui():
    """CUI (sin ceros) -> set de ACTIVO normalizados (Rep_Activos_F7)."""
    print("  · Cargando Rep_Activos_F7 (activos por CUI)...")
    import zipfile
    with zipfile.ZipFile(ACTIVOS_ZIP) as z:
        with z.open(ACTIVOS_CSV) as f:
            a = pd.read_csv(f, sep=";", encoding="utf-8-sig", dtype=str,
                            on_bad_lines="skip", usecols=["COD_UNICO", "ACTIVO"])
    a = a.dropna(subset=["COD_UNICO", "ACTIVO"])
    cui2act = {}
    for cui, act in zip(a["COD_UNICO"], a["ACTIVO"]):
        k = _dig(cui).lstrip("0") or "0"
        cui2act.setdefault(k, set()).add(na(act))
    print(f"    -> {len(cui2act):,} CUI con activos")
    return cui2act


def cargar_componentes_inversion():
    """HR Inversiones (gz): CUI (sin ceros) -> set de descripciones de
    producto/accion/componente. Es detalle adicional de activos por CUI que
    complementa a Rep_Activos_F7 para la deteccion de subcomponentes."""
    print("  · Cargando HR Inversiones (componentes por CUI)...")
    cui2comp = {}
    with gzip.open(INVERS_GZ, "rt", encoding="latin-1") as f:
        r = csv.reader(f, delimiter=";")
        h = next(r)
        idx = {c: i for i, c in enumerate(h)}
        iCu = idx.get("CODIGO_UNICO")
        cols = [idx[c] for c in ("DES_PRODUCTO", "DES_ACCION",
                                 "DES_TIPO_COMPONENTE") if c in idx]
        for row in r:
            if iCu is None or len(row) <= iCu:
                continue
            k = _dig(row[iCu]).lstrip("0") or "0"
            s = cui2comp.setdefault(k, set())
            for j in cols:
                if j < len(row) and row[j] and row[j].strip():
                    s.add(na(row[j]))
    print(f"    -> {len(cui2comp):,} CUI con componentes de inversion")
    return cui2comp


def main():
    print("=" * 70)
    print(" Subcomponentes de cierre de brecha - BASE CONSOLIDADA TOTAL")
    print("=" * 70)

    df = pd.read_parquet(CONS_PATH, columns=[
        "cod_local", "cui", "cod_modular", "grupo", "grupo_nombre",
        "nombre_ie", "departamento", "provincia", "distrito", "estado"])
    df = df[df["cod_local"].notna()].copy()
    df["cod_local"] = df["cod_local"].astype(str)
    print(f"  · Base consolidada: {len(df):,} filas | "
          f"{df['cod_local'].nunique():,} cod_local")

    cui2act = cargar_activos_por_cui()
    cui2comp = cargar_componentes_inversion()

    # patrones por subcomponente
    pats = {code: re.compile("|".join(re.escape(k) for k in kws))
            for code, _, _, kws in SUBCOMPONENTES}

    # ---- agregacion por cod_local ----
    def juntar(s):
        vals = [str(x) for x in s if pd.notna(x) and str(x).strip()]
        return " | ".join(dict.fromkeys(vals)) if vals else np.nan

    def cui_set(s):
        return sorted({_dig(x).lstrip("0") or "0" for x in s
                       if pd.notna(x) and _dig(x)})

    agg = df.groupby("cod_local").agg(
        cui_list=("cui", cui_set),
        cod_modular=("cod_modular", juntar),
        grupos=("grupo", lambda s: ",".join(map(str, sorted(set(s.dropna()))))),
        nombre_ie=("nombre_ie", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        departamento=("departamento", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        provincia=("provincia", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        distrito=("distrito", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        estado_base=("estado", juntar),
    ).reset_index()
    print(f"  · {len(agg):,} cod_local unicos (todos los grupos)")

    # activos e inversion consolidados por cod_local (via sus CUI)
    matriz_rows, evidencia_rows = [], []
    for _, r in agg.iterrows():
        cl = r["cod_local"]
        cuis = r["cui_list"]
        activos = set()
        for k in cuis:
            activos |= cui2act.get(k, set())     # activos F7
            activos |= cui2comp.get(k, set())    # componentes de inversion (gz)
        texto_act = " | ".join(sorted(activos))
        # estado de cierre: viene de la base consolidada (variable ESTADO)
        estado_base = r["estado_base"] if pd.notna(r["estado_base"]) else ""

        fila = {"cod_local": cl,
                "cui": ";".join(cuis) if len(cuis) else np.nan,
                "cod_modular": r["cod_modular"],
                "grupos": r["grupos"],
                "nombre_ie": r["nombre_ie"],
                "departamento": r["departamento"],
                "provincia": r["provincia"],
                "distrito": r["distrito"]}
        cumplidos = 0
        for code, gp, nombre, _ in SUBCOMPONENTES:
            pat = pats[code]
            matched = [a for a in activos if pat.search(a)]
            cumple = 1 if matched else 0
            fila[code] = cumple
            cumplidos += cumple
            evidencia_rows.append({
                "cod_local": cl,
                "cui": ";".join(cuis) if len(cuis) else np.nan,
                "cod_modular": r["cod_modular"],
                "grupos": r["grupos"],
                "subcomponente_codigo": code,
                "grupo_pnie": gp,
                "subcomponente": nombre,
                "cumple": cumple,
                "activo_que_valida": " | ".join(sorted(matched)) if matched else np.nan,
                "estado": estado_base or np.nan,
            })
        fila["subcomponentes_cumplidos"] = cumplidos
        fila["activos_f7"] = texto_act or np.nan
        fila["estado"] = estado_base or np.nan
        matriz_rows.append(fila)

    matriz = pd.DataFrame(matriz_rows)
    evidencia = pd.DataFrame(evidencia_rows)

    # orden de columnas de la matriz
    sc_cols = [c for c, *_ in SUBCOMPONENTES]
    meta = ["cod_local", "cui", "cod_modular", "grupos", "nombre_ie",
            "departamento", "provincia", "distrito"]
    cola = ["subcomponentes_cumplidos", "activos_f7", "estado"]
    matriz = matriz[meta + sc_cols + cola]

    os.makedirs(FINAL_DIR, exist_ok=True)
    def pq(d, ruta):
        d2 = d.copy()
        for c in d2.columns:
            if d2[c].dtype == object:
                d2[c] = d2[c].astype("string")
        d2.to_parquet(ruta, index=False)

    matriz.to_excel(os.path.join(FINAL_DIR, "subcomponentes_total_matriz.xlsx"), index=False)
    pq(matriz, os.path.join(FINAL_DIR, "subcomponentes_total_matriz.parquet"))
    if len(evidencia) < 1_000_000:
        evidencia.to_excel(os.path.join(FINAL_DIR, "subcomponentes_total_evidencia.xlsx"), index=False)
    pq(evidencia, os.path.join(FINAL_DIR, "subcomponentes_total_evidencia.parquet"))

    dic = pd.DataFrame([{"subcomponente_codigo": c, "grupo_pnie": g,
                         "subcomponente": n, "activos_clave": ", ".join(k)}
                        for c, g, n, k in SUBCOMPONENTES])
    dic.to_excel(os.path.join(FINAL_DIR, "subcomponentes_total_diccionario.xlsx"), index=False)

    print(f"\n  · Matriz   : {len(matriz):,} cod_local x {len(sc_cols)} subcomponentes")
    print(f"  · Evidencia: {len(evidencia):,} filas (cod_local x subcomponente)")
    print("\n  % de cod_local que CUMPLEN cada subcomponente:")
    for code, gp, nombre, _ in SUBCOMPONENTES:
        print(f"    {gp:3s} {code:9s} {matriz[code].mean()*100:5.1f}%  {nombre}")
    print("\n" + "=" * 70)
    print(" LISTO. Salidas en:", FINAL_DIR)
    print("=" * 70)


if __name__ == "__main__":
    main()
