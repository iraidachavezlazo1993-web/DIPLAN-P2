# -*- coding: utf-8 -*-
"""
================================================================================
 DIPLAN-P2 | Verificacion de SUBCOMPONENTES de cierre de brecha en Grupo 2 (PI)
================================================================================
Insumo de subcomponentes:
  Fichas_Indicadores_Cierre_Brecha_PNIE_060526.xlsx
  (hoja "Matriz transpuesta v26": 17 subcomponentes de costo ct_* / b_safil,
   agrupados en G1..G5 y ET, con sus variables operativas y definiciones).

Base evaluada:
  output/bases_limpias/Grupo_02_PI/df_pi_listados_final_v7.parquet

Salida (por cod_local):
  - output/base_final/subcomponentes_grupo2_matriz.(xlsx|parquet)
      1 fila por cod_local; 1 columna 1/0 por subcomponente; total cumplidos.
  - output/base_final/subcomponentes_grupo2_largo.(xlsx|parquet)
      1 fila por (cod_local, subcomponente): cod_local, grupo_pnie, codigo,
      subcomponente, cumple (1/0).
  - output/base_final/subcomponentes_diccionario.xlsx

"cumple" = 1 si en el texto de las intervenciones del cod_local (nombre del
proyecto + intervencion detectada + tipo de inversion) aparece alguna palabra
clave del subcomponente; 0 en caso contrario.
================================================================================
"""
import os
import re
import unicodedata
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
FINAL_DIR = os.path.join(BASE_DIR, "output", "base_final")
G2_PATH   = os.path.join(BASE_DIR, "output", "bases_limpias",
                         "Grupo_02_PI", "df_pi_listados_final_v7.parquet")
GRUPO_NUM, GRUPO_NOMBRE = 2, "PI"

# 17 subcomponentes (codigo, grupo PNIE, nombre, palabras clave de deteccion)
SUBCOMPONENTES = [
    ("ct_st_d", "G1", "Demolicion de locales para sustitucion total",
     ["demolicion", "sustitucion total", "sust_total", "sustitucion"]),
    ("ct_sp_d", "G1", "Demolicion de edificaciones para sustitucion parcial",
     ["demolicion", "sustitucion parcial", "sust_parcial"]),
    ("ct_ri", "G1", "Reforzamiento incremental",
     ["reforzamiento", "reforzar", "reforzamiento incremental"]),
    ("ct_ic", "G1", "Intervencion contingente",
     ["contingente", "contingencia", "intervencion contingente"]),
    ("ct_cp", "G1", "Cercos perimetricos",
     ["cerco perimetrico", "cerco_perimetrico", "cerco"]),
    ("ct_aad", "G2", "Acceso a agua y desague",
     ["agua", "desague", "desagee", "conexion a la red", "red publica",
      "saneamiento basico", "alcantarillado"]),
    ("ct_cad", "G2", "Mejora de la calidad del servicio de agua y saneamiento",
     ["cisterna", "tanque elevado", "saneamiento", "tratamiento de agua",
      "agua potable"]),
    ("ct_ce", "G3", "Mejora de la calidad del servicio de energia electrica",
     ["instalaciones electricas", "instalacion electrica", "electrico",
      "circuito electrico", "tablero"]),
    ("ct_me", "G3", "Mobiliario y equipamiento de la infraestructura existente",
     ["mobiliario", "equipamiento", "equipo", "carpeta", "pizarra"]),
    ("ct_ene", "G3", "Mantenimiento correctivo de elementos no estructurales",
     ["mantenimiento correctivo", "mantenimiento", "pisos", "carpinteria",
      "pintura", "elementos no estructurales"]),
    ("ct_sp_r", "G4", "Reposicion de edificaciones por demoler (sustitucion parcial)",
     ["sustitucion parcial", "sust_parcial", "reposicion"]),
    ("ct_rc", "G4", "Reforzamiento convencional",
     ["reforzamiento convencional", "reforzamiento", "reforzar"]),
    ("ct_ic_r", "G4", "Reposicion posterior a intervencion contingente",
     ["intervencion contingente", "contingente", "reposicion"]),
    ("ct_ae", "G4", "Acceso al servicio de energia electrica",
     ["energia electrica", "electrificacion", "acceso a energia",
      "red electrica", "panel solar"]),
    ("ct_acc", "G4", "Accesibilidad para personas con discapacidad",
     ["accesibilidad", "rampa", "discapacidad", "ascensor"]),
    ("ct_st", "G5", "Sustitucion total integral del local",
     ["sustitucion total", "sust_total", "integral", "sustitucion"]),
    ("b_safil", "ET", "Saneamiento fisico legal",
     ["saneamiento fisico legal", "fisico legal", "sfl", "saneamiento legal",
      "titulacion"]),
]


def na(texto):
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"\s+", " ", t).strip()


def main():
    print("=" * 70)
    print(" Verificacion de SUBCOMPONENTES de cierre de brecha - Grupo 2 (PI)")
    print("=" * 70)
    print(f"  · {len(SUBCOMPONENTES)} subcomponentes (G1-G5, ET)")

    g2 = pd.read_parquet(G2_PATH)
    g2 = g2.loc[:, ~g2.columns.duplicated()].copy()
    print(f"  · Base Grupo 2: {len(g2):,} filas")

    cols_txt = [c for c in ["nombre_pi", "tipo_intervencion_detectada",
                            "tipo_inversion", "criterio_que_falla", "nombre_ie"]
                if c in g2.columns]
    g2["_cl"] = g2["cod_local"].astype(str)
    # une el texto de las columnas relevantes (robusto a NaN)
    g2["_txt"] = g2[cols_txt].fillna("").astype(str).agg(" ".join, axis=1).map(na)

    por_local = g2.groupby("_cl").agg(
        _txt=("_txt", lambda s: " | ".join(s)),
        nombre_ie=("nombre_ie", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        departamento=("departamento", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        provincia=("provincia", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        distrito=("distrito", lambda s: next((x for x in s if pd.notna(x)), np.nan)),
        n_inversiones=("_txt", "size"),
    ).reset_index().rename(columns={"_cl": "cod_local"})
    print(f"  · {len(por_local):,} cod_local unicos en el Grupo 2")

    textos = por_local["_txt"].tolist()

    matriz = por_local[["cod_local", "nombre_ie", "departamento",
                        "provincia", "distrito", "n_inversiones"]].copy()
    matriz.insert(1, "grupo", GRUPO_NUM)
    matriz.insert(2, "grupo_nombre", GRUPO_NOMBRE)

    col_por_sc = {}
    for code, gp, nombre, kws in SUBCOMPONENTES:
        patron = re.compile("|".join(re.escape(k) for k in kws))
        matriz[code] = [1 if patron.search(t) else 0 for t in textos]
        col_por_sc[code] = code

    cols_sc = [c for c, *_ in SUBCOMPONENTES]
    matriz["subcomponentes_cumplidos"] = matriz[cols_sc].sum(axis=1)

    # formato largo
    largo_rows = []
    arr = {code: matriz[code].values for code, *_ in SUBCOMPONENTES}
    cls = matriz["cod_local"].values
    for i, cl in enumerate(cls):
        for code, gp, nombre, _ in SUBCOMPONENTES:
            largo_rows.append({
                "cod_local": cl, "grupo": GRUPO_NUM, "grupo_nombre": GRUPO_NOMBRE,
                "grupo_pnie": gp, "subcomponente_codigo": code,
                "subcomponente": nombre, "cumple": int(arr[code][i]),
            })
    largo = pd.DataFrame(largo_rows)

    os.makedirs(FINAL_DIR, exist_ok=True)
    def pq(df, ruta):
        df2 = df.copy()
        for c in df2.columns:
            if df2[c].dtype == object:
                df2[c] = df2[c].astype("string")
        df2.to_parquet(ruta, index=False)

    matriz.to_excel(os.path.join(FINAL_DIR, "subcomponentes_grupo2_matriz.xlsx"), index=False)
    pq(matriz, os.path.join(FINAL_DIR, "subcomponentes_grupo2_matriz.parquet"))
    if len(largo) < 1_000_000:
        largo.to_excel(os.path.join(FINAL_DIR, "subcomponentes_grupo2_largo.xlsx"), index=False)
    pq(largo, os.path.join(FINAL_DIR, "subcomponentes_grupo2_largo.parquet"))

    dic = pd.DataFrame([{"subcomponente_codigo": c, "grupo_pnie": g,
                         "subcomponente": n, "palabras_clave": ", ".join(k)}
                        for c, g, n, k in SUBCOMPONENTES])
    dic.to_excel(os.path.join(FINAL_DIR, "subcomponentes_diccionario.xlsx"), index=False)

    print(f"\n  · Matriz: {len(matriz):,} cod_local x {len(cols_sc)} subcomponentes")
    print(f"  · Largo : {len(largo):,} filas")
    print("\n  % de cod_local que CUMPLEN cada subcomponente:")
    for code, gp, nombre, _ in SUBCOMPONENTES:
        pct = matriz[code].mean() * 100
        print(f"    {gp:3s} {code:9s} {pct:5.1f}%  {nombre}")
    print("\n" + "=" * 70)
    print(" LISTO. Salidas en:", FINAL_DIR)
    print("=" * 70)


if __name__ == "__main__":
    main()
