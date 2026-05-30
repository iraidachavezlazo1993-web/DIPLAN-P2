# -*- coding: utf-8 -*-
# Completa tipo_activo y fecha_fin en la base armonizada usando las fuentes de
# inversiones (Rep_Activos_F7 -> activo por CUI; Rep_Inversiones -> fecha de
# culminacion / fin de ejecucion / cierre / F9 por CUI). Solo rellena vacios.
import os
import re
import zipfile
import pandas as pd

# El proyecto se ejecuta desde la raiz (insumos y carpeta de salida estan ahi).
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

FINAL = os.path.join("DIPLAN_OUTPUT", "base_final")
ARM = os.path.join(FINAL, "base_final_armonizada")
ACTIVOS_ZIP = "Rep_Activos_F7_02SET2024.zip"
ACTIVOS_CSV = "Rep_Activos_F7_02SET2024.csv"
INVERS = "Rep_Inversiones_13ABR2026_EDU.xlsb"


def cui_key(x):
    d = re.sub(r"\D", "", str(x)) if x is not None else ""
    return (d.lstrip("0") or "0") if d else None


def serial_a_fecha(v):
    # convierte el serial de fecha de Excel (p.ej. 45291.0) a 'DD/MM/YYYY'
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if not (f > 0):
        return None
    return (pd.Timestamp("1899-12-30") + pd.to_timedelta(int(f), unit="D")).strftime("%d/%m/%Y")


def activos_por_cui():
    with zipfile.ZipFile(ACTIVOS_ZIP) as z:
        with z.open(ACTIVOS_CSV) as f:
            a = pd.read_csv(f, sep=";", encoding="utf-8-sig", dtype=str,
                            on_bad_lines="skip", usecols=["COD_UNICO", "ACTIVO"])
    a = a.dropna(subset=["COD_UNICO", "ACTIVO"])
    m = {}
    for cu, ac in zip(a["COD_UNICO"], a["ACTIVO"]):
        k = cui_key(cu)
        ac = ac.strip()
        if k and ac:
            s = m.setdefault(k, [])
            if ac not in s:
                s.append(ac)
    return {k: " | ".join(v) for k, v in m.items()}


def fecha_fin_por_cui():
    # fecha de fin = culminacion fisica, o fin de ejecucion, o fecha de cierre
    # (registro de cierre / F9). La fecha de cierre tambien cuenta como fin.
    cols = ["CODIGO_UNICO", "CULMINACION_EJEC_FISICA", "FEC_FIN_EJUCION",
            "FEC_REGISTRO_CIERRE", "FEC_REG_F9"]
    x = pd.read_excel(INVERS, engine="pyxlsb", usecols=cols)
    m = {}
    for cu, c1, c2, c3, c4 in zip(x["CODIGO_UNICO"], x["CULMINACION_EJEC_FISICA"],
                                  x["FEC_FIN_EJUCION"], x["FEC_REGISTRO_CIERRE"],
                                  x["FEC_REG_F9"]):
        k = cui_key(cu)
        if not k or k in m:
            continue
        for v in (c1, c2, c3, c4):
            fecha = serial_a_fecha(v)
            if fecha:
                m[k] = fecha
                break
    return m


def main():
    df = pd.read_parquet(ARM + ".parquet")
    print("armonizada:", len(df), "filas")

    act = activos_por_cui()
    fin = fecha_fin_por_cui()
    print("activos por CUI:", len(act), "| fechas fin por CUI:", len(fin))

    cuik = df["cui"].map(cui_key)

    falta_ta = df["tipo_activo"].isna()
    df.loc[falta_ta, "tipo_activo"] = cuik[falta_ta].map(act).values

    falta_ff = df["fecha_fin"].isna()
    df.loc[falta_ff, "fecha_fin"] = cuik[falta_ff].map(fin).values

    print("tipo_activo lleno: %.1f%%" % (df["tipo_activo"].notna().mean() * 100))
    print("fecha_fin lleno:   %.1f%%" % (df["fecha_fin"].notna().mean() * 100))

    df.to_excel(ARM + ".xlsx", index=False)
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype("string")
    out.to_parquet(ARM + ".parquet", index=False)
    print("guardado.")


if __name__ == "__main__":
    main()
