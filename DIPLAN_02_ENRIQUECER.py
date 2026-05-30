# -*- coding: utf-8 -*-
# Completa tipo_activo y fecha_fin en la base armonizada usando las fuentes
# de inversiones (Rep_Activos_F7 y Rep_Inversiones). Solo rellena vacios.
import os
import re
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
FINAL = os.path.join(BASE, "DIPLAN-OUTPUT", "base_final")
ARM = os.path.join(FINAL, "base_final_armonizada")
ACTIVOS = os.path.join(BASE, "_tmp_zip", "Rep_Activos_F7_02SET2024.csv")
INVERS = os.path.join(BASE, "Rep_Inversiones_13ABR2026_EDU.xlsb")


def cui_key(x):
    d = re.sub(r"\D", "", str(x)) if x is not None else ""
    return (d.lstrip("0") or "0") if d else None


def activos_por_cui():
    a = pd.read_csv(ACTIVOS, sep=";", encoding="utf-8-sig", dtype=str,
                    on_bad_lines="skip", usecols=["COD_UNICO", "ACTIVO"])
    a = a.dropna(subset=["COD_UNICO", "ACTIVO"])
    m = {}
    for cu, ac in zip(a["COD_UNICO"], a["ACTIVO"]):
        k = cui_key(cu)
        if k:
            s = m.setdefault(k, [])
            ac = ac.strip()
            if ac and ac not in s:
                s.append(ac)
    return {k: " | ".join(v) for k, v in m.items()}


def _serial_a_fecha(v):
    # convierte serial de Excel (p.ej. 45291.0) a 'DD/MM/YYYY'
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if not (f > 0):          # descarta NaN, 0 y negativos
        return None
    base = pd.Timestamp("1899-12-30")
    return (base + pd.to_timedelta(int(f), unit="D")).strftime("%d/%m/%Y")


def fecha_fin_por_cui():
    # prioridad: culminacion fisica > fin de ejecucion > registro de cierre > F9
    cols = ["CODIGO_UNICO", "CULMINACION_EJEC_FISICA", "FEC_FIN_EJUCION",
            "FEC_REGISTRO_CIERRE", "FEC_REG_F9"]
    x = pd.read_excel(INVERS, engine="pyxlsb", usecols=cols)
    pri = cols[1:]
    m = {}
    for _, r in x.iterrows():
        k = cui_key(r["CODIGO_UNICO"])
        if not k or k in m:
            continue
        for c in pri:
            fecha = _serial_a_fecha(r[c])
            if fecha:
                m[k] = fecha
                break
    return m


def main():
    df = pd.read_parquet(ARM + ".parquet")
    print("armonizada:", len(df), "filas")

    act = activos_por_cui()
    fin = fecha_fin_por_cui()
    print("activos:", len(act), "| fechas fin:", len(fin))

    cuik = df["cui"].map(cui_key)

    # tipo_activo: rellena vacios con el activo del CUI
    falta_ta = df["tipo_activo"].isna()
    df.loc[falta_ta, "tipo_activo"] = cuik[falta_ta].map(act)

    # fecha_fin: rellena vacios con la fecha de cierre/termino del CUI
    falta_ff = df["fecha_fin"].isna()
    df.loc[falta_ff, "fecha_fin"] = cuik[falta_ff].map(fin)

    print("tipo_activo lleno:", f"{df['tipo_activo'].notna().mean()*100:.1f}%")
    print("fecha_fin lleno:", f"{df['fecha_fin'].notna().mean()*100:.1f}%")

    df.to_excel(ARM + ".xlsx", index=False)
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype("string")
    out.to_parquet(ARM + ".parquet", index=False)
    print("guardado.")


if __name__ == "__main__":
    main()
