# DIPLAN-P2 — Limpieza, consolidación y subcomponentes

Procesa las bases de las distintas unidades de PRONIED: las limpia, normaliza,
organiza en 11 grupos, consolida y evalúa los subcomponentes de cierre de brecha.

## Orden de ejecución

Los scripts están en `DIPLAN_SCRIPTS/` y se ejecutan en orden (cada uno hace
`chdir` a la raíz del proyecto, donde están los insumos):

```bash
pip install pandas openpyxl pyarrow pyxlsb
python DIPLAN_SCRIPTS/DIPLAN_01_LIMPIEZA.py        # limpia, consolida y arma bases
python DIPLAN_SCRIPTS/DIPLAN_02_ENRIQUECER.py      # completa tipo_activo y fecha_fin
python DIPLAN_SCRIPTS/DIPLAN_03_SUBCOMPONENTES.py  # subcomponentes sobre la consolidada
```

> Importante: respetar el orden. El paso 02 completa la base armonizada con
> las fuentes de inversiones; si se vuelve a correr el paso 01 hay que correr
> nuevamente 02 y 03.

## Salidas (`DIPLAN_OUTPUT/`)

Equivale a la carpeta de resultados del proyecto (local: `…/MINEDU/03_output`).

```
DIPLAN_OUTPUT/
├── bases_limpias/Grupo_XX/                  data limpia individual (xlsx + parquet)
├── base_final/
│   ├── consolidado_por_grupo/Grupo_XX...    consolidada por grupo (todos los campos)
│   ├── base_final_armonizada.xlsx/.parquet  consolidada total (columnas estándar)
│   ├── base_final_union_completa.parquet    unión total (todas las columnas)
│   ├── subcomponentes_total_matriz...       subcomponentes 1/0 por cod_local
│   ├── subcomponentes_total_evidencia...    evidencia (activo y estado para cotejo)
│   └── subcomponentes_total_diccionario.xlsx
├── diccionario_datos.xlsx
└── reporte_limpieza.xlsx
```

## Reglas de limpieza

- Fechas en `DD/MM/YYYY`; montos/devengado numéricos.
- `cod_local` 6 dígitos (imputado por cui/cod_mod/codinst con vinculaciones y
  padrón); `cod_mod` 7 dígitos (apilable); `cui` 7 dígitos.
- Texto normalizado; comentarios conservados.
- Geografía (departamento/provincia/distrito/ubigeo/DRE/UGEL), `nombre_ie` y
  `area` (Urbano/Rural) desde los padrones (cod_local y CUI).
- Variables canonizadas (minúscula/snake_case); columnas con contenido
  duplicado eliminadas (cada variable una sola vez).
- Campo `fuente`: unidad que reporta; en Anexos 1/2 el Remitente; en PI el
  campo `fuentes_concat`.

## Base armonizada — campos

`cod_local, cui, cod_modular, grupo, df_name, fuente, nombre_ie, departamento,
provincia, distrito, ubigeo, dre, ugel, area, tipo_intervencion, tipo_activo,
estado, monto, devengado, avance_fisico, fecha_inicio, fecha_fin, anio,
comentario`.

`tipo_intervencion` = nombre del grupo (en PI, el tipo de inversión);
`tipo_activo` = activo intervenido. `tipo_activo` y `fecha_fin` se completan en
el paso 02 con Rep_Activos_F7 (activo por CUI) y Rep_Inversiones (fecha de
culminación / fin de ejecución / cierre / F9 por CUI; la fecha de cierre cuenta
como fin).

## Subcomponentes

17 subcomponentes de cierre de brecha (Fichas_Indicadores_Cierre_Brecha_PNIE,
G1–G5/ET) evaluados por cod_local sobre la consolidada total: 1 = cumple,
0 = no, según el activo presente en las fuentes de inversiones. La base de
evidencia adjunta el activo que valida y el estado, con las llaves
(cui, cod_modular), para cotejo.
