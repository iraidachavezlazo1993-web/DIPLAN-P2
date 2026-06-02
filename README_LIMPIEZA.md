# DIPLAN-P2 — Limpieza, consolidación y subcomponentes

Proyecto para limpiar, normalizar y consolidar las bases de intervenciones de
PRONIED en 11 grupos, y validar los subcomponentes de cierre de brecha.

## Estructura del proyecto

```
DIPLAN_INPUTS/     Insumos de entrada (bases fuente y de cruce)
DIPLAN_SCRIPTS/    Scripts en orden de ejecución
DIPLAN_OUTPUT/     Resultados generados
```

### DIPLAN_SCRIPTS/ (ejecutar en orden)

1. **DIPLAN_01_LIMPIEZA.py** — Limpia y normaliza cada base, las organiza en los
   11 grupos, imputa `cod_local`, enriquece geografía (ubigeo/UGEL/padrón),
   asigna `fuente`, canoniza nombres de variables, elimina columnas duplicadas y
   genera: bases limpias por grupo, consolidados por grupo, base final
   armonizada, unión completa, diccionario y reporte.
2. **DIPLAN_02_ENRIQUECER.py** — Completa `tipo_activo` y `fecha_fin`
   (culminación / fin de ejecución / cierre / F9) en la base armonizada, cruzando
   por CUI con `Rep_Activos_F7` y `Rep_Inversiones`.
3. **DIPLAN_03_SUBCOMPONENTES.py** — Evalúa los 17 subcomponentes de cierre de
   brecha por `cod_local` sobre la base consolidada, validando con los activos de
   `Rep_Activos_F7` y los componentes de `HR Inversiones`.

### DIPLAN_INPUTS/ — Insumos

**Bases fuente de los grupos**
| Archivo | Grupos |
|---|---|
| ANIN.xlsx | 6 |
| FONCODES.xlsx | 6 |
| PEIP.xlsx / PEIP_solomant.xlsx | 3 / 6 |
| UGM.xlsx | 4, 6 |
| UGME.xlsx | 1, 3 |
| UGRD.xlsx | 3 |
| UGSC.xlsx | 5, 10 |
| 2026.03.30 Zonales_UZ.xlsx | 5, 10 |
| DIGEGED.xlsx | 6, 7, 8, 9 |
| DRELM.xlsx | 1, 3, 6, 7, 9 |
| Anexo 01 … 300426.xlsx | (referencia) |
| Anexo 02 … 300426_depurada.xlsx | 1, 3, 4, 6, 11 |
| df_pi_listados_final_v7.xlsx | 2 (PI, base ya consolidada) |

**Insumos de cruce / enriquecimiento**
| Archivo | Para qué |
|---|---|
| Ordenamiento_ultimo.xlsx | define los 11 grupos |
| df_vinculaciones_updated_20260527.xlsx | cui/cod_mod → cod_local |
| Copia de Padron_web.csv | imputar cod_local + geografía |
| ubigeo_UGEL.xlsx | departamento/provincia/distrito/DRE/UGEL/nombre IE |
| ubigeo_pronied.xlsx | geografía + ruralidad + cod_mod |
| ubigeo_cui.csv | geografía por CUI |
| Fichas_Indicadores_Cierre_Brecha_PNIE_060526.xlsx | catálogo de subcomponentes |
| Rep_Activos_F7_02SET2024.zip | activo por CUI (subcomponentes) |
| Rep_Inversiones_13ABR2026_EDU.xlsb | fecha fin/cierre por CUI |
| HR 079179-2026 - Inversiones.csv.gz | componentes por CUI |

### DIPLAN_OUTPUT/ — Resultados

```
bases_limpias/Grupo_XX_.../          base limpia por hoja (xlsx + parquet)
base_final/
  base_final_armonizada.(xlsx|parquet)        tabla estándar de todos los grupos
  base_final_union_completa.parquet           todas las columnas (solo parquet)
  consolidado_por_grupo/                       1 consolidado por grupo (xlsx+parquet)
  subcomponentes_total_matriz.(xlsx|parquet)   1/0 por subcomponente por cod_local
  subcomponentes_total_evidencia.(xlsx|parquet) detalle por cod_local x subcomp
  subcomponentes_total_diccionario.xlsx        subcomponentes + palabras clave
diccionario_datos.xlsx
reporte_limpieza.xlsx
```

## Subcomponentes — replicabilidad

La matriz y la evidencia incluyen el **texto crudo** del que se extraen las
palabras clave de identificación, para que el cumplimiento pueda revisarse,
replicarse y mejorarse:

- `activos_f7` — activos del local según `Rep_Activos_F7` (campo `ACTIVO`).
- `componentes_inversion` — componentes según `HR Inversiones`
  (`DES_PRODUCTO`, `DES_ACCION`, `DES_TIPO_COMPONENTE`).
- `activo_que_valida` (evidencia) — qué activo concreto disparó el `cumple=1`.
- `estado` — estado de la inversión (de la base), para cotejo.

El diccionario lista, por subcomponente, las `palabras_clave`, su
`fuente_palabras_clave` y la `regla` aplicada.

> El criterio de cumplimiento es un proxy por palabras clave sobre los nombres de
> activos/componentes; no aplica las fórmulas formales (áreas, metros lineales)
> de la ficha PNIE. Las llaves (`cui`, `cod_local`, `cod_modular`) y los textos
> crudos permiten refinarlo.
