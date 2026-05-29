# DIPLAN-P2 — Limpieza, normalización y consolidación de bases

Este proyecto toma todas las bases (libros de Excel) entregadas por las distintas
unidades de PRONIED, las **limpia**, **normaliza** y las **organiza en 11 grupos**
según el insumo `Ordenamiento_ultimo.xlsx`. Adicionalmente genera un
**diccionario de datos** y dos **bases finales consolidadas**.

> El **Grupo 2 (PI)** se incorpora con su base final ya consolidada
> (`df_pi_listados_final_v7.xlsx`, hoja `con_cod_local`), pasándola por el mismo
> pipeline para estandarizar códigos, fechas, montos, ubigeo y `fuente`.
> Los **insumos** (`Ordenamiento_ultimo.xlsx`, `df_vinculaciones_*.xlsx`,
> `ubigeo_UGEL.xlsx`, `Copia de Padron_web.csv`) **no se limpian**; solo se usan
> como referencia/cruce.

### Campo `fuente`
Todas las bases (individuales y consolidadas) incluyen una columna **`fuente`**
que indica la unidad/insumo que reporta principalmente el registro
(ANIN, UGM, DIGEGED, PI, …). En los **Anexos 1 y 2** la fuente se toma **por
fila** del campo *Remitente* (p. ej. `Anexo 2 - MP Tambopata`).

### Padrón web (imputación + respaldo geográfico)
El insumo `Copia de Padron_web.csv` (padrón nacional de locales) se usa como
fuente **adicional** para imputar `cod_local` (vía `cod_mod` / `codinst`) y como
**respaldo** de ubicación (departamento/provincia/distrito/ubigeo/DRE/UGEL)
cuando el local no está en `ubigeo_UGEL`. Cobertura resultante: **99.8% de
`cod_local`** y **99.6% de geografía/UGEL**.

---

## 1. Cómo ejecutar

```bash
pip install pandas openpyxl pyarrow
python3 limpieza_diplan_p2.py
```

Todas las salidas quedan en la carpeta `output/`.

---

## 2. Reglas de limpieza aplicadas

| Tipo de campo | Regla |
|---|---|
| **Fechas** | Formato `DD/MM/YYYY` (sin hora). |
| **Monto / Devengado / Costo / PIM** | Número puro (sin letras ni `S/`, sin separador de miles). |
| **cod_local** | Un solo código por celda, **6 dígitos** (ceros a la izquierda). Si falta, se imputa cruzando con `vinculaciones` (vía `cui` o `cod_modular`). |
| **cod_modular** | **7 dígitos**. Se permite **apilado** (varios códigos por celda, separados por ` \| `) porque así corresponde a la realidad modular. |
| **cui** | 7 dígitos (ceros a la izquierda). |
| **Texto disperso** | Normalizado: recortado, espacios colapsados, sin saltos de línea; placeholders de vacío (`-`, `·`, `N/A`, `SIN DATO`, …) → nulo. |
| **Comentarios / Observaciones** | Se **conservan completos** (solo se recortan espacios) porque suelen contener mayor información. |

### Enriquecimiento geográfico (ubigeo + UGEL)
A partir del `cod_local` limpio, **todas las bases** reciben las columnas
canónicas `departamento`, `provincia`, `distrito`, `centro_poblado`, `ubigeo`,
`dre` y `ugel`, tomadas del insumo `ubigeo_UGEL.xlsx`.

---

## 3. Salidas (`output/`)

```
output/
├── bases_limpias/
│   ├── Grupo_01_MOBILIARIO_Y_EQUIPO/   <df>.xlsx + <df>.parquet
│   ├── Grupo_03_MODULOS/
│   ├── Grupo_04_ACONDICIONAMIENTO/
│   ├── Grupo_05_ASISTENCIAS_SIMILAR/
│   ├── Grupo_06_MANTENIMIENTO/
│   ├── Grupo_07_SANEAMIENTO/
│   ├── Grupo_08_PROGRAMAS_PRESUPUESTALES/
│   ├── Grupo_09_OTROS/
│   ├── Grupo_10_INSPECCIONES/
│   └── Grupo_11_DEMOLICIONES/
├── base_final/
│   ├── base_final_armonizada.xlsx / .parquet      (columnas estándar)
│   └── base_final_union_completa.parquet          (todas las columnas originales)
├── diccionario_datos.xlsx
└── reporte_limpieza.xlsx
```

- **base_final_armonizada**: una sola tabla con columnas estándar
  (`cod_local`, `cod_modular`, `cui`, `grupo`, `df_name`, `nombre_ie`,
  `departamento`, `provincia`, `distrito`, `ubigeo`, `tipo_intervencion`,
  `estado`, `monto`, `devengado`, `avance_fisico`, `fecha`, `anio`,
  `comentario`). Ideal para análisis transversal.
- **base_final_union_completa**: apila todas las bases conservando **todas** sus
  columnas originales + metadatos (`_grupo`, `_df_name`, `_libro`, `_hoja`).
  Se entrega en Parquet por su gran tamaño (tabla muy ancha).
- **diccionario_datos.xlsx**: por cada base, lista de columnas, tipo de limpieza
  aplicada, conteo de no nulos y un valor de ejemplo.
- **reporte_limpieza.xlsx**: resumen por base (filas, columnas, cod_local
  imputados, filas con ubigeo) y resumen por grupo.

---

## 4. Mapa de los 11 grupos (insumo `Ordenamiento_ultimo.xlsx`)

| Grupo | Nombre | Bases incluidas (de los archivos disponibles) |
|---|---|---|
| 1 | MOBILIARIO Y EQUIPO | `df_ugme_mobiliario`, `df_drelm_mobiliario`, `df_anexo2_b` |
| 2 | PI | *(omitido — ya limpio)* |
| 3 | MODULOS | `df_ugme_modulares`, `df_ugme_conservacion`, `df_ugrd_mbr`, `df_ugrd_me`, `df_ugrd_pircc_pircc`, `df_peip_contingencia`, `df_drelm_modulo`, `df_anexo2_a` |
| 4 | ACONDICIONAMIENTO | `df_ugm_acondicionamiento`, `df_ugm_accesibilidad_2026`, `df_anexo2_e` |
| 5 | ASISTENCIAS-SIMILAR | `df_ugsc_asitec`, `df_uz_asesoramiento` |
| 6 | MANTENIMIENTO | `df_anin_mantenimiento`, `df_foncode_2025`, `df_foncode_2026`, `df_peip_mantenimiento`, `df_ugm_mantenimiento_2025`, `df_ugm_mantenimiento_2026`, `df_digeged_mto`, `df_drelm_mantenimiento`, `df_anexo2_c` |
| 7 | SANEAMIENTO | `df_digeged_sfl`, `df_drelm_sfl` |
| 8 | PROGRAMAS PRESUPUESTALES | `df_digeged_programa` |
| 9 | OTROS | `df_digeged_otros25`, `df_drelm_servicios` |
| 10 | INSPECCIONES | `df_uz_inspecciones`, `df_ugsc_seguimiento` |
| 11 | DEMOLICIONES | `df_anexo2_d` |

---

## 5. Observaciones / pendientes

- Algunas bases referenciadas en el Ordenamiento **no fueron entregadas** como
  archivo (p. ej. `MANTENIMIENTO_PRONIED`, `Base de Inversiones 2026`,
  `UGEO_2027-2025`). Pertenecen mayormente al Grupo 2 (omitido) o al Grupo 6.
  Ver `reporte_limpieza.xlsx`.
- `df_peip_mantenimiento` ahora usa el archivo corregido `PEIP_solomant.xlsx`,
  que separa `cui` y `cod_local`; con ello cruza al 100% con ubigeo.
- Hojas como `df_uz_asesoramiento` no tienen código de local (el registro es por
  entidad/UGEL, no por local educativo); en esos casos `cod_local` y el ubigeo
  quedan vacíos por naturaleza del dato.
- Las hojas de **DRELM** y de **mantenimiento UGM** traen bloques de columnas
  multinivel; se conservó el encabezado de mayor detalle. Las transformaciones de
  *pivoteo/transposición* descritas en las observaciones del Ordenamiento se
  consideran una etapa posterior de modelado, no de limpieza.
