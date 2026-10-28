# Herramientas de campo: Cronómetro SPM + Control de pruebas

Este repositorio publica dos herramientas con un menú común:

| Dirección | Herramienta | Archivos |
|---|---|---|
| `https://feror1989-ui.github.io/SPM/` | Cronómetro SPM | `index.html` |
| `https://feror1989-ui.github.io/SPM/pruebas/` | Control de pruebas de pozos | `pruebas/index.html`, `pruebas/app.js`, `pruebas/styles.css` |

## Actualizar el repositorio SPM (sin Git)

1. Entra a `https://github.com/feror1989-ui/SPM`.
2. **Add file → Upload files**.
3. Arrastra desde esta carpeta: `index.html`, `README.md`, `.gitignore` y la **carpeta `pruebas` completa** (arrastra la carpeta, no solo sus archivos, para que GitHub conserve la ruta `pruebas/`).
4. **Commit changes**. En 1–2 minutos ambas páginas quedan actualizadas.

---

# Control de pruebas de pozos (LPT / MPFM)

Herramienta web para revisar, a partir del **consolidado de pruebas** en Excel, qué pozos y clusters **no se han medido en el mes** respecto a su historial, y cuántos días tomaría medirlos.

Todo se procesa en el navegador: los archivos Excel **no se suben a ningún servidor**.

## Qué hace

- **Carga uno o varios consolidados** (`.xlsx`), por ejemplo Girasol y Jazmín al tiempo. Busca automáticamente la hoja con las columnas `CLUSTER` y `Pozo`.
- **Pendientes del mes**: un pozo queda *pendiente* si tiene pruebas en el historial de referencia (por defecto **todo el historial**) y ninguna en el mes revisado.
- **Vista por cluster**: pozos totales, medidos, pendientes, % de avance, **horas y días estimados** para medir lo que falta, máximo de días sin medir y crudo pendiente. Clic en un cluster para ver sus pozos.
- **Vista por pozo**: estado, última medición, días sin medir, meses medidos, número de pruebas, tiempo estimado, crudo/líquido de la última prueba y último equipo. Buscador y orden por columna.
- **Plan por equipo**: reparte los clusters pendientes entre los equipos disponibles y estima fechas de inicio/fin y si se alcanza dentro del mes.
- **Avance mensual**: pozos distintos medidos por mes.
- **Alertas de calidad de datos**: fechas escritas como texto, fechas fuera de rango (p. ej. 2016 en vez de 2026), clusters o pozos escritos de distintas formas (`PH3` = `PH03`, `GIRASOL-H0712` = `GIRASOL H0712`), pozos en varios clusters.
- **Exportar a Excel** con hojas: Por cluster, Pozos pendientes, Todos los pozos, Plan por equipo y Alertas.

## Cómo se estima el tiempo

Para cada pozo:

```
tiempo por pozo = promedio histórico de (fecha finalización − fecha alineación)
                  + tiempo típico entre el fin de una prueba y la alineación de la siguiente (mismo equipo)
```

Si un pozo no tiene fechas suficientes se usa el promedio de su cluster, y si no, el del campo.

```
días del cluster = suma de tiempos de sus pozos pendientes / horas operativas por día
```

Parámetros ajustables en pantalla: historial de referencia, ocultar pozos sin medir hace más de X días (posibles pozos inactivos), horas operativas por día, equipos disponibles (por defecto, los equipos usados en el mes), horas de movilización entre clusters, fecha de inicio y prioridad del plan.

## Formato esperado del Excel

Una hoja (preferiblemente `PRUEBAS`) con una fila de encabezados que incluya al menos:

| Columna | Obligatoria |
|---|---|
| CLUSTER | Sí |
| Pozo | Sí |
| Fecha inicio de prueba | Sí (o fecha de alineación / finalización) |
| Fecha de alineación, Fecha finalización de prueba, Duración [horas] | Para estimar tiempos |
| Equipo, Total líquido, BS&W, Crudo COR., Prueba Válida | Opcionales |

El nombre del campo se toma del nombre del archivo (`... CAMPO GIRASOL.xlsx` → `GIRASOL`).
Si la columna *Prueba Válida* dice `NO`, esa prueba se ignora.

> **No subas los Excel con datos de producción al repositorio**: si el repositorio es público, cualquiera podría verlos. El `.gitignore` ya excluye los archivos `.xlsx`.

## Probar en tu computador

Basta con abrir `pruebas/index.html` en el navegador (requiere internet para cargar la librería SheetJS desde cdnjs).
