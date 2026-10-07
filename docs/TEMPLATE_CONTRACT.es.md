# Contrato de plantillas monetarias BAP

Cada diseño monetario vive en su propia carpeta dentro de `templates/money/`.
El programa descubre automáticamente todas las carpetas que contienen `template.json`.
No es necesario recompilar el EXE para agregar, quitar o reemplazar un diseño.

## Estructura mínima

```
templates/money/Mi_Diseno/
  template.json
  report.html
  style.css
  assets/              # opcional
```

`template.json` mínimo:

```json
{
  "name": "Nombre visible",
  "version": "1.0",
  "description": "Descripción opcional",
  "entry": "report.html",
  "stylesheet": "style.css"
}
```

El HTML es Jinja2 y recibe los mismos datos del generador monetario: `donor`, `year`, `report_title`, `tagline`, `donor_logo`, `pages`, `closing_title`, `closing_photos`, `logos` y `css`.

## Reglas
- `entry` y `stylesheet` deben apuntar a archivos dentro de la misma carpeta de la plantilla.
- Una carpeta inválida no aparece como opción en la interfaz.
- Los borradores guardan el ID de carpeta de la plantilla seleccionada.
- Puede agregar una nueva plantilla copiando una carpeta válida dentro de `templates/money/` y reiniciando la aplicación.
- Para máxima compatibilidad offline, mantenga imágenes y CSS dentro de la carpeta de la plantilla o use las variables de imagen proporcionadas por el generador.

## Novedades V2.9 (opcionales, retrocompatibles)
- `template.json` puede incluir `kpi_boxes`: `{ "results_photos": {"1": [[x,y,ancho,alto]], "2": [...]}, "results_numbers": {...} }` con las cajas de indicadores en **mm sobre A4**. El generador las entrega a cada página como `p.kpi_boxes` (una caja por indicador) para centrar valor + descripción exactamente dentro de las cajas del fondo.
- El HTML también recibe `p.text_class`, `p.title_class`, `cover_title_class`, `closing_title_class` (tamaños base por longitud) y `donor_logo_max_mm`.
- Autoajuste: cualquier `<div class="fit" data-max="20" data-min="10"><div class="fit-inner">…</div></div>` toma el mayor tamaño (pt) entre `data-min` y `data-max` que cabe en su caja (ver el script al final de `report.html`).

## Novedades V3.0 (opcionales, retrocompatibles)
- `fonts/` dentro de la carpeta de la plantilla: archivos .otf/.ttf/.woff/.woff2 que se incrustan como familia `BAP Avenir` (el peso se deduce del nombre: Roman/Book=400, Medium=500, Heavy=800, etc.). El HTML recibe el CSS resultante en `font_css`.
- `p.manual` (páginas) y `manual` (portada/cierre): tamaños de fuente en pt elegidos por el usuario (`title`, `text`, `kpi_value`, `kpi_label`, `cover_title`, `donor`, `tagline`, `closing`). En el macro `fit(...)` un tamaño manual añade la clase `fixed`, que el script de autoajuste omite.

## Novedades V4.0 (plantilla monetaria)
- `kpi_boxes` ahora cubre `results_photos` (2-5) y `results_numbers` (1-4): rectángulos `[x, y, ancho, alto]` en mm. El generador entrega a la plantilla `p.kpi_items` (valor, descripción, color, icono como data-URI y medidas del círculo).
- `closing_styles` en `template.json` y fondos `closing_a/b/c.png`. Datos: `closing_style` (`closing_a|closing_b|closing_c`) y `closing_photos` (5, 5 o 3).
- `program_logo` (data-URI) y `p.show_logo` para el logo junto al título.
- Los indicadores aceptan una tercera columna: nombre de archivo de la biblioteca de iconos (`icon_dir`) o ruta completa.
