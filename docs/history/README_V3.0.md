# Generador de Reportes BAP V3.0

Cambios en los reportes **monetarios**. Los reportes de alimentos (código de Cam) no se tocaron en esta versión.

## Novedades
- **Fondos Canva nuevos** (sin línea naranja junto a las fotos ni puntos naranjas en la portada; franjas superiores más cortas). Reemplazan a los anteriores en `templates/money/Plantilla_BAP_Oficial/backgrounds/`.
- **Títulos y textos centrados.** Como las franjas ya terminan a 61 mm y solo tocan x > 193 mm, el título y el texto usan márgenes simétricos (antes se desplazaban a la izquierda para esquivarlas). La altura de cada título sigue la guía de Canva (65 mm en unas páginas, 68 mm en otras).
- **Posiciones de fotos tomadas de la guía de Canva** (`temp-30.pdf`): portadas, texto + fotos, fotos de distintos tamaños, resultados + fotos, galerías de 6/9/12 y cierre 3x3. Las cajas de indicadores no cambiaron (se remidieron y coinciden).
- **Fuente Avenir.** Ver `templates/money/Plantilla_BAP_Oficial/fonts/LEEME_FUENTES.txt`. Avenir es comercial: no viene incluida. Si copia `AvenirLTPro-Roman` y `AvenirLTPro-Heavy` (.otf/.ttf/.woff2) a esa carpeta, se incrustan solas en cada PDF; si tiene Avenir / Avenir Next instalada en Windows, también se usa. Sin ninguna, se usa una alternativa (Nunito Sans / Mulish / Arial).
- **Tamaño de fuente manual (opcional).** Sigue activo el ajuste automático por defecto. En la portada, el cierre y cada página hay un cuadro "Tamaño de fuente (opcional)". Al marcarlo aparece la advertencia "¿Está seguro de que desea cambiar el tamaño de la fuente?" (por defecto: No). Si acepta, esos textos dejan de ajustarse solos y pueden salirse o cortarse: es responsabilidad de quien los cambia. Desmarcar la casilla restaura el ajuste automático. Los tamaños se guardan en el borrador. Rangos permitidos por campo (pt) se validan antes de generar.

## Archivos modificados
`templates/money/Plantilla_BAP_Oficial/{report.html,style.css,template.json,backgrounds/*,fonts/}`, `engine/renderers.py`, `app.py`, `BUILD_EXE.bat` (título).

## Compilar
`BUILD_EXE.bat` como siempre; entregue la carpeta `dist\Generador de Reportes BAP` completa. Las plantillas (`templates\money`, incluida `fonts`) se copian junto al EXE y se pueden cambiar sin recompilar.

`samples/V3.0_money_QA.pdf` es un ejemplo con esta plantilla.
