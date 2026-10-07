# Generador de Reportes BAP V2.9

Refinamiento de apariencia del PDF monetario. Los reportes de alimentos (diseño de Cam) no se tocaron.

## Qué cambió en el PDF monetario
- **Indicadores centrados de verdad.** Las cajas naranjas/blancas vienen dibujadas en los fondos de Canva. Sus coordenadas se midieron píxel a píxel y ahora están en `template.json` (`kpi_boxes`), así que cada valor + descripción queda centrado en su caja (antes la fila inferior y la caja central quedaban pegadas arriba).
- **Texto autoajustable.** Título, texto, valores, descripciones, nombre del donante, lema y mensaje de cierre usan el tamaño más grande que cabe en su zona (entre un mínimo y un máximo). Los textos cortos ya no se ven diminutos y los largos siguen sin salirse.
- **Mejor uso del espacio.** Zonas de texto más altas, fotos más grandes con esquinas redondeadas y recorte sesgado hacia la parte superior (caras).
- **Contraste.** Texto oscuro (#1F2933) sobre las cajas naranjas, con descripciones más grandes.
- **Página de texto + fotos de distintos tamaños:** cuadrícula ordenada (1 alta + 2 apiladas alineadas) en lugar de tres fotos con bordes desiguales.
- **Portada 1:** los logos aliados van sobre una cápsula blanca centrada en la línea naranja (antes la línea atravesaba los logos). Fotos de portada más altas. Portada 2 reacomodada.
- **Galería:** el título ahora tiene tamaño propio (antes heredaba el tamaño por defecto del navegador y se veía diminuto).
- **Logo del donante** no se amplía más allá de ~150 dpi (evita logos borrosos); logo BAP de cada página más grande.
- **Fotos optimizadas:** se corrige la rotación EXIF y se reducen fotos muy grandes (máx. 2000 px) antes de incrustarlas: PDFs más livianos y generación más rápida. Solo se cargan los fondos que el reporte realmente usa.
- Los saltos de línea del texto se respetan.

## Archivos modificados
`templates/money/Plantilla_BAP_Oficial/{report.html,style.css,template.json}`, `engine/renderers.py`, `app.py` (solo versión), `BAP_Integration.spec` (Pillow), `BUILD_EXE.bat` (título).

## Importante al compilar
Compile con `BUILD_EXE.bat` como siempre y entregue/pruebe la **carpeta completa** `dist\Generador de Reportes BAP`. El autoajuste de texto requiere Chromium (Playwright), que el build ya incluye. Si Chromium fallara y se usara WeasyPrint como respaldo, el PDF sale con los tamaños base por longitud (sin autoajuste).
Si ya tiene una carpeta `dist` anterior, no copie solo el EXE: las plantillas (`templates\money`) también cambiaron.

`samples/V2.9_money_QA.pdf` es un ejemplo generado con esta plantilla.
