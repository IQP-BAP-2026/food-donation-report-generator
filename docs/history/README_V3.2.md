# Generador de Reportes BAP V3.2

## Alimentos: interfaz restaurada, motor del compañero intacto
La pantalla de alimentos vuelve a tener lo mismo que en V3.0:
- **Cargar Excel** (hoja MASTER-{AÑO} más reciente; mismo criterio que `food_generator.py`).
- **Donantes**: selección múltiple; se genera un PDF por donante.
- **Fotos: exactamente 7**, ni más ni menos. Se valida al seleccionarlas (con menos o más de 7 no se aceptan) y otra vez antes de generar. Modo "las mismas 7 fotos para todos" o "7 fotos distintas por donante".
- **Vista previa** (primer donante seleccionado) y **GENERAR PDF** (elige carpeta; abre la carpeta al terminar).
- **Guardar borrador** / Abrir borrador (compatible con borradores de V3.0).
- Abrir carpeta de reportes.

### Las 7 fotos
Fotos 1-4 -> `closing_photo1-4` (franja inferior de la página 2). Fotos 5-7 -> `custom_photo1-3` (zona central de la página 2 para donantes sin mapa regional).
Los JPG se copian sin modificar; PNG/WEBP se convierten a JPG. `food_generator.py` ignora fotos de 10 KB o menos, por eso cada foto debe pesar más de 10 KB (se avisa).

### Sin cambios en el código del compañero
`food_generator.py`, `food/food-traceability-template.html`, `food/food-traceability.css`, `food/MASTER-SHEET.xlsx` y `food/assets/` son idénticos byte a byte a su ZIP (verificado con SHA-256). Cada PDF se genera en un proceso aparte con `--generate`, UTF-8 y límite de 120 s, como en su aplicación. Los nombres de archivo son los de su generador (`Food_Traceability_<Donante>_<Mes>_<Año>.pdf`).

### Novedades de la interfaz
- La generación corre en segundo plano con ventana de progreso ("Generando 2 de 5: ...") y no congela la aplicación.
- Si un donante falla, los demás se generan y al final se muestra cuáles fallaron y por qué (detalle completo en Registros).
- Datos editables de alimentos: `%LOCALAPPDATA%\BAP Report Generator V2.8\alimentos_datos`.

## Monetarios
Sin cambios respecto a V3.0.
