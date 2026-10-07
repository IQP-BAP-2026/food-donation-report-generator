# Generador de Reportes BAP V4.0

## Alimentos
- Código actualizado del compañero (`food_generator.py`, plantilla y CSS): copiado **sin modificar** (verificado por SHA-256). El resto de su carpeta no cambió. No se incluye su `test_food_metrics.py` (solo son pruebas).
- La plantilla y el CSS se actualizan solos en cada equipo cuando el programa trae una versión nueva (el Excel y las fotos del usuario no se tocan). Antes se copiaban solo la primera vez, así que un equipo con V3.2 habría seguido usando la plantilla vieja con el generador nuevo.
- Interfaz sin cambios: Excel, varios donantes, exactamente 7 fotos, vista previa, PDF y borradores.

## Monetarios
- **Fondos y posiciones nuevos** (`bg-10-1.zip` y `graphic-10-1.pdf`): verde #7BC24D, portadas con título y año centrados, logo BAP arriba a la izquierda, fotos y textos medidos de la guía.
- **Indicadores nuevos con icono:** tarjetas de color (verde, naranja, azul) con círculo de icono, valor y descripción. Páginas con fotos: de **2 a 5** indicadores. Páginas de números grandes: de **1 a 4** (tarjetas más grandes).
- **Icono por indicador:** en el editor de página cada indicador tiene una lista de iconos y un botón "Subir…". Los iconos se guardan en una carpeta (botón "Abrir carpeta de iconos"): `%LOCALAPPDATA%\BAP Report Generator V2.8\iconos`. Vienen incluidos los 6 iconos del reporte de alimentos. Formatos: PNG, JPG, WEBP, SVG. Opción por página para mostrarlos en blanco dentro del círculo (recomendado). Un indicador sin icono se muestra sin círculo.
- **Logo del programa** (opcional, pestaña Portada): aparece junto al título de las páginas de contenido; se puede desactivar por página. No se usa en galerías ni en el cierre.
- **Tres diseños de cierre** (pestaña Cierre): Ondas y corazones (5 fotos), Marco verde y franjas (5 fotos), Hexágonos (3 fotos). La cantidad de fotos es exacta y se valida. El año del informe se muestra en el cierre.
- Autoajuste: los valores muy largos ya no se cortan (la fuente puede bajar hasta 5 pt).

## Correcciones
- Se restauró la creación de las carpetas de datos y del archivo de registro al iniciar (se había perdido en V3.1 y V3.2; en equipos que ya tenían las carpetas de versiones anteriores no se notaba, pero en un equipo nuevo fallaban el guardado de borradores y la carga de Excel).

## Cambios que afectan borradores y páginas existentes
- Páginas de resultados con 1 indicador y fotos ya no existen en el diseño (ahora son 2 a 5). Los borradores con esa página piden agregar un indicador.
- El cierre ya no usa 9 fotos: según el diseño son 5 o 3. Los borradores antiguos piden volver a seleccionar las fotos de cierre.
- Los indicadores antiguos (valor y descripción) siguen funcionando; solo les falta el icono.

## Fuentes
Ver `templates/money/Plantilla_BAP_Oficial/fonts/LEEME_FUENTES.txt` (Avenir y, opcionalmente, una serif para el cierre de hexágonos).
