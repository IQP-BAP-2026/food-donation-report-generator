# Generador de Reportes BAP V3.1

Une la aplicación de alimentos del compañero (fuente de verdad) con la interfaz del Generador BAP. Los reportes monetarios son los de V3.0.

## Alimentos: qué es del compañero y qué es de la interfaz
**Sin modificar (idénticos byte a byte al ZIP del compañero, verificado con SHA-256):**
`food_generator.py`, `food/food-traceability-template.html`, `food/food-traceability.css`, `food/MASTER-SHEET.xlsx` y todo `food/assets/`.
Por eso el PDF de alimentos sale con el mismo motor (WeasyPrint), la misma plantilla, el mismo CSS y las mismas reglas de datos.

**Lo que cambió (solo la interfaz y su conexión):**
- La pantalla de alimentos usa el estilo de este programa (barra superior, pestañas, botones) con las mismas funciones de la interfaz del compañero:
  Actualizar Excel, Actualizar lista, búsqueda de donante escribiendo, Fotos (4 fotos de cierre), GENERAR REPORTE, Abrir último PDF y Abrir carpeta de reportes.
- Se genera igual que en su aplicación: en un proceso aparte (`--generate`), con UTF-8 forzado y límite de 120 segundos.
- Reemplaza al motor de alimentos anterior (`engine/food_engine.py`, `engine/food_adapter.py`, plantilla y assets viejos), que era la causa del PDF distinto. Se eliminó la selección de varios donantes y de 7 fotos porque no forman parte de la interfaz del compañero.
- Datos editables de alimentos: `%LOCALAPPDATA%\BAP Report Generator V2.8\alimentos_datos` (se copian del programa la primera vez y no se sobrescriben).
  PDFs: `Documentos\BAP Donor Traceability Reports` (igual que su aplicación).
- Adaptación necesaria solo para modo desarrollo (`python app.py`): el proceso aparte se lanza como `python app.py --generate <donante>`. En el EXE es idéntico a su aplicación.

## Compilar
`BUILD_EXE.bat` como siempre (ya incluye la carpeta `food` y las DLL de WeasyPrint). Entregue la carpeta `dist\Generador de Reportes BAP` completa.
Nota: los borradores antiguos de alimentos ya no se usan (abrirlos lleva a la pantalla de alimentos).

`samples/` conserva los PDFs de QA anteriores.
