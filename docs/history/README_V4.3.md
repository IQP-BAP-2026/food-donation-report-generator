# Generador de Reportes BAP V4.3

## Logos de programas listos para elegir (monetarios)
- Se incluyen 4 logos de programas ya preparados (PNG con fondo transparente, recortados): **Desayunos Felices**, **Fondo Lázaro**, **Nutriendo Vidas** y **Padrinos de Bodega - Corazón Verde**. Los archivos recibidos ya tenían fondo transparente (se verificó el canal alfa); solo se recortaron los márgenes vacíos, sin modificar ningún otro píxel.
- Viven en `assets\program_logos` (dentro del programa) y se copian a la carpeta editable `%LOCALAPPDATA%\BAP Report Generator V2.8\logos_programas` al iniciar (sin sobrescribir ni borrar lo que el usuario ya tenga). Para agregar un logo a todos los equipos: poner el PNG preparado en `assets\program_logos` antes de compilar. Para un solo equipo: copiarlo a `logos_programas`.
- En el editor de página: lista desplegable de logos, botón **Abrir carpeta de logos**, **Usar otro archivo…**, Quitar y Usar en todas las páginas.
- **Aviso:** al pulsar «Usar otro archivo…» aparece un aviso que remite a la sección 5.6 del manual («Cómo preparar un logo del programa») con los botones Abrir el manual / Cancelar / Ya preparé mi logo, continuar.
- Manual actualizado (Versión 4.3): nueva sección 5.6 con requisitos, cómo quitar el fondo, cómo comprobarlo y cómo agregarlo a la lista; tabla de carpetas, solución de problemas y apéndice actualizados.

## Sin cambios
Todo lo demás es igual a V4.2 (alimentos con el código del compañero intacto, logos con recorte automático, indicadores con icono, tres cierres, etc.).
