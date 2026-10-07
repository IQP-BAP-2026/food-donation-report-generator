# BAP Report Generator

Windows desktop app that produces the donor reports of **Banco de Alimentos Panamá (BAP)**:

- **Food-donation traceability reports** (PDF per donor, generated from an Excel master sheet).
- **Monetary-donation impact reports** (cover, content pages with photos and KPI cards, galleries, closing page), built from an HTML/CSS template.

Current version: **App V4.3 / money template 7.4**. Developed for BAP by WPI students (IQP 2026).
The end-user manual (written in Spanish for BAP staff) is `Guia_Generador_Informes_Trazabilidad_BAP_Espanol.pdf`.

## How it works

| Report | Code | PDF engine |
|---|---|---|
| Food | `food_generator.py` (original generator, kept unmodified) + `food/` template and CSS | WeasyPrint |
| Money | `engine/renderers.py` + `templates/money/Plantilla_BAP_Oficial/` (Jinja2 HTML/CSS) | Playwright (Chromium) |

`app.py` is the Tkinter interface for both. The money template is plain HTML/CSS/JSON, so design changes do not need a rebuild
(see `TEMPLATE_CONTRACT.md`). Money reports use Playwright because it renders the approved design exactly.

## Repository layout

```
app.py                  Tkinter app (UI for food and money reports)
engine/                 money report validation + rendering
food/, food_generator.py  food report template, assets and generator
templates/money/        editable money template (backgrounds, report.html, style.css, template.json, fonts/)
assets/                 icons, program logos, photos bundled with the app
BUILD_EXE.bat           builds the Windows executable
START_DEVELOPMENT.bat   runs the app from source
BAP_Integration.spec    PyInstaller recipe
docs/                   Spanish template contract, old per-version notes (docs/history/)
```

## Run from source (Windows)

Requirements: **Windows 10/11 x64**, **Python 3.13 x64**, and **MSYS2 (UCRT64)** installed in `C:\msys64`
(WeasyPrint needs its Pango/GTK libraries; see the WeasyPrint docs for the Windows install).

```
START_DEVELOPMENT.bat
```
This creates `.build-venv`, installs `requirements.txt` and starts `app.py`.
The first run of the money reports needs the Playwright browser: `python -m playwright install chromium`
(`BUILD_EXE.bat` does this for you).

## Build the executable

```
BUILD_EXE.bat
```
Result: `dist\Generador de Reportes BAP\`. **Deliver the whole folder**, not only the `.exe`. The build bundles Chromium,
so the folder is several hundred MB; that is why `dist/`, `_internal/` and `*.exe` are in `.gitignore` and **must not be committed**.
Share the built folder as a ZIP (Drive/OneDrive or a GitHub Release).

## Fonts

Money reports use **Avenir** (commercial font, not included). Put `AvenirLTPro-Roman` and `AvenirLTPro-Heavy`
(`.otf/.ttf/.woff2`) in `templates/money/Plantilla_BAP_Oficial/fonts/`. Without them the report falls back to Arial Bold.
See `fonts/LEEME_FUENTES.txt`.

## Where the app stores data (on the user's PC)

`%LOCALAPPDATA%\BAP Report Generator V2.8\` (drafts, reports, icon and program-logo libraries, logs) and
`Documents\BAP Donor Traceability Reports\` (food reports).

## Known limitations

- "Vista previa" can fail with *Permission denied* if the previous preview PDF is still open in a viewer; close it first.
- The `.exe` is not code-signed, so Windows SmartScreen shows a warning on first run.
- Money-report layout rules (text limits, photo counts) are listed in the manual and enforced by `engine/renderers.py`.

## Tests

`TEST_CHECKLIST.md` lists the manual checks done before each delivery.
