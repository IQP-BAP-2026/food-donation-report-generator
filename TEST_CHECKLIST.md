# Manual test checklist (before each delivery)

Run from the built folder `dist\Generador de Reportes BAP\`, ideally on a PC without Python.

## General
- [ ] The EXE opens; Settings shows the expected app version and template path.
- [ ] The "Manual de usuario" button opens the PDF.

## Food reports
- [ ] Load the Excel; the summary shows the right sheet (`MASTER-<year>`) and year.
- [ ] Select several donors (Shift for a range, Ctrl for separate ones).
- [ ] Add exactly 7 photos (shared or per donor); 6 or 8 is rejected.
- [ ] Preview opens a PDF; Generate creates one PDF per donor.

## Money reports
- [ ] One page of each type: text + 3 photos, mixed photos, results + photos, results numbers, gallery (6/9/12).
- [ ] Every program logo and several icons; check logos are not too close to the orange stripes.
- [ ] The three closing designs (5, 5 and 3 photos).
- [ ] Cover: donor name above the logo, 4 photos.
- [ ] Save a draft, reopen it, and use a manual font-size override.
- [ ] Preview and Generate; open the PDF at full size and check clipping and colours (brand green #7BC24D).
- [ ] Avenir fonts are embedded (PDF properties) if the font files are present.

## Packaging
- [ ] The delivered folder contains the EXE, `_internal`, `templates`, the manual and the customer README.
- [ ] Unzip it somewhere else and run one report of each kind.
