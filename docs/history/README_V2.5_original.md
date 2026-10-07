# BAP Report Generator V2.5

## What changed
- The green/orange 16-page BAP design is used only by monetary reports.
- Food reports continue to use Cam's food generator/template.
- Monetary page choices now match the approved design: text + 3 photos, mixed-orientation photos, results + photos, numerical results, and 6/9/12-photo galleries.
- Cover requires exactly 4 photos; closing requires exactly 9.
- Result pages require 1-4 indicators; indicator values/descriptions have enforced text limits.
- Every monetary page has page-specific minimum/maximum text and photo requirements. Generation is blocked with a Spanish error when a rule is not met.
- Pages can be edited, duplicated, deleted, and moved up/down.
- Food workflow supports multiple donors and creates one PDF per donor.
- Food reports require exactly 7 photos, either shared across selected donors or assigned separately per donor.
- MASTER-SHEET.xlsx is no longer a PyInstaller build dependency; BAP selects the current workbook at runtime.
- User manual remains accessible from the app.

## Monetary limits
- Report title: 3-70 characters
- Donor: 2-80
- Tagline: 5-100
- Cover: exactly 4 photos
- Text + 3 photos: 40-500 characters, exactly 3 photos
- Mixed photos: 40-450 characters, exactly 3 photos (vertical, square, horizontal guidance)
- Results + photos: 30-350 characters, exactly 3 photos, 1-4 indicators
- Numerical results: 20-300 characters, 0 photos, 1-4 indicators
- Gallery: exactly 6, 9, or 12 photos
- Indicator value: 1-25 characters
- Indicator description: 2-60 characters
- Closing statement: 5-120 characters
- Closing: exactly 9 photos

## Build
Run BUILD_EXE.bat on Windows. Deliver the entire `dist/Generador de Reportes BAP` folder, not only the EXE.
