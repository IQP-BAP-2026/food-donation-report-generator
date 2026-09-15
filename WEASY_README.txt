WEASYPRINT VERSION
==================

This version of food_generator.py uses WeasyPrint instead of Playwright/Chromium
for PDF generation.

FILES
-----
food_generator.py              Report generator using WeasyPrint
food-traceability-template.html Current report template
food-traceability.css          Current report CSS
MASTER-SHEET.xlsx              Current workbook
assets/                         Report assets
requirements-weasy.txt          Python dependencies

INSTALL
-------
1. Create/activate your Python environment.
2. Install dependencies:
   py -m pip install -r requirements-weasy.txt

RUN
---
   py food_generator.py

The generator keeps the existing donor/data logic. Only the PDF rendering layer
has been changed from Playwright/Chromium to WeasyPrint.

TESTED
------
The included version was successfully tested against the NESTLE PANAMA donor and
produced a 2-page A4 PDF with the current HTML/CSS.
