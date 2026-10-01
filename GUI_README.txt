Donor Traceability Report GUI - Unicode fix

This GUI is a separate launcher and does not modify food_generator.py, the HTML, or the CSS.

Important fix:
The launcher forces UTF-8 for the child Python process because the existing generator prints Unicode characters such as ✅ and ❌. This prevents Windows cp1252 UnicodeEncodeError crashes.

Files:
- report_gui.py
- build_report_gui.bat

For testing:
  py -3 report_gui.py

Build EXE:
  build_report_gui.bat

The GUI expects the existing report project structure, including:
- food_generator.py
- food-traceability-template.html
- MASTER-SHEET.xlsx
- assets/
- output/
