# BAP money-template contract

(Spanish original: `docs/TEMPLATE_CONTRACT.es.md`)

Each money design lives in its own folder inside `templates/money/`.
The app automatically discovers every folder that contains a `template.json`.
You do not need to rebuild the EXE to add, remove or replace a design.

## Minimum structure

```
templates/money/My_Design/
  template.json
  report.html
  style.css
  assets/              # optional
```

Minimal `template.json`:

```json
{
  "name": "Visible name",
  "version": "1.0",
  "description": "Optional description",
  "entry": "report.html",
  "stylesheet": "style.css"
}
```

The HTML is Jinja2 and receives the same data as the money generator: `donor`, `year`, `report_title`, `tagline`, `donor_logo`, `pages`, `closing_title`, `closing_photos`, `logos` and `css`.

## Rules
- `entry` and `stylesheet` must point to files inside the same template folder.
- An invalid folder does not appear as an option in the interface.
- Drafts store the folder ID of the selected template.
- To add a new template, copy a valid folder into `templates/money/` and restart the app.
- For maximum offline compatibility, keep images and CSS inside the template folder or use the image variables provided by the generator.

## V2.9 additions (optional, backward compatible)
- `template.json` may include `kpi_boxes`: `{ "results_photos": {"1": [[x,y,width,height]], "2": [...]}, "results_numbers": {...} }` with the indicator boxes in **mm on A4**. The generator passes them to each page as `p.kpi_boxes` (one box per indicator) so that value + description are centred exactly inside the boxes of the background.
- The HTML also receives `p.text_class`, `p.title_class`, `cover_title_class`, `closing_title_class` (base sizes by text length) and `donor_logo_max_mm`.
- Auto-fit: any `<div class="fit" data-max="20" data-min="10"><div class="fit-inner">…</div></div>` takes the largest size (pt) between `data-min` and `data-max` that fits its box (see the script at the end of `report.html`).

## V3.0 additions (optional, backward compatible)
- `fonts/` inside the template folder: .otf/.ttf/.woff/.woff2 files that are embedded as the `BAP Avenir` family (the weight is inferred from the file name: Roman/Book=400, Medium=500, Heavy=800, etc.). The HTML receives the resulting CSS as `font_css`.
- `p.manual` (pages) and `manual` (cover/closing): font sizes in pt chosen by the user (`title`, `text`, `kpi_value`, `kpi_label`, `cover_title`, `donor`, `tagline`, `closing`). In the `fit(...)` macro, a manual size adds the class `fixed`, which the auto-fit script skips.

## V4.0 additions (money template)
- `kpi_boxes` now covers `results_photos` (2-5) and `results_numbers` (1-4): rectangles `[x, y, width, height]` in mm. The generator gives the template `p.kpi_items` (value, description, colour, icon as a data URI, and circle dimensions).
- `closing_styles` in `template.json` and backgrounds `closing_a/b/c.png`. Data: `closing_style` (`closing_a|closing_b|closing_c`) and `closing_photos` (5, 5 or 3).
- `program_logo` (data URI) and `p.show_logo` for the logo next to the title.
- Indicators accept a third column: a file name from the icon library (`icon_dir`) or a full path.
