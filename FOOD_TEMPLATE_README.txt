FOOD TRACEABILITY TEMPLATE
===========================

Files:
- food-traceability-template.html
- food-traceability.css
- assets/food/*

This template is a reusable, data-driven HTML recreation of the food traceability report style supplied by the user.

Expected optional Jinja variables:
- company_name
- month_name
- year
- kilos_donated
- kilos_share
- kilos_usable
- usable_pct
- kilos_waste
- waste_pct
- meals_served
- orgs_helped
- beneficiaries
- monthly_data: {kilos, kilos_total, organizations, organizations_total, population, population_total}
- month_labels
- national_delivered_kg
- regional_table: [{region, kg, ben, pct}]
- map_points: [{x, y, value}]
- distribution_photo (optional data URL or file URL)

Missing values simply hide the associated module. No placeholder warning text is generated.

The supplied example is a one-page tall-format report (600 x 1500 pt), so the CSS intentionally uses a custom 211.667mm x 529.167mm page size instead of A4.
