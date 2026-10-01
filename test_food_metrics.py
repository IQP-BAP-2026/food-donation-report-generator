import contextlib
import io
import unittest

import pandas as pd

from food_generator import build_report_data


class FoodDonationImpactTests(unittest.TestCase):
    def report(self, food_value=100, food_header="KG DONADOS   ALIMENTOS", **changes):
        row = {
            "DONANTE": "Test donor",
            "KG DONADOS EN EL MES": 1000,
            "KG APROVECHABES EN EL MES": 900,
            "KG MERMA EN EL MES": 100,
            "FÓRMULA CONVERSIÓN A PLATOS DE COMIDA": 2700,
            "Kilos_Ene": 1000,
        }
        if food_header is not None:
            row[food_header] = food_value
        row.update(changes)
        with contextlib.redirect_stdout(io.StringIO()):
            return build_report_data(pd.DataFrame([row]), "Test donor", 2026)

    def test_food_only_amount_overrides_old_meal_formula(self):
        report = self.report()
        self.assertEqual(report["meals_served"], "300")
        self.assertEqual(report["carbon_saved"], "275")
        self.assertTrue(report["plates_has_data"])
        self.assertTrue(report["carbon_saved_has_data"])
        self.assertEqual(report["kilos_donated"], "1,000")
        self.assertEqual(report["kilos_usable"], "900")

    def test_non_food_donations_cannot_increase_food_impact(self):
        baseline = self.report()
        changed = self.report(**{
            "KG DONADOS EN EL MES": 10000,
            "KG APROVECHABES EN EL MES": 9900,
            "FÓRMULA CONVERSIÓN A PLATOS DE COMIDA": 29700,
        })
        for key in ("meals_served", "carbon_saved"):
            self.assertEqual(baseline[key], changed[key])

    def test_header_spacing_and_case(self):
        for header in ("KG DONADOS ALIMENTOS", " KG DONADOS   ALIMENTOS ", "kg donados alimentos"):
            with self.subTest(header=header):
                self.assertEqual(self.report(food_header=header)["meals_served"], "300")

    def test_no_food_never_falls_back_to_overall_donations(self):
        for value in (0, None, float("nan"), "", -20):
            with self.subTest(value=value):
                report = self.report(food_value=value)
                self.assertFalse(report["plates_has_data"])
                self.assertFalse(report["carbon_saved_has_data"])
        report = self.report(food_header=None)
        self.assertFalse(report["plates_has_data"])
        self.assertFalse(report["carbon_saved_has_data"])

    def test_food_impact_does_not_require_usable_kilos(self):
        report = self.report(**{"KG APROVECHABES EN EL MES": None})
        self.assertEqual(report["meals_served"], "300")
        self.assertEqual(report["carbon_saved"], "275")

    def test_fractional_food_kilos_preserve_whole_meal_conversion(self):
        self.assertEqual(self.report(food_value=10.6)["meals_served"], "31")


if __name__ == "__main__":
    unittest.main()
