"""Pruebas unitarias del cálculo mensual del SLA."""

from datetime import datetime, timedelta
import importlib.util
from pathlib import Path
import unittest


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "consolidar_inventario.py"
SPEC = importlib.util.spec_from_file_location("consolidar_inventario_sla", SCRIPT_PATH)
inventory = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(inventory)


class SlaTests(unittest.TestCase):
    def test_risk_eight_uses_30_days_and_due_date_is_compliant(self):
        self.assertEqual(
            inventory.calculate_sla(8, "2026-01-01", "2026-01-31"),
            ("SI", "", "", "2026-01-01T00:00:00"),
        )

    def test_risk_seven_uses_90_days(self):
        self.assertEqual(
            inventory.calculate_sla(7, "2026-01-01", "2026-03-31"),
            ("SI", "", "", "2026-03-01T00:00:00"),
        )

    def test_late_closed_finding_is_imputed_from_first_breach_month_to_close_month(self):
        self.assertEqual(
            inventory.calculate_sla(10, "2026-01-01", "2026-02-02"),
            ("NO", "2026-02-01T00:00:00", "2026-02-01T00:00:00", ""),
        )

    def test_open_overdue_finding_uses_sentinel_end_month(self):
        self.assertEqual(
            inventory.calculate_sla(10, "2020-01-01", ""),
            ("NO", "2020-02-01T00:00:00", "9999-12-31T00:00:00", ""),
        )

    def test_open_finding_inside_sla_complies_so_far(self):
        active = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d")
        self.assertEqual(inventory.calculate_sla(9, active, ""), ("En plazo", "", "", ""))

    def test_missing_or_invalid_activation_date_is_not_applicable(self):
        for active in ("", "N/A", "not-a-date"):
            with self.subTest(active=active):
                self.assertEqual(inventory.calculate_sla(9, active, ""), ("N/A", "", "", ""))


if __name__ == "__main__":
    unittest.main()
