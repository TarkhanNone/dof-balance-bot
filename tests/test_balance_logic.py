import unittest
from datetime import date

from balance_logic import (
    bal1,
    bal2,
    balance_status,
    balc1,
    balc2,
    infer_report_period,
    is_consecutive_period,
    rolling_snapshots,
    sum_period,
)


def row(report_date, kv4, output):
    return {
        "report_date": report_date,
        "year": int(report_date[:4]),
        "month": int(report_date[5:7]),
        "day_num": int(report_date[8:10]),
        "kv4": kv4,
        "kv34": output,
    }


class SlidingBalanceTests(unittest.TestCase):
    def test_all_four_formulas_use_absolute_tonnage(self):
        data = {
            "kv4": 1000,
            "kv102": 100,
            "kv34": 850,
            "kv24p": 100,
            "kv24hv": 50,
            "kv28a1": 100,
            "kv14": 700,
            "kv32": 500,
            "kv3": 1000,
            "kv101": 100,
            "kv33": 850,
            "kv28a2": 0,
            "kv15": 700,
            "kv31": 500,
        }
        self.assertAlmostEqual(bal1(data), 0.0)
        self.assertAlmostEqual(balc1(data), -5.0)
        self.assertAlmostEqual(bal2(data), -5.0)
        self.assertAlmostEqual(balc2(data), -10.0)

    def test_tonnage_is_summed_before_percentage(self):
        rows = [
            row("2026-07-31", 1000, 700),  # -30%
            row("2026-08-01", 800, 1100),  # +37.5%
        ]
        total = sum_period(rows)
        self.assertAlmostEqual(bal1(total), 0.0)

    def test_three_day_window_crosses_month_boundary(self):
        rows = [
            row("2026-07-31", 1000, 990),
            row("2026-08-01", 1000, 1000),
            row("2026-08-02", 1000, 1010),
        ]
        result = rolling_snapshots(rows)[-1]
        self.assertEqual(result["days"], 3)
        self.assertEqual(result["label"], "31.07–02.08.2026")
        self.assertTrue(result["consecutive"])
        self.assertAlmostEqual(result["balances"]["b1"], 0.0)

    def test_gap_is_detected(self):
        rows = [
            row("2026-07-31", 1000, 1000),
            row("2026-08-02", 1000, 1000),
            row("2026-08-03", 1000, 1000),
        ]
        self.assertFalse(is_consecutive_period(rows))

    def test_stable_critical_deviation(self):
        rows = [
            row("2026-08-01", 1000, 1060),
            row("2026-08-02", 1000, 1060),
            row("2026-08-03", 1000, 1060),
        ]
        value = rolling_snapshots(rows)[-1]["balances"]["b1"]
        self.assertAlmostEqual(value, 6.0)
        self.assertEqual(balance_status(value), "crit")

    def test_report_period_is_detected(self):
        self.assertEqual(
            infer_report_period("Отчёт за июль 2026", date(2026, 8, 3)), (2026, 7, True)
        )
        self.assertEqual(
            infer_report_period("01.07.2026–31.07.2026", date(2026, 8, 3)),
            (2026, 7, True),
        )
        self.assertEqual(infer_report_period("", date(2026, 8, 3)), (2026, 8, False))

    def test_month_name_without_year_uses_nearest_past_year(self):
        self.assertEqual(
            infer_report_period("Отчёт за июль", date(2026, 8, 3)),
            (2026, 7, True),
        )
        self.assertEqual(
            infer_report_period("Отчёт за декабрь", date(2026, 1, 3)),
            (2025, 12, True),
        )

    def test_non_finite_values_do_not_poison_calculations(self):
        self.assertIsNone(bal1({"kv4": float("nan"), "kv34": 100}))
        self.assertEqual(balance_status(float("nan")), "none")


if __name__ == "__main__":
    unittest.main()
