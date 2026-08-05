import math
import os
import unittest

os.environ.setdefault(
    "BOT_TOKEN",
    "123456789:ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmno",
)

import main


class HelpersAndAlertsTests(unittest.TestCase):
    def test_zero_output_with_nonzero_input_is_critical(self):
        for key in main.NORMS:
            with self.subTest(key=key):
                status, percentage = main.check_norm(0, 1000, key)
                self.assertEqual(status, "crit")
                self.assertEqual(percentage, 0)

    def test_missing_duplicate_is_not_hidden(self):
        status, percentage, difference = main.check_doubles(1000, 0)
        self.assertEqual(status, "crit")
        self.assertEqual(percentage, 100)
        self.assertEqual(difference, 1000)

    def test_both_zero_duplicates_are_not_an_alert(self):
        self.assertEqual(main.check_doubles(0, 0), ("none", 0.0, 0.0))

    def test_user_number_accepts_comma_and_rejects_nonfinite(self):
        self.assertEqual(main.parse_user_number(" 12 345,5 "), 12345.5)
        for value in ("nan", "inf", "-inf"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                main.parse_user_number(value)

    def test_previous_day_label_crosses_year_boundary(self):
        self.assertEqual(main.previous_day_label(2026, 1, 1), "31.12.2025")

    def test_long_messages_are_split_without_data_loss(self):
        text = "\n".join(f"строка {index}: " + "x" * 60 for index in range(150))
        chunks = main.split_message(text, limit=500)

        self.assertGreater(len(chunks), 1)
        self.assertTrue(all(len(chunk) <= 500 for chunk in chunks))
        self.assertEqual("\n".join(chunks), text)

    def test_finite_helpers_do_not_emit_nan(self):
        data = {"kv44": math.nan, "kv44d": 100, "kv46d": math.inf}
        self.assertEqual(main.calc_produced(data), 50)
        self.assertEqual(main.fmt(math.nan), "—")
        self.assertEqual(main.sign(math.inf), "—")

    def test_ai_context_contains_absolute_daily_and_three_day_data(self):
        rows = []
        for day, output in ((30, 900), (31, 1000)):
            rows.append(
                {
                    "report_date": f"2026-07-{day:02d}",
                    "year": 2026,
                    "month": 7,
                    "day_num": day,
                    "kv4": 1000,
                    "kv34": output,
                    "kv3": 1000,
                    "kv33": output,
                }
            )
        rolling_rows = [
            {
                "report_date": "2026-07-29",
                "year": 2026,
                "month": 7,
                "day_num": 29,
                "kv4": 1000,
                "kv34": 1100,
                "kv3": 1000,
                "kv33": 1100,
            },
            *rows,
        ]

        context = main.make_ai_context(rows, rolling_rows)

        self.assertIn("ПОСУТОЧНЫЕ АБСОЛЮТНЫЕ ПОКАЗАНИЯ", context)
        self.assertIn("Дата 30.07.2026", context)
        self.assertIn("Окно 3 сут.; 29–31.07.2026", context)
        self.assertIn("К4=3000.00 т", context)
        self.assertIn("последовательные даты: да", context)


if __name__ == "__main__":
    unittest.main()
