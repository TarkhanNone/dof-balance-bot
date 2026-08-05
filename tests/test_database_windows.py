import asyncio
import os
import tempfile
import unittest
from datetime import date
from pathlib import Path

os.environ.setdefault(
    "BOT_TOKEN",
    "123456789:ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmno",
)

import main
from balance_logic import rolling_snapshots


def shifts(kv4, kv34):
    return {
        1: {"kv4": kv4 / 2, "kv34": kv34 / 2},
        2: {"kv4": kv4 / 2, "kv34": kv34 / 2},
    }


class DatabaseWindowTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="dof-db-test-")
        self.original_db_path = main.DB_PATH
        main.DB_PATH = str(Path(self.temp_dir.name) / "test.db")
        main.init_db()

    def tearDown(self):
        main.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    @classmethod
    def tearDownClass(cls):
        asyncio.run(main.bot.session.close())

    def test_completed_days_cross_month_boundary_and_current_day_is_excluded(self):
        july = {
            "period": "Отчёт за июль 2026",
            "daily_by_shift": {31: shifts(1000, 990)},
        }
        self.assertEqual(
            main.db_save_daily(
                july,
                user_id=1,
                filename="july.xlsx",
                year=2026,
                month=7,
                today=date(2026, 8, 3),
            ),
            (1, 31, None),
        )

        august = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: shifts(1000, 1000),
                2: shifts(1000, 1010),
                3: {
                    1: {"kv4": 500, "kv34": 450},
                    2: {},
                },
            },
        }
        self.assertEqual(
            main.db_save_daily(
                august,
                user_id=1,
                filename="august.xlsx",
                year=2026,
                month=8,
                today=date(2026, 8, 3),
            ),
            (2, 3, 3),
        )

        latest = main.db_get_latest_completed_days(3)
        self.assertEqual(
            [row["report_date"] for row in latest],
            ["2026-07-31", "2026-08-01", "2026-08-02"],
        )

        snapshot = rolling_snapshots(latest, periods=(3,))[0]
        self.assertTrue(snapshot["consecutive"])
        self.assertEqual(snapshot["label"], "31.07–02.08.2026")
        self.assertAlmostEqual(snapshot["balances"]["b1"], 0.0)

        night = main.db_get_night_shift(2026, 8)
        self.assertIsNotNone(night)
        self.assertEqual(night["day_num"], 3)

    def test_menu_contains_requested_action(self):
        labels = [
            [button.text for button in row] for row in main.main_keyboard().keyboard
        ]
        self.assertIn(["📉 Просмотр проскальзывания"], labels)

    def test_buttons_are_scoped_to_the_last_uploaded_report(self):
        old_report = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: shifts(1000, 900),
                2: shifts(1000, 900),
                3: shifts(1000, 900),
            },
        }
        main.db_save_daily(
            old_report,
            user_id=77,
            filename="old_august.xlsx",
            year=2026,
            month=8,
            today=date(2026, 9, 10),
        )

        fresh_report = {
            "period": "Отчёт за июль 2026",
            "daily_by_shift": {
                30: shifts(2000, 2000),
                31: shifts(2000, 2000),
            },
        }
        main.db_save_daily(
            fresh_report,
            user_id=77,
            filename="fresh_july.xlsx",
            year=2026,
            month=7,
            today=date(2026, 9, 10),
        )

        active = main.db_get_active_report(77)
        self.assertEqual((active["year"], active["month"]), (2026, 7))
        self.assertEqual(active["filename"], "fresh_july.xlsx")
        self.assertEqual(
            [row["report_date"] for row in main.db_get_active_month_data(active)],
            ["2026-07-30", "2026-07-31"],
        )
        self.assertEqual(
            [
                row["report_date"]
                for row in main.db_get_latest_completed_days(
                    3, end_date=main.active_report_end_date(active)
                )
            ],
            ["2026-07-30", "2026-07-31"],
        )

    def test_authorized_users_share_one_consistent_active_report(self):
        main.db_save_daily(
            {
                "period": "Отчёт за июль 2026",
                "daily_by_shift": {31: shifts(1000, 1000)},
            },
            user_id=101,
            filename="first-user.xlsx",
            year=2026,
            month=7,
            today=date(2026, 8, 5),
        )
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(2000, 2000)},
            },
            user_id=202,
            filename="second-user.xlsx",
            year=2026,
            month=8,
            today=date(2026, 8, 5),
        )

        active_for_first = main.db_get_active_report(101)
        active_for_second = main.db_get_active_report(202)
        self.assertEqual(active_for_first, active_for_second)
        self.assertEqual(active_for_first["filename"], "second-user.xlsx")
        self.assertEqual(active_for_first["uploaded_by"], 202)
        self.assertEqual(active_for_first["user_id"], main.ACTIVE_REPORT_SCOPE_ID)

    def test_stale_days_from_older_snapshot_are_not_shown(self):
        first = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: shifts(1000, 900),
                2: shifts(1000, 900),
                3: shifts(1000, 900),
            },
        }
        main.db_save_daily(first, 5, "old.xlsx", 2026, 8, today=date(2026, 9, 1))

        second = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: shifts(2000, 2000),
                2: shifts(2000, 2000),
            },
        }
        main.db_save_daily(second, 5, "fresh.xlsx", 2026, 8, today=date(2026, 9, 1))

        active = main.db_get_active_report(5)
        rows = main.db_get_active_month_data(active)
        self.assertEqual([row["day_num"] for row in rows], [1, 2])
        self.assertEqual(rows[0]["kv4"], 2000)
        self.assertEqual(active["last_completed_day"], 2)

    def test_incomplete_only_upload_does_not_fall_back_to_old_rows(self):
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(1000, 1000)},
            },
            9,
            "old.xlsx",
            2026,
            8,
            today=date(2026, 8, 10),
        )
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {
                    10: {1: {"kv4": 500, "kv34": 500}, 2: {}},
                },
            },
            9,
            "fresh.xlsx",
            2026,
            8,
            today=date(2026, 8, 10),
        )

        active = main.db_get_active_report(9)
        self.assertIsNone(active["last_completed_day"])
        self.assertEqual(main.db_get_active_month_data(active), [])

    def test_future_template_days_are_ignored_but_completed_zero_days_are_kept(
        self,
    ):
        report = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: shifts(1000, 1000),
                2: {1: {"kv4": 0, "kv34": 0}, 2: {"kv4": 0, "kv34": 0}},
                3: shifts(1000, 1000),
                4: {1: {"kv4": 0, "kv34": 0}, 2: {"kv4": 0, "kv34": 0}},
                31: {1: {"kv4": 0, "kv34": 0}, 2: {}},
            },
        }

        saved, max_day, incomplete = main.db_save_daily(
            report,
            1,
            "zeros.xlsx",
            2026,
            8,
            today=date(2026, 8, 10),
        )

        self.assertEqual((saved, max_day, incomplete), (4, 4, None))
        rows = main.db_get_month_data(2026, 8)
        self.assertEqual([row["day_num"] for row in rows], [1, 2, 3, 4])
        self.assertEqual(rows[1]["kv4"], 0)
        self.assertEqual(rows[3]["kv4"], 0)

    def test_nonzero_future_day_is_rejected_and_old_data_is_preserved(self):
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(1000, 1000)},
            },
            11,
            "old.xlsx",
            2026,
            8,
            today=date(2026, 8, 5),
        )

        with self.assertRaises(main.ReportDataError):
            main.db_save_daily(
                {
                    "period": "Отчёт за август 2026",
                    "daily_by_shift": {
                        1: shifts(2000, 2000),
                        6: shifts(100, 100),
                    },
                },
                11,
                "future.xlsx",
                2026,
                8,
                today=date(2026, 8, 5),
            )

        rows = main.db_get_month_data(2026, 8)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["kv4"], 1000)
        self.assertEqual(main.db_get_active_report(11)["filename"], "old.xlsx")

    def test_transaction_rolls_back_complete_month_replacement_on_sql_error(self):
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(1000, 1000)},
            },
            12,
            "old.xlsx",
            2026,
            8,
            today=date(2026, 8, 5),
        )
        with main.db_connection() as connection:
            connection.execute(
                """
                CREATE TRIGGER reject_second_day
                BEFORE INSERT ON daily_data
                WHEN NEW.day_num = 2
                BEGIN
                    SELECT RAISE(ABORT, 'test rollback');
                END
                """
            )

        with self.assertRaises(main.sqlite3.IntegrityError):
            main.db_save_daily(
                {
                    "period": "Отчёт за август 2026",
                    "daily_by_shift": {
                        1: shifts(2000, 2000),
                        2: shifts(2000, 2000),
                    },
                },
                12,
                "broken.xlsx",
                2026,
                8,
                today=date(2026, 8, 5),
            )

        rows = main.db_get_month_data(2026, 8)
        self.assertEqual([(row["day_num"], row["kv4"]) for row in rows], [(1, 1000)])
        self.assertEqual(main.db_get_active_report(12)["filename"], "old.xlsx")

    def test_stock_derived_values_are_refreshed_after_corrected_report(self):
        first = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: {
                    1: {"kv4": 500, "kv44": 100, "kv44d": 100},
                    2: {"kv4": 500, "kv44": 100, "kv44d": 100},
                }
            },
        }
        main.db_save_daily(first, 13, "first.xlsx", 2026, 8, today=date(2026, 8, 5))
        initial_stock = main.db_save_stock(2026, 8, 1, 1000, 1100, 13)
        self.assertEqual(initial_stock["produced"], 200)
        self.assertEqual(initial_stock["nesovpadenie"], 100)

        corrected = {
            "period": "Отчёт за август 2026",
            "daily_by_shift": {
                1: {
                    1: {"kv4": 500, "kv44": 200, "kv44d": 200},
                    2: {"kv4": 500, "kv44": 200, "kv44d": 200},
                }
            },
        }
        main.db_save_daily(
            corrected,
            13,
            "corrected.xlsx",
            2026,
            8,
            today=date(2026, 8, 5),
        )

        active = main.db_get_active_report(13)
        refreshed = main.db_get_active_stock_history(active)
        self.assertEqual(len(refreshed), 1)
        self.assertEqual(refreshed[0]["produced"], 400)
        self.assertEqual(refreshed[0]["marksh_izm"], 100)
        self.assertEqual(refreshed[0]["nesovpadenie"], 300)

    def test_stock_rows_without_a_day_in_fresh_report_are_hidden_not_deleted(self):
        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(1000, 1000), 2: shifts(1000, 1000)},
            },
            14,
            "full.xlsx",
            2026,
            8,
            today=date(2026, 8, 5),
        )
        main.db_save_stock(2026, 8, 2, 1000, 1000, 14)

        main.db_save_daily(
            {
                "period": "Отчёт за август 2026",
                "daily_by_shift": {1: shifts(1000, 1000)},
            },
            14,
            "short.xlsx",
            2026,
            8,
            today=date(2026, 8, 5),
        )

        active = main.db_get_active_report(14)
        self.assertEqual(main.db_get_active_stock_history(active), [])
        self.assertEqual(len(main.db_get_stock_history(2026, 8)), 1)

    def test_init_db_migrates_missing_stock_columns(self):
        with main.db_connection() as connection:
            connection.execute("DROP TABLE stock_data")
            connection.execute(
                "CREATE TABLE stock_data (id INTEGER PRIMARY KEY AUTOINCREMENT)"
            )

        main.init_db()

        with main.db_connection() as connection:
            columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(stock_data)")
            }
        self.assertTrue(
            {
                "year",
                "month",
                "day_num",
                "stock_prev",
                "stock_curr",
                "produced",
                "shipped",
                "nesovpadenie",
            }.issubset(columns)
        )


if __name__ == "__main__":
    unittest.main()
