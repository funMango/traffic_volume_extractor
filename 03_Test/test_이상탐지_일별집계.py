#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import io
import json
import sqlite3
import sys
import threading
import types
import unittest
from contextlib import redirect_stdout
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch


if "tqdm" not in sys.modules:
    tqdm_stub = types.ModuleType("tqdm")

    class _DummyTqdm:
        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            return False

        def set_postfix_str(self, *args, **kwargs):
            return None

        def update(self, *args, **kwargs):
            return None

        @staticmethod
        def write(*args, **kwargs):
            return None

    tqdm_stub.tqdm = _DummyTqdm
    sys.modules["tqdm"] = tqdm_stub


sys.path.insert(0, str(Path(__file__).parent))
import 이상탐지_일별집계 as daily  # noqa: E402


def _mk_slot(d: str, hour: int, aid: int, anm: str, vol: int | None, anomaly: str | None) -> dict:
    return {
        "date": d,
        "hour": hour,
        "approach_id": aid,
        "approach_name": anm,
        "traffic_volume": vol,
        "anomaly_type": anomaly,
    }


class TestHttpHelpers(unittest.TestCase):
    def test_api_get_returns_json(self):
        cm = MagicMock()
        cm.__enter__.return_value = cm
        cm.__exit__.return_value = False
        cm.read.return_value = json.dumps({"ok": True}).encode()
        with patch("이상탐지_일별집계.urllib.request.urlopen", return_value=cm):
            result = daily._api_get("/x")
        self.assertEqual(result, {"ok": True})

    def test_api_post_sends_json_request(self):
        cm = MagicMock()
        cm.__enter__.return_value = cm
        cm.__exit__.return_value = False
        cm.read.return_value = json.dumps({"count": 1}).encode()
        with patch("이상탐지_일별집계.urllib.request.Request") as mock_req, \
             patch("이상탐지_일별집계.urllib.request.urlopen", return_value=cm):
            mock_req.return_value = MagicMock()
            result = daily._api_post("/corrected-traffic", {"a": 1})
        self.assertEqual(result, {"count": 1})
        _, kwargs = mock_req.call_args
        self.assertEqual(kwargs["headers"]["Content-Type"], "application/json")
        self.assertEqual(kwargs["method"], "POST")


class TestParsing(unittest.TestCase):
    def test_parse_short_yymmdd(self):
        self.assertEqual(daily._parse_short("260101"), date(2026, 1, 1))

    def test_parse_short_yyyymmdd(self):
        self.assertEqual(daily._parse_short("20260101"), date(2026, 1, 1))

    def test_parse_cli_year(self):
        ds, de = daily.parse_cli_arg("2026")
        self.assertEqual((ds, de), (date(2026, 1, 1), date(2026, 12, 31)))

    def test_parse_cli_month(self):
        ds, de = daily.parse_cli_arg("202602")
        self.assertEqual((ds, de), (date(2026, 2, 1), date(2026, 2, 28)))

    def test_parse_cli_range(self):
        ds, de = daily.parse_cli_arg("260101~260131")
        self.assertEqual((ds, de), (date(2026, 1, 1), date(2026, 1, 31)))


class TestDbUtils(unittest.TestCase):
    def test_init_db_creates_schema(self):
        db_path = Path(__file__).parent / "_tmp_init_db_test.sqlite"
        if db_path.exists():
            db_path.unlink()
        conn = daily.init_db(db_path)
        try:
            rows = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name IN ('processed_dates','anomaly_daily')"
            ).fetchall()
            self.assertEqual({r[0] for r in rows}, {"processed_dates", "anomaly_daily"})
        finally:
            conn.close()
            if db_path.exists():
                db_path.unlink()

    def test_date_range(self):
        days = daily.date_range(date(2026, 1, 1), date(2026, 1, 3))
        self.assertEqual(days, [date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 3)])

    def test_get_unprocessed_dates(self):
        conn = sqlite3.connect(":memory:")
        conn.executescript(daily._SCHEMA_SQL)
        conn.execute(
            "INSERT INTO processed_dates (date, processed_at) VALUES (?, ?)",
            ("2026-01-02", "2026-01-03T00:00:00"),
        )
        conn.commit()
        all_dates = [date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 3)]
        todo = daily.get_unprocessed_dates(conn, all_dates)
        self.assertEqual(todo, [date(2026, 1, 1), date(2026, 1, 3)])
        conn.close()

    def test_clean_unprocessed_slots(self):
        conn = sqlite3.connect(":memory:")
        conn.executescript(daily._SCHEMA_SQL)
        conn.executemany(
            "INSERT INTO anomaly_daily (date, node_id, node_name, approach_id, approach_name, missing_count) VALUES (?,?,?,?,?,?)",
            [
                ("2026-01-01", 1, "A", None, None, 1),
                ("2026-01-02", 1, "A", None, None, 1),
            ],
        )
        conn.commit()
        daily.clean_unprocessed_slots(conn, [date(2026, 1, 1)])
        left = conn.execute("SELECT date FROM anomaly_daily ORDER BY date").fetchall()
        self.assertEqual(left, [("2026-01-02",)])
        conn.close()

    def test_mark_dates_processed(self):
        conn = sqlite3.connect(":memory:")
        conn.executescript(daily._SCHEMA_SQL)
        daily.mark_dates_processed(conn, [date(2026, 1, 1), date(2026, 1, 2)])
        rows = conn.execute("SELECT date FROM processed_dates ORDER BY date").fetchall()
        self.assertEqual(rows, [("2026-01-01",), ("2026-01-02",)])
        conn.close()


class TestGrouping(unittest.TestCase):
    def test_group_consecutive_dates(self):
        groups = daily.group_consecutive_dates(
            [date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 4), date(2026, 1, 5)]
        )
        self.assertEqual(groups, [(date(2026, 1, 1), date(2026, 1, 2)), (date(2026, 1, 4), date(2026, 1, 5))])

    def test_split_by_year(self):
        segments = daily.split_by_year(date(2025, 12, 30), date(2026, 1, 2))
        self.assertEqual(segments, [(date(2025, 12, 30), date(2025, 12, 31)), (date(2026, 1, 1), date(2026, 1, 2))])


class TestParseApiSlots(unittest.TestCase):
    def test_parse_api_slots_for_daily_filters_only_ab(self):
        slots = [
            _mk_slot("2026-01-01", 8, 11, "북", 0, "A형"),
            _mk_slot("2026-01-01", 9, 12, "남", 10, "B형"),
            _mk_slot("2026-01-01", 10, 13, "동", 20, None),
        ]
        rows = daily._parse_api_slots_for_daily(slots, "테스트교차로")
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["판정"], "A형")
        self.assertEqual(rows[1]["판정"], "B형")
        self.assertEqual(rows[0]["_date"], date(2026, 1, 1))


class TestAggregateAndInsert(unittest.TestCase):
    def test_aggregate_and_insert_daily(self):
        conn = sqlite3.connect(":memory:")
        conn.executescript(daily._SCHEMA_SQL)
        lock = threading.Lock()
        seg_day = date(2026, 1, 1)

        all_slots = [
            _mk_slot("2026-01-01", 8, 1, "북", 0, None),
            _mk_slot("2026-01-01", 8, 2, "남", 0, None),
            _mk_slot("2026-01-01", 9, 1, "북", 0, None),
            _mk_slot("2026-01-01", 9, 2, "남", 0, None),
        ]
        node_results = [
            {"판정": "A형", "_date": seg_day, "_hour": 8, "_acsr_id": 1, "방향": "북"},
            {"판정": "A형", "_date": seg_day, "_hour": 8, "_acsr_id": 2, "방향": "남"},
            {"판정": "A형", "_date": seg_day, "_hour": 9, "_acsr_id": 1, "방향": "북"},
        ]

        rows = daily.aggregate_and_insert_daily(
            conn,
            lock,
            node_id=100,
            node_name="테스트",
            node_results=node_results,
            all_slots=all_slots,
            seg_ds=seg_day,
            seg_de=seg_day,
        )
        self.assertEqual(len(rows), 2)
        db_rows = conn.execute(
            "SELECT date, node_id, approach_id, missing_count FROM anomaly_daily ORDER BY approach_id IS NULL DESC, approach_id"
        ).fetchall()
        self.assertEqual(len(db_rows), 2)
        conn.close()


class TestInsertDailyRows(unittest.TestCase):
    def test_insert_daily_rows(self):
        conn = sqlite3.connect(":memory:")
        conn.executescript(daily._SCHEMA_SQL)
        lock = threading.Lock()
        rows = [
            {
                "date": "2026-01-01",
                "node_id": 1,
                "node_name": "N",
                "approach_id": None,
                "approach_name": None,
                "missing_count": 2,
            },
            {
                "date": "2026-01-01",
                "node_id": 1,
                "node_name": "N",
                "approach_id": 11,
                "approach_name": "북",
                "missing_count": 1,
            },
        ]
        inserted = daily.insert_daily_rows(conn, lock, rows)
        self.assertEqual(inserted, rows)
        db_rows = conn.execute("SELECT COUNT(*) FROM anomaly_daily").fetchone()[0]
        self.assertEqual(db_rows, 2)
        conn.close()


class TestProcessNode(unittest.TestCase):
    def test_process_node_success(self):
        with patch("이상탐지_일별집계._api_post", return_value={"rows": [{
                 "date": "2026-01-01",
                 "node_id": 1,
                 "node_name": "N",
                 "approach_id": None,
                 "approach_name": None,
                 "missing_count": 1,
             }]}):
            inserted, node_name, summary, failed, pending_rows = daily._process_node(
                1, "N", [(date(2026, 1, 1), date(2026, 1, 1))]
            )
        self.assertEqual(inserted, 1)
        self.assertEqual(node_name, "N")
        self.assertFalse(failed)
        self.assertEqual(summary[date(2026, 1, 1)], 1)
        self.assertEqual(len(pending_rows), 1)

    def test_process_node_api_failure_marks_failed(self):
        with patch("이상탐지_일별집계._api_post", side_effect=RuntimeError("boom")), \
             patch("이상탐지_일별집계.tqdm.write"):
            inserted, _, summary, failed, pending_rows = daily._process_node(
                1, "N", [(date(2026, 1, 1), date(2026, 1, 1))]
            )
        self.assertEqual(inserted, 0)
        self.assertEqual(dict(summary), {})
        self.assertTrue(failed)
        self.assertEqual(pending_rows, [])


class TestOutputAndMain(unittest.TestCase):
    def test_print_summary_outputs_total(self):
        buf = io.StringIO()
        with redirect_stdout(buf):
            daily.print_summary({"노드A": {date(2026, 1, 1): 2}}, total_inserted=2)
        out = buf.getvalue()
        self.assertIn("처리 결과 요약", out)
        self.assertIn("완료: 2건 저장", out)

    def test_main_early_exit_when_no_todo_dates(self):
        fake_conn = MagicMock()
        fake_conn.close = MagicMock()
        with patch.object(sys, "argv", ["prog", "2026"]), \
             patch("이상탐지_일별집계.parse_cli_arg", return_value=(date(2026, 1, 1), date(2026, 1, 1))), \
             patch("이상탐지_일별집계.init_db", return_value=fake_conn), \
             patch("이상탐지_일별집계.date_range", return_value=[date(2026, 1, 1)]), \
             patch("이상탐지_일별집계.get_unprocessed_dates", return_value=[]), \
             patch("이상탐지_일별집계._api_get") as mock_api_get:
            daily.main()
        fake_conn.close.assert_called_once()
        mock_api_get.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
