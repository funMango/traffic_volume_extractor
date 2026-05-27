from __future__ import annotations

import importlib.util
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


MODULE_PATH = Path(__file__).with_name("export_major_intersections_raw_aadt.py")


def _load_module():
    spec = importlib.util.spec_from_file_location("export_major_intersections_raw_aadt", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class ExportMajorIntersectionsRawAadtTest(unittest.TestCase):
    def test_ensure_select_sql_allows_select_only(self):
        module = _load_module()

        module._ensure_select_sql("SELECT NODE_ID FROM M_CRSRD_INF WHERE NODE_ID = :node_id")

    def test_ensure_select_sql_blocks_write_and_transaction_sql(self):
        module = _load_module()
        blocked_sql = [
            "UPDATE T SET A = 1",
            "DELETE FROM T",
            "MERGE INTO T USING S ON (T.ID = S.ID) WHEN MATCHED THEN UPDATE SET A = 1",
            "DROP TABLE T",
            "TRUNCATE TABLE T",
            "ALTER TABLE T ADD A NUMBER",
            "CREATE TABLE T (A NUMBER)",
            "COMMIT",
            "SELECT 1 FROM DUAL; DROP TABLE T",
        ]

        for sql in blocked_sql:
            with self.subTest(sql=sql):
                with self.assertRaises(RuntimeError):
                    module._ensure_select_sql(sql)

    def test_connect_db_sets_transaction_read_only(self):
        module = _load_module()
        executed_sql = []

        class DummyCursor:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def execute(self, sql):
                executed_sql.append(sql)

        class DummyConnection:
            def cursor(self):
                return DummyCursor()

        class DummyOracleDb:
            @staticmethod
            def connect(**kwargs):
                self.assertEqual(kwargs["user"], "reader")
                self.assertEqual(kwargs["password"], "secret")
                self.assertEqual(kwargs["host"], "localhost")
                self.assertEqual(kwargs["port"], 1521)
                self.assertEqual(kwargs["service_name"], "xe")
                return DummyConnection()

        env = {
            "DB_USER": "reader",
            "DB_PASSWORD": "secret",
            "DB_HOST": "localhost",
            "DB_PORT": "1521",
            "DB_SERVICE": "xe",
        }

        with (
            patch.object(module, "oracledb", DummyOracleDb),
            patch.object(module, "load_dotenv", lambda path: None),
            patch.object(module.os, "getenv", lambda key, default=None: env.get(key, default)),
        ):
            module._connect_db(Path("dummy.env"))

        self.assertEqual(executed_sql, ["SET TRANSACTION READ ONLY"])

    def test_save_excel_uses_requested_sheet_headers(self):
        module = _load_module()
        output_path = MODULE_PATH.with_name("_tmp_test_major_intersections_raw_aadt.xlsx")
        if output_path.exists():
            output_path.unlink()

        module._save_excel(
            [{"교차로이름": "문예사거리", "2024": 100, "2023": 90}],
            [
                {
                    "교차로ID": "BC001N0122",
                    "교차로이름": "문예사거리",
                    "연도": 2024,
                    "수집일 수": 366,
                    "제외일 수": 0,
                    "일교통량 합계": 36600,
                    "평균 산정값": 100,
                    "DB table": module.TRAFFIC_TABLE,
                    "평균 기준": module.AVERAGE_CRITERIA,
                }
            ],
            output_path,
        )

        with zipfile.ZipFile(output_path) as workbook:
            traffic_xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
            info_xml = workbook.read("xl/worksheets/sheet2.xml").decode("utf-8")

        self.assertIn("<t>교차로이름</t>", traffic_xml)
        self.assertIn("<t>2024</t>", traffic_xml)
        self.assertIn("<t>2023</t>", traffic_xml)
        self.assertIn("<t>수집일 수</t>", info_xml)
        self.assertIn("<t>제외일 수</t>", info_xml)


if __name__ == "__main__":
    unittest.main()
