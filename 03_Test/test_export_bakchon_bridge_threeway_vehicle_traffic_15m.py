from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from unittest.mock import patch

import pytest


MODULE_PATH = (
    Path(__file__).resolve().parents[1]
    / "01_Program"
    / "09_else"
    / "export_bakchon_bridge_threeway_vehicle_traffic_15m.py"
)
SPEC = importlib.util.spec_from_file_location("bakchon_vehicle_traffic", MODULE_PATH)
module = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = module
SPEC.loader.exec_module(module)


def test_ensure_select_sql_allows_one_select_statement() -> None:
    module.ensure_select_sql("SELECT NODE_ID FROM M_CRSRD_INF WHERE NODE_ID = :node_id")


@pytest.mark.parametrize(
    "sql",
    [
        "UPDATE M_CRSRD_INF SET CRSRD_NM = 'x'",
        "DELETE FROM M_CRSRD_INF",
        "SELECT 1 FROM DUAL; DROP TABLE M_CRSRD_INF",
        "SELECT 1 FROM DUAL;",
        "COMMIT",
    ],
)
def test_ensure_select_sql_blocks_non_read_only_statements(sql: str) -> None:
    with pytest.raises(RuntimeError):
        module.ensure_select_sql(sql)


def test_connect_read_only_sets_transaction_read_only() -> None:
    executed_sql: list[str] = []

    class DummyCursor:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return False

        def execute(self, sql: str):
            executed_sql.append(sql)

    class DummyConnection:
        def cursor(self):
            return DummyCursor()

    class DummyOracleDb:
        @staticmethod
        def connect(**kwargs):
            assert kwargs == {
                "user": "reader",
                "password": "secret",
                "host": "localhost",
                "port": 1521,
                "service_name": "xe",
            }
            return DummyConnection()

    values = {
        "DB_USER": "reader",
        "DB_PASSWORD": "secret",
        "DB_HOST": "localhost",
        "DB_PORT": "1521",
        "DB_SERVICE": "xe",
    }
    with (
        patch.object(module, "oracledb", DummyOracleDb),
        patch.object(module, "load_dotenv", lambda path: None),
        patch.object(module.os, "getenv", lambda key, default=None: values.get(key, default)),
        patch.dict(module.os.environ, values, clear=True),
    ):
        module.connect_read_only(Path("unused.env"))

    assert executed_sql == ["SET TRANSACTION READ ONLY"]


def test_save_and_verify_workbook_reconciles_all_vehicle_classes(tmp_path: Path) -> None:
    target = module.ResolvedTarget(
        node_id="NODE-1",
        intersection_name="박촌교삼거리",
        approach_id="ACSR-1",
        approach_name="박촌교삼거리-서 (동향)",
    )
    rows = [
        module.VehicleTraffic("10", "승용차", 123),
        module.VehicleTraffic("20", "버스", 45),
    ]
    output_path = tmp_path / "bakchon.xlsx"

    total = module.save_workbook(target, rows, output_path)
    module.verify_workbook(output_path, rows, total)

    assert total == 168
