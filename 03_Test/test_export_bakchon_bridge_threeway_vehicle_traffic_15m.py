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


def test_save_and_verify_workbook_reconciles_direction_vehicle_combinations(
    tmp_path: Path,
) -> None:
    target = module.ResolvedTarget(
        node_id="NODE-1",
        intersection_name="박촌교삼거리",
        approach_id="ACSR-1",
        approach_name="박촌교삼거리-서 (동향)",
    )
    rows = [
        module.DirectionVehicleTraffic("01", "좌회전", "10", "승용차", 40),
        module.DirectionVehicleTraffic("02", "직진", "10", "승용차", 100),
        module.DirectionVehicleTraffic("03", "우회전", "20", "버스", 28),
    ]
    output_path = tmp_path / "bakchon.xlsx"

    total = module.save_workbook(target, rows, output_path)
    module.verify_workbook(output_path, target, rows, total)

    assert total == 168


def test_load_direction_vehicle_traffic_uses_15_minute_vehicle_source() -> None:
    executed: list[tuple[str, dict[str, object]]] = []

    class DummyCursor:
        def execute(self, sql: str, params: dict[str, object]) -> None:
            executed.append((sql, params))

        def fetchall(self):
            return [("01", "10", 40), ("02", "10", 100), ("03", "20", 28)]

    target = module.ResolvedTarget("NODE-1", "박촌교삼거리", "ACSR-1", "서 (동향)")
    result = module.load_direction_vehicle_traffic(
        DummyCursor(),
        target,
        {"01": "좌회전", "02": "직진", "03": "우회전"},
        {"10": "승용차", "20": "버스"},
    )

    assert [row.direction_name for row in result] == ["좌회전", "직진", "우회전"]
    assert [row.vehicle_name for row in result] == ["승용차", "승용차", "버스"]
    assert module.TRAFFIC_TABLE in executed[0][0]
    assert "< :end_at" in executed[0][0]
