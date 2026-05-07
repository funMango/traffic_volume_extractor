#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
api_server.py 단위 테스트
대상:
  - JobRequest 유효성 검증 (Pydantic 모델)
  - resolve_hours() 시간대 해석
  - _serialize_slot() 슬롯 직렬화
  - _serialize_job() 작업 직렬화
  - GET /intersections 엔드포인트
  - GET /jobs/{job_id} 엔드포인트
  - POST /jobs 엔드포인트 (유효성 검증 + 생성)
  - _run_job() 백그라운드 작업 (성공·실패)
  - analyse_node() in 이상탐지.py
"""

import os
import sys
import types
import unittest
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

import numpy as np

# ════════════════════════════════════════════════════════════════════════════════
# 외부 의존성 스텁 (oracledb / plotly / openpyxl / dotenv)
# — test_이상탐지.py 와 동일한 패턴 사용
# ════════════════════════════════════════════════════════════════════════════════
sys.path.insert(0, os.path.dirname(__file__))


def _make_module(name: str) -> types.ModuleType:
    m = types.ModuleType(name)
    sys.modules[name] = m
    return m


if "oracledb" not in sys.modules:
    _make_module("oracledb")

if "dotenv" not in sys.modules:
    _dotenv = _make_module("dotenv")
    _dotenv.load_dotenv = lambda *a, **kw: None  # type: ignore

for _pmod in ("plotly", "plotly.graph_objects"):
    if _pmod not in sys.modules:
        _m = _make_module(_pmod)
        _m.Figure = type("Figure", (), {
            "__init__": lambda s, **kw: None,
            "add_trace": lambda s, *a, **kw: None,
            "update_layout": lambda s, **kw: None,
            "write_html": lambda s, *a, **kw: None,
        })
        _m.Scatter = type("Scatter", (), {"__init__": lambda s, **kw: None})

_openpyxl = _make_module("openpyxl")
_openpyxl.Workbook = type("Workbook", (), {
    "__init__": lambda s: None,
    "active": property(lambda s: type("WS", (), {
        "title": "",
        "append": lambda s, r: None,
        "column_dimensions": defaultdict(lambda: type("CD", (), {"width": 0})()),
    })()),
    "create_sheet": lambda s, t: None,
    "save": lambda s, p: None,
    "remove": lambda s, ws: None,
})

_styles = _make_module("openpyxl.styles")
for _cls in ("Font", "PatternFill", "Alignment", "Border", "Side"):
    setattr(_styles, _cls, type(_cls, (), {"__init__": lambda s, **kw: None}))

_utils = _make_module("openpyxl.utils")
_utils.get_column_letter = lambda i: chr(64 + i)

# ── 실제 모듈 import ──────────────────────────────────────────────────────────
import 이상탐지 as ad  # noqa: E402
import api_server      # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

# ── 공통 클라이언트 ──────────────────────────────────────────────────────────
_client = TestClient(api_server.app, raise_server_exceptions=True)


# ════════════════════════════════════════════════════════════════════════════════
# 목(Mock) 데이터
# ════════════════════════════════════════════════════════════════════════════════

_MOCK_INTERSECTIONS = [(1001, "계남교차로"), (1002, "중앙교차로")]

_MOCK_NODE_RESULTS = [
    {
        "날짜":     "2026.03.15",
        "시간":     "08:00",
        "교차로":   "계남교차로",
        "방향":     "북",
        "교통량":   0,
        "판정":     "A형",
        "_acsr_id": 2001,
        "_date":    date(2026, 3, 15),
        "_hour":    8,
        "_cell":    (2001, 1, 8, 3),  # (acsr_id, day_type, hour, month)
        "보정값":   320,
        "보정방법": "median",
        "신뢰도":   "OK",
    }
]

_MOCK_BASELINES = {
    (2001, 1, 8, 3): {
        "zero_rate": 0.05,
        "expansion_stage": 0,
        "stats": {"median": 320.0, "q1": 280.0, "q3": 360.0, "n_clean": 15},
    }
}

_BASE_JOB: dict = {
    "job_id":      "test-job-id",
    "status":      "pending",
    "request": {
        "node_ids":   [1001],
        "date_start": "2026-03-01",
        "date_end":   "2026-03-15",
        "hours":      list(range(24)),
    },
    "created_at":  datetime(2026, 3, 1, 10, 0, 0),
    "started_at":  None,
    "finished_at": None,
    "progress":    None,
    "result":      None,
    "error":       None,
}


def _fresh_job() -> dict:
    """job 딕셔너리의 독립적인 복사본 반환"""
    import copy
    return copy.deepcopy(_BASE_JOB)


# ════════════════════════════════════════════════════════════════════════════════
# 1. JobRequest 유효성 검증
# ════════════════════════════════════════════════════════════════════════════════

class TestJobRequestValidation(unittest.TestCase):

    def _req(self, **kwargs) -> dict:
        base = {
            "node_ids":   [1001],
            "date_start": "2026-03-01",
            "date_end":   "2026-03-31",
        }
        base.update(kwargs)
        return base

    # ── 정상 케이스 ──────────────────────────────────────────────
    def test_valid_minimal(self):
        """필수 필드만 포함한 최소 요청 — 성공"""
        req = api_server.JobRequest(**self._req())
        self.assertEqual(req.node_ids, [1001])

    def test_valid_with_hours(self):
        """hours 직접 지정 — 성공"""
        req = api_server.JobRequest(**self._req(hours=[7, 8, 17]))
        self.assertEqual(sorted(req.hours), [7, 8, 17])

    def test_valid_with_hours_preset_peak(self):
        """hours_preset=peak — 성공"""
        req = api_server.JobRequest(**self._req(hours_preset="peak"))
        self.assertEqual(req.hours_preset, "peak")

    def test_valid_with_hours_preset_all(self):
        """hours_preset=all — 성공"""
        req = api_server.JobRequest(**self._req(hours_preset="all"))
        self.assertEqual(req.hours_preset, "all")

    def test_valid_same_date(self):
        """date_start == date_end — 성공 (당일 1일치 분석)"""
        req = api_server.JobRequest(**self._req(date_start="2026-03-15", date_end="2026-03-15"))
        self.assertEqual(req.date_start, req.date_end)

    def test_valid_multiple_nodes(self):
        """node_ids 여러 개 — 성공"""
        req = api_server.JobRequest(**self._req(node_ids=[1001, 1002, 1003]))
        self.assertEqual(len(req.node_ids), 3)

    # ── 경계값 케이스 ─────────────────────────────────────────────
    def test_hours_boundary_zero_and_23(self):
        """hours 최솟값(0)과 최댓값(23) — 허용"""
        req = api_server.JobRequest(**self._req(hours=[0, 23]))
        self.assertIn(0, req.hours)
        self.assertIn(23, req.hours)

    def test_hours_deduplicated(self):
        """hours 중복 값 — resolve_hours에서 중복 제거"""
        req = api_server.JobRequest(**self._req(hours=[7, 7, 8, 8]))
        self.assertEqual(sorted(req.resolve_hours()), [7, 8])

    # ── 오류 케이스 ──────────────────────────────────────────────
    def test_invalid_empty_node_ids(self):
        """node_ids 빈 리스트 — 400/422"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(node_ids=[]))

    def test_invalid_date_order(self):
        """date_start > date_end — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(date_start="2026-04-01", date_end="2026-03-01"))

    def test_invalid_date_format(self):
        """날짜 포맷 오류 (YYYYMMDD) — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(date_start="20260301"))

    def test_invalid_hours_out_of_range(self):
        """hours에 24 포함 — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(hours=[7, 24]))

    def test_invalid_negative_hours(self):
        """hours에 음수 포함 — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(hours=[-1, 8]))

    def test_invalid_hours_and_preset_both(self):
        """hours와 hours_preset 동시 지정 — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(hours=[7, 8], hours_preset="peak"))

    def test_invalid_hours_preset_unknown(self):
        """알 수 없는 hours_preset — 오류"""
        with self.assertRaises(Exception):
            api_server.JobRequest(**self._req(hours_preset="custom"))


# ════════════════════════════════════════════════════════════════════════════════
# 2. resolve_hours()
# ════════════════════════════════════════════════════════════════════════════════

class TestResolveHours(unittest.TestCase):

    def _req(self, **kwargs):
        base = {"node_ids": [1], "date_start": "2026-03-01", "date_end": "2026-03-31"}
        base.update(kwargs)
        return api_server.JobRequest(**base)

    def test_no_hours_no_preset_returns_all(self):
        """hours·preset 모두 없으면 0–23 반환"""
        self.assertEqual(self._req().resolve_hours(), list(range(24)))

    def test_preset_all_returns_all(self):
        """preset=all → 0–23"""
        self.assertEqual(self._req(hours_preset="all").resolve_hours(), list(range(24)))

    def test_preset_peak_returns_peak_hours(self):
        """preset=peak → 7,8,12,13,17,18"""
        self.assertEqual(self._req(hours_preset="peak").resolve_hours(), [7, 8, 12, 13, 17, 18])

    def test_explicit_hours_returned_sorted(self):
        """직접 지정 hours → 정렬된 값"""
        self.assertEqual(self._req(hours=[17, 7, 8]).resolve_hours(), [7, 8, 17])

    def test_explicit_hours_deduplicated(self):
        """중복 hours → 중복 제거"""
        self.assertEqual(self._req(hours=[8, 8, 7]).resolve_hours(), [7, 8])


# ════════════════════════════════════════════════════════════════════════════════
# 3. _serialize_slot()
# ════════════════════════════════════════════════════════════════════════════════

class TestSerializeSlot(unittest.TestCase):

    def _make_result(self, **overrides) -> dict:
        base = _MOCK_NODE_RESULTS[0].copy()
        base.update(overrides)
        return base

    def test_date_format_conversion(self):
        """'2026.03.15' → '2026-03-15' 변환 확인"""
        slot = api_server._serialize_slot(self._make_result(), node_id=1001, baselines=_MOCK_BASELINES)
        self.assertEqual(slot["date"], "2026-03-15")

    def test_hour_is_int_not_string(self):
        """hour 필드가 '08:00' 문자열이 아닌 정수 8이어야 함"""
        slot = api_server._serialize_slot(self._make_result(), node_id=1001, baselines=_MOCK_BASELINES)
        self.assertIsInstance(slot["hour"], int)
        self.assertEqual(slot["hour"], 8)

    def test_node_id_injected_correctly(self):
        """node_id는 결과 dict가 아닌 파라미터로 주입"""
        slot = api_server._serialize_slot(self._make_result(), node_id=9999, baselines=_MOCK_BASELINES)
        self.assertEqual(slot["node_id"], 9999)

    def test_traffic_volume_zero_for_missing_record(self):
        """A형(결측) 슬롯의 traffic_volume은 0"""
        slot = api_server._serialize_slot(self._make_result(교통량=0), node_id=1001, baselines=_MOCK_BASELINES)
        self.assertEqual(slot["traffic_volume"], 0)

    def test_anomaly_type_mapped(self):
        """판정 → anomaly_type 매핑"""
        for 판정 in ("A형", "B형", "정상"):
            slot = api_server._serialize_slot(self._make_result(판정=판정), node_id=1001, baselines=_MOCK_BASELINES)
            self.assertEqual(slot["anomaly_type"], 판정)

    def test_baseline_fields_populated(self):
        """baseline 필드가 올바르게 채워지는지 확인"""
        slot = api_server._serialize_slot(self._make_result(), node_id=1001, baselines=_MOCK_BASELINES)
        bl = slot["baseline"]
        self.assertAlmostEqual(bl["median"], 320.0)
        self.assertAlmostEqual(bl["q1"], 280.0)
        self.assertAlmostEqual(bl["q3"], 360.0)
        self.assertEqual(bl["n_clean"], 15)
        self.assertAlmostEqual(bl["zero_rate"], 0.05)
        self.assertEqual(bl["expansion_stage"], 0)

    def test_baseline_none_when_cell_missing(self):
        """baselines에 해당 셀이 없으면 모든 baseline 필드 None"""
        slot = api_server._serialize_slot(self._make_result(), node_id=1001, baselines={})
        bl = slot["baseline"]
        self.assertIsNone(bl["median"])
        self.assertIsNone(bl["q1"])
        self.assertIsNone(bl["zero_rate"])

    def test_corrected_value_none_for_정상(self):
        """정상 슬롯은 보정값이 None"""
        r = self._make_result(판정="정상", 보정값=None, 보정방법=None, 신뢰도=None)
        slot = api_server._serialize_slot(r, node_id=1001, baselines=_MOCK_BASELINES)
        self.assertIsNone(slot["corrected_value"])
        self.assertIsNone(slot["correction_method"])
        self.assertIsNone(slot["confidence"])

    def test_correction_method_and_confidence(self):
        """보정방법·신뢰도가 올바르게 직렬화"""
        slot = api_server._serialize_slot(self._make_result(), node_id=1001, baselines=_MOCK_BASELINES)
        self.assertEqual(slot["correction_method"], "median")
        self.assertEqual(slot["confidence"], "OK")


# ════════════════════════════════════════════════════════════════════════════════
# 4. _serialize_job()
# ════════════════════════════════════════════════════════════════════════════════

class TestSerializeJob(unittest.TestCase):

    def test_pending_job_serialized(self):
        """pending 상태 job 직렬화 — started_at·finished_at이 None"""
        job = _fresh_job()
        result = api_server._serialize_job(job)
        self.assertEqual(result["status"], "pending")
        self.assertIsNone(result["started_at"])
        self.assertIsNone(result["finished_at"])

    def test_datetime_converted_to_isoformat(self):
        """created_at이 ISO 8601 문자열로 변환"""
        job = _fresh_job()
        result = api_server._serialize_job(job)
        self.assertIsInstance(result["created_at"], str)
        self.assertIn("T", result["created_at"])  # ISO 형식 확인

    def test_done_job_serialized(self):
        """done 상태 job — result 포함"""
        job = _fresh_job()
        job["status"] = "done"
        job["started_at"]  = datetime(2026, 3, 1, 10, 0, 1)
        job["finished_at"] = datetime(2026, 3, 1, 10, 2, 0)
        job["result"] = {"summary": {}, "slots": []}
        result = api_server._serialize_job(job)
        self.assertEqual(result["status"], "done")
        self.assertIsNotNone(result["started_at"])
        self.assertIsNotNone(result["finished_at"])
        self.assertIsNotNone(result["result"])

    def test_failed_job_serialized(self):
        """failed 상태 job — error 포함"""
        job = _fresh_job()
        job["status"] = "failed"
        job["error"]  = "DB 연결 실패"
        result = api_server._serialize_job(job)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["error"], "DB 연결 실패")

    def test_all_keys_present(self):
        """직렬화 결과에 필수 키가 모두 포함"""
        required = {"job_id", "status", "created_at", "started_at",
                    "finished_at", "progress", "result", "error"}
        result = api_server._serialize_job(_fresh_job())
        self.assertEqual(set(result.keys()), required)


# ════════════════════════════════════════════════════════════════════════════════
# 5. GET /intersections
# ════════════════════════════════════════════════════════════════════════════════

class TestIntersectionsEndpoint(unittest.TestCase):

    def setUp(self):
        api_server._intersections_cache = None

    def test_returns_list_from_db(self):
        """DB에서 교차로 목록을 정상 반환"""
        with patch("api_server.ad.connect_db") as mock_conn, \
             patch("api_server.ad.load_intersections", return_value=_MOCK_INTERSECTIONS):
            mock_conn.return_value.__enter__ = MagicMock(return_value=MagicMock())
            mock_conn.return_value.close = MagicMock()
            resp = _client.get("/intersections")

        self.assertEqual(resp.status_code, 200)
        body = resp.json()
        self.assertEqual(body["count"], 2)
        self.assertEqual(body["items"][0]["node_id"], 1001)
        self.assertEqual(body["items"][0]["name"], "계남교차로")

    def test_cache_used_on_second_call(self):
        """두 번째 호출은 DB를 조회하지 않고 캐시 사용"""
        api_server._intersections_cache = [{"node_id": 9999, "name": "캐시교차로"}]
        with patch("api_server.ad.connect_db") as mock_conn:
            resp = _client.get("/intersections")
            mock_conn.assert_not_called()

        self.assertEqual(resp.json()["items"][0]["node_id"], 9999)

    def test_refresh_true_bypasses_cache(self):
        """refresh=true면 캐시를 무시하고 DB 재조회"""
        api_server._intersections_cache = [{"node_id": 9999, "name": "캐시교차로"}]
        with patch("api_server.ad.connect_db") as mock_conn, \
             patch("api_server.ad.load_intersections", return_value=_MOCK_INTERSECTIONS):
            mock_conn.return_value.close = MagicMock()
            resp = _client.get("/intersections?refresh=true")

        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json()["items"][0]["node_id"], 1001)

    def test_empty_intersection_list(self):
        """교차로가 없는 경우 — count=0"""
        with patch("api_server.ad.connect_db"), \
             patch("api_server.ad.load_intersections", return_value=[]):
            resp = _client.get("/intersections")

        self.assertEqual(resp.json()["count"], 0)
        self.assertEqual(resp.json()["items"], [])


# ════════════════════════════════════════════════════════════════════════════════
# 6. GET /jobs/{job_id}
# ════════════════════════════════════════════════════════════════════════════════

class TestGetJobEndpoint(unittest.TestCase):

    def setUp(self):
        api_server.job_store.clear()

    def test_existing_job_returned(self):
        """존재하는 job_id — 200 OK"""
        job = _fresh_job()
        api_server.job_store["test-job-id"] = job
        resp = _client.get("/jobs/test-job-id")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json()["job_id"], "test-job-id")
        self.assertEqual(resp.json()["status"], "pending")

    def test_nonexistent_job_returns_404(self):
        """존재하지 않는 job_id — 404"""
        resp = _client.get("/jobs/nonexistent-id")
        self.assertEqual(resp.status_code, 404)

    def test_running_job_has_progress(self):
        """running 상태 job은 progress 필드 포함"""
        job = _fresh_job()
        job["status"]   = "running"
        job["progress"] = {"total": 3, "done": 1, "current_node": "계남교차로"}
        api_server.job_store["test-job-id"] = job
        resp = _client.get("/jobs/test-job-id")
        self.assertEqual(resp.json()["progress"]["done"], 1)

    def test_done_job_has_result(self):
        """done 상태 job은 result 포함"""
        job = _fresh_job()
        job["status"] = "done"
        job["result"] = {"summary": {"total_anomalies": 5}, "slots": []}
        api_server.job_store["test-job-id"] = job
        resp = _client.get("/jobs/test-job-id")
        body = resp.json()
        self.assertEqual(body["status"], "done")
        self.assertEqual(body["result"]["summary"]["total_anomalies"], 5)


# ════════════════════════════════════════════════════════════════════════════════
# 7. POST /jobs
# ════════════════════════════════════════════════════════════════════════════════

class TestCreateJobEndpoint(unittest.TestCase):

    def setUp(self):
        api_server.job_store.clear()

    def _valid_body(self, **kwargs) -> dict:
        base = {
            "node_ids":   [1001],
            "date_start": "2026-03-01",
            "date_end":   "2026-03-31",
        }
        base.update(kwargs)
        return base

    def test_valid_request_returns_201(self):
        """유효한 요청 — 201 Created, job_id 포함"""
        with patch("api_server._run_job"):  # 백그라운드 실행 방지
            resp = _client.post("/jobs", json=self._valid_body())
        self.assertEqual(resp.status_code, 201)
        body = resp.json()
        self.assertIn("job_id", body)
        self.assertEqual(body["status"], "pending")
        self.assertIn("created_at", body)

    def test_job_added_to_store(self):
        """생성 후 job_store에 추가되었는지 확인"""
        with patch("api_server._run_job"):
            resp = _client.post("/jobs", json=self._valid_body())
        job_id = resp.json()["job_id"]
        self.assertIn(job_id, api_server.job_store)

    def test_hours_resolved_before_store(self):
        """hours_preset=peak → job_store에 실제 시간 리스트로 저장"""
        with patch("api_server._run_job"):
            resp = _client.post("/jobs", json=self._valid_body(hours_preset="peak"))
        job_id = resp.json()["job_id"]
        self.assertEqual(api_server.job_store[job_id]["request"]["hours"], [7, 8, 12, 13, 17, 18])

    def test_hours_none_resolves_to_all_24(self):
        """hours·preset 없으면 0–23 전체 저장"""
        with patch("api_server._run_job"):
            resp = _client.post("/jobs", json=self._valid_body())
        job_id = resp.json()["job_id"]
        self.assertEqual(api_server.job_store[job_id]["request"]["hours"], list(range(24)))

    def test_invalid_empty_node_ids_returns_422(self):
        """node_ids 빈 리스트 — 422 Unprocessable Entity"""
        resp = _client.post("/jobs", json=self._valid_body(node_ids=[]))
        self.assertEqual(resp.status_code, 422)

    def test_invalid_date_order_returns_422(self):
        """date_start > date_end — 422"""
        resp = _client.post("/jobs", json=self._valid_body(
            date_start="2026-04-01", date_end="2026-03-01"
        ))
        self.assertEqual(resp.status_code, 422)

    def test_invalid_hours_out_of_range_returns_422(self):
        """hours에 24 포함 — 422"""
        resp = _client.post("/jobs", json=self._valid_body(hours=[7, 24]))
        self.assertEqual(resp.status_code, 422)

    def test_hours_and_preset_both_returns_422(self):
        """hours + hours_preset 동시 지정 — 422"""
        resp = _client.post("/jobs", json=self._valid_body(hours=[8], hours_preset="peak"))
        self.assertEqual(resp.status_code, 422)


# ════════════════════════════════════════════════════════════════════════════════
# 8. _run_job() — 백그라운드 작업
# ════════════════════════════════════════════════════════════════════════════════

class TestRunJob(unittest.TestCase):

    def setUp(self):
        api_server.job_store.clear()
        api_server._holiday_dates = set()  # 스타트업 완료 상태 시뮬레이션

    def _register_job(self, node_ids=(1001,)) -> str:
        job_id = "run-test-job"
        api_server.job_store[job_id] = {
            "job_id":      job_id,
            "status":      "pending",
            "request": {
                "node_ids":   list(node_ids),
                "date_start": "2026-03-01",
                "date_end":   "2026-03-15",
                "hours":      [7, 8],
            },
            "created_at":  datetime.now(),
            "started_at":  None,
            "finished_at": None,
            "progress":    None,
            "result":      None,
            "error":       None,
        }
        return job_id

    def test_successful_run_sets_done(self):
        """정상 실행 — status가 done으로 변경"""
        job_id = self._register_job()
        mock_conn = MagicMock()

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024, 2025]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[(2026, 2)]), \
             patch("api_server.ad.load_intersections", return_value=[(1001, "계남교차로")]), \
             patch("api_server.ad.analyse_node", return_value=(_MOCK_NODE_RESULTS, _MOCK_BASELINES, {})):
            api_server._run_job(job_id)

        job = api_server.job_store[job_id]
        self.assertEqual(job["status"], "done")
        self.assertIsNotNone(job["finished_at"])
        self.assertIsNone(job["error"])

    def test_result_summary_counts(self):
        """완료 후 result.summary의 count_A 등 집계 확인"""
        job_id = self._register_job()
        mock_conn = MagicMock()

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024, 2025]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[(1001, "계남교차로")]), \
             patch("api_server.ad.analyse_node", return_value=(_MOCK_NODE_RESULTS, _MOCK_BASELINES, {})):
            api_server._run_job(job_id)

        summary = api_server.job_store[job_id]["result"]["summary"]
        self.assertEqual(summary["count_A"], 1)
        self.assertEqual(summary["count_B"], 0)
        self.assertEqual(summary["total_anomalies"], 1)

    def test_db_failure_sets_failed(self):
        """DB 연결 실패 — status가 failed로 변경, error 메시지 포함"""
        job_id = self._register_job()

        with patch("api_server.ad.connect_db", side_effect=RuntimeError("ORA-12541: TNS:no listener")):
            api_server._run_job(job_id)

        job = api_server.job_store[job_id]
        self.assertEqual(job["status"], "failed")
        self.assertIn("ORA-12541", job["error"])
        self.assertIsNotNone(job["finished_at"])

    def test_analyse_node_failure_sets_failed(self):
        """analyse_node 실패 — status가 failed로 변경"""
        job_id = self._register_job()
        mock_conn = MagicMock()

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[(1001, "계남교차로")]), \
             patch("api_server.ad.analyse_node", side_effect=ValueError("베이스라인 계산 실패")):
            api_server._run_job(job_id)

        self.assertEqual(api_server.job_store[job_id]["status"], "failed")

    def test_progress_updated_per_node(self):
        """노드 2개 처리 시 progress.done이 순서대로 증가"""
        job_id = self._register_job(node_ids=[1001, 1002])
        mock_conn = MagicMock()
        done_snapshots = []

        def fake_analyse(conn, node_id, *args, **kwargs):
            done_snapshots.append(api_server.job_store[job_id]["progress"]["done"])
            return [], {}, {}

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections",
                   return_value=[(1001, "계남교차로"), (1002, "중앙교차로")]), \
             patch("api_server.ad.analyse_node", side_effect=fake_analyse):
            api_server._run_job(job_id)

        # 첫 번째 노드 처리 전 done=0, 후 done=1
        self.assertEqual(done_snapshots[0], 0)   # 두 번째 호출 시 첫 번째는 이미 done=1
        final = api_server.job_store[job_id]["progress"]["done"]
        self.assertEqual(final, 2)

    def test_unknown_node_id_uses_str_fallback(self):
        """load_intersections에 없는 node_id — str(node_id) 를 이름으로 사용"""
        job_id = self._register_job(node_ids=[9999])
        mock_conn = MagicMock()

        captured_names = []

        def fake_analyse(conn, node_id, node_name, *args, **kwargs):
            captured_names.append(node_name)
            return [], {}, {}

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[]), \
             patch("api_server.ad.analyse_node", side_effect=fake_analyse):
            api_server._run_job(job_id)

        self.assertEqual(captured_names[0], "9999")

    def test_conn_closed_on_success(self):
        """정상 완료 후 DB 연결이 닫히는지 확인"""
        job_id = self._register_job()
        mock_conn = MagicMock()

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", return_value=[2024]), \
             patch("api_server.ad.select_fallback_year", return_value=2025), \
             patch("api_server.ad.get_adjacent_month_periods", return_value=[]), \
             patch("api_server.ad.load_intersections", return_value=[]), \
             patch("api_server.ad.analyse_node", return_value=([], {}, {})):
            api_server._run_job(job_id)

        mock_conn.close.assert_called_once()

    def test_conn_closed_on_failure(self):
        """실패 시에도 DB 연결이 닫히는지 확인 (finally 블록)"""
        job_id = self._register_job()
        mock_conn = MagicMock()

        with patch("api_server.ad.connect_db", return_value=mock_conn), \
             patch("api_server.ad.select_baseline_years", side_effect=RuntimeError("오류")):
            api_server._run_job(job_id)

        mock_conn.close.assert_called_once()

    def test_holiday_dates_none_sets_failed(self):
        """_holiday_dates=None(서버 초기화 미완료) 상태에서 _run_job → failed"""
        job_id = self._register_job()
        api_server._holiday_dates = None  # 초기화 미완료 상태 시뮬레이션

        api_server._run_job(job_id)

        job = api_server.job_store[job_id]
        self.assertEqual(job["status"], "failed")
        self.assertIsNotNone(job["error"])
        self.assertIn("초기화", job["error"])

    def test_holiday_dates_none_does_not_call_connect_db(self):
        """_holiday_dates=None이면 connect_db를 호출하지 않고 즉시 실패"""
        job_id = self._register_job()
        api_server._holiday_dates = None

        with patch("api_server.ad.connect_db") as mock_conn:
            api_server._run_job(job_id)
            mock_conn.assert_not_called()

    def test_status_set_to_running_before_execution(self):
        """_run_job 호출 직후 status가 running으로 변경되는지 확인"""
        job_id = self._register_job()
        captured_status = []

        def fake_connect():
            captured_status.append(api_server.job_store[job_id]["status"])
            raise RuntimeError("중단")  # 이후 실행 방지

        with patch("api_server.ad.connect_db", side_effect=fake_connect):
            api_server._run_job(job_id)

        self.assertEqual(captured_status[0], "running")


# ════════════════════════════════════════════════════════════════════════════════
# 9. analyse_node() — 이상탐지.py
# ════════════════════════════════════════════════════════════════════════════════

class TestAnalyseNode(unittest.TestCase):
    """이상탐지.ad.analyse_node() 단위 테스트. 내부 DB 함수를 patch로 대체."""

    _CONN   = MagicMock()
    _APPROACHES = [(2001, "북"), (2002, "남")]

    def _call(self, **overrides):
        defaults = dict(
            conn           = self._CONN,
            node_id        = 1001,
            node_name      = "계남교차로",
            date_start     = date(2026, 3, 15),
            date_end       = date(2026, 3, 15),
            hours          = [8],
            baseline_years = [2024, 2025],
            fallback_year  = 2025,
            adj_periods    = [],
            holiday_dates  = set(),
        )
        defaults.update(overrides)
        return ad.analyse_node(**defaults)

    def test_returns_tuple_of_three(self):
        """정상 실행 — (results, baselines, target_data) 튜플 반환"""
        with patch("이상탐지.load_approaches", return_value=self._APPROACHES), \
             patch("이상탐지.load_baseline_data", return_value=[]), \
             patch("이상탐지.load_adjacent_month_data", return_value=[]), \
             patch("이상탐지.build_baselines", return_value=({}, {})), \
             patch("이상탐지.load_target_data", return_value={}), \
             patch("이상탐지.detect_anomalies", return_value=[]), \
             patch("이상탐지.apply_corrections"):
            result = self._call()

        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

    def test_empty_approaches_returns_empty_tuple(self):
        """접근로 정보 없으면 ([], {}, {}) 반환"""
        with patch("이상탐지.load_approaches", return_value=[]):
            result = self._call()

        self.assertEqual(result, ([], {}, {}))

    def test_apply_corrections_called(self):
        """apply_corrections가 호출되는지 확인"""
        mock_results = [{"판정": "A형"}]
        with patch("이상탐지.load_approaches", return_value=self._APPROACHES), \
             patch("이상탐지.load_baseline_data", return_value=[]), \
             patch("이상탐지.load_adjacent_month_data", return_value=[]), \
             patch("이상탐지.build_baselines", return_value=({}, {})), \
             patch("이상탐지.load_target_data", return_value={}), \
             patch("이상탐지.detect_anomalies", return_value=mock_results), \
             patch("이상탐지.apply_corrections") as mock_apply:
            self._call()

        mock_apply.assert_called_once()

    def test_fallback_baseline_data_loaded(self):
        """fallback_year 데이터도 별도 조회되는지 확인 (load_baseline_data 2회 호출)"""
        with patch("이상탐지.load_approaches", return_value=self._APPROACHES), \
             patch("이상탐지.load_baseline_data", return_value=[]) as mock_load_bl, \
             patch("이상탐지.load_adjacent_month_data", return_value=[]), \
             patch("이상탐지.build_baselines", return_value=({}, {})), \
             patch("이상탐지.load_target_data", return_value={}), \
             patch("이상탐지.detect_anomalies", return_value=[]), \
             patch("이상탐지.apply_corrections"):
            self._call(baseline_years=[2024, 2025], fallback_year=2025)

        # baseline_years([2024,2025])와 fallback_year([2025]) 두 번 호출
        self.assertEqual(mock_load_bl.call_count, 2)

    def test_results_from_detect_anomalies_returned(self):
        """detect_anomalies 결과가 첫 번째 반환값으로 전달"""
        sentinel = [{"판정": "B형", "_cell": None}]
        with patch("이상탐지.load_approaches", return_value=self._APPROACHES), \
             patch("이상탐지.load_baseline_data", return_value=[]), \
             patch("이상탐지.load_adjacent_month_data", return_value=[]), \
             patch("이상탐지.build_baselines", return_value=({}, {})), \
             patch("이상탐지.load_target_data", return_value={}), \
             patch("이상탐지.detect_anomalies", return_value=sentinel), \
             patch("이상탐지.apply_corrections"):
            results, _, _ = self._call()

        self.assertIs(results, sentinel)


# ════════════════════════════════════════════════════════════════════════════════
# 신규: _fetch_corrected_traffic 및 POST /corrected-traffic 테스트
# ════════════════════════════════════════════════════════════════════════════════

from datetime import timedelta  # noqa: E402 (이미 상단에 있을 수 있음)

_DATE_START = date(2026, 3, 22)
_DATE_END   = date(2026, 3, 22)
_HOURS      = [6, 7, 8, 9, 10, 11, 12, 13]
_APPROACHES = [(2001, "북향"), (2002, "남향")]

# 보정 전 원본값 목
_TARGET_DATA = {
    (2001, _DATE_START, 6):  300,
    (2001, _DATE_START, 7):  310,
    # 8시: A형 결측 — target_data에 없음
    (2001, _DATE_START, 9):  290,
    (2002, _DATE_START, 6):  250,
    (2002, _DATE_START, 7):  260,
    (2002, _DATE_START, 8):  None,  # NULL 값 (슬롯 제외 대상)
}

# analyse_node 반환 목 (A형 결측 보정)
_NODE_RESULTS_WITH_CORRECTION = [
    {
        "날짜": "2026.03.22", "시간": "08:00",
        "교차로": "송내사거리", "방향": "북향",
        "교통량": 0, "판정": "A형",
        "_acsr_id": 2001, "_date": _DATE_START, "_hour": 8,
        "_cell": (2001, 0, 8, 3),
        "보정값": 305, "보정방법": "선형보간", "신뢰도": "OK",
    }
]

# stage2 "정상" 판정 — 보정값 None
_NODE_RESULTS_STAGE2 = [
    {
        "날짜": "2026.03.22", "시간": "08:00",
        "교차로": "송내사거리", "방향": "북향",
        "교통량": 0, "판정": "정상",
        "_acsr_id": 2001, "_date": _DATE_START, "_hour": 8,
        "_cell": (2001, 0, 8, 3),
        "보정값": None, "보정방법": None, "신뢰도": None,
    }
]


def _patch_analyse(node_results, target_data=None):
    """analyse_node 목 컨텍스트 패치"""
    td = target_data if target_data is not None else _TARGET_DATA
    return patch.object(ad, "analyse_node",
                        return_value=(node_results, {}, td))


def _patch_approaches(approaches=None):
    ap = approaches if approaches is not None else _APPROACHES
    return patch.object(ad, "load_approaches", return_value=ap)


def _patch_db():
    """connect_db / load_intersections 목"""
    return (
        patch.object(ad, "connect_db", return_value=MagicMock()),
        patch.object(ad, "load_intersections",
                     return_value=[(260322, "송내사거리")]),
    )


# ── _fetch_corrected_traffic 단위 테스트 ─────────────────────────────────────

class TestFetchCorrectedTraffic(unittest.TestCase):

    def _call(self, node_results=None, target_data=None):
        """_fetch_corrected_traffic 직접 호출 헬퍼"""
        nr = node_results if node_results is not None else _NODE_RESULTS_WITH_CORRECTION
        api_server._holiday_dates = set()  # 초기화
        req = {
            "node_ids":   [260322],
            "date_start": _DATE_START.isoformat(),
            "date_end":   _DATE_END.isoformat(),
            "hours":      _HOURS,
        }
        db_patch, inter_patch = _patch_db()
        with db_patch, inter_patch, \
             _patch_approaches(), \
             _patch_analyse(nr, target_data), \
             patch.object(ad, "select_baseline_years", return_value=[2024, 2025]), \
             patch.object(ad, "select_fallback_year",  return_value=2025), \
             patch.object(ad, "get_adjacent_month_periods", return_value=[]):
            return api_server._fetch_corrected_traffic(req)

    # 정상 케이스

    def test_corrected_slot_included_with_corrected_value(self):
        """A형 보정값이 있으면 해당 슬롯에 보정값이 반환되는지 확인"""
        result = self._call()
        slots = result["slots"]
        h8_north = next(
            (s for s in slots if s["hour"] == 8 and s["approach_name"] == "북향"),
            None
        )
        self.assertIsNotNone(h8_north)
        self.assertEqual(h8_north["traffic_volume"], 305)
        self.assertTrue(h8_north["is_corrected"])

    def test_original_slot_included_unchanged(self):
        """원본 데이터 슬롯은 is_corrected=False로 반환되는지 확인"""
        result = self._call()
        h6_north = next(
            s for s in result["slots"]
            if s["hour"] == 6 and s["approach_name"] == "북향"
        )
        self.assertEqual(h6_north["traffic_volume"], 300)
        self.assertFalse(h6_north["is_corrected"])

    def test_null_target_data_slot_excluded(self):
        """target_data 값이 None인 슬롯은 결과에서 제외되는지 확인"""
        result = self._call()
        h8_south = [
            s for s in result["slots"]
            if s["hour"] == 8 and s["approach_name"] == "남향"
        ]
        # 남향 8시: target_data=None, 보정 없음 → 슬롯 제외
        self.assertEqual(len(h8_south), 0)

    def test_stage2_normal_slot_not_in_corrected_index(self):
        """정상(stage2) 판정은 corrected 인덱스에 포함되지 않음"""
        # 정상 판정 슬롯의 보정값=None이므로 보정 인덱스 미포함
        # → 해당 시간 target_data 없으면 슬롯 제외
        result = self._call(
            node_results=_NODE_RESULTS_STAGE2,
            target_data=_TARGET_DATA,  # 8시 없음
        )
        h8 = [s for s in result["slots"] if s["hour"] == 8]
        self.assertEqual(len(h8), 0)

    def test_response_has_slots_key(self):
        """반환값에 'slots' 키가 있는지 확인"""
        result = self._call()
        self.assertIn("slots", result)

    def test_slot_has_required_fields(self):
        """슬롯에 필수 필드가 모두 포함되는지 확인"""
        result = self._call()
        required = {"date", "hour", "approach_id", "approach_name",
                    "traffic_volume", "is_corrected"}
        for slot in result["slots"]:
            for field in required:
                self.assertIn(field, slot, f"필드 '{field}' 누락")

    def test_date_format_is_iso(self):
        """date 필드가 ISO 형식(YYYY-MM-DD)인지 확인"""
        result = self._call()
        for slot in result["slots"]:
            d = slot["date"]
            self.assertRegex(d, r"^\d{4}-\d{2}-\d{2}$")

    # 경계값 케이스

    def test_empty_approaches_returns_empty_slots(self):
        """접근로가 없으면 슬롯이 비어있는지 확인"""
        api_server._holiday_dates = set()
        req = {
            "node_ids":   [260322],
            "date_start": _DATE_START.isoformat(),
            "date_end":   _DATE_END.isoformat(),
            "hours":      _HOURS,
        }
        db_patch, inter_patch = _patch_db()
        with db_patch, inter_patch, \
             patch.object(ad, "load_approaches", return_value=[]), \
             _patch_analyse([]), \
             patch.object(ad, "select_baseline_years", return_value=[2024, 2025]), \
             patch.object(ad, "select_fallback_year",  return_value=2025), \
             patch.object(ad, "get_adjacent_month_periods", return_value=[]):
            result = api_server._fetch_corrected_traffic(req)
        self.assertEqual(result["slots"], [])

    def test_server_not_initialized_raises(self):
        """_holiday_dates=None이면 RuntimeError 발생"""
        api_server._holiday_dates = None
        req = {
            "node_ids":   [260322],
            "date_start": _DATE_START.isoformat(),
            "date_end":   _DATE_END.isoformat(),
            "hours":      _HOURS,
        }
        with self.assertRaises(RuntimeError):
            api_server._fetch_corrected_traffic(req)

    # 오류 케이스

    def test_no_corrected_value_missing_slot_excluded(self):
        """보정값=None인 A형 → target_data에도 없으면 슬롯 제외"""
        no_corr = [{**_NODE_RESULTS_WITH_CORRECTION[0], "보정값": None}]
        td = {k: v for k, v in _TARGET_DATA.items() if k[2] != 8}  # 8시 제거
        result = self._call(node_results=no_corr, target_data=td)
        h8 = [s for s in result["slots"] if s["hour"] == 8]
        self.assertEqual(len(h8), 0)


# ── POST /corrected-traffic 엔드포인트 통합 테스트 ───────────────────────────

class TestCorrectedTrafficEndpoint(unittest.TestCase):

    def setUp(self):
        api_server._holiday_dates = set()

    def _post(self, payload: dict):
        return _client.post("/corrected-traffic", json=payload)

    def _mock_fetch(self, slots=None):
        """_fetch_corrected_traffic 목"""
        return patch.object(
            api_server, "_fetch_corrected_traffic",
            return_value={"slots": slots or []}
        )

    # 정상 케이스

    def test_returns_200_with_slots(self):
        """정상 요청 → 200 + slots 반환"""
        slots = [{"date": "2026-03-22", "hour": 6, "approach_id": 1,
                  "approach_name": "북향", "traffic_volume": 300,
                  "is_corrected": False}]
        with self._mock_fetch(slots):
            resp = self._post({
                "node_ids": [260322],
                "date_start": "2026-03-22",
                "date_end": "2026-03-22",
            })
        self.assertEqual(resp.status_code, 200)
        self.assertIn("slots", resp.json())
        self.assertEqual(len(resp.json()["slots"]), 1)

    def test_hours_preset_all_resolved(self):
        """hours_preset='all' 요청 → resolve_hours() 정상 처리"""
        with self._mock_fetch() as mock_fn:
            resp = self._post({
                "node_ids": [260322],
                "date_start": "2026-03-22",
                "date_end": "2026-03-22",
                "hours_preset": "all",
            })
        self.assertEqual(resp.status_code, 200)
        # _fetch_corrected_traffic에 hours=list(range(24)) 전달됐는지 확인
        call_req = mock_fn.call_args[0][0]
        self.assertEqual(call_req["hours"], list(range(24)))

    def test_hours_preset_peak_resolved(self):
        """hours_preset='peak' → [7,8,12,13,17,18] 전달"""
        with self._mock_fetch() as mock_fn:
            resp = self._post({
                "node_ids": [260322],
                "date_start": "2026-03-22",
                "date_end": "2026-03-22",
                "hours_preset": "peak",
            })
        self.assertEqual(resp.status_code, 200)
        call_req = mock_fn.call_args[0][0]
        self.assertEqual(call_req["hours"], [7, 8, 12, 13, 17, 18])

    def test_empty_slots_response(self):
        """슬롯 없음 → 200 + 빈 슬롯 배열"""
        with self._mock_fetch([]):
            resp = self._post({
                "node_ids": [260322],
                "date_start": "2026-03-22",
                "date_end": "2026-03-22",
            })
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.json()["slots"], [])

    # 유효성 오류 케이스

    def test_missing_node_ids_returns_422(self):
        """node_ids 누락 → 422"""
        resp = self._post({"date_start": "2026-03-22", "date_end": "2026-03-22"})
        self.assertEqual(resp.status_code, 422)

    def test_invalid_date_format_returns_422(self):
        """날짜 형식 오류 → 422"""
        resp = self._post({"node_ids": [1], "date_start": "20260322", "date_end": "20260322"})
        self.assertEqual(resp.status_code, 422)

    def test_start_after_end_returns_422(self):
        """시작일 > 종료일 → 422"""
        resp = self._post({"node_ids": [1], "date_start": "2026-03-23", "date_end": "2026-03-22"})
        self.assertEqual(resp.status_code, 422)

    def test_both_hours_and_preset_returns_422(self):
        """hours + hours_preset 동시 지정 → 422"""
        resp = self._post({
            "node_ids": [1], "date_start": "2026-03-22", "date_end": "2026-03-22",
            "hours": [7, 8], "hours_preset": "peak",
        })
        self.assertEqual(resp.status_code, 422)


# ════════════════════════════════════════════════════════════════════════════════
# 실행
# ════════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    unittest.main(verbosity=2)
