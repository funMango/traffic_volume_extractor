#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
get_traffic.py 단위 테스트

대상 함수:
  - _api_get() / _api_post()     HTTP 헬퍼 (신규)
  - input_intersections()        교차로 입력 (API 경유, 신규)
  - fetch_data()                 API 조회 (신규)
  - _parse_custom_time()         시간대 문자열 파싱
  - aggregate_by_node()          방향별 교통량 합산
  - _cell_value()                속성 → 셀 값 변환
  - _safe_sheet_name()           Excel 시트 이름 정규화
  - make_filename()              출력 파일명 생성
  - _node_total_rows()           교차로 방향 합산 시계열
  - save_excel()                 Excel 저장 (스텁)
  - input_sheet_type()           시트 설정 로직
"""

import io
import json
import os
import sys
import types
import unittest
from collections import defaultdict
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch, call

# ════════════════════════════════════════════════════════════════
# 외부 의존성 스텁 — 실제 라이브러리 없이 테스트 실행 가능하도록
# ════════════════════════════════════════════════════════════════

sys.path.insert(0, os.path.dirname(__file__))


def _make_module(name: str) -> types.ModuleType:
    m = types.ModuleType(name)
    sys.modules[name] = m
    return m


# oracledb 스텁
if "oracledb" not in sys.modules:
    _make_module("oracledb")

# dotenv 스텁
if "dotenv" not in sys.modules:
    _dotenv = _make_module("dotenv")
    _dotenv.load_dotenv = lambda *a, **kw: None  # type: ignore

# plotly 스텁
for _pmod in ("plotly", "plotly.graph_objects"):
    if _pmod not in sys.modules:
        _m = _make_module(_pmod)

        class _FakeFigure:
            def __init__(self, **kw): pass
            def add_trace(self, *a, **kw): pass
            def update_layout(self, **kw): pass
            def write_html(self, *a, **kw): pass

        class _FakeScatter:
            def __init__(self, **kw): pass

        _m.Figure = _FakeFigure
        _m.Scatter = _FakeScatter

# openpyxl 스텁
_openpyxl = _make_module("openpyxl")

class _FakeCell:
    def __init__(self):
        self.value = None
        self.fill  = None
        self.font  = None
        self.alignment = None
        self.column = 1

class _FakeWS:
    def __init__(self, title=""):
        self.title  = title
        self._rows: list[list] = []
        self.column_dimensions: dict = defaultdict(
            lambda: type("CD", (), {"width": 0})()
        )
        self._cells: dict[tuple, _FakeCell] = {}

    def append(self, row):
        self._rows.append(list(row))

    def cell(self, row, column):
        key = (row, column)
        if key not in self._cells:
            c = _FakeCell()
            c.column = column
            self._cells[key] = c
        return self._cells[key]

    @property
    def columns(self):
        if not self._rows:
            return []
        ncols = len(self._rows[0]) if self._rows else 0
        cols = []
        for ci in range(1, ncols + 1):
            col_cells = []
            for ri, row in enumerate(self._rows, 1):
                c = _FakeCell()
                c.column = ci
                c.value  = row[ci - 1] if ci - 1 < len(row) else None
                col_cells.append(c)
            cols.append(col_cells)
        return cols

class _FakeWorkbook:
    def __init__(self):
        self._sheets: list[_FakeWS] = [_FakeWS("Sheet")]
        self.active = self._sheets[0]
        self._saved_path = None

    def create_sheet(self, title=""):
        ws = _FakeWS(title)
        self._sheets.append(ws)
        return ws

    def remove(self, ws):
        if ws in self._sheets:
            self._sheets.remove(ws)

    def save(self, path):
        self._saved_path = path

    @property
    def sheet_names(self):
        return [ws.title for ws in self._sheets]

_openpyxl.Workbook = _FakeWorkbook

_styles = _make_module("openpyxl.styles")
for _cls_name in ("Font", "PatternFill", "Alignment", "Border", "Side"):
    setattr(_styles, _cls_name,
            type(_cls_name, (), {"__init__": lambda s, *a, **kw: None}))

_utils = _make_module("openpyxl.utils")
_utils.get_column_letter = lambda i: chr(64 + i) if i <= 26 else f"A{chr(64 + i - 26)}"

# numpy 스텁
if "numpy" not in sys.modules:
    _np = _make_module("numpy")


# ── get_traffic import ────────────────────────────────────────────────────────
import get_traffic as gt  # noqa: E402


# ════════════════════════════════════════════════════════════════
# 테스트 픽스처 (목 값)
# ════════════════════════════════════════════════════════════════

D1 = date(2026, 4, 13)
D2 = date(2026, 4, 14)

SAMPLE_ROWS = [
    {"date": D1, "hour": 7,  "node_name": "구지사거리", "approach_name": "북(남향)", "traffic_volume": 100},
    {"date": D1, "hour": 7,  "node_name": "구지사거리", "approach_name": "남(북향)", "traffic_volume": 80},
    {"date": D1, "hour": 8,  "node_name": "구지사거리", "approach_name": "북(남향)", "traffic_volume": 120},
    {"date": D1, "hour": 8,  "node_name": "구지사거리", "approach_name": "남(북향)", "traffic_volume": 90},
    {"date": D2, "hour": 7,  "node_name": "구지사거리", "approach_name": "북(남향)", "traffic_volume": 110},
    {"date": D2, "hour": 7,  "node_name": "구지사거리", "approach_name": "남(북향)", "traffic_volume": None},  # NULL
    {"date": D1, "hour": 7,  "node_name": "멀뫼사거리", "approach_name": "동(서향)", "traffic_volume": 200},
    {"date": D1, "hour": 7,  "node_name": "멀뫼사거리", "approach_name": "서(동향)", "traffic_volume": 150},
]


# ── HTTP 목 헬퍼 ─────────────────────────────────────────────────────────────

def _make_url_response(data: dict) -> MagicMock:
    """urllib.request.urlopen 반환값 목(컨텍스트 매니저)"""
    cm = MagicMock()
    cm.__enter__ = lambda s: s
    cm.__exit__  = MagicMock(return_value=False)
    cm.read.return_value = json.dumps(data, ensure_ascii=False).encode()
    return cm


# ════════════════════════════════════════════════════════════════
# 1. HTTP 헬퍼 (_api_get / _api_post)
# ════════════════════════════════════════════════════════════════

class TestApiHelpers(unittest.TestCase):

    # ── _api_get ────────────────────────────────────────────────
    def test_api_get_returns_dict(self):
        """_api_get: 정상 응답이 dict로 파싱되는지 확인"""
        resp_data = {"count": 2, "items": [{"node_id": 1, "name": "A"}]}
        cm = _make_url_response(resp_data)
        with patch("get_traffic.urllib.request.urlopen", return_value=cm):
            result = gt._api_get("/intersections")
        self.assertEqual(result["count"], 2)
        self.assertEqual(result["items"][0]["name"], "A")

    def test_api_get_builds_correct_url(self):
        """_api_get: API_BASE_URL + path 조합 확인"""
        cm = _make_url_response({})
        with patch("get_traffic.urllib.request.urlopen", return_value=cm) as mock_open:
            gt._api_get("/intersections")
        called_url = mock_open.call_args[0][0]
        self.assertTrue(called_url.startswith("http://localhost:8000"))
        self.assertIn("/intersections", called_url)

    def test_api_get_empty_response(self):
        """_api_get: 빈 dict 응답 정상 처리"""
        cm = _make_url_response({})
        with patch("get_traffic.urllib.request.urlopen", return_value=cm):
            result = gt._api_get("/anything")
        self.assertEqual(result, {})

    # ── _api_post ───────────────────────────────────────────────
    def test_api_post_returns_dict(self):
        """_api_post: 정상 응답이 dict로 파싱되는지 확인"""
        resp_data = {"slots": [{"date": "2026-04-13", "hour": 7}]}
        cm = _make_url_response(resp_data)
        with patch("get_traffic.urllib.request.urlopen", return_value=cm), \
             patch("get_traffic.urllib.request.Request") as mock_req:
            mock_req.return_value = MagicMock()
            result = gt._api_post("/corrected-traffic", {"node_ids": [1]})
        self.assertIn("slots", result)

    def test_api_post_sends_json_content_type(self):
        """_api_post: Content-Type: application/json 헤더 포함 확인"""
        cm = _make_url_response({})
        with patch("get_traffic.urllib.request.urlopen", return_value=cm), \
             patch("get_traffic.urllib.request.Request") as mock_req:
            mock_req.return_value = MagicMock()
            gt._api_post("/path", {"key": "val"})
        _, kwargs = mock_req.call_args
        headers = kwargs.get("headers", {})
        self.assertEqual(headers.get("Content-Type"), "application/json")

    def test_api_post_payload_encoded_as_json(self):
        """_api_post: payload가 JSON으로 인코딩되어 전달되는지 확인"""
        payload = {"node_ids": [260322], "date_start": "2026-03-22"}
        cm = _make_url_response({})
        with patch("get_traffic.urllib.request.urlopen", return_value=cm), \
             patch("get_traffic.urllib.request.Request") as mock_req:
            mock_req.return_value = MagicMock()
            gt._api_post("/path", payload)
        _, kwargs = mock_req.call_args
        data_bytes = kwargs.get("data", b"")
        decoded = json.loads(data_bytes.decode())
        self.assertEqual(decoded["node_ids"], [260322])


# ════════════════════════════════════════════════════════════════
# 2. input_intersections (API 경유)
# ════════════════════════════════════════════════════════════════

class TestInputIntersections(unittest.TestCase):

    API_ITEMS = [
        {"node_id": 1001, "name": "구지사거리"},
        {"node_id": 1002, "name": "멀뫼사거리"},
        {"node_id": 1003, "name": "송내사거리"},
    ]

    def _patch_api(self):
        """GET /intersections 목"""
        return patch.object(gt, "_api_get",
                            return_value={"count": 3, "items": self.API_ITEMS})

    def test_exact_name_match(self):
        """정확한 이름 입력 시 node_id 반환"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["구지사거리", "y"]):
            result = gt.input_intersections()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["node_id"], 1001)
        self.assertEqual(result[0]["name"], "구지사거리")

    def test_multiple_intersections(self):
        """쉼표로 복수 교차로 선택"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["구지사거리,멀뫼사거리", "y"]):
            result = gt.input_intersections()
        self.assertEqual(len(result), 2)
        names = [r["name"] for r in result]
        self.assertIn("구지사거리", names)
        self.assertIn("멀뫼사거리", names)

    def test_fuzzy_match_and_select(self):
        """유사 이름 → 후보 제시 → 사용자가 1번 선택"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["구지사거", "1", "y"]):
            result = gt.input_intersections()
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["name"], "구지사거리")

    def test_duplicate_removal(self):
        """같은 교차로 중복 입력 시 한 번만 포함"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["구지사거리,구지사거리", "y"]):
            result = gt.input_intersections()
        self.assertEqual(len(result), 1)

    def test_no_match_retry(self):
        """완전히 다른 이름 → 재입력 요구 → 두 번째 시도 성공"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["zzzzzzzz", "구지사거리", "y"]):
            result = gt.input_intersections()
        self.assertEqual(result[0]["name"], "구지사거리")

    def test_confirm_no_retry(self):
        """y/n 확인에서 n 선택 시 재입력"""
        with self._patch_api(), \
             patch("builtins.input", side_effect=["구지사거리", "n", "멀뫼사거리", "y"]):
            result = gt.input_intersections()
        self.assertEqual(result[0]["name"], "멀뫼사거리")


# ════════════════════════════════════════════════════════════════
# 3. fetch_data (API 스텁)
# ════════════════════════════════════════════════════════════════

class TestFetchData(unittest.TestCase):

    INTERSECTIONS = [{"node_id": 260322, "name": "송내사거리"}]
    PERIODS       = [(D1, D1, "260413~260413")]
    HOURS         = [7, 8]

    # API 응답 목 값 (보정 완료 슬롯)
    API_SLOTS = [
        {"date": "2026-04-13", "hour": 7, "approach_id": 1,
         "approach_name": "북향", "traffic_volume": 150, "is_corrected": False},
        {"date": "2026-04-13", "hour": 7, "approach_id": 2,
         "approach_name": "남향", "traffic_volume": 130, "is_corrected": False},
        {"date": "2026-04-13", "hour": 8, "approach_id": 1,
         "approach_name": "북향", "traffic_volume": 200, "is_corrected": True},  # 보정값
    ]

    def _patch_post(self, slots=None):
        data = {"slots": slots if slots is not None else self.API_SLOTS}
        return patch.object(gt, "_api_post", return_value=data)

    def test_returns_correct_structure(self):
        """rows의 각 항목에 필수 키가 포함되는지 확인"""
        with self._patch_post():
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        self.assertTrue(len(rows) > 0)
        for r in rows:
            self.assertIn("date",           r)
            self.assertIn("hour",           r)
            self.assertIn("node_name",      r)
            self.assertIn("approach_name",  r)
            self.assertIn("traffic_volume", r)

    def test_date_is_date_object(self):
        """date 필드가 date 객체로 변환되는지 확인"""
        with self._patch_post():
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        for r in rows:
            self.assertIsInstance(r["date"], date)

    def test_node_name_set_correctly(self):
        """node_name이 intersections의 name으로 설정되는지 확인"""
        with self._patch_post():
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        for r in rows:
            self.assertEqual(r["node_name"], "송내사거리")

    def test_corrected_value_used(self):
        """보정값(is_corrected=True)이 traffic_volume으로 반환되는지 확인"""
        with self._patch_post():
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        h8_north = next(
            r for r in rows
            if r["hour"] == 8 and r["approach_name"] == "북향"
        )
        self.assertEqual(h8_north["traffic_volume"], 200)

    def test_api_called_with_correct_payload(self):
        """_api_post 호출 시 올바른 payload가 전달되는지 확인"""
        with self._patch_post() as mock_post:
            gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        call_args = mock_post.call_args
        path    = call_args[0][0]
        payload = call_args[0][1]
        self.assertEqual(path, "/corrected-traffic")
        self.assertIn(260322, payload["node_ids"])
        self.assertEqual(payload["date_start"], "2026-04-13")
        self.assertEqual(payload["hours"], [7, 8])

    def test_multiple_periods_merged(self):
        """복수 기간 조회 시 rows가 합산되는지 확인"""
        periods = [
            (D1, D1, "260413~260413"),
            (D2, D2, "260414~260414"),
        ]
        slots_d2 = [
            {"date": "2026-04-14", "hour": 7, "approach_id": 1,
             "approach_name": "북향", "traffic_volume": 160, "is_corrected": False},
        ]

        call_count = {"n": 0}
        def fake_post(path, payload):
            call_count["n"] += 1
            if payload["date_start"] == "2026-04-13":
                return {"slots": self.API_SLOTS}
            return {"slots": slots_d2}

        with patch.object(gt, "_api_post", side_effect=fake_post):
            rows = gt.fetch_data(self.INTERSECTIONS, periods, self.HOURS)

        self.assertEqual(call_count["n"], 2)
        dates = {r["date"] for r in rows}
        self.assertIn(D1, dates)
        self.assertIn(D2, dates)

    def test_empty_slots_returns_empty(self):
        """API가 슬롯을 반환하지 않으면 빈 리스트"""
        with self._patch_post(slots=[]):
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        self.assertEqual(rows, [])

    def test_sorted_by_node_date_hour_approach(self):
        """결과가 교차로→날짜→시간→방향 순 정렬되는지 확인"""
        with self._patch_post():
            rows = gt.fetch_data(self.INTERSECTIONS, self.PERIODS, self.HOURS)
        sort_keys = [
            (r["node_name"], r["date"], r["hour"], r["approach_name"])
            for r in rows
        ]
        self.assertEqual(sort_keys, sorted(sort_keys))

    def test_multiple_intersections(self):
        """복수 교차로 각각 API 호출"""
        inters = [
            {"node_id": 100, "name": "A교차로"},
            {"node_id": 200, "name": "B교차로"},
        ]
        call_node_ids = []
        def fake_post(path, payload):
            call_node_ids.extend(payload["node_ids"])
            return {"slots": []}

        with patch.object(gt, "_api_post", side_effect=fake_post):
            gt.fetch_data(inters, self.PERIODS, self.HOURS)

        self.assertIn(100, call_node_ids)
        self.assertIn(200, call_node_ids)


# ════════════════════════════════════════════════════════════════
# 4. _parse_custom_time
# ════════════════════════════════════════════════════════════════

class TestParseCustomTime(unittest.TestCase):

    def test_single_range(self):
        self.assertEqual(gt._parse_custom_time("18:00~20:00"), [18, 19])

    def test_multiple_ranges(self):
        self.assertEqual(gt._parse_custom_time("18:00~20:00, 22:00~24:00"), [18, 19, 22, 23])

    def test_24_hour_end(self):
        self.assertEqual(gt._parse_custom_time("22:00~24:00"), [22, 23])

    def test_duplicate_removal(self):
        self.assertEqual(gt._parse_custom_time("07:00~09:00, 08:00~10:00"), [7, 8, 9])

    def test_single_hour(self):
        self.assertEqual(gt._parse_custom_time("7:00~8:00"), [7])

    def test_zero_start(self):
        self.assertEqual(gt._parse_custom_time("0:00~2:00"), [0, 1])

    def test_full_day(self):
        self.assertEqual(gt._parse_custom_time("0:00~24:00"), list(range(24)))

    def test_wrong_format(self):
        self.assertIsNone(gt._parse_custom_time("1800~2000"))

    def test_wrong_minutes(self):
        self.assertIsNone(gt._parse_custom_time("18:30~20:00"))

    def test_start_equals_end(self):
        self.assertIsNone(gt._parse_custom_time("7:00~7:00"))

    def test_start_after_end(self):
        self.assertIsNone(gt._parse_custom_time("20:00~18:00"))

    def test_empty_string(self):
        self.assertIsNone(gt._parse_custom_time(""))

    def test_out_of_range_hour(self):
        self.assertIsNone(gt._parse_custom_time("23:00~25:00"))


# ════════════════════════════════════════════════════════════════
# 5. aggregate_by_node
# ════════════════════════════════════════════════════════════════

class TestAggregateByNode(unittest.TestCase):

    def test_basic_aggregation(self):
        result = gt.aggregate_by_node(SAMPLE_ROWS)
        entry = next(r for r in result
                     if r["node_name"] == "구지사거리"
                     and r["date"] == D1 and r["hour"] == 7)
        self.assertEqual(entry["traffic_volume"], 180)

    def test_null_ignored_in_sum(self):
        entry = next(r for r in gt.aggregate_by_node(SAMPLE_ROWS)
                     if r["node_name"] == "구지사거리"
                     and r["date"] == D2 and r["hour"] == 7)
        self.assertEqual(entry["traffic_volume"], 110)

    def test_all_null_excluded(self):
        rows = [
            {"date": D1, "hour": 9, "node_name": "A", "approach_name": "북", "traffic_volume": None},
            {"date": D1, "hour": 9, "node_name": "A", "approach_name": "남", "traffic_volume": None},
        ]
        self.assertEqual(gt.aggregate_by_node(rows), [])

    def test_multiple_nodes_separate(self):
        result = gt.aggregate_by_node(SAMPLE_ROWS)
        node_names = {r["node_name"] for r in result}
        self.assertIn("구지사거리", node_names)
        self.assertIn("멀뫼사거리", node_names)

    def test_approach_name_is_empty(self):
        result = gt.aggregate_by_node(SAMPLE_ROWS)
        for r in result:
            self.assertEqual(r["approach_name"], "")

    def test_sorted_by_node_date_hour(self):
        result = gt.aggregate_by_node(SAMPLE_ROWS)
        keys = [(r["node_name"], r["date"], r["hour"]) for r in result]
        self.assertEqual(keys, sorted(keys))

    def test_empty_input(self):
        self.assertEqual(gt.aggregate_by_node([]), [])


# ════════════════════════════════════════════════════════════════
# 6. _cell_value
# ════════════════════════════════════════════════════════════════

class TestCellValue(unittest.TestCase):

    SAMPLE_ROW = {
        "date": D1, "hour": 8,
        "node_name": "구지사거리", "approach_name": "북(남향)",
        "traffic_volume": 120,
    }

    def test_시간(self):
        self.assertEqual(gt._cell_value(self.SAMPLE_ROW, "시간"), "2026-04-13 08:00")

    def test_교차로명(self):
        self.assertEqual(gt._cell_value(self.SAMPLE_ROW, "교차로명"), "구지사거리")

    def test_방향명(self):
        self.assertEqual(gt._cell_value(self.SAMPLE_ROW, "방향명"), "북(남향)")

    def test_교통량(self):
        self.assertEqual(gt._cell_value(self.SAMPLE_ROW, "교통량"), 120)

    def test_교통량_none(self):
        row = dict(self.SAMPLE_ROW, traffic_volume=None)
        self.assertIsNone(gt._cell_value(row, "교통량"))

    def test_unknown_attr(self):
        self.assertIsNone(gt._cell_value(self.SAMPLE_ROW, "없는속성"))

    def test_시간_hour_padding(self):
        row = dict(self.SAMPLE_ROW, hour=7)
        self.assertEqual(gt._cell_value(row, "시간"), "2026-04-13 07:00")


# ════════════════════════════════════════════════════════════════
# 7. _safe_sheet_name
# ════════════════════════════════════════════════════════════════

class TestSafeSheetName(unittest.TestCase):

    def test_normal_name_unchanged(self):
        self.assertEqual(gt._safe_sheet_name("구지사거리"), "구지사거리")

    def test_forbidden_chars_replaced(self):
        for ch in r'\/*?:[]':
            self.assertEqual(gt._safe_sheet_name(ch), "_", f"'{ch}' not replaced")

    def test_long_name_truncated(self):
        self.assertEqual(len(gt._safe_sheet_name("가" * 40)), 31)

    def test_exactly_31_chars_not_truncated(self):
        name = "가" * 31
        self.assertEqual(gt._safe_sheet_name(name), name)

    def test_mixed(self):
        name = "구지사거리_북향/" + "A" * 30
        result = gt._safe_sheet_name(name)
        self.assertNotIn("/", result)
        self.assertLessEqual(len(result), 31)


# ════════════════════════════════════════════════════════════════
# 8. make_filename
# ════════════════════════════════════════════════════════════════

class TestMakeFilename(unittest.TestCase):

    def test_single_intersection(self):
        inter = [{"node_id": 1001, "name": "구지사거리"}]
        self.assertEqual(gt.make_filename(inter, "260413~260414"),
                         "구지사거리_260413~260414")

    def test_multiple_intersections(self):
        inter = [{"node_id": 1001, "name": "구지사거리"},
                 {"node_id": 1002, "name": "멀뫼사거리"}]
        self.assertEqual(gt.make_filename(inter, "260413~260414"),
                         "구지사거리외_1개_260413~260414")

    def test_period_preserved(self):
        inter = [{"node_id": 1001, "name": "A"}]
        self.assertTrue(gt.make_filename(inter, "260101~260131").endswith("260101~260131"))


# ════════════════════════════════════════════════════════════════
# 9. _node_total_rows
# ════════════════════════════════════════════════════════════════

class TestNodeTotalRows(unittest.TestCase):

    def test_sums_directions(self):
        result = gt._node_total_rows(SAMPLE_ROWS, "구지사거리")
        entry = next(r for r in result if r["date"] == D1 and r["hour"] == 7)
        self.assertEqual(entry["traffic_volume"], 180)

    def test_filters_other_nodes(self):
        result = gt._node_total_rows(SAMPLE_ROWS, "멀뫼사거리")
        dates_hours = {(r["date"], r["hour"]) for r in result}
        self.assertEqual(dates_hours, {(D1, 7)})

    def test_null_ignored(self):
        result = gt._node_total_rows(SAMPLE_ROWS, "구지사거리")
        entry = next(r for r in result if r["date"] == D2 and r["hour"] == 7)
        self.assertEqual(entry["traffic_volume"], 110)

    def test_sorted(self):
        result = gt._node_total_rows(SAMPLE_ROWS, "구지사거리")
        keys = [(r["date"], r["hour"]) for r in result]
        self.assertEqual(keys, sorted(keys))

    def test_unknown_node_empty(self):
        self.assertEqual(gt._node_total_rows(SAMPLE_ROWS, "없는교차로"), [])


# ════════════════════════════════════════════════════════════════
# 10. save_excel (openpyxl 스텁 활용)
# ════════════════════════════════════════════════════════════════

class TestSaveExcel(unittest.TestCase):

    def _sample_data(self):
        return [
            {"date": D1, "hour": 7, "node_name": "구지사거리",
             "approach_name": "북(남향)", "traffic_volume": 100},
            {"date": D1, "hour": 7, "node_name": "구지사거리",
             "approach_name": "남(북향)", "traffic_volume": 80},
            {"date": D1, "hour": 7, "node_name": "멀뫼사거리",
             "approach_name": "동(서향)", "traffic_volume": 200},
        ]

    def _captured_wb(self):
        captured = []

        class CapturingWB(_FakeWorkbook):
            def save(self, p):
                super().save(p)
                captured.append(self)

        return CapturingWB, captured

    def test_교차로별_sheet_count(self):
        CapturingWB, captured = self._captured_wb()
        with patch("get_traffic.Workbook", CapturingWB), \
             patch.object(Path, "mkdir", return_value=None):
            gt.save_excel(self._sample_data(),
                          ["시간", "교차로명", "방향명", "교통량"],
                          "교차로별", Path("/tmp/t.xlsx"))
        sheet_names = [ws.title for ws in captured[0]._sheets]
        self.assertIn("구지사거리", sheet_names)
        self.assertIn("멀뫼사거리", sheet_names)
        self.assertEqual(len(sheet_names), 2)

    def test_단일시트_sheet(self):
        CapturingWB, captured = self._captured_wb()
        with patch("get_traffic.Workbook", CapturingWB), \
             patch.object(Path, "mkdir", return_value=None):
            gt.save_excel(self._sample_data(), ["시간", "교통량"],
                          "단일시트", Path("/tmp/t.xlsx"))
        self.assertEqual([ws.title for ws in captured[0]._sheets], ["교통량"])

    def test_단일시트_정렬_시간_교차로명(self):
        data = [
            {"date": D1, "hour": 8, "node_name": "멀뫼사거리", "approach_name": "동", "traffic_volume": 1},
            {"date": D1, "hour": 7, "node_name": "원미사거리", "approach_name": "북", "traffic_volume": 2},
            {"date": D1, "hour": 7, "node_name": "구지사거리", "approach_name": "남", "traffic_volume": 3},
        ]
        CapturingWB, captured = self._captured_wb()
        with patch("get_traffic.Workbook", CapturingWB), \
             patch.object(Path, "mkdir", return_value=None):
            gt.save_excel(data, ["시간", "교차로명", "교통량"], "단일시트", Path("/tmp/t.xlsx"))

        ws = captured[0]._sheets[0]
        # row0: header, row1~: data
        self.assertEqual(ws._rows[1][1], "구지사거리")
        self.assertEqual(ws._rows[2][1], "원미사거리")
        self.assertEqual(ws._rows[3][1], "멀뫼사거리")

    def test_header_written_as_first_row(self):
        captured_ws = []

        class CapturingWB(_FakeWorkbook):
            def create_sheet(self, title=""):
                ws = _FakeWS(title)
                self._sheets.append(ws)
                captured_ws.append(ws)
                return ws
            def save(self, p): pass

        attrs = ["시간", "교차로명", "교통량"]
        with patch("get_traffic.Workbook", CapturingWB), \
             patch.object(Path, "mkdir", return_value=None):
            gt.save_excel(
                [{"date": D1, "hour": 7, "node_name": "A",
                  "approach_name": "북", "traffic_volume": 50}],
                attrs, "교차로별", Path("/tmp/t.xlsx")
            )
        self.assertEqual(captured_ws[0]._rows[0], attrs)


# ════════════════════════════════════════════════════════════════
# 11. input_sheet_type
# ════════════════════════════════════════════════════════════════

class TestInputSheetType(unittest.TestCase):

    def test_default_returns_단일시트(self):
        with patch("builtins.input", return_value=""):
            self.assertEqual(gt.input_sheet_type(has_node=False, has_dir=False), "단일시트")

    def test_user_selects_교차로별(self):
        with patch("builtins.input", return_value="2"):
            self.assertEqual(gt.input_sheet_type(has_node=True, has_dir=True), "교차로별")

    def test_user_selects_단일시트(self):
        with patch("builtins.input", return_value="1"):
            self.assertEqual(gt.input_sheet_type(has_node=True, has_dir=True), "단일시트")

    def test_invalid_then_valid_input(self):
        with patch("builtins.input", side_effect=["9", "abc", "2"]):
            self.assertEqual(gt.input_sheet_type(has_node=True, has_dir=True), "교차로별")


# ════════════════════════════════════════════════════════════════
# 실행
# ════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    unittest.main(verbosity=2)
