#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
교통량 이상탐지 시스템
이상값(A형: 결측, B형: 비정상 저값)을 탐지·보정하여 엑셀로 출력하고 시각화 HTML을 생성
"""

import os
import sys
import json
import difflib
import re
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from collections import defaultdict, OrderedDict
from threading import Event, Lock, Thread

import numpy as np
import oracledb
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import anomaly_core as core

# ── 경로 설정 ────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "00_Data"
RESULT_DIR = BASE_DIR / "02_Result"
ENV_PATH = DATA_DIR / ".env"
HOLIDAYS_PATH = DATA_DIR / "holidays_2026.json"

# ── API 설정 ─────────────────────────────────────────────────────────────────
API_BASE_URL = "http://localhost:8000"


def _api_get(path: str) -> dict:
    with urllib.request.urlopen(f"{API_BASE_URL}{path}") as resp:
        return json.loads(resp.read())


def _api_post(path: str, payload: dict) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode()
    req  = urllib.request.Request(
        f"{API_BASE_URL}{path}",
        data    = data,
        headers = {"Content-Type": "application/json"},
        method  = "POST",
    )
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"\n[DEBUG] POST {path} → HTTP {e.code}")
        print(f"[DEBUG] 요청 payload: {json.dumps(payload, ensure_ascii=False)}")
        print(f"[DEBUG] 서버 응답: {body}")
        raise


def _parse_api_response(slots: list, node_name: str) -> tuple:
    """API /corrected-traffic 응답 슬롯 → (node_results, target_data, approaches)

    node_results : 이상 슬롯 목록 (export_excel / export_visualization 호환 구조)
    target_data  : {(acsr_id, date, hour): original_volume} (시각화용 원본 데이터)
    approaches   : [(acsr_id, acsr_nm), ...] (등장 순서 기준)
    """
    node_results: list = []
    target_data: dict = {}
    approaches_seen: dict = {}  # acsr_id → acsr_nm (삽입 순서 보장)

    for slot in slots:
        d       = date.fromisoformat(slot["date"])
        h       = slot["hour"]
        acsr_id = slot["approach_id"]
        acsr_nm = slot["approach_name"]

        if acsr_id not in approaches_seen:
            approaches_seen[acsr_id] = acsr_nm

        # 원본 교통량 → target_data (정상 슬롯 포함 전체)
        orig_val = slot.get("traffic_volume")
        if orig_val is not None:
            target_data[(acsr_id, d, h)] = orig_val

        # A형/B형/A+B혼합 이상 슬롯만 node_results에 추가
        anomaly_type = slot.get("anomaly_type")
        if anomaly_type in ("A형", "B형", "A+B혼합"):
            node_results.append({
                "날짜":     d.strftime("%Y.%m.%d"),
                "시간":     f"{h:02d}:00",
                "교차로":   node_name,
                "방향":     acsr_nm,
                "교통량":   orig_val if orig_val is not None else 0,
                "판정":     anomaly_type,
                "보정값":   slot.get("corrected_value"),
                "보정방법": slot.get("correction_method"),
                "신뢰도":   slot.get("confidence"),
                "_acsr_id": acsr_id,
                "_date":    d,
                "_hour":    h,
                "_cell":    None,
            })

    approaches = list(approaches_seen.items())
    return node_results, target_data, approaches


def _sanitize_sheet_title(text: str) -> str:
    """Excel 시트명 제약(31자/금지문자)에 맞춰 정규화."""
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(text)).strip()
    return cleaned or "Sheet"


def _sanitize_path_component(text: str) -> str:
    """Windows 경로 구성요소 금지문자를 치환."""
    cleaned = re.sub(r"[<>:\"/\\|?*]", "_", str(text)).strip()
    return cleaned or "_"


SPINNER_FRAMES = ("|", "/", "-", "\\")


def _print_inline_progress(message: str, min_width: int = 100) -> None:
    print(f"\r{message.ljust(min_width)}", end="", flush=True)


class _SpinnerProgress:
    """단일 작업 구간용 스피너+퍼센트 진행 표시."""

    def __init__(self, label: str, start_percent: int = 0, max_percent: int = 95, interval_sec: float = 0.12):
        self.label = label
        self.percent = max(0, min(start_percent, 100))
        self.max_percent = max(0, min(max_percent, 100))
        self.interval_sec = interval_sec
        self._stop_event = Event()
        self._lock = Lock()
        self._thread: Thread | None = None

    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = Thread(target=self._run, daemon=True)
        self._thread.start()

    def set_percent(self, value: int) -> None:
        with self._lock:
            self.percent = max(0, min(value, 100))

    def _run(self) -> None:
        idx = 0
        while not self._stop_event.is_set():
            with self._lock:
                pct = self.percent
                if pct < self.max_percent:
                    self.percent += 1
            frame = SPINNER_FRAMES[idx % len(SPINNER_FRAMES)]
            idx += 1
            _print_inline_progress(f"  {self.label}... {frame} {pct:3d}%")
            time.sleep(self.interval_sec)

    def stop(self, final_percent: int = 100, done_text: str = "완료") -> None:
        self.set_percent(final_percent)
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join()
        _print_inline_progress(f"  {self.label}... {done_text} {max(0, min(final_percent, 100)):3d}%")
        print()


def _make_step_progress_callback(label: str, start_percent: int, end_percent: int):
    """다중 스텝 작업용 진행 콜백 생성."""
    state = {"idx": 0}
    start = max(0, min(start_percent, 100))
    end = max(start, min(end_percent, 100))

    def _callback(done: int, total: int, detail: str = "") -> None:
        bounded_total = max(total, 1)
        bounded_done = max(0, min(done, bounded_total))
        ratio = bounded_done / bounded_total
        pct = start + int((end - start) * ratio)
        frame = SPINNER_FRAMES[state["idx"] % len(SPINNER_FRAMES)]
        state["idx"] += 1

        detail_txt = f" ({bounded_done}/{bounded_total})"
        if detail:
            detail_txt += f" {detail}"

        if bounded_done >= bounded_total:
            _print_inline_progress(f"  {label}... 완료 {pct:3d}%{detail_txt}")
            print()
        else:
            _print_inline_progress(f"  {label}... {frame} {pct:3d}%{detail_txt}")

    return _callback


def _parse_api_response_drct(drct_slots: list, node_name: str) -> tuple:
    """API /corrected-traffic-drct 응답 파싱.

    Returns:
      - node_results: A/B 이상 슬롯 목록
      - target_data: {(acsr_id, drct_cd, date, hour): original_volume}
      - approaches: [(acsr_id, acsr_nm), ...]
      - drct_options_by_acsr: {
            acsr_id: {
                "approach_name": acsr_nm,
                "drcts": [(drct_cd, drct_name), ...],
            }
        }
    """
    node_results: list = []
    target_data: dict = {}
    approaches_seen: OrderedDict = OrderedDict()
    drct_options_by_acsr: OrderedDict = OrderedDict()
    drct_seen: dict = defaultdict(set)

    for slot in drct_slots:
        d = date.fromisoformat(slot["date"])
        h = int(slot["hour"])
        acsr_id = slot["approach_id"]
        acsr_nm = slot.get("approach_name") or str(acsr_id)
        drct_cd = _normalize_drct_cd(slot.get("drct_cd"))
        drct_nm = slot.get("drct_name") or drct_cd

        # DRCT 00(미분류)은 전체 흐름에서 제외
        if drct_cd == "00":
            continue

        if acsr_id not in approaches_seen:
            approaches_seen[acsr_id] = acsr_nm

        if acsr_id not in drct_options_by_acsr:
            drct_options_by_acsr[acsr_id] = {
                "approach_name": acsr_nm,
                "drcts": [],
            }
        if drct_cd not in drct_seen[acsr_id]:
            drct_seen[acsr_id].add(drct_cd)
            drct_options_by_acsr[acsr_id]["drcts"].append((drct_cd, drct_nm))

        orig_val = slot.get("traffic_volume")
        if orig_val is not None:
            target_data[(acsr_id, drct_cd, d, h)] = orig_val

        anomaly_type = slot.get("anomaly_type")
        if anomaly_type not in ("A형", "B형"):
            continue

        node_results.append({
            "날짜":      d.strftime("%Y.%m.%d"),
            "시간":      f"{h:02d}:00",
            "교차로":    node_name,
            "방향":      acsr_nm,
            "접근로방향": drct_nm,
            "교통량":    orig_val if orig_val is not None else 0,
            "판정":      anomaly_type,
            "보정값":    slot.get("corrected_value"),
            "보정방법":  slot.get("correction_method"),
            "신뢰도":    slot.get("confidence"),
            "_acsr_id":  acsr_id,
            "_drct_cd":  drct_cd,
            "_drct_name": drct_nm,
            "_date":     d,
            "_hour":     h,
            "_cell":     None,
        })

    approaches = list(approaches_seen.items())
    return node_results, target_data, approaches, drct_options_by_acsr


def _parse_api_response_acsr_from_drct(drct_slots: list, node_name: str) -> tuple:
    """DRCT 슬롯을 ACSR+시각 단위로 집계해 ACSR 출력 포맷으로 변환."""
    target_data: dict = {}
    node_results: list = []
    approaches_seen: OrderedDict = OrderedDict()
    aggregated: dict = {}

    for slot in drct_slots:
        d = date.fromisoformat(slot["date"])
        h = int(slot["hour"])
        acsr_id = slot["approach_id"]
        acsr_nm = slot.get("approach_name") or str(acsr_id)
        drct_cd = _normalize_drct_cd(slot.get("drct_cd"))

        # DRCT 00(미분류)은 전체 흐름에서 제외
        if drct_cd == "00":
            continue

        if acsr_id not in approaches_seen:
            approaches_seen[acsr_id] = acsr_nm

        key = (acsr_id, d, h)
        if key not in aggregated:
            aggregated[key] = {
                "acsr_nm": acsr_nm,
                "raw_sum": 0,
                "corr_sum": 0,
                "has_raw": False,
                "has_corr": False,
                "has_a": False,
                "has_b": False,
            }
        item = aggregated[key]

        raw_val = slot.get("traffic_volume")
        if raw_val is not None:
            item["raw_sum"] += raw_val
            item["has_raw"] = True

        corr_val = slot.get("corrected_value")
        if corr_val is None:
            corr_val = raw_val
        if corr_val is not None:
            item["corr_sum"] += corr_val
            item["has_corr"] = True

        anomaly_type = slot.get("anomaly_type")
        if anomaly_type == "A형":
            item["has_a"] = True
        elif anomaly_type == "B형":
            item["has_b"] = True
        elif anomaly_type == "A+B혼합":
            item["has_a"] = True
            item["has_b"] = True

    for (acsr_id, d, h), item in sorted(
        aggregated.items(),
        key=lambda kv: (kv[0][1], kv[0][2], str(kv[0][0])),
    ):
        raw_sum = item["raw_sum"] if item["has_raw"] else None
        corr_sum = item["corr_sum"] if item["has_corr"] else None
        if raw_sum is not None:
            target_data[(acsr_id, d, h)] = raw_sum

        anomaly_type = None
        if item["has_a"] and item["has_b"]:
            anomaly_type = "A+B혼합"
        elif item["has_a"]:
            anomaly_type = "A형"
        elif item["has_b"]:
            anomaly_type = "B형"

        if anomaly_type is None:
            continue

        node_results.append({
            "날짜": d.strftime("%Y.%m.%d"),
            "시간": f"{h:02d}:00",
            "교차로": node_name,
            "방향": item["acsr_nm"],
            "교통량": raw_sum if raw_sum is not None else 0,
            "판정": anomaly_type,
            "보정값": corr_sum,
            "보정방법": None,
            "신뢰도": None,
            "_acsr_id": acsr_id,
            "_date": d,
            "_hour": h,
            "_cell": None,
        })

    approaches = list(approaches_seen.items())
    return node_results, target_data, approaches


# ── 상수 ────────────────────────────────────────────────────────────────────
AVAILABLE_YEARS = [2022, 2024, 2025, 2026]
DATA_START_YEAR = 2023          # 데이터가 제대로 수집되기 시작한 연도
ZERO_RATE_THRESHOLD = 0.2
RATIO_THRESHOLD = 0.5
IQR_MULTIPLIER = 1.5
MEDIAN_RATIO = 0.3
MIN_CLEAN_SAMPLES = 5
MIN_CONFIDENCE_SAMPLES = 8  # n_clean 미만이면 신뢰도 LOW
MAX_INTERP_DAYS = 14        # 선형보간 최대 앞뒤 거리 합(일)
MAX_INTERP_DAYS_OK = 7      # 이하면 신뢰도 OK, 초과~MAX_INTERP_DAYS이면 LOW
ZERO_RATE_EXPANSION_THRESHOLD = 0.8   # Stage 1/2 확장 진입 임계값


# ════════════════════════════════════════════════════════════════
# DB / 공휴일
# ════════════════════════════════════════════════════════════════

PROFILE_DB_BREAKDOWN = os.getenv("ANOMALY_PROFILE_DB_BREAKDOWN", "1").strip().lower() in (
    "1", "true", "yes", "y", "on"
)

# Keep compatibility names in this module, but source values from core.
AVAILABLE_YEARS = core.AVAILABLE_YEARS
DATA_START_YEAR = core.DATA_START_YEAR
ZERO_RATE_THRESHOLD = core.ZERO_RATE_THRESHOLD
RATIO_THRESHOLD = core.RATIO_THRESHOLD
IQR_MULTIPLIER = core.IQR_MULTIPLIER
MEDIAN_RATIO = core.MEDIAN_RATIO
MIN_CLEAN_SAMPLES = core.MIN_CLEAN_SAMPLES
MIN_CONFIDENCE_SAMPLES = core.MIN_CONFIDENCE_SAMPLES
MAX_INTERP_DAYS = core.MAX_INTERP_DAYS
MAX_INTERP_DAYS_OK = core.MAX_INTERP_DAYS_OK
ZERO_RATE_EXPANSION_THRESHOLD = core.ZERO_RATE_EXPANSION_THRESHOLD

BASELINE_CACHE_ENABLED = os.getenv("ANOMALY_BASELINE_CACHE", "1").strip().lower() in (
    "1", "true", "yes", "y", "on"
)
BASELINE_CACHE_MAX_ENTRIES = max(1, int(os.getenv("ANOMALY_BASELINE_CACHE_MAX_ENTRIES", "64")))
_baseline_query_cache: OrderedDict[tuple, list] = OrderedDict()
_baseline_cache_lock = Lock()
_baseline_cache_hits = 0
_baseline_cache_misses = 0


def _baseline_cache_key(
    node_id,
    baseline_years: list,
    hours: list | None,
    months: list[int] | None,
    acsr_ids: list | None,
    source: str = "ACSR",
) -> tuple:
    years_key = tuple(sorted({int(y) for y in baseline_years}))
    hour_vals = None
    if hours:
        norm_hours = sorted({int(h) for h in hours if 0 <= int(h) <= 23})
        if norm_hours and set(norm_hours) != set(range(24)):
            hour_vals = tuple(norm_hours)
    month_vals = tuple(sorted({int(m) for m in months})) if months else None
    acsr_vals = tuple(sorted({str(a) for a in acsr_ids})) if acsr_ids else None
    return ("baseline", source, str(node_id), years_key, hour_vals, month_vals, acsr_vals)


def _baseline_cache_get(cache_key: tuple):
    global _baseline_cache_hits, _baseline_cache_misses
    if not BASELINE_CACHE_ENABLED:
        return None
    with _baseline_cache_lock:
        rows = _baseline_query_cache.get(cache_key)
        if rows is not None:
            _baseline_query_cache.move_to_end(cache_key)
            _baseline_cache_hits += 1
            return rows
        _baseline_cache_misses += 1
        return None


def _baseline_cache_set(cache_key: tuple, rows: list) -> None:
    if not BASELINE_CACHE_ENABLED:
        return
    with _baseline_cache_lock:
        _baseline_query_cache[cache_key] = rows
        _baseline_query_cache.move_to_end(cache_key)
        while len(_baseline_query_cache) > BASELINE_CACHE_MAX_ENTRIES:
            _baseline_query_cache.popitem(last=False)


def _baseline_cache_stats() -> tuple[int, int, int]:
    with _baseline_cache_lock:
        return _baseline_cache_hits, _baseline_cache_misses, len(_baseline_query_cache)


def connect_db():
    load_dotenv(ENV_PATH)
    return oracledb.connect(
        user=os.getenv("DB_USER").strip(),
        password=os.getenv("DB_PASSWORD").strip(),
        host=os.getenv("DB_HOST").strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.getenv("DB_SERVICE").strip()
    )


def load_holidays() -> set:
    """2026년 공휴일 날짜 set (문자열 'YYYY-MM-DD')"""
    if not HOLIDAYS_PATH.exists():
        return set()
    with open(HOLIDAYS_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return {h["date"] for h in data.get("holidays", [])}


def get_day_type(d: date, holiday_dates: set) -> int:
    """0=평일, 1=토요일, 2=일요일·공휴일"""
    if d.strftime("%Y-%m-%d") in holiday_dates:
        return 2
    wd = d.weekday()  # 0=Mon … 6=Sun
    if wd == 6:
        return 2
    if wd == 5:
        return 1
    return 0


# ════════════════════════════════════════════════════════════════
# DB 조회
# ════════════════════════════════════════════════════════════════

def load_intersections(conn) -> list:
    """[(node_id, crsrd_nm), ...]"""
    with conn.cursor() as cur:
        cur.execute("SELECT NODE_ID, CRSRD_NM FROM M_CRSRD_INF ORDER BY CRSRD_NM")
        return [(r[0], r[1]) for r in cur.fetchall()]


def load_approaches(conn, node_id) -> list:
    """[(acsr_id, acsr_nm), ...]"""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT ACSR_ID, ACSR_NM FROM M_CRSRD_ACSR_INF "
            "WHERE NODE_ID = :nid ORDER BY ACSR_ID",
            nid=node_id
        )
        return [(r[0], r[1]) for r in cur.fetchall()]


def _normalize_drct_cd(drct_cd) -> str:
    if drct_cd is None:
        return ""
    value = str(drct_cd).strip()
    if not value:
        return ""
    if value.isdigit():
        return value.zfill(2)
    return value


def load_drct_code_map(conn) -> dict:
    """{drct_cd: drct_name} from M_CD_INF where GRP_CD='DRCT_CD'."""
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT CD, CD_NM
            FROM M_CD_INF
            WHERE GRP_CD = 'DRCT_CD'
            ORDER BY CD
            """
        )
        return {_normalize_drct_cd(cd): cd_nm for cd, cd_nm in cur.fetchall()}


def load_drct_approaches(conn, node_id, drct_code_map: dict | None = None) -> tuple[list, dict]:
    """Return DRCT analysis units and metadata.

    Returns:
      - approaches: [((acsr_id, drct_cd), drct_name), ...]
      - meta_by_key: {
            (acsr_id, drct_cd): {
                "approach_id": acsr_id,
                "approach_name": acsr_nm,
                "drct_cd": drct_cd,
                "drct_name": drct_nm,
            }
        }
    """
    if drct_code_map is None:
        drct_code_map = load_drct_code_map(conn)
    acsr_name_map = {aid: anm for aid, anm in load_approaches(conn, node_id)}
    sql = """
        SELECT DISTINCT ACSR_ID, LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') AS DRCT_CD
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE NODE_ID = :nid
          AND TOT_DT < SYSDATE
        ORDER BY ACSR_ID, DRCT_CD
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, nid=node_id)
        rows = cur.fetchall()

    approaches: list = []
    meta_by_key: dict = {}
    for acsr_id, drct_cd in rows:
        norm_drct_cd = _normalize_drct_cd(drct_cd)
        key = (acsr_id, norm_drct_cd)
        acsr_nm = acsr_name_map.get(acsr_id, str(acsr_id))
        drct_nm = drct_code_map.get(norm_drct_cd, norm_drct_cd)
        approaches.append((key, drct_nm))
        meta_by_key[key] = {
            "approach_id": acsr_id,
            "approach_name": acsr_nm,
            "drct_cd": norm_drct_cd,
            "drct_name": drct_nm,
        }
    return approaches, meta_by_key


def _build_drct_pair_clause(drct_keys: list | None, bind_params: dict) -> str:
    if not drct_keys:
        return ""
    unique_keys = sorted(
        {(key[0], _normalize_drct_cd(key[1])) for key in drct_keys},
        key=lambda x: (str(x[0]), str(x[1])),
    )
    clauses: list[str] = []
    for idx, (acsr_id, drct_cd) in enumerate(unique_keys):
        acsr_name = f"drct_acsr_{idx}"
        code_name = f"drct_cd_{idx}"
        bind_params[acsr_name] = acsr_id
        bind_params[code_name] = drct_cd
        clauses.append(
            f"(ACSR_ID = :{acsr_name} AND LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') = :{code_name})"
        )
    return f"\n          AND ({' OR '.join(clauses)})"


def load_drct_baseline_data(
    conn,
    node_id,
    baseline_years: list,
    hours: list | None = None,
    months: list[int] | None = None,
    drct_keys: list | None = None,
) -> list:
    """DRCT 베이스라인 조회: [((acsr_id, drct_cd), tot_dt, trf_qnty), ...]."""
    if not baseline_years:
        return []

    drct_tokens = None
    if drct_keys:
        drct_tokens = [f"{key[0]}|{_normalize_drct_cd(key[1])}" for key in drct_keys]
    cache_key = _baseline_cache_key(
        node_id,
        baseline_years,
        hours,
        months,
        drct_tokens,
        source="DRCT",
    )
    cached_rows = _baseline_cache_get(cache_key)
    if cached_rows is not None:
        return cached_rows

    years = sorted(set(baseline_years))
    if months:
        month_vals = sorted({m for m in months if 1 <= m <= 12})
        if not month_vals:
            return []

        def _next_month(y: int, m: int) -> tuple[int, int]:
            return (y + 1, 1) if m == 12 else (y, m + 1)

        range_conditions = []
        for y in years:
            for m in month_vals:
                ny, nm = _next_month(y, m)
                range_conditions.append(
                    f"(TOT_DT >= TO_DATE('{y}-{m:02d}-01','YYYY-MM-DD') "
                    f"AND TOT_DT < TO_DATE('{ny}-{nm:02d}-01','YYYY-MM-DD'))"
                )
        date_clause = " OR ".join(range_conditions)
    else:
        date_clause = " OR ".join(
            f"(TOT_DT >= TO_DATE('{y}-01-01','YYYY-MM-DD') "
            f"AND TOT_DT < TO_DATE('{y + 1}-01-01','YYYY-MM-DD'))"
            for y in years
        )

    hour_clause = ""
    if hours and set(hours) != set(range(24)):
        hour_in = ",".join(str(h) for h in sorted(set(hours)))
        hour_clause = f"\n          AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})"

    bind_params = {"nid": node_id}
    drct_clause = _build_drct_pair_clause(drct_keys, bind_params)

    sql = f"""
        SELECT ACSR_ID, LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') AS DRCT_CD, TOT_DT, TRF_QNTY
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE NODE_ID = :nid
          AND ({date_clause})
          {drct_clause}
          {hour_clause}
          AND TOT_DT < SYSDATE
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, bind_params)
        rows = [
            ((acsr_id, _normalize_drct_cd(drct_cd)), tot_dt, trf_qnty)
            for acsr_id, drct_cd, tot_dt, trf_qnty in cur.fetchall()
        ]

    _baseline_cache_set(cache_key, rows)
    return rows


def load_drct_adjacent_month_data(
    conn,
    node_id: int,
    adj_periods: list[tuple[int, int]],
    hours: list | None = None,
    drct_keys: list | None = None,
) -> list:
    """DRCT 인접 월 조회: [((acsr_id, drct_cd), tot_dt, trf_qnty), ...]."""
    if not adj_periods:
        return []

    def _next_month(y, m):
        return (y + 1, 1) if m == 12 else (y, m + 1)

    conditions = " OR ".join(
        f"(TOT_DT >= TO_DATE('{y}-{m:02d}-01','YYYY-MM-DD') "
        f"AND TOT_DT < TO_DATE('{_next_month(y, m)[0]}-{_next_month(y, m)[1]:02d}-01','YYYY-MM-DD'))"
        for y, m in adj_periods
    )
    hour_clause = ""
    if hours and set(hours) != set(range(24)):
        hour_in = ",".join(str(h) for h in sorted(set(hours)))
        hour_clause = f"\n          AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})"
    bind_params = {"nid": node_id}
    drct_clause = _build_drct_pair_clause(drct_keys, bind_params)
    sql = f"""
        SELECT ACSR_ID, LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') AS DRCT_CD, TOT_DT, TRF_QNTY
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE NODE_ID = :nid
          AND ({conditions})
          {drct_clause}
          {hour_clause}
          AND TOT_DT < SYSDATE
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, bind_params)
        return [
            ((acsr_id, _normalize_drct_cd(drct_cd)), tot_dt, trf_qnty)
            for acsr_id, drct_cd, tot_dt, trf_qnty in cur.fetchall()
        ]


def load_drct_target_data(
    conn,
    node_id,
    date_start: date,
    date_end: date,
    hours: list,
) -> dict:
    """{((acsr_id, drct_cd), date, hour): trf_qnty}"""
    if not hours:
        return {}
    ds_dt = datetime(date_start.year, date_start.month, date_start.day)
    de_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)
    if set(hours) == set(range(24)):
        sql = """
            SELECT ACSR_ID, LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') AS DRCT_CD, TOT_DT, TRF_QNTY
            FROM S_CRSRD_DRCT_TRF_1HH
            WHERE NODE_ID = :nid
              AND TOT_DT >= :ds
              AND TOT_DT < :de_next
        """
        with conn.cursor() as cur:
            cur.arraysize = 10000
            cur.execute(sql, nid=node_id, ds=ds_dt, de_next=de_next)
            result = {}
            for acsr_id, drct_cd, tot_dt, trf_qnty in cur.fetchall():
                d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
                h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
                key = (acsr_id, _normalize_drct_cd(drct_cd))
                result[(key, d, h)] = trf_qnty
            return result
    hour_in = ",".join(str(h) for h in hours)
    sql = f"""
        SELECT ACSR_ID, LPAD(TRIM(TO_CHAR(DRCT_CD)), 2, '0') AS DRCT_CD, TOT_DT, TRF_QNTY
        FROM S_CRSRD_DRCT_TRF_1HH
        WHERE NODE_ID = :nid
          AND TOT_DT >= :ds
          AND TOT_DT < :de_next
          AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, nid=node_id, ds=ds_dt, de_next=de_next)
        result = {}
        for acsr_id, drct_cd, tot_dt, trf_qnty in cur.fetchall():
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            key = (acsr_id, _normalize_drct_cd(drct_cd))
            result[(key, d, h)] = trf_qnty
        return result


def load_baseline_data(
    conn,
    node_id,
    baseline_years: list,
    hours: list | None = None,
    months: list[int] | None = None,
    acsr_ids: list | None = None,
) -> list:
    """베이스라인 연도 시간별 접근로 데이터: [(acsr_id, tot_dt, trf_qnty), ...]"""
    if not baseline_years:
        return []

    cache_key = _baseline_cache_key(node_id, baseline_years, hours, months, acsr_ids)
    cached_rows = _baseline_cache_get(cache_key)
    if cached_rows is not None:
        return cached_rows

    years = sorted(set(baseline_years))
    if months:
        month_vals = sorted({m for m in months if 1 <= m <= 12})
        if not month_vals:
            return []

        def _next_month(y: int, m: int) -> tuple[int, int]:
            return (y + 1, 1) if m == 12 else (y, m + 1)

        range_conditions = []
        for y in years:
            for m in month_vals:
                ny, nm = _next_month(y, m)
                range_conditions.append(
                    f"(TOT_DT >= TO_DATE('{y}-{m:02d}-01','YYYY-MM-DD') "
                    f"AND TOT_DT < TO_DATE('{ny}-{nm:02d}-01','YYYY-MM-DD'))"
                )
        date_clause = " OR ".join(range_conditions)
    else:
        date_clause = " OR ".join(
            f"(TOT_DT >= TO_DATE('{y}-01-01','YYYY-MM-DD') "
            f"AND TOT_DT < TO_DATE('{y + 1}-01-01','YYYY-MM-DD'))"
            for y in years
        )

    hour_clause = ""
    if hours and set(hours) != set(range(24)):
        hour_in = ",".join(str(h) for h in sorted(set(hours)))
        hour_clause = f"\n          AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})"

    acsr_clause = ""
    bind_params = {"nid": node_id}
    if acsr_ids:
        uniq_acsr = sorted(set(acsr_ids), key=lambda x: str(x))
        bind_names: list[str] = []
        for idx, acsr_id in enumerate(uniq_acsr):
            name = f"acsr_{idx}"
            bind_names.append(f":{name}")
            bind_params[name] = acsr_id
        acsr_clause = f"\n          AND ACSR_ID IN ({', '.join(bind_names)})"

    sql = f"""
        SELECT ACSR_ID, TOT_DT, TRF_QNTY
        FROM S_CRSRD_ACSR_TRF_1HH
        WHERE NODE_ID = :nid
          AND ({date_clause})
          {acsr_clause}
          {hour_clause}
          AND TOT_DT < SYSDATE
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, bind_params)
        rows = cur.fetchall()

    _baseline_cache_set(cache_key, rows)
    return rows


def get_adjacent_month_periods(date_start: date, date_end: date) -> list[tuple[int, int]]:
    """대상 기간 인접 월 (연도, 월) 목록 반환 — 대상 기간 자체 및 미래 제외"""
    # 대상 기간에 포함된 (연도, 월) 집합
    target_ym = set()
    cur = date_start.replace(day=1)
    while cur <= date_end:
        target_ym.add((cur.year, cur.month))
        cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)

    adj = set()
    for y, m in target_ym:
        prev_ym = (y - 1, 12) if m == 1  else (y, m - 1)
        next_ym = (y + 1, 1)  if m == 12 else (y, m + 1)
        adj.add(prev_ym)
        adj.add(next_ym)

    adj -= target_ym  # 대상 기간 자체 제외
    today = date.today()
    adj = {(y, m) for y, m in adj if date(y, m, 1) <= today}
    return sorted(adj)


def load_adjacent_month_data(
    conn,
    node_id: int,
    adj_periods: list[tuple[int, int]],
    hours: list | None = None,
) -> list:
    """C형 인접 월 데이터 조회: [(acsr_id, tot_dt, trf_qnty), ...]"""
    if not adj_periods:
        return []

    def _next_month(y, m):
        return (y + 1, 1) if m == 12 else (y, m + 1)

    conditions = " OR ".join(
        f"(TOT_DT >= TO_DATE('{y}-{m:02d}-01','YYYY-MM-DD') "
        f"AND TOT_DT < TO_DATE('{_next_month(y, m)[0]}-{_next_month(y, m)[1]:02d}-01','YYYY-MM-DD'))"
        for y, m in adj_periods
    )
    hour_clause = ""
    if hours and set(hours) != set(range(24)):
        hour_in = ",".join(str(h) for h in sorted(set(hours)))
        hour_clause = f"\n          AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})"
    sql = f"""
        SELECT ACSR_ID, TOT_DT, TRF_QNTY
        FROM S_CRSRD_ACSR_TRF_1HH
        WHERE NODE_ID = :nid
          AND ({conditions})
          {hour_clause}
          AND TOT_DT < SYSDATE
    """
    with conn.cursor() as cur:
        cur.arraysize = 10000
        cur.execute(sql, nid=node_id)
        return cur.fetchall()


def load_target_data(conn, node_id, date_start: date, date_end: date, hours: list) -> dict:
    """{(acsr_id, date, hour): trf_qnty}"""
    if not hours:
        return {}
    ds_dt   = datetime(date_start.year, date_start.month, date_start.day)
    de_next = datetime(date_end.year, date_end.month, date_end.day) + timedelta(days=1)
    if set(hours) == set(range(24)):
        sql = """
            SELECT ACSR_ID, TOT_DT, TRF_QNTY
            FROM S_CRSRD_ACSR_TRF_1HH
            WHERE NODE_ID = :nid
              AND TOT_DT >= :ds
              AND TOT_DT < :de_next
        """
        with conn.cursor() as cur:
            cur.arraysize = 10000
            cur.execute(sql, nid=node_id, ds=ds_dt, de_next=de_next)
            result = {}
            for acsr_id, tot_dt, trf_qnty in cur.fetchall():
                d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
                h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
                result[(acsr_id, d, h)] = trf_qnty
            return result
    else:
        hour_in = ",".join(str(h) for h in hours)
        sql = f"""
            SELECT ACSR_ID, TOT_DT, TRF_QNTY
            FROM S_CRSRD_ACSR_TRF_1HH
            WHERE NODE_ID = :nid
              AND TOT_DT >= :ds
              AND TOT_DT < :de_next
              AND TO_NUMBER(TO_CHAR(TOT_DT, 'HH24')) IN ({hour_in})
        """
        with conn.cursor() as cur:
            cur.arraysize = 10000
            cur.execute(sql, nid=node_id, ds=ds_dt, de_next=de_next)
            result = {}
            for acsr_id, tot_dt, trf_qnty in cur.fetchall():
                d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
                h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
                result[(acsr_id, d, h)] = trf_qnty
            return result


# ════════════════════════════════════════════════════════════════
# 베이스라인 연도 선택
# ════════════════════════════════════════════════════════════════

def select_baseline_years(target_year: int) -> list:
    """보유 연도에서 대상 연도를 제외하고 양옆 최대 2개 선택"""
    candidates = [y for y in AVAILABLE_YEARS if y != target_year]
    before = sorted([y for y in candidates if y < target_year], reverse=True)
    after  = sorted([y for y in candidates if y > target_year])
    result = []
    if before: result.append(before[0])
    if after:  result.append(after[0])
    if len(result) < 2:
        if not before and len(after)  >= 2: result.append(after[1])
        if not after  and len(before) >= 2: result.append(before[1])
    return sorted(result)


def select_fallback_year(target_year: int) -> int:
    """데이터 부족 셀 폴백에 사용할 연도 반환.

    - 일반: 직전 연도 (target_year - 1)
    - target_year == DATA_START_YEAR: 직전에 데이터 없음 → 다음 연도 (target_year + 1)
    """
    if target_year > DATA_START_YEAR:
        return target_year - 1
    return target_year + 1


# ════════════════════════════════════════════════════════════════
# Iterative Cleaning & 베이스라인 구축
# ════════════════════════════════════════════════════════════════

def compute_baseline(values: list) -> dict | None:
    clean = [v for v in values if v is not None and v >= 0]
    for _ in range(3):
        if len(clean) < MIN_CLEAN_SAMPLES:
            break
        arr = np.array(clean, dtype=float)
        median = float(np.median(arr))
        q1, q3 = float(np.percentile(arr, 25)), float(np.percentile(arr, 75))
        iqr = q3 - q1
        lower = max(0.0, q1 - IQR_MULTIPLIER * iqr, median * MEDIAN_RATIO)
        new_clean = [v for v in clean if v >= lower]
        if len(new_clean) == len(clean):
            break
        clean = new_clean

    if len(clean) < MIN_CLEAN_SAMPLES:
        return None
    arr = np.array(clean, dtype=float)
    return {
        "median":  float(np.median(arr)),
        "q1":      float(np.percentile(arr, 25)),
        "q3":      float(np.percentile(arr, 75)),
        "n_clean": len(clean),
    }


def build_baselines(raw_rows: list, baseline_years: list, holiday_dates: set,
                    approaches: list, date_start: date, date_end: date,
                    hours: list, adj_month_rows: list = None,
                    fallback_rows: list = None,
                    expand_stages: bool = True) -> tuple:
    """
    대상 기간에서 필요한 셀만 베이스라인 계산.
    adj_month_rows: 대상 연도 인접 월 데이터 (샘플 보강용)
    fallback_rows:  데이터 부족 셀용 폴백 연도 전체 월 데이터
    반환: ({(acsr_id, day_type, hour, month): {zero_rate, stats, n_adj, n_fallback}},
           {"n_cells": int, "n_values": int})
    """
    # 필요한 셀 목록 구성
    needed_cells = set()
    cur_d = date_start
    while cur_d <= date_end:
        day_type = get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for acsr_id, _ in approaches:
                needed_cells.add((acsr_id, day_type, hour, month))
        cur_d += timedelta(days=1)

    # 베이스라인 데이터 셀별 분류
    cell_values = defaultdict(list)
    cell_actual_dates = defaultdict(set)
    cell_zero_counts: dict[tuple, int] = defaultdict(int)   # 베이스라인 내 0값 레코드 수
    # Stage 1/2 확장용 전체 월 조회 (needed_cells 필터 없음)
    all_month_values: dict = defaultdict(list)
    all_month_actual_dates: dict = defaultdict(set)

    for acsr_id, tot_dt, trf_qnty in raw_rows:
        d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
        h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
        month = d.month
        day_type = get_day_type(d, holiday_dates)
        cell = (acsr_id, day_type, h, month)
        # 전체 월 조회 (Stage 1/2용)
        if trf_qnty is not None:
            all_month_values[cell].append(trf_qnty)
        all_month_actual_dates[cell].add(d)
        # needed_cells 필터 조회 (Stage 0용)
        if cell in needed_cells:
            if trf_qnty is not None:
                cell_values[cell].append(trf_qnty)
                if trf_qnty == 0:
                    cell_zero_counts[cell] += 1
            cell_actual_dates[cell].add(d)

    # C형: 인접 월 데이터 병합 — 인접 월 값을 가장 가까운 대상 월 셀에 추가
    cell_adj_counts: dict[tuple, int] = defaultdict(int)
    if adj_month_rows:
        target_months = sorted({month for _, _, _, month in needed_cells})
        # (acsr_id, day_type, hour) 키가 needed_cells에 있는지 빠르게 확인
        needed_keys = {(a, dt, h) for a, dt, h, _ in needed_cells}
        for acsr_id, tot_dt, trf_qnty in adj_month_rows:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (acsr_id, day_type, h) not in needed_keys:
                continue
            adj_month = d.month
            nearest_month = min(target_months, key=lambda m: abs(m - adj_month))
            cell = (acsr_id, day_type, h, nearest_month)
            if cell in needed_cells:
                cell_values[cell].append(trf_qnty)
                cell_adj_counts[cell] += 1

    # 폴백: 1차+2차 후에도 데이터 부족 셀에 대해 폴백 연도 전체 월 병합
    cell_fallback_counts: dict[tuple, int] = defaultdict(int)
    if fallback_rows:
        fallback_cells = {
            cell for cell in needed_cells
            if len(cell_values.get(cell, [])) < MIN_CLEAN_SAMPLES
        }
        if fallback_cells:
            # (acsr_id, day_type, hour) → [trf_qnty, ...] (month 무관)
            fallback_by_key: dict[tuple, list] = defaultdict(list)
            for acsr_id, tot_dt, trf_qnty in fallback_rows:
                if trf_qnty is None:
                    continue
                d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
                h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
                day_type = get_day_type(d, holiday_dates)
                fallback_by_key[(acsr_id, day_type, h)].append((d, trf_qnty))

            for cell in fallback_cells:
                acsr_id, day_type, hour, month = cell
                already_seen = cell_actual_dates.get(cell, set())
                for d, val in fallback_by_key.get((acsr_id, day_type, hour), []):
                    if d not in already_seen:
                        cell_values[cell].append(val)
                        cell_fallback_counts[cell] += 1

    # 베이스라인 연도별 expected 날짜 수 계산
    expected_counts: dict[tuple, int] = defaultdict(int)
    today = date.today()
    for year in baseline_years:
        cur_d = date(year, 1, 1)
        end_y = date(year, 12, 31)
        while cur_d <= end_y:
            if cur_d >= today:
                cur_d += timedelta(days=1)
                continue
            day_type = get_day_type(cur_d, holiday_dates)
            month = cur_d.month
            for h in range(24):
                expected_counts[(day_type, h, month)] += 1
            cur_d += timedelta(days=1)

    # 셀별 베이스라인 계산 (Stage 0)
    baselines = {}
    for cell in needed_cells:
        acsr_id, day_type, hour, month = cell
        exp_count = expected_counts.get((day_type, hour, month), 0)
        if exp_count == 0:
            continue
        act_count = len(cell_actual_dates.get(cell, set()))
        zero_rate = (exp_count - act_count) / exp_count
        stats = compute_baseline(cell_values.get(cell, []))
        n_adj = cell_adj_counts.get(cell, 0)
        n_fallback = cell_fallback_counts.get(cell, 0)
        baselines[cell] = {
            "zero_rate":        zero_rate,
            "stage0_zr":        zero_rate,
            "stage1_zr":        None,
            "stage2_zr":        None,
            "stats":            stats,
            "n_adj":            n_adj,
            "n_fallback":       n_fallback,
            "expansion_stage":  0,
            "exp_count":        exp_count,
            "act_count":        act_count,
            "n_zero_baseline":  cell_zero_counts.get(cell, 0),
            "n_total_baseline": len(cell_values.get(cell, [])),
        }

    if expand_stages:
        # ── Stage 1/2 zero_rate 확장 ──────────────────────────────────
        # zero_rate >= ZERO_RATE_EXPANSION_THRESHOLD(0.8)인 셀에 대해
        # 인접 월(Stage 1) → 연간 전체(Stage 2) 순으로 범위를 확장하여
        # zero_rate 및 stats 재계산
        for cell, bl in baselines.items():
            if bl["zero_rate"] < ZERO_RATE_EXPANSION_THRESHOLD:
                continue

            acsr_id, day_type, hour, month = cell
            prev_m = ((month - 2) % 12) + 1   # 1월 → 12월, 12월 → 11월
            next_m = (month % 12) + 1          # 12월 → 1월, 1월 → 2월
            stage1_months = [prev_m, month, next_m]

            # Stage 1: ±1개월 확장
            exp_1 = sum(expected_counts.get((day_type, hour, m), 0) for m in stage1_months)
            if exp_1 > 0:
                act_dates_1: set = set()
                for m in stage1_months:
                    act_dates_1 |= all_month_actual_dates.get((acsr_id, day_type, hour, m), set())
                zr_1 = (exp_1 - len(act_dates_1)) / exp_1
                vals_1 = []
                for m in stage1_months:
                    vals_1.extend(all_month_values.get((acsr_id, day_type, hour, m), []))
                bl["zero_rate"]       = zr_1
                bl["stage1_zr"]       = zr_1
                bl["stats"]           = compute_baseline(vals_1)
                bl["expansion_stage"] = 1
                if zr_1 < ZERO_RATE_EXPANSION_THRESHOLD and bl["stats"] is not None:
                    continue

            # Stage 2: 연간 전체 확장
            exp_2 = sum(expected_counts.get((day_type, hour, m), 0) for m in range(1, 13))
            if exp_2 > 0:
                act_dates_2: set = set()
                for m in range(1, 13):
                    act_dates_2 |= all_month_actual_dates.get((acsr_id, day_type, hour, m), set())
                zr_2 = (exp_2 - len(act_dates_2)) / exp_2
                vals_2 = []
                for m in range(1, 13):
                    vals_2.extend(all_month_values.get((acsr_id, day_type, hour, m), []))
                bl["zero_rate"]       = zr_2
                bl["stage2_zr"]       = zr_2
                bl["stats"]           = compute_baseline(vals_2)
                bl["expansion_stage"] = 2

    fallback_applied = {c for c, v in cell_fallback_counts.items() if v > 0}
    n_stage1 = sum(1 for bl in baselines.values() if bl["expansion_stage"] >= 1)
    n_stage2 = sum(1 for bl in baselines.values() if bl["expansion_stage"] == 2)
    fallback_summary = {
        "n_cells":   len(fallback_applied),
        "n_values":  sum(cell_fallback_counts.values()),
        "n_stage1":  n_stage1,
        "n_stage2":  n_stage2,
    }
    return baselines, fallback_summary


# ════════════════════════════════════════════════════════════════
# 진단 출력
# ════════════════════════════════════════════════════════════════

def export_zero_rate_diagnostic(node_name: str, approaches: list, baselines: dict,
                                target_data: dict, date_start: date, date_end: date,
                                hours: list, holiday_dates: set, diag_dir: Path) -> None:
    """zero_rate 단계별 진단 CSV 2개 출력

    CSV 1 (셀별): acsr_nm/day_type/hour 단위 stage0/1/2 zero_rate 및 베이스라인 현황
    CSV 2 (슬롯별): 대상 기간 각 날짜×hour×접근로 슬롯의 레코드 상태 및 판정 가능 여부
    """
    import csv

    diag_dir.mkdir(parents=True, exist_ok=True)
    period = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    acsr_map = {acsr_id: acsr_nm for acsr_id, acsr_nm in approaches}
    day_type_label = {0: "평일", 1: "토요일", 2: "일·공휴일"}

    # ── CSV 1: 셀별 zero_rate ─────────────────────────────────
    csv1_path = diag_dir / f"진단_zero_rate_셀별_{node_name}_{period}.csv"
    headers1 = [
        "acsr_nm", "day_type", "hour",
        "stage0_zr", "stage1_zr", "stage2_zr", "final_stage",
        "exp_count", "act_count",
        "n_zero_baseline", "n_total_baseline", "baseline_median",
    ]
    rows1 = []
    for (acsr_id, day_type, hour, month), bl in sorted(baselines.items()):
        acsr_nm = acsr_map.get(acsr_id, str(acsr_id))
        stats = bl.get("stats")
        rows1.append({
            "acsr_nm":          acsr_nm,
            "day_type":         day_type_label.get(day_type, day_type),
            "hour":             f"{hour:02d}",
            "stage0_zr":        f"{bl['stage0_zr']:.4f}",
            "stage1_zr":        f"{bl['stage1_zr']:.4f}" if bl.get("stage1_zr") is not None else "-",
            "stage2_zr":        f"{bl['stage2_zr']:.4f}" if bl.get("stage2_zr") is not None else "-",
            "final_stage":      bl.get("expansion_stage", 0),
            "exp_count":        bl.get("exp_count", ""),
            "act_count":        bl.get("act_count", ""),
            "n_zero_baseline":  bl.get("n_zero_baseline", 0),
            "n_total_baseline": bl.get("n_total_baseline", 0),
            "baseline_median":  f"{stats['median']:.1f}" if stats else "-",
        })
    with open(csv1_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers1)
        writer.writeheader()
        writer.writerows(rows1)
    print(f"  진단 CSV(셀별) 저장: {csv1_path}")

    # ── CSV 2: 슬롯별 상태 ────────────────────────────────────
    csv2_path = diag_dir / f"진단_슬롯별_{node_name}_{period}.csv"
    headers2 = [
        "날짜", "요일유형", "hour", "acsr_nm",
        "레코드여부", "trf_qnty",
        "판정가능여부", "final_stage", "final_zero_rate",
    ]
    rows2 = []
    cur_d = date_start
    while cur_d <= date_end:
        day_type = get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for acsr_id, acsr_nm in approaches:
                cell = (acsr_id, day_type, hour, month)
                bl = baselines.get(cell)
                if bl is None:
                    continue
                key = (acsr_id, cur_d, hour)
                has_rec = key in target_data
                trf_val = target_data.get(key)
                zr = bl["zero_rate"]
                stage = bl.get("expansion_stage", 0)

                if has_rec:
                    rec_label = "있음(0값)" if trf_val == 0 else "있음(비0)"
                    judgment  = "B형 검토"
                else:
                    rec_label = "없음(NULL)"
                    trf_val   = None
                    if zr < ZERO_RATE_THRESHOLD:
                        judgment = "A형 가능"
                    elif stage == 2 and zr >= ZERO_RATE_EXPANSION_THRESHOLD:
                        judgment = "정상처리"
                    else:
                        judgment = "B형"

                rows2.append({
                    "날짜":          cur_d.strftime("%Y-%m-%d"),
                    "요일유형":      day_type_label.get(day_type, day_type),
                    "hour":          f"{hour:02d}",
                    "acsr_nm":       acsr_nm,
                    "레코드여부":    rec_label,
                    "trf_qnty":      trf_val if trf_val is not None else "",
                    "판정가능여부":  judgment,
                    "final_stage":   stage,
                    "final_zero_rate": f"{zr:.4f}",
                })
        cur_d += timedelta(days=1)

    with open(csv2_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers2)
        writer.writeheader()
        writer.writerows(rows2)
    print(f"  진단 CSV(슬롯별) 저장: {csv2_path}")


# ════════════════════════════════════════════════════════════════
# 이상 탐지
# ════════════════════════════════════════════════════════════════

def detect_anomalies(node_name: str, approaches: list, baselines: dict,
                     target_data: dict, date_start: date, date_end: date,
                     hours: list, holiday_dates: set) -> list:
    results = []
    cur_d = date_start
    while cur_d <= date_end:
        day_type = get_day_type(cur_d, holiday_dates)
        month = cur_d.month
        for hour in hours:
            for acsr_id, acsr_nm in approaches:
                cell = (acsr_id, day_type, hour, month)
                bl = baselines.get(cell)
                if bl is None:
                    continue

                zero_rate = bl["zero_rate"]
                stats = bl["stats"]
                has_record = (acsr_id, cur_d, hour) in target_data
                trf_qnty = target_data.get((acsr_id, cur_d, hour))

                anomaly_type = None
                trf_val = 0

                if not has_record:
                    # A형: 레코드 없음 + zero_rate < 임계값
                    if zero_rate < ZERO_RATE_THRESHOLD:
                        anomaly_type = "A형"
                        trf_val = 0
                    # 정상: zero_rate >= 0.8 (베이스라인에서 자주 비는 구간 → 정상)
                    elif zero_rate >= ZERO_RATE_EXPANSION_THRESHOLD:
                        anomaly_type = "정상"
                        trf_val = 0
                    # B형: 0.2 ≤ zero_rate < 0.8 (데이터 있어야 하는데 없음)
                    else:
                        anomaly_type = "B형"
                        trf_val = 0
                elif stats is not None:
                    # B형: 비정상 저값
                    median = stats["median"]
                    q1 = stats["q1"]
                    q3 = stats["q3"]
                    iqr = q3 - q1
                    lower_iqr = q1 - IQR_MULTIPLIER * iqr
                    ratio = (trf_qnty / median) if median > 0 else 1.0
                    if ratio < RATIO_THRESHOLD or trf_qnty < lower_iqr:
                        anomaly_type = "B형"
                        trf_val = trf_qnty

                if anomaly_type:
                    results.append({
                        "날짜":     cur_d.strftime("%Y.%m.%d"),
                        "시간":     f"{hour:02d}:00",
                        "교차로":   node_name,
                        "방향":     acsr_nm,
                        "교통량":   trf_val,
                        "판정":     anomaly_type,
                        "_acsr_id": acsr_id,
                        "_date":    cur_d,
                        "_hour":    hour,
                        "_cell":    cell,
                    })
        cur_d += timedelta(days=1)
    return results


# ════════════════════════════════════════════════════════════════
# 보정
# ════════════════════════════════════════════════════════════════

def apply_corrections(results: list, baselines: dict, target_data: dict) -> None:
    """results를 in-place로 수정: 보정값, 보정방법, 신뢰도 키 추가

    우선순위:
      1순위 선형보간: 같은 (acsr_id, hour) 시계열에서 앞뒤 정상값으로 보간
                      앞뒤 거리 합 <= MAX_INTERP_DAYS 조건 충족 시 적용
      2순위 median:   베이스라인 중앙값으로 대체
    """
    anomaly_set = {(r["_acsr_id"], r["_date"], r["_hour"]) for r in results}

    for r in results:
        if r["판정"] == "정상":
            r["보정값"]   = None
            r["보정방법"] = None
            r["신뢰도"]   = None
            continue

        acsr_id = r["_acsr_id"]
        cur_d   = r["_date"]
        hour    = r["_hour"]
        cell    = r["_cell"]

        bl    = baselines.get(cell)
        stats = bl["stats"] if bl else None

        # 이전 정상값 탐색 (최대 MAX_INTERP_DAYS일 이전까지)
        prev_val, d_before = None, None
        for d in range(1, MAX_INTERP_DAYS + 1):
            key = (acsr_id, cur_d - timedelta(days=d), hour)
            if key in anomaly_set:
                continue
            if key in target_data:
                prev_val = target_data[key]
                d_before = d
                break

        # 이후 정상값 탐색 (최대 MAX_INTERP_DAYS일 이후까지)
        next_val, d_after = None, None
        for d in range(1, MAX_INTERP_DAYS + 1):
            key = (acsr_id, cur_d + timedelta(days=d), hour)
            if key in anomaly_set:
                continue
            if key in target_data:
                next_val = target_data[key]
                d_after = d
                break

        corrected, method = None, None
        interp_dist = None
        if (prev_val is not None and next_val is not None
                and d_before + d_after <= MAX_INTERP_DAYS):
            corrected = prev_val + (next_val - prev_val) * d_before / (d_before + d_after)
            method = "선형보간"
            interp_dist = d_before + d_after
        elif stats is not None:
            corrected = stats["median"]
            method = "median"

        if corrected is not None:
            r["보정값"]   = round(corrected)
            r["보정방법"] = method
            is_low = (
                (stats and stats["n_clean"] < MIN_CONFIDENCE_SAMPLES)
                or bl.get("n_fallback", 0) > 0
                or (interp_dist is not None and interp_dist > MAX_INTERP_DAYS_OK)
            )
            r["신뢰도"]   = "LOW" if is_low else "OK"
        else:
            r["보정값"]   = None
            r["보정방법"] = None
            r["신뢰도"]   = None


# ════════════════════════════════════════════════════════════════
# 단일 교차로 분석 (API·CLI 공용)
# ════════════════════════════════════════════════════════════════

# Core calculation functions are sourced from anomaly_core.
# Keep these names in this module for backward compatibility.
get_day_type = core.get_day_type
select_baseline_years = core.select_baseline_years
select_fallback_year = core.select_fallback_year
compute_baseline = core.compute_baseline
build_baselines = core.build_baselines
detect_anomalies = core.detect_anomalies
apply_corrections = core.apply_corrections


def analyse_node(
    conn,
    node_id: int,
    node_name: str,
    date_start: date,
    date_end: date,
    hours: list,
    baseline_years: list,
    fallback_year: int,
    adj_periods: list,
    holiday_dates: set,
) -> tuple[list, dict, dict]:
    """단일 교차로 이상탐지+보정 수행. I/O 없이 순수 계산만 수행.

    Returns:
        (node_results, baselines, target_data)
        - node_results : detect_anomalies + apply_corrections 적용된 슬롯 리스트
        - baselines    : build_baselines 결과 (시각화·진단에서 재사용 가능)
        - target_data  : 대상 기간 원본 데이터 (시각화에서 재사용 가능)
        접근로 정보가 없으면 ([], {}, {}) 반환.
    """
    approaches = load_approaches(conn, node_id)
    if not approaches:
        return [], {}, {}
    cache_h0, cache_m0, _ = _baseline_cache_stats()

    # 대상 월 및 Stage1(±1개월) 월 집합
    target_months: set[int] = set()
    cur_m = date_start.replace(day=1)
    while cur_m <= date_end:
        target_months.add(cur_m.month)
        cur_m = date(cur_m.year + 1, 1, 1) if cur_m.month == 12 else date(cur_m.year, cur_m.month + 1, 1)

    stage1_months: set[int] = set(target_months)
    for m in target_months:
        stage1_months.add(((m - 2) % 12) + 1)  # prev
        stage1_months.add((m % 12) + 1)         # next

    approach_ids = [acsr_id for acsr_id, _ in approaches]

    t0 = time.perf_counter()
    raw_rows_core = load_baseline_data(
        conn,
        node_id,
        baseline_years,
        hours=hours,
        months=sorted(stage1_months),
        acsr_ids=approach_ids,
    )
    t_raw_core = time.perf_counter() - t0

    t1 = time.perf_counter()
    adj_rows = load_adjacent_month_data(conn, node_id, adj_periods, hours=hours)
    t_adj = time.perf_counter() - t1

    t2 = time.perf_counter()
    baselines_pre, _ = build_baselines(
        raw_rows_core, baseline_years, holiday_dates,
        approaches, date_start, date_end, hours,
        adj_month_rows=adj_rows,
        fallback_rows=None,
        expand_stages=False,
    )
    t_pre = time.perf_counter() - t2

    stage2_keys = {
        (acsr_id, day_type, hour)
        for (acsr_id, day_type, hour, _), bl in baselines_pre.items()
        if bl["zero_rate"] >= ZERO_RATE_EXPANSION_THRESHOLD
    }

    t3 = time.perf_counter()
    raw_rows_stage2: list = []
    raw_rows_stage2_fetched = 0
    if stage2_keys:
        stage2_acsr_ids = sorted({acsr_id for acsr_id, _, _ in stage2_keys})
        annual_rows = load_baseline_data(
            conn,
            node_id,
            baseline_years,
            hours=hours,
            months=None,
            acsr_ids=stage2_acsr_ids,
        )
        raw_rows_stage2_fetched = len(annual_rows)
        for acsr_id, tot_dt, trf_qnty in annual_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (acsr_id, day_type, h) in stage2_keys:
                raw_rows_stage2.append((acsr_id, tot_dt, trf_qnty))
    t_stage2 = time.perf_counter() - t3

    # core + stage2 rows 병합 (중복 제거)
    seen_positions: set[tuple[int, date, int]] = set()
    raw_rows: list = []
    for source_rows in (raw_rows_core, raw_rows_stage2):
        for acsr_id, tot_dt, trf_qnty in source_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            key = (acsr_id, d, h)
            if key in seen_positions:
                continue
            seen_positions.add(key)
            raw_rows.append((acsr_id, tot_dt, trf_qnty))

    fallback_keys = {
        (acsr_id, day_type, hour)
        for (acsr_id, day_type, hour, _), bl in baselines_pre.items()
        if bl.get("n_total_baseline", 0) < MIN_CLEAN_SAMPLES
    }

    t4 = time.perf_counter()
    fallback_raw_rows: list = []
    fallback_raw_rows_fetched = 0
    if fallback_keys:
        fallback_acsr_ids = sorted({acsr_id for acsr_id, _, _ in fallback_keys})
        fallback_rows_all = load_baseline_data(
            conn,
            node_id,
            [fallback_year],
            hours=hours,
            months=None,
            acsr_ids=fallback_acsr_ids,
        )
        fallback_raw_rows_fetched = len(fallback_rows_all)
        for acsr_id, tot_dt, trf_qnty in fallback_rows_all:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (acsr_id, day_type, h) in fallback_keys:
                fallback_raw_rows.append((acsr_id, tot_dt, trf_qnty))
    t_fallback = time.perf_counter() - t4
    t_db = t_raw_core + t_adj + t_stage2 + t_fallback

    t5 = time.perf_counter()
    baselines, fallback_summary = build_baselines(
        raw_rows, baseline_years, holiday_dates,
        approaches, date_start, date_end, hours,
        adj_month_rows=adj_rows,
        fallback_rows=fallback_raw_rows,
    )
    t_baselines = (time.perf_counter() - t5) + t_pre

    t6 = time.perf_counter()
    target_data = load_target_data(conn, node_id, date_start, date_end, hours)
    t_target = time.perf_counter() - t6

    node_results = detect_anomalies(
        node_name, approaches, baselines, target_data,
        date_start, date_end, hours, holiday_dates,
    )
    apply_corrections(node_results, baselines, target_data)

    n_cells = len(baselines)
    n_s1 = fallback_summary.get("n_stage1", 0)
    n_s2 = fallback_summary.get("n_stage2", 0)
    print(
        f"[PROFILE] {node_name} | "
        f"DB쿼리={t_db:.3f}s | baselines={t_baselines:.3f}s | target_data={t_target:.3f}s | "
        f"cells={n_cells} stage1={n_s1} stage2={n_s2}"
    )

    if PROFILE_DB_BREAKDOWN:
        cache_h1, cache_m1, cache_size = _baseline_cache_stats()
        print(
            f"[PROFILE_DB] {node_name} | "
            f"baseline_core={t_raw_core:.3f}s/{len(raw_rows_core):,}rows | "
            f"adj={t_adj:.3f}s/{len(adj_rows):,}rows | "
            f"stage2={t_stage2:.3f}s/{len(raw_rows_stage2):,}rows(filtered:{raw_rows_stage2_fetched:,}) | "
            f"fallback={t_fallback:.3f}s/{len(fallback_raw_rows):,}rows(filtered:{fallback_raw_rows_fetched:,}) | "
            f"target={t_target:.3f}s/{len(target_data):,}rows | "
            f"period={date_start}~{date_end} hours={len(hours)} "
            f"approaches={len(approaches)} baseline_years={baseline_years} "
            f"fallback_year={fallback_year} adj_months={len(adj_periods)} "
            f"stage2_keys={len(stage2_keys)} fallback_keys={len(fallback_keys)} "
            f"cache(hits={cache_h1-cache_h0}, misses={cache_m1-cache_m0}, size={cache_size})"
        )

    return node_results, baselines, target_data


def analyse_node_drct(
    conn,
    node_id: int,
    node_name: str,
    date_start: date,
    date_end: date,
    hours: list,
    baseline_years: list,
    fallback_year: int,
    adj_periods: list,
    holiday_dates: set,
) -> tuple[list, dict, dict, dict]:
    """DRCT 단위 이상탐지+보정 수행.

    Returns:
      (node_results, baselines, target_data, drct_meta)
      - node_results : detect_anomalies + apply_corrections 적용 슬롯
      - baselines    : DRCT 단위 베이스라인
      - target_data  : {((acsr_id, drct_cd), date, hour): trf_qnty}
      - drct_meta    : {(acsr_id, drct_cd): {approach_id, approach_name, drct_cd, drct_name}}
    """
    drct_code_map = load_drct_code_map(conn)
    approaches, drct_meta = load_drct_approaches(conn, node_id, drct_code_map=drct_code_map)
    if not approaches:
        return [], {}, {}, {}
    cache_h0, cache_m0, _ = _baseline_cache_stats()

    target_months: set[int] = set()
    cur_m = date_start.replace(day=1)
    while cur_m <= date_end:
        target_months.add(cur_m.month)
        cur_m = date(cur_m.year + 1, 1, 1) if cur_m.month == 12 else date(cur_m.year, cur_m.month + 1, 1)

    stage1_months: set[int] = set(target_months)
    for m in target_months:
        stage1_months.add(((m - 2) % 12) + 1)  # prev
        stage1_months.add((m % 12) + 1)         # next

    drct_keys = [drct_key for drct_key, _ in approaches]

    t0 = time.perf_counter()
    raw_rows_core = load_drct_baseline_data(
        conn,
        node_id,
        baseline_years,
        hours=hours,
        months=sorted(stage1_months),
        drct_keys=drct_keys,
    )
    t_raw_core = time.perf_counter() - t0

    t1 = time.perf_counter()
    adj_rows = load_drct_adjacent_month_data(conn, node_id, adj_periods, hours=hours, drct_keys=drct_keys)
    t_adj = time.perf_counter() - t1

    t2 = time.perf_counter()
    baselines_pre, _ = build_baselines(
        raw_rows_core, baseline_years, holiday_dates,
        approaches, date_start, date_end, hours,
        adj_month_rows=adj_rows,
        fallback_rows=None,
        expand_stages=False,
    )
    t_pre = time.perf_counter() - t2

    stage2_keys = {
        (drct_key, day_type, hour)
        for (drct_key, day_type, hour, _), bl in baselines_pre.items()
        if bl["zero_rate"] >= ZERO_RATE_EXPANSION_THRESHOLD
    }

    t3 = time.perf_counter()
    raw_rows_stage2: list = []
    raw_rows_stage2_fetched = 0
    if stage2_keys:
        stage2_drct_keys = sorted(
            {drct_key for drct_key, _, _ in stage2_keys},
            key=lambda x: (str(x[0]), str(x[1])),
        )
        annual_rows = load_drct_baseline_data(
            conn,
            node_id,
            baseline_years,
            hours=hours,
            months=None,
            drct_keys=stage2_drct_keys,
        )
        raw_rows_stage2_fetched = len(annual_rows)
        for drct_key, tot_dt, trf_qnty in annual_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (drct_key, day_type, h) in stage2_keys:
                raw_rows_stage2.append((drct_key, tot_dt, trf_qnty))
    t_stage2 = time.perf_counter() - t3

    seen_positions: set[tuple[tuple, date, int]] = set()
    raw_rows: list = []
    for source_rows in (raw_rows_core, raw_rows_stage2):
        for drct_key, tot_dt, trf_qnty in source_rows:
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            key = (drct_key, d, h)
            if key in seen_positions:
                continue
            seen_positions.add(key)
            raw_rows.append((drct_key, tot_dt, trf_qnty))

    fallback_keys = {
        (drct_key, day_type, hour)
        for (drct_key, day_type, hour, _), bl in baselines_pre.items()
        if bl.get("n_total_baseline", 0) < MIN_CLEAN_SAMPLES
    }

    t4 = time.perf_counter()
    fallback_raw_rows: list = []
    fallback_raw_rows_fetched = 0
    if fallback_keys:
        fallback_drct_keys = sorted(
            {drct_key for drct_key, _, _ in fallback_keys},
            key=lambda x: (str(x[0]), str(x[1])),
        )
        fallback_rows_all = load_drct_baseline_data(
            conn,
            node_id,
            [fallback_year],
            hours=hours,
            months=None,
            drct_keys=fallback_drct_keys,
        )
        fallback_raw_rows_fetched = len(fallback_rows_all)
        for drct_key, tot_dt, trf_qnty in fallback_rows_all:
            if trf_qnty is None:
                continue
            d = tot_dt.date() if isinstance(tot_dt, datetime) else tot_dt
            h = tot_dt.hour if isinstance(tot_dt, datetime) else 0
            day_type = get_day_type(d, holiday_dates)
            if (drct_key, day_type, h) in fallback_keys:
                fallback_raw_rows.append((drct_key, tot_dt, trf_qnty))
    t_fallback = time.perf_counter() - t4
    t_db = t_raw_core + t_adj + t_stage2 + t_fallback

    t5 = time.perf_counter()
    baselines, fallback_summary = build_baselines(
        raw_rows, baseline_years, holiday_dates,
        approaches, date_start, date_end, hours,
        adj_month_rows=adj_rows,
        fallback_rows=fallback_raw_rows,
    )
    t_baselines = (time.perf_counter() - t5) + t_pre

    t6 = time.perf_counter()
    target_data = load_drct_target_data(conn, node_id, date_start, date_end, hours)
    t_target = time.perf_counter() - t6

    node_results = detect_anomalies(
        node_name, approaches, baselines, target_data,
        date_start, date_end, hours, holiday_dates,
    )
    apply_corrections(node_results, baselines, target_data)

    n_cells = len(baselines)
    n_s1 = fallback_summary.get("n_stage1", 0)
    n_s2 = fallback_summary.get("n_stage2", 0)
    print(
        f"[PROFILE][DRCT] {node_name} | "
        f"DB쿼리={t_db:.3f}s | baselines={t_baselines:.3f}s | target_data={t_target:.3f}s | "
        f"cells={n_cells} stage1={n_s1} stage2={n_s2}"
    )

    if PROFILE_DB_BREAKDOWN:
        cache_h1, cache_m1, cache_size = _baseline_cache_stats()
        print(
            f"[PROFILE_DB][DRCT] {node_name} | "
            f"baseline_core={t_raw_core:.3f}s/{len(raw_rows_core):,}rows | "
            f"adj={t_adj:.3f}s/{len(adj_rows):,}rows | "
            f"stage2={t_stage2:.3f}s/{len(raw_rows_stage2):,}rows(filtered:{raw_rows_stage2_fetched:,}) | "
            f"fallback={t_fallback:.3f}s/{len(fallback_raw_rows):,}rows(filtered:{fallback_raw_rows_fetched:,}) | "
            f"target={t_target:.3f}s/{len(target_data):,}rows | "
            f"period={date_start}~{date_end} hours={len(hours)} "
            f"drct_keys={len(drct_keys)} baseline_years={baseline_years} "
            f"fallback_year={fallback_year} adj_months={len(adj_periods)} "
            f"stage2_keys={len(stage2_keys)} fallback_keys={len(fallback_keys)} "
            f"cache(hits={cache_h1-cache_h0}, misses={cache_m1-cache_m0}, size={cache_size})"
        )

    return node_results, baselines, target_data, drct_meta


# ════════════════════════════════════════════════════════════════
# 시각화
# ════════════════════════════════════════════════════════════════

def _write_plotly_html(fig, out_path: Path, verbose: bool = True) -> None:
    plotly_div = fig.to_html(include_plotlyjs="cdn", full_html=False)
    html = (
        "<!DOCTYPE html>\n<html>\n<head><meta charset=\"utf-8\">\n"
        "<style>body{margin:0;} .scroll-wrap{overflow-x:auto; width:100%;}</style>\n"
        "</head>\n<body>\n"
        f"<div class=\"scroll-wrap\">{plotly_div}</div>\n"
        "</body>\n</html>"
    )
    out_path.write_text(html, encoding="utf-8")
    if verbose:
        print(f"  시각화 저장: {out_path}")


def export_acsr_visualizations(
    node_name: str,
    approaches: list,
    target_data: dict,
    node_results: list,
    date_start: date,
    date_end: date,
    hours: list,
    viz_dir: Path,
    progress_callback=None,
    verbose: bool = True,
) -> None:
    """ACSR 기준 시각화 생성.

    - 교차로 전체 교통량(보정 전/후) 1개
    - 교차로 방향별 개별 그래프(보정 전/후) N개
    """
    period = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    sub_dir = viz_dir / _sanitize_path_component(f"{node_name}_{period}")
    sub_dir.mkdir(parents=True, exist_ok=True)
    hours_sorted = sorted(hours)

    total_tasks = len(approaches) + 1
    if progress_callback:
        progress_callback(0, total_tasks, "준비 중")

    anomaly_map = {(r["_acsr_id"], r["_date"], r["_hour"]): r for r in node_results}

    x_all, y_orig, y_corr = [], [], []
    a_x, a_y = [], []
    b_x, b_y = [], []

    cur_d = date_start
    while cur_d <= date_end:
        for hour in hours_sorted:
            dt = datetime(cur_d.year, cur_d.month, cur_d.day, hour)
            total_orig = 0
            total_corr = 0
            has_a = False
            has_b = False
            for acsr_id, _ in approaches:
                pos = (acsr_id, cur_d, hour)
                if pos in anomaly_map:
                    r = anomaly_map[pos]
                    orig_val = r["교통량"]
                    corr_val = r["보정값"] if r["보정값"] is not None else orig_val
                    anomaly_type = r["판정"]
                    if anomaly_type in ("A형", "A+B혼합"):
                        has_a = True
                    if anomaly_type in ("B형", "A+B혼합"):
                        has_b = True
                else:
                    orig_val = target_data.get(pos, 0) or 0
                    corr_val = orig_val
                total_orig += orig_val
                total_corr += corr_val

            x_all.append(dt)
            y_orig.append(total_orig)
            y_corr.append(total_corr)
            if has_a:
                a_x.append(dt)
                a_y.append(total_orig)
            if has_b:
                b_x.append(dt)
                b_y.append(total_orig)
        cur_d += timedelta(days=1)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x_all, y=y_orig, mode="lines", name="보정 전",
        line=dict(color="#4C72B0", width=1.5, dash="dot", shape="spline"),
        connectgaps=False,
    ))
    if a_x:
        fig.add_trace(go.Scatter(
            x=a_x, y=a_y, mode="markers", name="A형(결측)",
            marker=dict(color="red", symbol="triangle-up", size=10),
        ))
    if b_x:
        fig.add_trace(go.Scatter(
            x=b_x, y=b_y, mode="markers", name="B형(이상저값)",
            marker=dict(color="orange", symbol="diamond", size=10),
        ))
    fig.add_trace(go.Scatter(
        x=x_all, y=y_corr, mode="lines", name="보정 후",
        line=dict(color="#DD4949", width=2, shape="spline"),
        connectgaps=False,
    ))
    fig.update_layout(
        title=f"{node_name} — 전체 교통량 보정 전/후",
        xaxis_title="일시",
        yaxis_title="교통량 (대/시)",
        hovermode="x unified",
        width=max(1200, len(x_all) * 6),
        height=675,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    _write_plotly_html(fig, sub_dir / "_전체.html", verbose=verbose)
    done = 1
    if progress_callback:
        progress_callback(done, total_tasks, "_전체")

    for acsr_id, acsr_nm in approaches:
        corrections_by_pos = {
            (r["_acsr_id"], r["_date"], r["_hour"]): r
            for r in node_results if r["_acsr_id"] == acsr_id
        }
        export_visualization(
            node_name=node_name,
            acsr_id=acsr_id,
            acsr_nm=acsr_nm,
            target_data=target_data,
            corrections_by_pos=corrections_by_pos,
            date_start=date_start,
            date_end=date_end,
            hours=hours_sorted,
            viz_dir=viz_dir,
            verbose=verbose,
        )
        done += 1
        if progress_callback:
            progress_callback(done, total_tasks, acsr_nm)


def export_drct_visualizations(
    node_name: str,
    drct_options_by_acsr: dict,
    drct_selection_by_acsr: dict,
    target_data: dict,
    node_results: list,
    date_start: date,
    date_end: date,
    hours: list,
    viz_dir: Path,
    progress_callback=None,
    verbose: bool = True,
) -> None:
    """DRCT 전용 시각화 생성.

    경로:
      02_Result/보정_시각화/<교차로명_기간>/<교차로방향>/
        - <drct_name>.html
        - _전체.html
    """
    period = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    node_dir = viz_dir / _sanitize_path_component(f"{node_name}_{period}")
    node_dir.mkdir(parents=True, exist_ok=True)
    hours_sorted = sorted(hours)

    anomaly_map = {
        (r["_acsr_id"], r["_drct_cd"], r["_date"], r["_hour"]): r
        for r in node_results
    }

    tasks_by_approach: list[tuple] = []
    total_tasks = 0
    for acsr_id, info in drct_options_by_acsr.items():
        selected_codes = drct_selection_by_acsr.get(acsr_id, set())
        selected_drcts = [
            (drct_cd, drct_name)
            for drct_cd, drct_name in info.get("drcts", [])
            if drct_cd in selected_codes
        ]
        if not selected_drcts:
            continue
        tasks_by_approach.append((acsr_id, info, selected_drcts))
        total_tasks += len(selected_drcts) + 1  # 개별 + 전체

    generated_tasks = 0
    if progress_callback and total_tasks == 0:
        progress_callback(1, 1, "생성 대상 없음")
        return
    if progress_callback:
        progress_callback(0, total_tasks, "준비 중")

    for acsr_id, info, selected_drcts in tasks_by_approach:
        approach_name = info.get("approach_name") or str(acsr_id)
        approach_dir = node_dir / _sanitize_path_component(approach_name)
        approach_dir.mkdir(parents=True, exist_ok=True)

        drct_series: list[dict] = []
        x_all_common = []
        for drct_cd, drct_name in selected_drcts:
            x_all, y_orig, y_corr = [], [], []
            a_x, a_y = [], []
            b_x, b_y = [], []

            cur_d = date_start
            while cur_d <= date_end:
                for hour in hours_sorted:
                    dt = datetime(cur_d.year, cur_d.month, cur_d.day, hour)
                    pos = (acsr_id, drct_cd, cur_d, hour)
                    anomaly = anomaly_map.get(pos)
                    orig_val = target_data.get(pos)
                    corr_val = orig_val
                    if anomaly:
                        corr_val = anomaly.get("보정값")
                        if corr_val is None:
                            corr_val = orig_val
                        if anomaly["판정"] == "A형":
                            a_x.append(dt)
                            a_y.append(orig_val)
                        elif anomaly["판정"] == "B형":
                            b_x.append(dt)
                            b_y.append(orig_val)
                    x_all.append(dt)
                    y_orig.append(orig_val)
                    y_corr.append(corr_val)
                cur_d += timedelta(days=1)

            x_all_common = x_all
            drct_series.append({
                "name": drct_name,
                "x": x_all,
                "y_orig": y_orig,
                "y_corr": y_corr,
            })

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=x_all, y=y_orig, mode="lines", name="보정 전",
                line=dict(color="#4C72B0", width=1.5, dash="dot", shape="spline"),
                connectgaps=False,
            ))
            if a_x:
                fig.add_trace(go.Scatter(
                    x=a_x, y=a_y, mode="markers", name="A형(결측)",
                    marker=dict(color="red", symbol="triangle-up", size=9),
                ))
            if b_x:
                fig.add_trace(go.Scatter(
                    x=b_x, y=b_y, mode="markers", name="B형(이상저값)",
                    marker=dict(color="orange", symbol="diamond", size=9),
                ))
            fig.add_trace(go.Scatter(
                x=x_all, y=y_corr, mode="lines", name="보정 후",
                line=dict(color="#DD4949", width=2, shape="spline"),
                connectgaps=False,
            ))

            fig.update_layout(
                title=f"{node_name} — {approach_name} — {drct_name}",
                xaxis_title="일시",
                yaxis_title="교통량 (대/시)",
                hovermode="x unified",
                width=max(1200, len(x_all) * 6),
                height=675,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            out_file = approach_dir / f"{_sanitize_path_component(drct_name)}.html"
            _write_plotly_html(fig, out_file, verbose=verbose)
            generated_tasks += 1
            if progress_callback:
                progress_callback(generated_tasks, total_tasks, f"{approach_name}/{drct_name}")

        y_total_orig = []
        y_total_corr = []
        for idx in range(len(x_all_common)):
            sum_orig = 0
            sum_corr = 0
            for series in drct_series:
                v_orig = series["y_orig"][idx]
                v_corr = series["y_corr"][idx]
                sum_orig += (v_orig if v_orig is not None else 0)
                sum_corr += (v_corr if v_corr is not None else 0)
            y_total_orig.append(sum_orig)
            y_total_corr.append(sum_corr)

        fig_total = go.Figure()
        fig_total.add_trace(go.Scatter(
            x=x_all_common, y=y_total_orig, mode="lines", name="보정 전",
            line=dict(color="#4C72B0", width=1.5, dash="dot", shape="spline"),
            connectgaps=False,
        ))
        fig_total.add_trace(go.Scatter(
            x=x_all_common, y=y_total_corr, mode="lines", name="보정 후",
            line=dict(color="#DD4949", width=2, shape="spline"),
            connectgaps=False,
        ))
        fig_total.update_layout(
            title=f"{node_name} — {approach_name} — 전체 교통량 보정 전/후",
            xaxis_title="일시",
            yaxis_title="교통량 (대/시)",
            hovermode="x unified",
            width=max(1200, len(x_all_common) * 6),
            height=675,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        _write_plotly_html(fig_total, approach_dir / "_전체.html", verbose=verbose)
        generated_tasks += 1
        if progress_callback:
            progress_callback(generated_tasks, total_tasks, f"{approach_name}/_전체")


def export_visualization(node_name: str, acsr_id, acsr_nm: str,
                          target_data: dict, corrections_by_pos: dict,
                          date_start: date, date_end: date, hours: list,
                          viz_dir: Path, verbose: bool = True) -> None:
    """접근로 하나의 보정 전/후 시계열 HTML 파일 생성"""
    x_all, y_orig, y_corr = [], [], []
    a_x, a_y = [], []  # A형 이상값 마커
    b_x, b_y = [], []  # B형 이상값 마커

    cur_d = date_start
    while cur_d <= date_end:
        for hour in sorted(hours):
            dt  = datetime(cur_d.year, cur_d.month, cur_d.day, hour)
            pos = (acsr_id, cur_d, hour)

            if pos in corrections_by_pos:
                r        = corrections_by_pos[pos]
                orig_val = r["교통량"]
                corr_val = r["보정값"] if r["보정값"] is not None else orig_val
                x_all.append(dt)
                y_orig.append(orig_val)
                y_corr.append(corr_val)
                anomaly_type = r["판정"]
                if anomaly_type in ("A형", "A+B혼합"):
                    a_x.append(dt)
                    a_y.append(orig_val)
                if anomaly_type in ("B형", "A+B혼합"):
                    b_x.append(dt)
                    b_y.append(orig_val)
            else:
                val = target_data.get(pos)  # None이면 plotly에서 gap 처리
                x_all.append(dt)
                y_orig.append(val)
                y_corr.append(val)
        cur_d += timedelta(days=1)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x_all, y=y_orig, mode="lines", name="보정 전",
        line=dict(color="#4C72B0", width=1.5, dash="dot", shape="spline"),
        connectgaps=False,
    ))
    if a_x:
        fig.add_trace(go.Scatter(
            x=a_x, y=a_y, mode="markers", name="A형(결측)",
            marker=dict(color="red", symbol="triangle-up", size=10),
        ))
    if b_x:
        fig.add_trace(go.Scatter(
            x=b_x, y=b_y, mode="markers", name="B형(이상저값)",
            marker=dict(color="orange", symbol="diamond", size=10),
        ))
    fig.add_trace(go.Scatter(
        x=x_all, y=y_corr, mode="lines", name="보정 후",
        line=dict(color="#DD4949", width=2, shape="spline"),
        connectgaps=False,
    ))

    period = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    chart_width = max(1200, len(x_all) * 6)
    fig.update_layout(
        title=f"{node_name} — {acsr_nm}  보정 전/후",
        xaxis_title="일시",
        yaxis_title="교통량 (대/시)",
        hovermode="x unified",
        width=chart_width,
        height=675,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )

    sub_dir = viz_dir / f"{node_name}_{period}"
    sub_dir.mkdir(parents=True, exist_ok=True)
    out_path = sub_dir / f"{acsr_nm}.html"
    plotly_div = fig.to_html(include_plotlyjs="cdn", full_html=False)
    html = (
        "<!DOCTYPE html>\n<html>\n<head><meta charset=\"utf-8\">\n"
        "<style>body{margin:0;} .scroll-wrap{overflow-x:auto; width:100%;}</style>\n"
        "</head>\n<body>\n"
        f"<div class=\"scroll-wrap\">{plotly_div}</div>\n"
        "</body>\n</html>"
    )
    out_path.write_text(html, encoding="utf-8")
    if verbose:
        print(f"  시각화 저장: {out_path}")


def export_total_visualization(node_name: str, approaches: list,
                               target_data: dict, node_results: list,
                               date_start: date, date_end: date, hours: list,
                               viz_dir: Path) -> None:
    """교차로 전체 방향 합산 + 방향별 개별 보정 전/후 시계열 HTML 파일 생성 (2개 그래프)"""
    PALETTE = [
        "#E63946",  # 빨강
        "#F4A261",  # 주황
        "#2A9D8F",  # 초록
        "#457B9D",  # 파랑
        "#9B5DE5",  # 보라
        "#F15BB5",  # 분홍
        "#E9C46A",  # 노랑
        "#00BBF9",  # 하늘
    ]

    # 이상값 딕셔너리: {(acsr_id, date, hour): result_dict}
    anomaly_map = {
        (r["_acsr_id"], r["_date"], r["_hour"]): r for r in node_results
    }

    # ── 그래프 1용: 방향별 합산 데이터 수집 ──────────────────────────────────
    x_all, y_orig, y_corr = [], [], []
    a_x, a_y = [], []
    b_x, b_y = [], []

    cur_d = date_start
    while cur_d <= date_end:
        for hour in sorted(hours):
            dt = datetime(cur_d.year, cur_d.month, cur_d.day, hour)
            total_orig, total_corr = 0, 0
            has_a = False
            has_b = False

            for acsr_id, _ in approaches:
                pos = (acsr_id, cur_d, hour)
                if pos in anomaly_map:
                    r = anomaly_map[pos]
                    total_orig += r["교통량"]
                    total_corr += r["보정값"] if r["보정값"] is not None else r["교통량"]
                    anomaly_type = r["판정"]
                    if anomaly_type in ("A형", "A+B혼합"):
                        has_a = True
                    if anomaly_type in ("B형", "A+B혼합"):
                        has_b = True
                else:
                    val = target_data.get(pos, 0) or 0
                    total_orig += val
                    total_corr += val

            x_all.append(dt)
            y_orig.append(total_orig)
            y_corr.append(total_corr)

            if has_a:
                a_x.append(dt); a_y.append(total_orig)
            if has_b:
                b_x.append(dt); b_y.append(total_orig)

        cur_d += timedelta(days=1)

    # ── 그래프 2용: 접근로별 개별 데이터 수집 ───────────────────────────────
    acsr_data: dict[str, dict] = {}  # acsr_nm → {x, y_orig, y_corr}
    for acsr_id, acsr_nm in approaches:
        xs, ys_orig, ys_corr = [], [], []
        cur_d = date_start
        while cur_d <= date_end:
            for hour in sorted(hours):
                dt = datetime(cur_d.year, cur_d.month, cur_d.day, hour)
                pos = (acsr_id, cur_d, hour)
                if pos in anomaly_map:
                    r = anomaly_map[pos]
                    orig_val = r["교통량"]
                    corr_val = r["보정값"] if r["보정값"] is not None else orig_val
                else:
                    orig_val = target_data.get(pos)
                    corr_val = orig_val
                xs.append(dt)
                ys_orig.append(orig_val)
                ys_corr.append(corr_val)
            cur_d += timedelta(days=1)
        acsr_data[acsr_nm] = {"x": xs, "y_orig": ys_orig, "y_corr": ys_corr}

    # ── 서브플롯 생성 ────────────────────────────────────────────────────────
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        subplot_titles=(
            f"{node_name} — 방향별 합산 교통량",
            f"{node_name} — 방향별 교통량",
        ),
    )

    # 그래프 1: 합산 보정 전/후
    fig.add_trace(go.Scatter(
        x=x_all, y=y_orig, mode="lines", name="합산 보정 전",
        line=dict(color="#4C72B0", width=1.5, dash="dot", shape="spline"),
        connectgaps=False,
    ), row=1, col=1)
    if a_x:
        fig.add_trace(go.Scatter(
            x=a_x, y=a_y, mode="markers", name="A형(결측)",
            marker=dict(color="red", symbol="triangle-up", size=10),
            showlegend=True,
        ), row=1, col=1)
    if b_x:
        fig.add_trace(go.Scatter(
            x=b_x, y=b_y, mode="markers", name="B형(이상저값)",
            marker=dict(color="orange", symbol="diamond", size=10),
            showlegend=True,
        ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=x_all, y=y_corr, mode="lines", name="합산 보정 후",
        line=dict(color="#DD4949", width=2, shape="spline"),
        connectgaps=False,
    ), row=1, col=1)

    # 그래프 2: 접근로별 보정 전/후
    for idx, (acsr_nm, data) in enumerate(acsr_data.items()):
        color = PALETTE[idx % len(PALETTE)]
        fig.add_trace(go.Scatter(
            x=data["x"], y=data["y_orig"],
            mode="lines", name=f"{acsr_nm} 보정 전",
            line=dict(color=color, width=1.5, dash="dot", shape="spline"),
            connectgaps=False,
            legendgroup=acsr_nm,
        ), row=2, col=1)
        fig.add_trace(go.Scatter(
            x=data["x"], y=data["y_corr"],
            mode="lines", name=f"{acsr_nm} 보정 후",
            line=dict(color=color, width=2, shape="spline"),
            connectgaps=False,
            legendgroup=acsr_nm,
        ), row=2, col=1)

    period = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    chart_width = max(1200, len(x_all) * 6)
    fig.update_layout(
        hovermode="x unified",
        width=chart_width,
        height=1100,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_yaxes(title_text="합산 교통량 (대/시)", row=1, col=1)
    fig.update_yaxes(title_text="방향별 교통량 (대/시)", row=2, col=1)
    fig.update_xaxes(title_text="일시", row=2, col=1)

    sub_dir = viz_dir / f"{node_name}_{period}"
    sub_dir.mkdir(parents=True, exist_ok=True)
    out_path = sub_dir / f"{node_name}.html"
    plotly_div = fig.to_html(include_plotlyjs="cdn", full_html=False)
    html = (
        "<!DOCTYPE html>\n<html>\n<head><meta charset=\"utf-8\">\n"
        "<style>body{margin:0;} .scroll-wrap{overflow-x:auto; width:100%;}</style>\n"
        "</head>\n<body>\n"
        f"<div class=\"scroll-wrap\">{plotly_div}</div>\n"
        "</body>\n</html>"
    )
    out_path.write_text(html, encoding="utf-8")
    print(f"  시각화 저장: {out_path}")


# ════════════════════════════════════════════════════════════════
# 엑셀 출력
# ════════════════════════════════════════════════════════════════

def _build_excel_sheet_groups(results: list, intersections_count: int, has_drct: bool) -> OrderedDict:
    """출력 정책에 맞는 시트 그룹을 구성."""
    grouped: OrderedDict = OrderedDict()
    multi_intersections = intersections_count >= 2

    for row in results:
        if row.get("판정") not in ("A형", "B형", "A+B혼합"):
            continue
        if has_drct:
            if multi_intersections:
                sheet_key = row["교차로"]
            else:
                sheet_key = f"{row['교차로']}-{row['방향']}"
        else:
            if multi_intersections:
                sheet_key = row["교차로"]
            else:
                sheet_key = row["방향"]
        if sheet_key not in grouped:
            grouped[sheet_key] = []
        grouped[sheet_key].append(row)
    return grouped


def _unique_sheet_title(raw_title: str, used_titles: set) -> str:
    base = _sanitize_sheet_title(raw_title)[:31]
    title = base
    suffix_idx = 1
    while title in used_titles:
        suffix = f"_{suffix_idx}"
        title = f"{base[:31 - len(suffix)]}{suffix}"
        suffix_idx += 1
    used_titles.add(title)
    return title


def export_excel(results: list, filepath: Path, intersections_count: int = 1):
    wb = Workbook()
    multi_intersections = intersections_count >= 2
    has_drct = any((r.get("접근로방향") or r.get("_drct_name")) for r in results)

    thin   = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    fill_a = PatternFill("solid", fgColor="FFD7D7")  # A형: 연빨강
    fill_b = PatternFill("solid", fgColor="FFF3CD")  # B형: 연노랑
    fill_m = PatternFill("solid", fgColor="FFE6CC")  # A+B혼합: 연주황
    fill_n = PatternFill("solid", fgColor="E8E8E8")  # 정상(Stage2): 연회색

    sheets = _build_excel_sheet_groups(results, intersections_count, has_drct)
    used_titles: set = set()

    for sheet_key, rows in sheets.items():
        ws = wb.create_sheet(title=_unique_sheet_title(sheet_key, used_titles))

        if has_drct:
            if multi_intersections:
                headers = ["날짜", "시간", "교차로 방향", "접근로 방향", "교통량",
                           "보정값", "보정방법", "신뢰도", "이상 판단"]
                col_widths = [14, 8, 20, 20, 10, 10, 12, 10, 12]
            else:
                headers = ["날짜", "시간", "접근로 방향", "교통량",
                           "보정값", "보정방법", "신뢰도", "이상 판단"]
                col_widths = [14, 8, 20, 10, 10, 12, 10, 12]
        else:
            if multi_intersections:
                headers = ["날짜", "시간", "교차로 이름", "교차로 방향", "교통량",
                           "보정값", "보정방법", "신뢰도", "이상 판단"]
                col_widths = [14, 8, 22, 18, 10, 10, 12, 10, 12]
            else:
                headers = ["날짜", "시간", "교차로 방향", "교통량",
                           "보정값", "보정방법", "신뢰도", "이상 판단"]
                col_widths = [14, 8, 18, 10, 10, 12, 10, 12]

        for col, (hdr, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=1, column=col, value=hdr)
            cell.font      = Font(bold=True, color="FFFFFF")
            cell.fill      = PatternFill("solid", fgColor="2F4F8F")
            cell.alignment = center
            cell.border    = border
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 20

        for row_idx, row in enumerate(rows, 2):
            if has_drct:
                drct_name = row.get("접근로방향") or row.get("_drct_name") or ""
                if multi_intersections:
                    vals = [
                        row["날짜"], row["시간"], row["방향"], drct_name, row["교통량"],
                        row.get("보정값"), row.get("보정방법"), row.get("신뢰도"), row["판정"],
                    ]
                else:
                    vals = [
                        row["날짜"], row["시간"], drct_name, row["교통량"],
                        row.get("보정값"), row.get("보정방법"), row.get("신뢰도"), row["판정"],
                    ]
            else:
                if multi_intersections:
                    vals = [
                        row["날짜"], row["시간"], row["교차로"], row["방향"], row["교통량"],
                        row.get("보정값"), row.get("보정방법"), row.get("신뢰도"), row["판정"],
                    ]
                else:
                    vals = [
                        row["날짜"], row["시간"], row["방향"], row["교통량"],
                        row.get("보정값"), row.get("보정방법"), row.get("신뢰도"), row["판정"],
                    ]
            if row["판정"] == "A형":
                fill = fill_a
            elif row["판정"] == "B형":
                fill = fill_b
            elif row["판정"] == "A+B혼합":
                fill = fill_m
            else:
                fill = fill_n
            for col, val in enumerate(vals, 1):
                cell           = ws.cell(row=row_idx, column=col, value=val)
                cell.alignment = center
                cell.border    = border
                cell.fill      = fill

    # 기본 빈 시트 제거
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    filepath.parent.mkdir(parents=True, exist_ok=True)
    wb.save(filepath)
    print(f"\n  저장 완료: {filepath}")


# ════════════════════════════════════════════════════════════════
# 사용자 입력
# ════════════════════════════════════════════════════════════════

def input_intersections() -> list:
    """[(node_id, name), ...] 반환 — API /intersections 사용"""
    items      = _api_get("/intersections")["items"]
    name_to_id = {item["name"]: item["node_id"] for item in items}
    all_names  = list(name_to_id.keys())

    print("\n교차로 이름을 입력하세요 (쉼표로 구분 가능, 빈 줄로 종료):")
    selected = []
    seen     = set()

    while True:
        raw = input(f"  교차로 [{len(selected) + 1}]: ").strip()
        if not raw:
            if selected:
                break
            print("  최소 1개 이상 입력하세요.")
            continue

        parts = [p.strip() for p in raw.split(",") if p.strip()]
        for part in parts:
            if part in name_to_id:
                matched = part
            else:
                candidates = difflib.get_close_matches(part, all_names, n=1, cutoff=0.6)
                if not candidates:
                    print(f"  [오류] '{part}'와 유사한 교차로를 찾을 수 없습니다.")
                    continue
                matched = candidates[0]
                ans = input(f"  '{matched}'이(가) 맞습니까? (Y/n): ").strip().lower()
                if ans == "n":
                    print(f"  '{part}' 입력을 건너뜁니다.")
                    continue

            if matched in seen:
                print(f"  '{matched}'은 이미 추가되었습니다.")
                continue
            seen.add(matched)
            selected.append((name_to_id[matched], matched))
            print(f"  ✓ '{matched}' 추가됨")

    return selected


def parse_date_range(text: str):
    today = date.today()

    def parse_one(s):
        s = s.replace(".", "").replace("-", "").replace("/", "")
        if len(s) == 8:
            return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        if len(s) == 6:
            return date(2000 + int(s[:2]), int(s[2:4]), int(s[4:6]))
        if len(s) == 4:
            return date(today.year, int(s[:2]), int(s[2:4]))
        raise ValueError(f"날짜 형식 오류: '{s}'")

    parts = text.strip().replace(" ", "").split("~")
    if len(parts) != 2:
        raise ValueError("'~'로 구분된 날짜 범위를 입력하세요.")
    return parse_one(parts[0]), parse_one(parts[1])


def input_date_range():
    print("\n기간을 입력하세요 (예: 260324~260512, 0324~0722):")
    while True:
        raw = input("  기간: ").strip()
        try:
            d_start, d_end = parse_date_range(raw)
            if d_start > d_end:
                print("  [오류] 시작일이 종료일보다 늦습니다.")
                continue
            print(f"  → {d_start.strftime('%Y.%m.%d')} ~ {d_end.strftime('%Y.%m.%d')}")
            return d_start, d_end
        except (ValueError, IndexError) as e:
            print(f"  [오류] {e}")


def parse_custom_hours(text: str) -> list:
    """'08:00~14:00, 17:00~22:00' → [8,9,...,13,17,...,21]"""
    hours = set()
    for part in text.split(","):
        part = part.strip()
        if "~" not in part:
            continue
        s, e = part.split("~")
        sh = int(s.strip().split(":")[0])
        eh = int(e.strip().split(":")[0])
        for h in range(sh, eh):
            hours.add(h)
    return sorted(hours)


def input_hours() -> list:
    print("\n시간대를 선택하세요:")
    print("  1. 24시간 전체 (00~23시)")
    print("  2. 첨두시 (오전 07~09, 비첨두 12~14, 오후 17~19)")
    print("  3. 직접 입력 (예: 08:00~14:00, 17:00~22:00)")
    while True:
        choice = input("  선택 (1/2/3): ").strip()
        if choice == "1":
            return list(range(24))
        if choice == "2":
            hours = list(range(7, 9)) + list(range(12, 14)) + list(range(17, 19))
            print(f"  → {hours}")
            return hours
        if choice == "3":
            raw = input("  시간대 입력: ").strip()
            hours = parse_custom_hours(raw)
            if not hours:
                print("  [오류] 유효한 시간대가 없습니다.")
                continue
            print(f"  → {hours}")
            return hours
        print("  1, 2, 3 중 선택하세요.")


def input_view_mode() -> str:
    print("\n출력 기준을 선택하세요:")
    print("  1. 방향별 교통량(ACSR)")
    print("  2. 접근로 방향별 교통량(DRCT)")
    while True:
        choice = input("  선택 (1/2): ").strip()
        if choice == "1":
            return "ACSR"
        if choice == "2":
            return "DRCT"
        print("  1 또는 2를 입력하세요.")


def input_approach_names(approaches: list, node_name: str) -> list:
    """각 접근로 방향명을 사용자에게 입력받음. Enter → 기존 이름 유지."""
    print(f"\n  [{node_name}] 접근로 방향명 입력 (Enter=기본값 유지):")
    result = []
    for acsr_id, acsr_nm in approaches:
        custom = input(f"    {acsr_nm} → ").strip()
        result.append((acsr_id, custom if custom else acsr_nm))
    return result


def input_drct_selection_by_approach(drct_options_by_acsr: dict) -> dict:
    """ACSR별 DRCT 선택 입력.

    Returns:
      {acsr_id: {drct_cd, ...}}
    """
    print("\n  접근로별 DRCT 선택 (콤마 입력, Enter=전체 선택):")
    selected_by_acsr: dict = {}

    for acsr_id, info in drct_options_by_acsr.items():
        approach_name = info.get("approach_name") or str(acsr_id)
        drcts = info.get("drcts") or []
        if not drcts:
            selected_by_acsr[acsr_id] = set()
            print(f"    - {approach_name}: 선택 가능한 DRCT 없음")
            continue

        print(f"\n    [{approach_name}]")
        idx_to_cd = {}
        for idx, (drct_cd, drct_name) in enumerate(drcts, 1):
            idx_to_cd[str(idx)] = drct_cd
            print(f"      {idx}. {drct_cd} - {drct_name}")

        all_codes = {drct_cd for drct_cd, _ in drcts}
        while True:
            raw = input("      선택: ").strip()
            if not raw:
                selected_by_acsr[acsr_id] = all_codes
                break

            tokens = [t.strip() for t in raw.split(",") if t.strip()]
            picked: set = set()
            invalid: list = []
            for token in tokens:
                norm_token = _normalize_drct_cd(token)
                if token in idx_to_cd:
                    picked.add(idx_to_cd[token])
                elif norm_token in all_codes:
                    picked.add(norm_token)
                else:
                    invalid.append(token)

            if invalid:
                print(f"      [오류] 유효하지 않은 선택: {', '.join(invalid)}")
                continue
            if not picked:
                print("      [오류] 최소 1개 이상 선택하세요.")
                continue

            selected_by_acsr[acsr_id] = picked
            break

    return selected_by_acsr


def apply_approach_name_overrides(
    node_results: list,
    approaches: list,
    drct_options_by_acsr: dict,
) -> tuple[list, dict]:
    """사용자 입력 접근로명(ACSR)을 결과/옵션에 반영."""
    acsr_name_map = {acsr_id: acsr_nm for acsr_id, acsr_nm in approaches}

    for r in node_results:
        acsr_id = r.get("_acsr_id")
        if acsr_id in acsr_name_map:
            r["방향"] = acsr_name_map[acsr_id]

    for acsr_id, info in drct_options_by_acsr.items():
        if acsr_id in acsr_name_map:
            info["approach_name"] = acsr_name_map[acsr_id]

    return node_results, drct_options_by_acsr


def filter_by_drct_selection(
    node_results: list,
    target_data: dict,
    drct_selection_by_acsr: dict,
) -> tuple[list, dict]:
    """ACSR별 DRCT 선택 필터를 이상 슬롯/원본 슬롯 모두에 적용."""
    filtered_results = [
        r for r in node_results
        if r.get("_drct_cd") in drct_selection_by_acsr.get(r.get("_acsr_id"), set())
    ]

    filtered_target_data = {
        key: val
        for key, val in target_data.items()
        if key[1] in drct_selection_by_acsr.get(key[0], set())
    }
    return filtered_results, filtered_target_data


def make_filename(intersections: list, date_start: date, date_end: date) -> str:
    names     = [nm for _, nm in intersections]
    name_part = names[0] if len(names) == 1 else f"{names[0]}_{names[-1]}"
    period    = f"{date_start.strftime('%y%m%d')}~{date_end.strftime('%y%m%d')}"
    return f"{name_part}_{period}.xlsx"


# ════════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  교통량 이상탐지 시스템")
    print("=" * 60)

    print("\nAPI 서버 연결 중...")
    try:
        _api_get("/intersections")
        print("  API 서버 연결 성공")
    except Exception as e:
        print(f"  [오류] API 서버 연결 실패: {e}")
        print(f"  (api_server.py가 실행 중인지 확인하세요)")
        sys.exit(1)

    holiday_dates = load_holidays()

    # Step 1: 교차로
    intersections = input_intersections()
    print(f"\n선택된 교차로 ({len(intersections)}개): {[nm for _, nm in intersections]}")

    # Step 2: 기간
    date_start, date_end = input_date_range()

    # Step 3: 시간대
    hours = input_hours()

    # Step 4: 기준 선택
    view_mode = input_view_mode()

    all_results = []
    viz_dir     = RESULT_DIR / "보정_시각화"
    excel_dir   = RESULT_DIR / "이상탐지_결과"

    for node_id, node_name in intersections:
        print(f"\n{'─'*50}")
        print(f"[{node_name}] 처리 중...")

        api_label_suffix = "API-DRCT"
        api_progress = _SpinnerProgress(
            label=f"[{node_name}] 데이터 조회 및 이상탐지 ({api_label_suffix})",
            start_percent=5,
            max_percent=38,
        )
        api_progress.start()
        try:
            api_path = "/corrected-traffic-drct"
            resp = _api_post(api_path, {
                "node_ids":   [node_id],
                "date_start": date_start.isoformat(),
                "date_end":   date_end.isoformat(),
                "hours":      hours,
            })
        finally:
            api_progress.stop(final_percent=40)

        drct_options_by_acsr = None
        drct_selection_by_acsr = None
        if view_mode == "ACSR":
            node_results, target_data, approaches = _parse_api_response_acsr_from_drct(
                resp.get("drct_slots", []), node_name
            )
            _print_inline_progress(f"  [{node_name}] DRCT→ACSR 집계 파싱... 완료 50%")
        else:
            node_results, target_data, approaches, drct_options_by_acsr = _parse_api_response_drct(
                resp.get("drct_slots", []), node_name
            )
            _print_inline_progress(f"  [{node_name}] DRCT 응답 파싱... 완료 50%")
        print()

        if not approaches:
            print(f"  접근로 정보 없음, 건너뜀")
            continue

        approaches = input_approach_names(approaches, node_name)
        if view_mode == "ACSR":
            acsr_name_map = {acsr_id: acsr_nm for acsr_id, acsr_nm in approaches}
            for r in node_results:
                if r.get("_acsr_id") in acsr_name_map:
                    r["방향"] = acsr_name_map[r["_acsr_id"]]
        else:
            node_results, drct_options_by_acsr = apply_approach_name_overrides(
                node_results, approaches, drct_options_by_acsr
            )
            drct_selection_by_acsr = input_drct_selection_by_approach(drct_options_by_acsr)
            node_results, target_data = filter_by_drct_selection(
                node_results, target_data, drct_selection_by_acsr
            )

        cnt_a = sum(1 for r in node_results if r["판정"] == "A형")
        cnt_b = sum(1 for r in node_results if r["판정"] == "B형")
        cnt_m = sum(1 for r in node_results if r["판정"] == "A+B혼합")
        if view_mode == "ACSR":
            filtered_slots = [
                s for s in resp.get("drct_slots", [])
                if _normalize_drct_cd(s.get("drct_cd")) != "00"
            ]
            print(f"  수신 DRCT 슬롯(00 제외): {len(filtered_slots):,}건")
            print(
                f"  이상 탐지(A/B/혼합): {len(node_results)}건  "
                f"(A형: {cnt_a}, B형: {cnt_b}, A+B혼합: {cnt_m})"
            )

            viz_progress = _make_step_progress_callback(
                label=f"[{node_name}] ACSR 시각화 생성",
                start_percent=60,
                end_percent=95,
            )
            export_acsr_visualizations(
                node_name=node_name,
                approaches=approaches,
                target_data=target_data,
                node_results=node_results,
                date_start=date_start,
                date_end=date_end,
                hours=hours,
                viz_dir=viz_dir,
                progress_callback=viz_progress,
                verbose=False,
            )
        else:
            filtered_slots = [
                s for s in resp.get("drct_slots", [])
                if _normalize_drct_cd(s.get("drct_cd")) != "00"
            ]
            print(f"  수신 DRCT 슬롯(00 제외): {len(filtered_slots):,}건")
            print(
                f"  이상 탐지(A/B/혼합): {len(node_results)}건  "
                f"(A형: {cnt_a}, B형: {cnt_b}, A+B혼합: {cnt_m})"
            )

            viz_progress = _make_step_progress_callback(
                label=f"[{node_name}] DRCT 시각화 생성",
                start_percent=60,
                end_percent=95,
            )
            export_drct_visualizations(
                node_name=node_name,
                drct_options_by_acsr=drct_options_by_acsr,
                drct_selection_by_acsr=drct_selection_by_acsr,
                target_data=target_data,
                node_results=node_results,
                date_start=date_start,
                date_end=date_end,
                hours=hours,
                viz_dir=viz_dir,
                progress_callback=viz_progress,
                verbose=False,
            )
        _print_inline_progress(f"  [{node_name}] 노드 처리 완료... 완료 100%")
        print()

        all_results.extend(node_results)

    print(f"\n{'='*60}")

    if not all_results:
        print("이상 슬롯이 발견되지 않았습니다.")
        return

    # 정렬: 날짜 → 시간 → 교차로 → 교차로방향 → 접근로방향
    all_results.sort(
        key=lambda r: (
            r["날짜"],
            r["시간"],
            r["교차로"],
            r["방향"],
            r.get("접근로방향") or r.get("_drct_name") or "",
        )
    )

    filename = make_filename(intersections, date_start, date_end)
    filepath = excel_dir / filename
    export_excel(all_results, filepath, intersections_count=len(intersections))
    print(f"총 {len(all_results)}건의 이상값이 탐지되었습니다.")


if __name__ == "__main__":
    main()
