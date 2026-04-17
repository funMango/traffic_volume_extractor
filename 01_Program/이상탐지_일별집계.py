#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""이상탐지 일별 집계 캐시 빌더

이상탐지 결과를 날짜별·교차로별로 SQLite에 집계·저장한다.
이미 처리된 날짜는 DB를 재사용하고, 미처리 날짜만 API를 통해 분석한다.

사용법:
  python 이상탐지_일별집계.py 2025           # 2025년 전체
  python 이상탐지_일별집계.py 202506         # 2025년 6월
  python 이상탐지_일별집계.py 250615~250712  # 2025-06-15 ~ 2025-07-12
"""

import calendar
import json
import sqlite3
import sys
import threading
import urllib.error
import urllib.request
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path

from tqdm import tqdm

sys.path.insert(0, str(Path(__file__).parent))
import 이상탐지 as ad

# ── API 설정 ──────────────────────────────────────────────────────────────────

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

# ── 경로 상수 ─────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH  = BASE_DIR / "02_Result" / "이상탐지_캐시" / "anomaly_cache.db"

HOURS = list(range(24))

# ── SQLite 스키마 ─────────────────────────────────────────────────────────────

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS processed_dates (
    date         TEXT PRIMARY KEY,
    processed_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS anomaly_daily (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    date          TEXT    NOT NULL,
    node_id       INTEGER NOT NULL,
    node_name     TEXT    NOT NULL,
    approach_id   INTEGER,
    approach_name TEXT,
    missing_count INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_daily_date ON anomaly_daily(date);
CREATE INDEX IF NOT EXISTS idx_daily_node ON anomaly_daily(node_id, date);
"""


# ── 입력 파싱 ─────────────────────────────────────────────────────────────────

def _parse_short(s: str) -> date:
    s = s.strip()
    if len(s) == 6:
        y, m, d = 2000 + int(s[:2]), int(s[2:4]), int(s[4:6])
    elif len(s) == 8:
        y, m, d = int(s[:4]), int(s[4:6]), int(s[6:8])
    else:
        raise ValueError(f"날짜 형식 오류: '{s}' (YYMMDD 또는 YYYYMMDD)")
    return date(y, m, d)


def parse_cli_arg(arg: str) -> tuple[date, date]:
    arg = arg.strip()
    if "~" in arg:
        left, right = arg.split("~", 1)
        return _parse_short(left), _parse_short(right)
    if len(arg) == 4:
        y = int(arg)
        return date(y, 1, 1), date(y, 12, 31)
    if len(arg) == 6:
        y, m = int(arg[:4]), int(arg[4:6])
        last = calendar.monthrange(y, m)[1]
        return date(y, m, 1), date(y, m, last)
    raise ValueError(f"지원하지 않는 입력 형식: '{arg}'")


# ── DB 유틸 ───────────────────────────────────────────────────────────────────

def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    # 구 스키마 제거
    conn.execute("DROP TABLE IF EXISTS anomaly_slots")
    conn.execute("DROP INDEX IF EXISTS idx_slots_date")
    conn.execute("DROP INDEX IF EXISTS idx_slots_node")
    conn.executescript(_SCHEMA_SQL)
    conn.commit()
    return conn


def date_range(ds: date, de: date) -> list[date]:
    days = []
    cur = ds
    while cur <= de:
        days.append(cur)
        cur += timedelta(days=1)
    return days


def get_unprocessed_dates(db_conn: sqlite3.Connection, all_dates: list[date]) -> list[date]:
    if not all_dates:
        return []
    placeholders = ",".join("?" * len(all_dates))
    rows = db_conn.execute(
        f"SELECT date FROM processed_dates WHERE date IN ({placeholders})",
        [d.isoformat() for d in all_dates],
    ).fetchall()
    done = {r[0] for r in rows}
    return [d for d in all_dates if d.isoformat() not in done]


def group_consecutive_dates(dates: list[date]) -> list[tuple[date, date]]:
    if not dates:
        return []
    sorted_dates = sorted(set(dates))
    groups = []
    seg_start = seg_end = sorted_dates[0]
    for d in sorted_dates[1:]:
        if (d - seg_end).days == 1:
            seg_end = d
        else:
            groups.append((seg_start, seg_end))
            seg_start = seg_end = d
    groups.append((seg_start, seg_end))
    return groups


def split_by_year(ds: date, de: date) -> list[tuple[date, date]]:
    segments = []
    cur = ds
    while cur.year < de.year:
        segments.append((cur, date(cur.year, 12, 31)))
        cur = date(cur.year + 1, 1, 1)
    segments.append((cur, de))
    return segments


def clean_unprocessed_slots(db_conn: sqlite3.Connection, dates: list[date]) -> None:
    if not dates:
        return
    placeholders = ",".join("?" * len(dates))
    db_conn.execute(
        f"DELETE FROM anomaly_daily WHERE date IN ({placeholders})",
        [d.isoformat() for d in dates],
    )
    db_conn.commit()


def mark_dates_processed(db_conn: sqlite3.Connection, dates: list[date]) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    db_conn.executemany(
        "INSERT OR IGNORE INTO processed_dates (date, processed_at) VALUES (?, ?)",
        [(d.isoformat(), now) for d in dates],
    )
    db_conn.commit()


# ── 날짜별 집계 + INSERT ──────────────────────────────────────────────────────

MAX_WORKERS = 6  # API 서버 부하에 따라 조정


def aggregate_and_insert_daily(
    db_conn: sqlite3.Connection,
    db_lock: threading.Lock,
    node_id: int,
    node_name: str,
    node_results: list,
    all_slots: list,
    seg_ds: date,
    seg_de: date,
) -> list[dict]:
    # (date, hour) → set[acsr_id] 인덱스 (전체 슬롯 기준 expected 계산용)
    all_slot_index: dict[tuple, set[int]] = defaultdict(set)
    for slot in all_slots:
        d_key = date.fromisoformat(slot["date"])
        all_slot_index[(d_key, slot["hour"])].add(slot["approach_id"])

    # (date, hour) → {acsr_id: result_dict} 인덱스 (A형 슬롯만)
    slot_index: dict[tuple, dict[int, dict]] = defaultdict(dict)
    for r in node_results:
        if r["판정"] == "A형":
            slot_index[(r["_date"], r["_hour"])][r["_acsr_id"]] = r

    # 날짜별 카운터
    node_counter: dict[date, int] = defaultdict(int)
    approach_counter: dict[tuple, list] = {}  # (date, acsr_id) → [name, count]

    cur_d = seg_ds
    while cur_d <= seg_de:
        for hour in HOURS:
            expected = all_slot_index.get((cur_d, hour), set())
            if not expected:
                continue

            anomaly_map = slot_index.get((cur_d, hour), {})
            if not anomaly_map:
                continue

            if set(anomaly_map.keys()) == expected:
                # 모든 접근로 이상 → 교차로 레벨
                node_counter[cur_d] += 1
            else:
                # 일부 접근로만 이상 → 방향별
                for acsr_id, r in anomaly_map.items():
                    key = (cur_d, acsr_id)
                    if key not in approach_counter:
                        approach_counter[key] = [r["방향"], 0]
                    approach_counter[key][1] += 1

        cur_d += timedelta(days=1)

    rows_to_insert: list[dict] = []
    for d, cnt in node_counter.items():
        rows_to_insert.append({
            "date":          d.isoformat(),
            "node_id":       node_id,
            "node_name":     node_name,
            "approach_id":   None,
            "approach_name": None,
            "missing_count": cnt,
        })
    for (d, acsr_id), (name, cnt) in approach_counter.items():
        rows_to_insert.append({
            "date":          d.isoformat(),
            "node_id":       node_id,
            "node_name":     node_name,
            "approach_id":   acsr_id,
            "approach_name": name,
            "missing_count": cnt,
        })

    if rows_to_insert:
        with db_lock:
            db_conn.executemany(
                """INSERT INTO anomaly_daily
                   (date, node_id, node_name, approach_id, approach_name, missing_count)
                   VALUES
                   (:date, :node_id, :node_name, :approach_id, :approach_name, :missing_count)
                """,
                rows_to_insert,
            )

    return rows_to_insert


# ── 교차로 단위 처리 (스레드 워커) ───────────────────────────────────────────

def _process_node(
    node_id: int,
    node_name: str,
    year_split_segments: list,
    db_conn: sqlite3.Connection,
    db_lock: threading.Lock,
) -> tuple:
    node_inserted = 0
    local_summary: dict[date, int] = defaultdict(int)
    failed = False

    for seg_ds, seg_de in year_split_segments:
        try:
            resp = _api_post("/corrected-traffic", {
                "node_ids":   [node_id],
                "date_start": seg_ds.isoformat(),
                "date_end":   seg_de.isoformat(),
                "hours":      HOURS,
            })
            all_slots = resp["slots"]
            node_results, _, _ = ad._parse_api_response(all_slots, node_name)
        except Exception as e:
            tqdm.write(f"  [경고] {node_name} {seg_ds}~{seg_de} 분석 실패: {e}")
            failed = True
            continue

        try:
            inserted = aggregate_and_insert_daily(
                db_conn, db_lock, node_id, node_name,
                node_results, all_slots,
                seg_ds, seg_de,
            )
        except sqlite3.Error as e:
            tqdm.write(f"  [오류] DB 쓰기 실패 ({node_name}): {e}")
            continue

        node_inserted += len(inserted)
        for row in inserted:
            d = date.fromisoformat(row["date"])
            local_summary[d] += 1

    return node_inserted, node_name, local_summary, failed


# ── 요약 출력 ─────────────────────────────────────────────────────────────────

def print_summary(summary: dict[str, dict[date, int]], total_inserted: int) -> None:
    print("\n" + "=" * 60)
    print("  처리 결과 요약")
    print("=" * 60)

    date_items: dict[date, list[str]] = defaultdict(list)
    for node_name, date_counts in summary.items():
        for d, count in date_counts.items():
            if count > 0:
                date_items[d].append(f"{node_name} {count}건")

    if not date_items:
        print("이상 없음")
    else:
        for d in sorted(date_items.keys()):
            print(f"{d}: {', '.join(date_items[d])}")

    print("=" * 60)
    print(f"완료: {total_inserted:,}건 저장")


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main() -> None:
    if len(sys.argv) < 2:
        arg = input("기간 입력 (예: 2025 / 202506 / 250615~250712): ").strip()
    else:
        arg = sys.argv[1]

    try:
        date_start, date_end = parse_cli_arg(arg)
    except ValueError as e:
        print(f"[오류] {e}")
        sys.exit(1)

    print("=" * 60)
    print("  교통량 이상탐지 일별 캐싱")
    print(f"  기간: {date_start} ~ {date_end}")
    print(f"  DB  : {DB_PATH}")
    print("=" * 60)

    db_conn = init_db(DB_PATH)

    all_dates  = date_range(date_start, date_end)
    todo_dates = get_unprocessed_dates(db_conn, all_dates)
    skip_count = len(all_dates) - len(todo_dates)

    print(f"기존 처리: {skip_count}일  /  미처리: {len(todo_dates)}일")

    if not todo_dates:
        print("모든 날짜가 이미 처리되어 있습니다.")
        db_conn.close()
        return

    consecutive_groups  = group_consecutive_dates(todo_dates)
    year_split_segments: list[tuple[date, date]] = []
    for ds, de in consecutive_groups:
        year_split_segments.extend(split_by_year(ds, de))

    clean_unprocessed_slots(db_conn, todo_dates)

    print(f"API 서버 연결 중... ({API_BASE_URL})")
    try:
        raw = _api_get("/intersections")
        intersections = [(item["node_id"], item["name"]) for item in raw["items"]]
        print(f"API 서버 연결 완료 (교차로 {len(intersections)}개)")
    except Exception as e:
        print(f"[오류] API 서버 연결 실패: {e}")
        print("  (api_server.py가 실행 중인지 확인하세요)")
        db_conn.close()
        sys.exit(1)

    summary: dict[str, dict[date, int]] = defaultdict(lambda: defaultdict(int))
    total_inserted = 0
    failed_nodes: list[str] = []

    db_lock = threading.Lock()

    with tqdm(total=len(intersections), desc="이상탐지 캐싱", unit="교차로", ncols=80) as pbar:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(_process_node, nid, nm, year_split_segments, db_conn, db_lock): nm
                for nid, nm in intersections
            }
            for fut in as_completed(futures):
                node_inserted, node_name, local_summary, failed = fut.result()
                if failed:
                    failed_nodes.append(node_name)
                total_inserted += node_inserted
                for d, cnt in local_summary.items():
                    summary[node_name][d] += cnt
                with db_lock:
                    db_conn.commit()
                pbar.set_postfix_str(node_name[:12])
                pbar.update(1)

    if failed_nodes:
        tqdm.write(f"\n[주의] 일부 교차로 실패: {', '.join(set(failed_nodes))}")
        tqdm.write("실패 교차로가 포함된 날짜는 다음 실행 시 재처리됩니다.")
    else:
        mark_dates_processed(db_conn, todo_dates)

    db_conn.close()

    print_summary(summary, total_inserted)


if __name__ == "__main__":
    main()
