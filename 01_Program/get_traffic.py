#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""교통량 추출 프로그램

사용법: python get_traffic.py

교차로 이름, 기간, 시간대, 속성, 시트 설정을 순서대로 입력하면
Excel(.xlsx) 파일과 선택적으로 HTML 그래프를 생성합니다.

저장 위치:
  Excel  : 02_Result/교통량_추출/교통량/[교차로이름]_[기간].xlsx
  그래프 : 02_Result/교통량_추출/그래프/[교차로이름]_[기간]/
"""

import difflib
import json
import re
import sys
import urllib.request
import urllib.error
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── API 서버 설정 ─────────────────────────────────────────────────────────────
API_BASE_URL = "http://localhost:8000"

BASE_DIR   = Path(__file__).resolve().parent.parent
RESULT_DIR = BASE_DIR / "02_Result" / "교통량_추출"
EXCEL_DIR  = RESULT_DIR / "교통량"
GRAPH_DIR  = RESULT_DIR / "그래프"

# 속성 한글 레이블
ALL_ATTRS = ["시간", "교차로명", "방향명", "교통량"]

# 그래프 방향별 색상 — 세련된 팔레트
DIR_COLORS = ["#2563EB", "#DC2626", "#059669", "#D97706", "#7C3AED"]
NAVY       = "#1E40AF"   # 단독 그래프용 기본색


# ════════════════════════════════════════════════════════════════
# API 헬퍼
# ════════════════════════════════════════════════════════════════

def _api_get(path: str) -> dict:
    """GET 요청 → dict 반환"""
    with urllib.request.urlopen(f"{API_BASE_URL}{path}") as resp:
        return json.loads(resp.read())


def _api_post(path: str, payload: dict) -> dict:
    """POST 요청 → dict 반환"""
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


# ════════════════════════════════════════════════════════════════
# 입력 처리
# ════════════════════════════════════════════════════════════════

def input_intersections() -> list[dict]:
    """교차로 이름 입력 (복수 가능, 퍼지 매칭 지원)
    반환: [{"node_id": int, "name": str}, ...]
    """
    items = _api_get("/intersections")["items"]  # [{"node_id": ..., "name": ...}, ...]
    all_names  = [it["name"] for it in items]
    name_to_id = {it["name"]: it["node_id"] for it in items}

    print("\n[1단계] 교차로 이름 입력")
    print("  여러 교차로는 쉼표로 구분하세요. (예: 구지사거리, 멀뫼사거리)")
    print(f"  전체 교차로 수: {len(all_names)}개")

    selected: list[dict] = []

    while True:
        raw = input("\n교차로 이름: ").strip()
        if not raw:
            print("  교차로 이름을 입력해 주세요.")
            continue

        inputs = [s.strip() for s in raw.split(",") if s.strip()]
        resolved: list[dict] = []
        failed = False

        for name in inputs:
            if name in name_to_id:
                resolved.append({"node_id": name_to_id[name], "name": name})
            else:
                candidates = difflib.get_close_matches(name, all_names, n=3, cutoff=0.3)
                if candidates:
                    print(f"\n  '{name}' 을(를) 찾을 수 없습니다. 비슷한 교차로:")
                    for i, c in enumerate(candidates, 1):
                        print(f"    {i}. {c}")
                    print(f"    0. 직접 다시 입력")
                    while True:
                        choice = input("  선택 (번호 입력): ").strip()
                        if choice == "0":
                            failed = True
                            break
                        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
                            chosen = candidates[int(choice) - 1]
                            resolved.append({"node_id": name_to_id[chosen], "name": chosen})
                            break
                        print("  올바른 번호를 입력해 주세요.")
                    if failed:
                        break
                else:
                    print(f"\n  '{name}' 과(와) 유사한 교차로를 찾을 수 없습니다.")
                    failed = True
                    break

        if failed:
            continue

        # 중복 제거 (node_id 기준)
        seen: set[str] = set()
        unique: list[dict] = []
        for item in resolved:
            if item["node_id"] not in seen:
                seen.add(item["node_id"])
                unique.append(item)

        names_str = ", ".join(d["name"] for d in unique)
        print(f"\n  선택된 교차로 ({len(unique)}개): {names_str}")
        confirm = input("  계속 진행하시겠습니까? (y/n) [y]: ").strip().lower()
        if confirm in ("", "y"):
            selected = unique
            break
        # n이면 처음부터

    return selected


def input_period() -> list[tuple[date, date, str]]:
    """기간 입력 (복수 구간 가능, 쉼표 구분)
    반환: [(date_start, date_end, "YYMMDD~YYMMDD"), ...]
    """
    print("\n[2단계] 기간 입력")
    print("  형식: YYMMDD~YYMMDD  (예: 260413~260414)")
    print("  복수 구간: 쉼표로 구분  (예: 260302~260304, 260312~260315)")

    def parse_yymmdd(s: str) -> date | None:
        try:
            yy, mm, dd = int(s[:2]), int(s[2:4]), int(s[4:6])
            return date(2000 + yy, mm, dd)
        except ValueError:
            return None

    while True:
        raw = input("\n기간: ").strip()
        segments = [s.strip() for s in raw.split(",") if s.strip()]
        if not segments:
            print("  기간을 입력해 주세요.")
            continue

        periods: list[tuple[date, date, str]] = []
        failed = False

        for seg in segments:
            m = re.fullmatch(r"(\d{6})~(\d{6})", seg)
            if not m:
                print(f"  형식이 올바르지 않습니다: '{seg}'  예) 260413~260414")
                failed = True
                break
            ds = parse_yymmdd(m.group(1))
            de = parse_yymmdd(m.group(2))
            if ds is None or de is None:
                print(f"  날짜 값이 올바르지 않습니다: '{seg}'")
                failed = True
                break
            if ds > de:
                print(f"  시작 날짜가 종료 날짜보다 늦습니다: '{seg}'")
                failed = True
                break
            periods.append((ds, de, seg))

        if failed:
            continue

        for ds, de, seg in periods:
            print(f"  {seg}  ({ds} ~ {de})")
        return periods


def _parse_custom_time(raw: str) -> list[int] | None:
    """'18:00~20:00, 22:00~24:00' → [18, 19, 22, 23]"""
    hours: list[int] = []
    segments = [s.strip() for s in raw.split(",") if s.strip()]
    for seg in segments:
        m = re.fullmatch(r"(\d{1,2}):00\s*~\s*(\d{1,2}):00", seg)
        if not m:
            return None
        h_start, h_end = int(m.group(1)), int(m.group(2))
        if not (0 <= h_start <= 24 and 0 <= h_end <= 24):
            return None
        if h_start >= h_end:
            return None
        # 24시는 23시까지만 포함
        end = min(h_end, 24)
        hours.extend(range(h_start, end))
    return sorted(set(hours)) if hours else None


def input_time_range() -> list[int]:
    """시간대 선택 → hours list"""
    print("\n[3단계] 시간대 입력")
    print("  1. 24시간 (0~23시)")
    print("  2. 첨두시간 (07~09시, 17~19시)")
    print("  3. 직접 입력 (예: 18:00~20:00, 22:00~24:00)")

    while True:
        choice = input("\n선택 (1/2/3): ").strip()
        if choice == "1":
            print("  → 0~23시 (24시간)")
            return list(range(24))
        elif choice == "2":
            hours = [7, 8, 17, 18]
            print("  → 7, 8, 17, 18시 (첨두시간)")
            return hours
        elif choice == "3":
            while True:
                raw = input("  시간대 입력: ").strip()
                hours = _parse_custom_time(raw)
                if hours is None:
                    print("  형식이 올바르지 않습니다. 예) 18:00~20:00, 22:00~24:00")
                    continue
                print(f"  → {hours}")
                return hours
        else:
            print("  1, 2, 3 중 하나를 입력하세요.")


def input_attributes() -> tuple[list[str], bool, bool]:
    """속성 선택
    반환: (attr_list, has_node, has_dir)
      attr_list : 출력할 속성 순서 리스트
      has_node  : '교차로명' 포함 여부
      has_dir   : '방향명' 포함 여부
    """
    print("\n[4단계] 속성 선택")
    print("  1. 시간, 교차로명, 방향명, 교통량")
    print("  2. 시간, 교차로명, 교통량  (방향별 합산)")
    print("  3. 직접 입력 (예: 교통량, 교차로명)")
    print(f"  선택 가능 속성: {', '.join(ALL_ATTRS)}")

    while True:
        choice = input("\n선택 (1/2/3): ").strip()
        if choice == "1":
            attrs = ["시간", "교차로명", "방향명", "교통량"]
            print(f"  → {', '.join(attrs)}")
            return attrs, True, True
        elif choice == "2":
            attrs = ["시간", "교차로명", "교통량"]
            print(f"  → {', '.join(attrs)}  (방향별 합산)")
            return attrs, True, False
        elif choice == "3":
            while True:
                raw = input("  속성 입력: ").strip()
                parts = [p.strip() for p in raw.split(",") if p.strip()]
                invalid = [p for p in parts if p not in ALL_ATTRS]
                if invalid:
                    print(f"  올바르지 않은 속성: {invalid}")
                    print(f"  선택 가능 속성: {', '.join(ALL_ATTRS)}")
                    continue
                if not parts:
                    print("  속성을 하나 이상 입력하세요.")
                    continue
                has_node = "교차로명" in parts
                has_dir  = "방향명"  in parts
                print(f"  → {', '.join(parts)}")
                return parts, has_node, has_dir
        else:
            print("  1, 2, 3 중 하나를 입력하세요.")


def input_sheet_type(has_node: bool, has_dir: bool) -> str:
    """시트 설정 선택
    반환: "교차로별" | "방향별" | "단일"
    """
    print("\n[5단계] 시트 설정")

    options: list[tuple[str, str]] = []
    if has_node:
        options.append(("교차로별", "교차로 하나당 시트 1개"))
    if has_dir:
        options.append(("방향별", "방향 하나당 시트 1개"))

    if not options:
        print("  속성에 교차로명/방향명이 없어 단일 시트로 저장합니다.")
        return "단일"

    if len(options) == 1:
        name, desc = options[0]
        print(f"  → {name} ({desc}) 으로 자동 설정됩니다.")
        return name

    for i, (name, desc) in enumerate(options, 1):
        print(f"  {i}. {name}  ({desc})")

    while True:
        choice = input("\n선택 (번호): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            chosen = options[int(choice) - 1][0]
            print(f"  → {chosen}")
            return chosen
        print(f"  1~{len(options)} 중 하나를 입력하세요.")


# ════════════════════════════════════════════════════════════════
# 데이터 조회
# ════════════════════════════════════════════════════════════════

def fetch_data(intersections: list[dict],
               periods: list[tuple[date, date, str]],
               hours: list[int]) -> list[dict]:
    """API에서 보정된 교통량 조회 → list[dict] 반환
    각 dict: {date, hour, node_name, approach_name, traffic_volume}
    복수 기간을 순환하여 조회하고 합산한다.
    """
    rows: list[dict] = []
    for inter in intersections:
        node_id   = inter["node_id"]
        node_name = inter["name"]

        for ds, de, _ in periods:
            resp = _api_post("/corrected-traffic", {
                "node_ids":   [node_id],
                "date_start": ds.isoformat(),
                "date_end":   de.isoformat(),
                "hours":      hours,
            })
            for slot in resp["slots"]:
                rows.append({
                    "date":           date.fromisoformat(slot["date"]),
                    "hour":           slot["hour"],
                    "node_name":      node_name,
                    "approach_name":  slot["approach_name"],
                    "traffic_volume": slot["traffic_volume"],
                })

    # 정렬: 교차로 → 날짜 → 시간 → 방향
    rows.sort(key=lambda r: (r["node_name"], r["date"], r["hour"], r["approach_name"]))
    return rows


def aggregate_by_node(rows: list[dict]) -> list[dict]:
    """방향별 교통량을 교차로+시간 기준으로 합산"""
    agg: dict[tuple, int | None] = defaultdict(lambda: None)

    for r in rows:
        key = (r["date"], r["hour"], r["node_name"])
        val = r["traffic_volume"]
        if val is not None:
            prev = agg[key]
            agg[key] = (prev or 0) + val

    result: list[dict] = []
    for (d, h, node), total in agg.items():
        result.append({
            "date":          d,
            "hour":          h,
            "node_name":     node,
            "approach_name": "",
            "traffic_volume": total,
        })
    result.sort(key=lambda r: (r["node_name"], r["date"], r["hour"]))
    return result


# ════════════════════════════════════════════════════════════════
# Excel 저장
# ════════════════════════════════════════════════════════════════

def _cell_value(row: dict, attr: str) -> str | int | None:
    if attr == "시간":
        return f"{row['date']} {row['hour']:02d}:00"
    if attr == "교차로명":
        return row["node_name"]
    if attr == "방향명":
        return row["approach_name"]
    if attr == "교통량":
        return row["traffic_volume"]
    return None


def _style_header(ws, ncols: int) -> None:
    fill = PatternFill("solid", fgColor="D9D9D9")
    font = Font(bold=True)
    align = Alignment(horizontal="center", vertical="center")
    for col in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill  = fill
        cell.font  = font
        cell.alignment = align


def _auto_width(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value is not None else ""
                # 한글 2바이트 고려
                length = sum(2 if ord(c) > 127 else 1 for c in val)
                max_len = max(max_len, length)
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)


def save_excel(data: list[dict], attrs: list[str], sheet_type: str, path: Path) -> None:
    """데이터를 Excel 파일로 저장"""
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    if sheet_type == "교차로별":
        # 교차로명별로 그룹화
        groups: dict[str, list[dict]] = defaultdict(list)
        for r in data:
            groups[r["node_name"]].append(r)
        for node_name, rows in groups.items():
            ws = wb.create_sheet(title=_safe_sheet_name(node_name))
            ws.append(attrs)
            _style_header(ws, len(attrs))
            for r in rows:
                ws.append([_cell_value(r, a) for a in attrs])
            _auto_width(ws)

    elif sheet_type == "방향별":
        # (교차로명_방향명)별로 그룹화
        groups: dict[str, list[dict]] = defaultdict(list)
        for r in data:
            key = f"{r['node_name']}_{r['approach_name']}"
            groups[key].append(r)
        for sheet_key, rows in groups.items():
            ws = wb.create_sheet(title=_safe_sheet_name(sheet_key))
            ws.append(attrs)
            _style_header(ws, len(attrs))
            for r in rows:
                ws.append([_cell_value(r, a) for a in attrs])
            _auto_width(ws)

    else:  # 단일 시트
        ws = wb.create_sheet(title="교통량")
        ws.append(attrs)
        _style_header(ws, len(attrs))
        for r in data:
            ws.append([_cell_value(r, a) for a in attrs])
        _auto_width(ws)

    wb.save(path)
    print(f"\n  Excel 저장 완료: {path}")


def _safe_sheet_name(name: str) -> str:
    """Excel 시트 이름 제한(31자, 특수문자) 처리"""
    forbidden = r'\/*?:[]'
    for ch in forbidden:
        name = name.replace(ch, "_")
    return name[:31]


# ════════════════════════════════════════════════════════════════
# 그래프 출력 — Chart.js 기반 커스텀 HTML
# ════════════════════════════════════════════════════════════════

def _node_total_rows(rows: list[dict], node_name: str) -> list[dict]:
    """특정 교차로의 방향 합산 시계열"""
    agg: dict[tuple, int | None] = defaultdict(lambda: None)
    for r in rows:
        if r["node_name"] != node_name:
            continue
        key = (r["date"], r["hour"])
        val = r["traffic_volume"]
        if val is not None:
            agg[key] = (agg[key] or 0) + val

    result = []
    for (d, h), total in sorted(agg.items()):
        result.append({"date": d, "hour": h, "traffic_volume": total})
    return result


def _make_chart_dataset(rows: list[dict], label: str, color: str,
                         fill: bool = False) -> dict:
    """Chart.js dataset dict 생성.
    빠진 시간대는 null로 채워 비연속 구간에서 선을 자동으로 끊는다.
    """
    val_map: dict[datetime, int | None] = {}
    for r in rows:
        dt = datetime(r["date"].year, r["date"].month, r["date"].day, r["hour"])
        val_map[dt] = r["traffic_volume"]

    if not val_map:
        return {"label": label, "data": [], "borderColor": color}

    data_pts = [
        {"x": dt.strftime("%m/%d %H:%M"), "y": v}
        for dt, v in sorted(val_map.items())
    ]

    h = color.lstrip("#")
    rv, gv, bv = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    return {
        "label":                    label,
        "data":                     data_pts,
        "borderColor":              color,
        "borderWidth":              2.5,
        "tension":                  0.35,
        "pointRadius":              0,
        "pointHoverRadius":         6,
        "pointHoverBorderWidth":    2.5,
        "pointHoverBorderColor":    "#FFFFFF",
        "pointHoverBackgroundColor": color,
        "spanGaps":                 True,
        "_hasFill":                 fill,
        "_fillTop":                 f"rgba({rv},{gv},{bv},0.18)",
        "_fillBot":                 f"rgba({rv},{gv},{bv},0)",
    }


def _calc_stats(rows: list[dict]) -> dict:
    """최대·평균·데이터 수 계산"""
    vals = [r["traffic_volume"] for r in rows if r.get("traffic_volume") is not None]
    if not vals:
        return {"max": 0, "avg": 0, "n": 0}
    return {"max": max(vals), "avg": round(sum(vals) / len(vals)), "n": len(vals)}


# ── HTML 템플릿 ───────────────────────────────────────────────────────────────

_PAGE_TEMPLATE = """\
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns@3.0.0/dist/chartjs-adapter-date-fns.bundle.min.js"></script>
<style>
:root {
  --bg:     #F0F2F5;
  --card:   #FFFFFF;
  --border: #E2E8F0;
  --t1:     #1A202C;
  --t2:     #4A5568;
  --t3:     #A0AEC0;
  --accent: #2B6CB0;
  --r:      14px;
  --sans:   'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
  --mono:   'JetBrains Mono', 'Consolas', monospace;
}
*,*::before,*::after { box-sizing:border-box; margin:0; padding:0; }
html { font-size:14px; -webkit-font-smoothing:antialiased; }
body {
  font-family: var(--sans);
  background: var(--bg);
  color: var(--t1);
  padding: 44px 36px;
  min-height: 100vh;
}
.wrap { max-width:1600px; margin:0 auto; }

/* ── 상단 헤더 ── */
.topbar {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 16px;
  margin-bottom: 28px;
}
.eyebrow {
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: var(--accent);
  margin-bottom: 8px;
}
.page-title {
  font-size: 28px;
  font-weight: 700;
  letter-spacing: -0.03em;
  color: var(--t1);
  line-height: 1;
}
.page-desc {
  font-size: 13px;
  color: var(--t2);
  margin-top: 6px;
  font-weight: 400;
}
.period-badge {
  align-self: flex-start;
  background: #EBF4FF;
  color: var(--accent);
  font-family: var(--mono);
  font-size: 12px;
  font-weight: 500;
  padding: 7px 16px;
  border-radius: 100px;
  border: 1px solid #BEE3F8;
  letter-spacing: 0.04em;
  white-space: nowrap;
  margin-top: 4px;
}

/* ── 카드 ── */
.card {
  background: var(--card);
  border-radius: var(--r);
  border: 1px solid var(--border);
  box-shadow:
    0 1px 2px rgba(0,0,0,0.04),
    0 4px 12px rgba(0,0,0,0.05),
    0 16px 40px rgba(0,0,0,0.04);
  padding: 32px 28px 24px;
  animation: rise .5s cubic-bezier(.22,1,.36,1) both;
}
@keyframes rise {
  from { opacity:0; transform:translateY(18px); }
  to   { opacity:1; transform:translateY(0); }
}

/* ── 범례 ── */
.legend {
  display: flex;
  flex-wrap: wrap;
  gap: 8px 24px;
  margin-bottom: 24px;
  padding-bottom: 20px;
  border-bottom: 1px solid var(--border);
}
.legend-item {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
  font-weight: 500;
  color: var(--t2);
}
.legend-swatch {
  width: 28px;
  height: 3px;
  border-radius: 99px;
  flex-shrink: 0;
}

/* ── 차트 제목 ── */
.card-title {
  font-size: 15px;
  font-weight: 600;
  color: var(--t1);
  letter-spacing: -0.02em;
  margin-bottom: 16px;
}

/* ── 차트 ── */
.chart-outer {
  overflow-x: auto;
  overflow-y: hidden;
  width: 100%;
}
.chart-inner {
  position: relative;
  height: 440px;
  min-width: 100%;
}
.chart-inner canvas { display: block; }

/* ── 통계 ── */
.stats {
  display: flex;
  flex-wrap: wrap;
  gap: 0;
  margin-top: 22px;
  padding-top: 20px;
  border-top: 1px solid var(--border);
}
.stat {
  padding-right: 32px;
  margin-right: 32px;
  border-right: 1px solid var(--border);
}
.stat:last-child { border-right:none; padding-right:0; margin-right:0; }
.stat-label {
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.1em;
  text-transform: uppercase;
  color: var(--t3);
  margin-bottom: 5px;
}
.stat-val {
  font-family: var(--mono);
  font-size: 24px;
  font-weight: 500;
  color: var(--t1);
  letter-spacing: -0.03em;
  line-height: 1;
}
.stat-unit {
  font-family: var(--sans);
  font-size: 12px;
  font-weight: 400;
  color: var(--t3);
  margin-left: 4px;
}

/* ── 푸터 ── */
.footer {
  text-align: center;
  font-size: 11px;
  color: var(--t3);
  margin-top: 20px;
  letter-spacing: 0.03em;
}
</style>
</head>
<body>
<div class="wrap">

  <div class="topbar">
    <div>
      <div class="eyebrow">교통량 분석 리포트</div>
      <h1 class="page-title">__NODE_NAME__</h1>
      <p class="page-desc">__DESC__</p>
    </div>
    <span class="period-badge">__PERIOD_STR__</span>
  </div>

  __CHARTS__

  <p class="footer">교통량 추출 시스템&nbsp;&nbsp;·&nbsp;&nbsp;__GENERATED__</p>
</div>
</body>
</html>"""


_CHART_JS = """\
(function(){
  var DS = __DATASETS__;
  var canvas = document.getElementById('__CID__');
  var inner  = document.getElementById('__CID__wrap');
  var nHours = __N_HOURS__;
  var PX_PER_HOUR = 20;
  var VIEWPORT_HOURS = 48;
  var MAX_CANVAS_PX = Math.floor(14000 / (window.devicePixelRatio || 1));
  if (nHours > VIEWPORT_HOURS) {
    var w = Math.min(nHours * PX_PER_HOUR, MAX_CANVAS_PX);
    inner.style.width = w + 'px';
  }
  var ctx = canvas.getContext('2d');
  DS.forEach(function(ds){
    if (ds._hasFill) {
      var g = ctx.createLinearGradient(0, 0, 0, 440);
      g.addColorStop(0, ds._fillTop);
      g.addColorStop(1, ds._fillBot);
      ds.backgroundColor = g;
      ds.fill = 'origin';
    }
  });
  new Chart(ctx, {
    type: 'line',
    data: { datasets: DS },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: 1,
      animation: { duration: 900, easing: 'easeInOutCubic' },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1A202C',
          titleColor: '#718096',
          bodyColor:  '#EDF2F7',
          borderColor: '#2D3748',
          borderWidth: 1,
          padding: { x:14, y:10 },
          cornerRadius: 10,
          titleFont: { family:"'JetBrains Mono',monospace", size:11 },
          bodyFont:  { family:"'Noto Sans KR',sans-serif", size:13, weight:'500' },
          callbacks: {
            title: function(items) {
              if (!items.length) return '';
              return items[0].label;
            },
            label: function(item) {
              return item.parsed.y != null
                ? '  '+item.dataset.label+': '+item.parsed.y.toLocaleString()+' 대'
                : '  '+item.dataset.label+': \u2014';
            },
          },
        },
      },
      scales: {
        x: {
          type: 'category',
          grid:   { color:'#EDF2F7', tickLength:4 },
          border: { color:'#E2E8F0' },
          ticks: {
            color: '#A0AEC0',
            font:  { family:"'JetBrains Mono',monospace", size:10 },
            maxRotation: 30,
            autoSkip: true,
            maxTicksLimit: 18,
          },
        },
        y: {
          beginAtZero: true,
          grid:   { color:'#EDF2F7' },
          border: { display:false },
          ticks: {
            color: '#A0AEC0',
            font:  { family:"'JetBrains Mono',monospace", size:10 },
            padding: 8,
            callback: function(v){ return v.toLocaleString(); },
          },
        },
      },
      elements: {
        line: { tension:0.35, borderWidth:2.5, borderCapStyle:'round', borderJoinStyle:'round' },
        point: { radius:0, hoverRadius:6, hoverBorderWidth:2.5, hoverBorderColor:'#fff' },
      },
      spanGaps: true,
    },
  });
})();"""


def _make_chart_block(chart_id: str, card_title: str,
                      datasets: list[dict], stat_rows: list[dict]) -> str:
    """단일 Chart.js 카드 HTML(+ 인라인 스크립트) 생성.
    48시간 초과 시 canvas 컨테이너를 넓혀 가로 스크롤 활성화.
    """
    # 범례 (2개 이상일 때만)
    if len(datasets) > 1:
        items = "".join(
            f'<div class="legend-item">'
            f'<div class="legend-swatch" style="background:{ds["borderColor"]}"></div>'
            f'{ds["label"]}</div>'
            for ds in datasets
        )
        legend_html = f'<div class="legend">{items}</div>'
    else:
        legend_html = ""

    # 통계
    s = _calc_stats(stat_rows)
    if s["n"] > 0:
        stats_html = (
            '<div class="stats">'
            f'<div class="stat"><div class="stat-label">최대 교통량</div>'
            f'<div class="stat-val">{s["max"]:,}<span class="stat-unit">대</span></div></div>'
            f'<div class="stat"><div class="stat-label">평균 교통량</div>'
            f'<div class="stat-val">{s["avg"]:,}<span class="stat-unit">대</span></div></div>'
            f'<div class="stat"><div class="stat-label">데이터 수</div>'
            f'<div class="stat-val">{s["n"]:,}<span class="stat-unit">개</span></div></div>'
            '</div>'
        )
    else:
        stats_html = ""

    # 시간 범위 계산 (null 채움 포함)
    n_hours = max((len(ds["data"]) for ds in datasets if ds.get("data")), default=0)

    ds_json = json.dumps(datasets, ensure_ascii=False)
    js = (
        _CHART_JS
        .replace("__CID__", chart_id)
        .replace("__N_HOURS__", str(n_hours))
        .replace("__DATASETS__", ds_json)
    )

    title_html = f'<div class="card-title">{card_title}</div>' if card_title else ""

    return (
        f'<div class="card">'
        f'{title_html}'
        f'{legend_html}'
        f'<div class="chart-outer">'
        f'<div class="chart-inner" id="{chart_id}wrap">'
        f'<canvas id="{chart_id}"></canvas>'
        f'</div></div>'
        f'{stats_html}'
        f'</div>'
        f'<script>{js}</script>'
    )


def _write_graph_html(
    path: Path,
    node_name: str,
    period_str: str,
    desc: str,
    charts: list[dict],
) -> None:
    """Chart.js 그래프 HTML 파일 저장.

    charts: [{"title": str, "datasets": [...], "stat_rows": [...]}, ...]
    """
    charts_html_parts = []
    for idx, c in enumerate(charts):
        cid = f"ch{idx}"
        block = _make_chart_block(cid, c["title"], c["datasets"], c["stat_rows"])
        charts_html_parts.append(block)

    charts_html = '\n\n  '.join(charts_html_parts)
    generated   = datetime.now().strftime("%Y.%m.%d %H:%M")

    html = (
        _PAGE_TEMPLATE
        .replace("__TITLE__",      f"{node_name} 교통량 ({period_str})")
        .replace("__NODE_NAME__",  node_name)
        .replace("__DESC__",       desc)
        .replace("__PERIOD_STR__", period_str)
        .replace("__CHARTS__",     charts_html)
        .replace("__GENERATED__",  generated)
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")


def build_graphs(rows: list[dict], has_dir: bool,
                 node_names: list[str], period_str: str) -> None:
    """그래프 HTML 파일 생성"""

    if not has_dir:
        # ── 교차로 합산 그래프 (방향 없음) ──────────────────────────────
        for node_name in node_names:
            node_rows = [r for r in rows if r["node_name"] == node_name]
            if not node_rows:
                continue

            ds = _make_chart_dataset(node_rows, node_name, NAVY, fill=True)
            out_path = GRAPH_DIR / f"{node_name}_{period_str}" / f"{node_name}_{period_str}.html"
            _write_graph_html(
                path       = out_path,
                node_name  = node_name,
                period_str = period_str,
                desc       = "교차로 전체 교통량",
                charts     = [{"title": "", "datasets": [ds], "stat_rows": node_rows}],
            )
            print(f"  그래프 저장: {out_path}")

    else:
        # ── 방향별 그래프 ──────────────────────────────────────────────
        for node_name in node_names:
            node_rows = [r for r in rows if r["node_name"] == node_name]
            if not node_rows:
                continue

            # 방향 목록 (순서 유지)
            dirs_seen: list[str] = []
            for r in node_rows:
                if r["approach_name"] not in dirs_seen:
                    dirs_seen.append(r["approach_name"])

            folder = GRAPH_DIR / f"{node_name}_{period_str}"

            # 1. 통합 그래프: 그래프①교차로 전체(채움) + 그래프②방향별 멀티라인
            total_rows = _node_total_rows(node_rows, node_name)
            total_ds   = _make_chart_dataset(total_rows, f"{node_name} 전체", NAVY, fill=True)

            dir_ds_list = []
            for i, dir_name in enumerate(dirs_seen):
                dir_rows = [r for r in node_rows if r["approach_name"] == dir_name]
                dir_ds_list.append(
                    _make_chart_dataset(dir_rows, dir_name, DIR_COLORS[i % len(DIR_COLORS)])
                )

            _write_graph_html(
                path       = folder / f"{node_name}_{period_str}.html",
                node_name  = node_name,
                period_str = period_str,
                desc       = "방향별 교통량 통합",
                charts     = [
                    {"title": "교차로 교통량",  "datasets": [total_ds],  "stat_rows": total_rows},
                    {"title": "방향별 교통량",  "datasets": dir_ds_list, "stat_rows": node_rows},
                ],
            )
            print(f"  그래프 저장: {folder / f'{node_name}_{period_str}.html'}")

            # 2. 방향별 개별 그래프 (채움 영역)
            for i, dir_name in enumerate(dirs_seen):
                dir_rows = [r for r in node_rows if r["approach_name"] == dir_name]
                color = DIR_COLORS[i % len(DIR_COLORS)]
                ds = _make_chart_dataset(dir_rows, dir_name, color, fill=True)
                safe_dir = re.sub(r'[\\/*?:"<>|]', "_", dir_name)
                dir_path = folder / f"{node_name}_{safe_dir}_{period_str}.html"
                _write_graph_html(
                    path       = dir_path,
                    node_name  = node_name,
                    period_str = period_str,
                    desc       = f"{dir_name} 방향 교통량",
                    charts     = [{"title": "", "datasets": [ds], "stat_rows": dir_rows}],
                )
                print(f"  그래프 저장: {dir_path}")


# ════════════════════════════════════════════════════════════════
# 파일명 생성
# ════════════════════════════════════════════════════════════════

def make_filename(intersections: list[dict], period_str: str) -> str:
    """[교차로이름]_[기간]  (교차로 이름은 언더스코어로 결합)"""
    names = "_".join(d["name"] for d in intersections)
    return f"{names}_{period_str}"


# ════════════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 55)
    print("   교통량 추출 프로그램")
    print("=" * 55)

    # API 서버 접속 확인
    print(f"\nAPI 서버 연결 확인 중... ({API_BASE_URL})")
    try:
        _api_get("/intersections")
    except urllib.error.URLError as e:
        print(f"\nAPI 서버에 접속할 수 없습니다: {e}")
        print(f"  → api_server.py가 실행 중인지 확인하세요. ({API_BASE_URL})")
        sys.exit(1)
    print("API 서버 연결 성공")

    # 1. 교차로 입력
    intersections = input_intersections()

    # 2. 기간 입력
    periods = input_period()
    period_str = "_".join(p for _, _, p in periods)

    # 3. 시간대 입력
    hours = input_time_range()

    # 4. 속성 선택
    attrs, has_node, has_dir = input_attributes()

    # 5. 시트 설정
    #    속성 2번 (방향 합산) → has_dir=False → 교차로별만 가능 → 자동 고정
    sheet_type = input_sheet_type(has_node, has_dir)

    # 6. 데이터 조회 (API 경유, 보정값 적용)
    print("\n데이터 조회 중...")
    raw_rows = fetch_data(intersections, periods, hours)
    if not raw_rows:
        print("조회된 데이터가 없습니다.")
        return

    # 방향 합산이 필요한 경우
    if not has_dir:
        data = aggregate_by_node(raw_rows)
    else:
        data = raw_rows

    print(f"  조회된 레코드 수: {len(data)}")

    # 7. Excel 저장
    filename = make_filename(intersections, period_str)
    excel_path = EXCEL_DIR / f"{filename}.xlsx"
    save_excel(data, attrs, sheet_type, excel_path)

    # 8. 그래프 출력
    print("\n[6단계] 그래프 출력")
    graph_yn = input("  그래프를 출력하시겠습니까? (y/n) [n]: ").strip().lower()
    if graph_yn == "y":
        node_names = [d["name"] for d in intersections]
        if len(periods) == 1:
            # 단일 기간: 기존과 동일
            build_graphs(raw_rows, has_dir, node_names, periods[0][2])
        else:
            # 복수 기간: 각 기간별로 별도 파일 생성
            for ds, de, p_str in periods:
                period_rows = [r for r in raw_rows if ds <= r["date"] <= de]
                if period_rows:
                    print(f"\n  기간 {p_str} 그래프 생성 중...")
                    build_graphs(period_rows, has_dir, node_names, p_str)
                else:
                    print(f"  기간 {p_str}: 데이터 없음, 건너뜁니다.")
        print("  그래프 저장 완료.")
    else:
        print("  그래프 생성을 건너뜁니다.")

    print("\n완료.")


if __name__ == "__main__":
    main()
