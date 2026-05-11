from __future__ import annotations

from datetime import date, datetime

import api_server as legacy


class LegacyTrafficAnalysisGateway:
    """Infrastructure adapter around the existing analysis module boundary."""

    def initialize(self) -> None:
        legacy._holiday_dates = legacy.ad.load_holidays()
        legacy._ensure_drct_direction_presence_cache()

    def corrected_traffic(self, request: dict) -> dict:
        return legacy._fetch_corrected_traffic(request)

    def corrected_traffic_drct(self, request) -> dict:
        return legacy._fetch_corrected_traffic_drct(request)

    def raw_traffic(self, request) -> dict:
        return legacy._fetch_raw_traffic(request)

    def raw_traffic_drct(self, request) -> dict:
        return legacy._fetch_raw_traffic_drct(request)

    def raw_traffic_vknd(self, request) -> dict:
        return legacy._fetch_raw_traffic_vknd(request)

    def corrected_traffic_vknd(self, request) -> dict:
        return legacy._fetch_corrected_traffic_vknd(request)

    def anomaly_daily_summary(self, request) -> dict:
        return legacy._fetch_anomaly_daily_summary(request)

    def run_job(self, job: dict) -> None:
        if legacy._holiday_dates is None:
            self._fail(job, "서버 초기화가 완료되지 않았습니다. 잠시 후 다시 시도하세요.")
            return

        job["status"] = "running"
        job["started_at"] = datetime.now()
        job["progress"] = {"current": 0, "total": len(job["request"]["node_ids"])}

        conn = legacy.ad.connect_db()
        try:
            req = job["request"]
            date_start = date.fromisoformat(req["date_start"])
            date_end = date.fromisoformat(req["date_end"])
            hours = req["hours"]
            baseline_years = legacy.ad.select_baseline_years(date_start.year)
            fallback_year = legacy.ad.select_fallback_year(date_start.year)
            adj_periods = legacy.ad.get_adjacent_month_periods(date_start, date_end)
            id_to_name = legacy._get_id_to_name_map(conn)

            all_slots: list[dict] = []
            for index, node_id in enumerate(req["node_ids"], start=1):
                norm_node_id = legacy._normalize_node_id(node_id)
                node_name = id_to_name.get(norm_node_id, str(node_id))
                node_results, baselines, _ = legacy.ad.analyse_node(
                    conn,
                    norm_node_id,
                    node_name,
                    date_start,
                    date_end,
                    hours,
                    baseline_years,
                    fallback_year,
                    adj_periods,
                    legacy._holiday_dates,
                )
                all_slots.extend(
                    legacy._serialize_slot(row, norm_node_id, baselines)
                    for row in node_results
                )
                job["progress"] = {"current": index, "total": len(req["node_ids"])}

            job["status"] = "done"
            job["finished_at"] = datetime.now()
            job["result"] = {"slots": all_slots}
        except Exception as exc:
            self._fail(job, str(exc))
        finally:
            conn.close()

    def _fail(self, job: dict, message: str) -> None:
        job["status"] = "failed"
        job["finished_at"] = datetime.now()
        job["error"] = message


class LegacyReferenceDataGateway:
    def intersections(self) -> list[dict]:
        return legacy._fetch_intersections()

    def vehicle_kinds(self) -> list[dict]:
        return legacy._fetch_vehicle_kinds()
