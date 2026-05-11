from __future__ import annotations

from .interfaces import ReferenceDataGateway, TrafficAnalysisGateway


class TrafficQueryUseCase:
    def __init__(self, gateway: TrafficAnalysisGateway) -> None:
        self._gateway = gateway

    def corrected_traffic(self, request) -> dict:
        payload = {
            "node_ids": request.node_ids,
            "date_start": request.date_start,
            "date_end": request.date_end,
            "hours": request.resolve_hours(),
        }
        return self._gateway.corrected_traffic(payload)

    def corrected_traffic_drct(self, request) -> dict:
        return self._gateway.corrected_traffic_drct(request)

    def raw_traffic(self, request) -> dict:
        return self._gateway.raw_traffic(request)

    def raw_traffic_vknd(self, request) -> dict:
        return self._gateway.raw_traffic_vknd(request)

    def corrected_traffic_vknd(self, request) -> dict:
        return self._gateway.corrected_traffic_vknd(request)

    def anomaly_daily_summary(self, request) -> dict:
        return self._gateway.anomaly_daily_summary(request)


class ReferenceDataUseCase:
    def __init__(self, gateway: ReferenceDataGateway) -> None:
        self._gateway = gateway
        self._intersections_cache: list[dict] | None = None

    def vehicle_kinds(self) -> dict:
        items = self._gateway.vehicle_kinds()
        return {"count": len(items), "items": items}

    def intersections(self, refresh: bool = False) -> dict:
        if self._intersections_cache is None or refresh:
            self._intersections_cache = self._gateway.intersections()
        return {
            "count": len(self._intersections_cache),
            "items": self._intersections_cache,
        }
