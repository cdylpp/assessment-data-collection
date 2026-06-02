from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Mapping, Optional, Sequence, Tuple


def parse_positive_int(value: Any, label: str) -> int:
    if isinstance(value, bool):
        raise ValueError("{0} must be a positive integer".format(label))
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        raise ValueError("{0} must be a positive integer".format(label))
    if parsed < 1:
        raise ValueError("{0} must be at least 1".format(label))
    return parsed


def raw_metric_ids_from_evolution(data: Mapping[str, Any]) -> Tuple[str, ...]:
    raw_metric_ids = data.get("metric_ids")
    if raw_metric_ids is None and "metric_id" in data:
        raw_metric_ids = data.get("metric_id")
    if raw_metric_ids is None:
        return tuple()
    if not isinstance(raw_metric_ids, list):
        raise ValueError(
            "Evolution '{0}' metric_ids must be a list".format(
                data.get("evolution_id", "unknown")
            )
        )
    return tuple(str(metric_id) for metric_id in raw_metric_ids)


def occurrence_map_from_evolution(data: Mapping[str, Any]) -> Dict[str, int]:
    raw_occurrences = data.get("metric_occurrences", {})
    if raw_occurrences is None:
        raw_occurrences = {}
    if not isinstance(raw_occurrences, dict):
        raise ValueError(
            "Evolution '{0}' metric_occurrences must be a mapping".format(
                data.get("evolution_id", "unknown")
            )
        )

    occurrence_map = {}  # type: Dict[str, int]
    for metric_id, raw_count in raw_occurrences.items():
        occurrence_map[str(metric_id)] = parse_positive_int(
            raw_count,
            "Evolution '{0}' metric_occurrences.{1}".format(
                data.get("evolution_id", "unknown"), metric_id
            ),
        )
    return occurrence_map


@dataclass(frozen=True)
class MetricDefinition:
    metric_id: str
    display_name: str
    metric_type: str
    input_kind: str
    raw: Mapping[str, Any]

    @classmethod
    def from_mapping(cls, data: Mapping[str, Any]) -> "MetricDefinition":
        metric_id = str(data.get("metric_id", ""))
        return cls(
            metric_id=metric_id,
            display_name=str(data.get("display_name", metric_id)),
            metric_type=str(data.get("type", "")),
            input_kind=str(data.get("input_kind", "")),
            raw=data,
        )

    @classmethod
    def undefined(cls, metric_id: str) -> "MetricDefinition":
        return cls(
            metric_id=metric_id,
            display_name=metric_id,
            metric_type="text",
            input_kind="undefined",
            raw={
                "metric_id": metric_id,
                "display_name": metric_id,
                "type": "text",
                "input_kind": "undefined",
            },
        )


@dataclass(frozen=True)
class EvolutionDefinition:
    evolution_id: str
    display_name: str
    sheet_name: str
    metric_ids: Tuple[str, ...]
    metric_occurrences: Mapping[str, int]
    raw: Mapping[str, Any]

    @classmethod
    def from_mapping(cls, data: Mapping[str, Any]) -> "EvolutionDefinition":
        evolution_id = str(data.get("evolution_id", ""))
        metric_ids = raw_metric_ids_from_evolution(data)
        display_name = str(data.get("display_name", evolution_id))
        sheet_name = str(data.get("sheet_name", display_name or evolution_id))
        return cls(
            evolution_id=evolution_id,
            display_name=display_name,
            sheet_name=sheet_name,
            metric_ids=metric_ids,
            metric_occurrences=occurrence_map_from_evolution(data),
            raw=data,
        )

    @classmethod
    def placeholder(cls, evolution_id: str) -> "EvolutionDefinition":
        return cls(
            evolution_id=evolution_id,
            display_name=evolution_id,
            sheet_name=evolution_id,
            metric_ids=tuple(),
            metric_occurrences={},
            raw={
                "evolution_id": evolution_id,
                "display_name": evolution_id,
                "sheet_name": evolution_id,
                "metric_ids": [],
            },
        )

    def expanded_metric_ids(self) -> Tuple[str, ...]:
        metric_ids = []
        for metric_id in self.metric_ids:
            metric_ids.extend([metric_id] * self.metric_occurrences.get(metric_id, 1))
        return tuple(metric_ids)

    def validate_shape(self) -> None:
        if not self.evolution_id:
            raise ValueError("Evolution is missing evolution_id")
        for metric_id in self.metric_occurrences:
            if metric_id not in set(self.metric_ids):
                raise ValueError(
                    "Evolution '{0}' defines metric_occurrences for metric_id '{1}' "
                    "that is not listed in metric_ids".format(
                        self.evolution_id, metric_id
                    )
                )

    def with_event_display(
        self,
        *,
        display_name: Optional[str] = None,
        occurrence_index: int = 1,
        occurrence_count: int = 1,
    ) -> "EvolutionDefinition":
        if display_name is None and occurrence_count <= 1:
            return self
        raw = dict(self.raw)
        if display_name is None:
            display_name = "{0} {1}".format(self.display_name, occurrence_index)
            sheet_name = "{0} {1}".format(self.sheet_name, occurrence_index)
        else:
            sheet_name = display_name
        raw["display_name"] = display_name
        raw["sheet_name"] = sheet_name
        return EvolutionDefinition(
            evolution_id=self.evolution_id,
            display_name=str(raw["display_name"]),
            sheet_name=str(raw["sheet_name"]),
            metric_ids=self.metric_ids,
            metric_occurrences=self.metric_occurrences,
            raw=raw,
        )


@dataclass(frozen=True)
class EventEvolutionReference:
    evolution_id: str
    display_name: Optional[str]
    source_is_mapping: bool
    raw: Any

    @classmethod
    def from_raw(cls, raw: Any) -> "EventEvolutionReference":
        if isinstance(raw, str):
            return cls(
                evolution_id=raw,
                display_name=None,
                source_is_mapping=False,
                raw=raw,
            )
        if isinstance(raw, (list, tuple)):
            if not raw:
                return cls(
                    evolution_id="",
                    display_name=None,
                    source_is_mapping=False,
                    raw=raw,
                )
            display_name = None
            if len(raw) > 1 and raw[1] is not None:
                display_name = str(raw[1])
            return cls(
                evolution_id=str(raw[0]) if raw[0] else "",
                display_name=display_name,
                source_is_mapping=False,
                raw=raw,
            )
        if not isinstance(raw, dict):
            return cls(
                evolution_id="",
                display_name=None,
                source_is_mapping=False,
                raw=raw,
            )
        evolution_id = raw.get("id", raw.get("evolution_id", ""))
        display_name = raw.get("display_name")
        return cls(
            evolution_id=str(evolution_id) if evolution_id else "",
            display_name=str(display_name) if display_name is not None else None,
            source_is_mapping=True,
            raw=raw,
        )


@dataclass(frozen=True)
class EventDefinition:
    event_id: str
    event_name: str
    evolution_refs: Tuple[EventEvolutionReference, ...]
    raw: Mapping[str, Any]

    @classmethod
    def from_mapping(cls, data: Mapping[str, Any]) -> "EventDefinition":
        event_id = str(data.get("id", data.get("event_id", "")))
        raw_entries = data.get("evolutions", [])
        if not isinstance(raw_entries, list):
            raw_entries = []
        return cls(
            event_id=event_id,
            event_name=str(data.get("name", data.get("event_name", event_id))),
            evolution_refs=tuple(
                EventEvolutionReference.from_raw(entry) for entry in raw_entries
            ),
            raw=data,
        )


@dataclass(frozen=True)
class EventEvolutionInstance:
    event_id: Optional[str]
    event_name: Optional[str]
    instance_id: str
    evolution: EvolutionDefinition
    event_occurrence_index: int
    event_occurrence_count: int

    @property
    def evolution_id(self) -> str:
        return self.evolution.evolution_id

    @property
    def sheet_name(self) -> str:
        return self.evolution.sheet_name[:31]

    @property
    def display_name(self) -> str:
        return self.evolution.display_name

    def to_evolution_mapping(self) -> Dict[str, Any]:
        data = dict(self.evolution.raw)
        data["evolution_id"] = self.evolution.evolution_id
        data["display_name"] = self.evolution.display_name
        data["sheet_name"] = self.evolution.sheet_name
        data["metric_ids"] = list(self.evolution.metric_ids)
        data["metric_occurrences"] = dict(self.evolution.metric_occurrences)
        data["event_id"] = self.event_id
        data["event_name"] = self.event_name
        data["event_instance_id"] = self.instance_id
        data["event_occurrence_index"] = self.event_occurrence_index
        data["event_occurrence_count"] = self.event_occurrence_count
        return data


@dataclass(frozen=True)
class WorkbookGenerationPlan:
    event_id: Optional[str]
    event_name: Optional[str]
    instances: Tuple[EventEvolutionInstance, ...]


@dataclass(frozen=True)
class AssessmentConfig:
    config_doc: Mapping[str, Any]
    metrics_doc: Mapping[str, Any]
    evolutions_doc: Mapping[str, Any]
    events_doc: Optional[Mapping[str, Any]]
    metrics_by_id: Mapping[str, MetricDefinition]
    evolutions_by_id: Mapping[str, EvolutionDefinition]
    events_by_id: Mapping[str, EventDefinition]

    @classmethod
    def from_documents(
        cls,
        *,
        config_doc: Mapping[str, Any],
        metrics_doc: Mapping[str, Any],
        evolutions_doc: Mapping[str, Any],
        events_doc: Optional[Mapping[str, Any]] = None,
    ) -> "AssessmentConfig":
        metrics = metrics_doc.get("metrics", [])
        if not isinstance(metrics, list):
            metrics = []
        evolutions = evolutions_doc.get("evolutions", [])
        if not isinstance(evolutions, list):
            evolutions = []
        events = (events_doc or {}).get("events", [])
        if not isinstance(events, list):
            events = []

        metrics_by_id = {}  # type: Dict[str, MetricDefinition]
        for metric in metrics:
            if isinstance(metric, dict) and metric.get("metric_id"):
                definition = MetricDefinition.from_mapping(metric)
                metrics_by_id[definition.metric_id] = definition

        evolutions_by_id = {}  # type: Dict[str, EvolutionDefinition]
        for evolution in evolutions:
            if isinstance(evolution, dict) and evolution.get("evolution_id"):
                definition = EvolutionDefinition.from_mapping(evolution)
                evolutions_by_id[definition.evolution_id] = definition

        events_by_id = {}  # type: Dict[str, EventDefinition]
        for event in events:
            if isinstance(event, dict) and (event.get("id") or event.get("event_id")):
                definition = EventDefinition.from_mapping(event)
                events_by_id[definition.event_id] = definition

        return cls(
            config_doc=config_doc,
            metrics_doc=metrics_doc,
            evolutions_doc=evolutions_doc,
            events_doc=events_doc,
            metrics_by_id=metrics_by_id,
            evolutions_by_id=evolutions_by_id,
            events_by_id=events_by_id,
        )
