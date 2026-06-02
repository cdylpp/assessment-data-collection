from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Set, Tuple


@dataclass(frozen=True)
class EvaluationIssue:
    severity: str
    origin: str
    message: str

    def format(self) -> str:
        return "{0}: {1} ({2})".format(self.severity.upper(), self.message, self.origin)


@dataclass(frozen=True)
class EvaluationResult:
    issues: Sequence[EvaluationIssue]

    @property
    def errors(self) -> List[EvaluationIssue]:
        return [issue for issue in self.issues if issue.severity == "error"]

    @property
    def warnings(self) -> List[EvaluationIssue]:
        return [issue for issue in self.issues if issue.severity == "warning"]

    @property
    def is_valid(self) -> bool:
        return not self.errors

    def error_message(self) -> str:
        if not self.issues:
            return "No evaluation issues found."
        return "\n".join(issue.format() for issue in self.issues)

    def raise_for_errors(self) -> None:
        if self.errors:
            raise ValueError(self.error_message())


class Evaluator:
    def __init__(
        self,
        *,
        metrics_doc: Mapping[str, Any],
        evolutions_doc: Mapping[str, Any],
        events_doc: Optional[Mapping[str, Any]] = None,
    ) -> None:
        self.metrics_doc = metrics_doc
        self.evolutions_doc = evolutions_doc
        self.events_doc = events_doc or {}

    def evaluate(self, event_id: Optional[str] = None) -> EvaluationResult:
        issues = []  # type: List[EvaluationIssue]
        metric_ids = self._metric_ids(issues)
        evolution_ids = self._evolution_ids(issues, metric_ids)
        self._event_issues(issues, evolution_ids, event_id)
        return EvaluationResult(issues=issues)

    def validate_event(self, event_id: Optional[str] = None) -> EvaluationResult:
        return self.evaluate(event_id=event_id)

    def _metric_ids(self, issues: List[EvaluationIssue]) -> Set[str]:
        metrics = self.metrics_doc.get("metrics")
        if not isinstance(metrics, list):
            issues.append(
                EvaluationIssue(
                    severity="error",
                    origin="metrics.yaml:metrics",
                    message="Expected a list of metric definitions.",
                )
            )
            return set()

        metric_ids = set()  # type: Set[str]
        for index, metric in enumerate(metrics):
            origin = "metrics.yaml:metrics[{0}]".format(index)
            if not isinstance(metric, dict):
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Metric entry must be a mapping.",
                    )
                )
                continue
            metric_id = metric.get("metric_id")
            if not metric_id:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Metric is missing required field 'metric_id'.",
                    )
                )
                continue
            metric_id = str(metric_id)
            if metric_id in metric_ids:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Duplicate metric_id '{0}'.".format(metric_id),
                    )
                )
            metric_ids.add(metric_id)
        return metric_ids

    def _evolution_ids(
        self, issues: List[EvaluationIssue], metric_ids: Set[str]
    ) -> Set[str]:
        evolutions = self.evolutions_doc.get("evolutions")
        if not isinstance(evolutions, list):
            issues.append(
                EvaluationIssue(
                    severity="error",
                    origin="evolutions.yaml:evolutions",
                    message="Expected a list of evolution definitions.",
                )
            )
            return set()

        evolution_ids = set()  # type: Set[str]
        for index, evolution in enumerate(evolutions):
            origin = "evolutions.yaml:evolutions[{0}]".format(index)
            if not isinstance(evolution, dict):
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Evolution entry must be a mapping.",
                    )
                )
                continue

            evolution_id = evolution.get("evolution_id")
            if not evolution_id:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Evolution is missing required field 'evolution_id'.",
                    )
                )
                continue
            evolution_id = str(evolution_id)
            if evolution_id in evolution_ids:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=origin,
                        message="Duplicate evolution_id '{0}'.".format(evolution_id),
                    )
                )
            evolution_ids.add(evolution_id)

            if "metric_id" in evolution and "metric_ids" not in evolution:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin="{0}.metric_id".format(origin),
                        message=(
                            "Evolution '{0}' uses 'metric_id'; expected "
                            "'metric_ids' list."
                        ).format(evolution_id),
                    )
                )

            raw_metric_ids = evolution.get("metric_ids")
            if not raw_metric_ids:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin="{0}.metric_ids".format(origin),
                        message="Evolution '{0}' has no metric_ids.".format(
                            evolution_id
                        ),
                    )
                )
                raw_metric_ids = []
            if raw_metric_ids and not isinstance(raw_metric_ids, list):
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin="{0}.metric_ids".format(origin),
                        message="Evolution '{0}' metric_ids must be a list.".format(
                            evolution_id
                        ),
                    )
                )
                raw_metric_ids = []

            configured_metric_ids = set(str(metric_id) for metric_id in raw_metric_ids)
            occurrences = evolution.get("metric_occurrences", {}) or {}
            if not isinstance(occurrences, dict):
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin="{0}.metric_occurrences".format(origin),
                        message="Evolution '{0}' metric_occurrences must be a mapping.".format(
                            evolution_id
                        ),
                    )
                )
                occurrences = {}
            for occurrence_metric_id, raw_count in occurrences.items():
                occurrence_metric_id = str(occurrence_metric_id)
                if occurrence_metric_id not in configured_metric_ids:
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin="{0}.metric_occurrences.{1}".format(
                                origin, occurrence_metric_id
                            ),
                            message=(
                                "Evolution '{0}' defines occurrences for metric "
                                "'{1}', but that metric is not listed in metric_ids."
                            ).format(evolution_id, occurrence_metric_id),
                        )
                    )
                if not self._is_positive_int(raw_count):
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin="{0}.metric_occurrences.{1}".format(
                                origin, occurrence_metric_id
                            ),
                            message="Occurrence count must be a positive integer.",
                        )
                    )

            for metric_id in configured_metric_ids:
                if metric_id not in metric_ids:
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin="{0}.metric_ids".format(origin),
                            message=(
                                "Evolution '{0}' references undefined metric_id "
                                "'{1}'."
                            ).format(evolution_id, metric_id),
                        )
                    )
        return evolution_ids

    def _event_issues(
        self,
        issues: List[EvaluationIssue],
        evolution_ids: Set[str],
        selected_event_id: Optional[str],
    ) -> None:
        events = self.events_doc.get("events")
        if events is None:
            return
        if not isinstance(events, list):
            issues.append(
                EvaluationIssue(
                    severity="error",
                    origin="event.yaml:events",
                    message="Expected a list of event definitions.",
                )
            )
            return

        seen_event_ids = set()  # type: Set[str]
        selected_found = selected_event_id is None
        for event_index, event in enumerate(events):
            event_origin = "event.yaml:events[{0}]".format(event_index)
            if not isinstance(event, dict):
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=event_origin,
                        message="Event entry must be a mapping.",
                    )
                )
                continue

            event_id = event.get("id", event.get("event_id"))
            if not event_id:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=event_origin,
                        message="Event is missing required field 'id'.",
                    )
                )
                continue
            event_id = str(event_id)
            if event_id in seen_event_ids:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin=event_origin,
                        message="Duplicate event_id '{0}'.".format(event_id),
                    )
                )
            seen_event_ids.add(event_id)
            if event_id == selected_event_id:
                selected_found = True

            entries = event.get("evolutions")
            if not isinstance(entries, list) or not entries:
                issues.append(
                    EvaluationIssue(
                        severity="error",
                        origin="{0}.evolutions".format(event_origin),
                        message="Event '{0}' must define a non-empty evolutions list.".format(
                            event_id
                        ),
                    )
                )
                continue

            for entry_index, entry in enumerate(entries):
                origin = "{0}.evolutions[{1}]".format(event_origin, entry_index)
                evolution_id, entry_errors = self._parse_event_entry(entry)
                for entry_error in entry_errors:
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin=origin,
                            message=entry_error.format(event_id),
                        )
                    )
                if not evolution_id:
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin=origin,
                            message=(
                                "Event '{0}' evolution entry must include an 'id'."
                            ).format(event_id),
                        )
                    )
                    continue
                if evolution_id not in evolution_ids:
                    issues.append(
                        EvaluationIssue(
                            severity="error",
                            origin=origin,
                            message=(
                                "Event '{0}' references undefined evolution_id "
                                "'{1}'."
                            ).format(event_id, evolution_id),
                        )
                    )

        if not selected_found:
            issues.append(
                EvaluationIssue(
                    severity="error",
                    origin="event.yaml:events",
                    message="Requested event_id '{0}' was not found.".format(
                        selected_event_id
                    ),
                )
            )

    def _parse_event_entry(self, entry: Any) -> Tuple[Optional[str], List[str]]:
        if isinstance(entry, str):
            return entry, []
        if isinstance(entry, (list, tuple)):
            errors = []  # type: List[str]
            if len(entry) < 1 or len(entry) > 2:
                errors.append(
                    "Event '{0}' evolution entry must be [evolution_id] or "
                    "[evolution_id, display_name]."
                )
                return None, errors
            if len(entry) == 2 and entry[1] is not None and not isinstance(entry[1], str):
                errors.append(
                    "Event '{0}' evolution display_name must be a string when provided."
                )
            evolution_id = entry[0]
            return str(evolution_id) if evolution_id else None, errors
        if not isinstance(entry, dict):
            return None, [
                "Event '{0}' evolution entry must be a string, tuple, or mapping."
            ]
        evolution_id = entry.get("id", entry.get("evolution_id"))
        errors = []  # type: List[str]
        display_name = entry.get("display_name")
        if display_name is not None and not isinstance(display_name, str):
            errors.append(
                "Event '{0}' evolution display_name must be a string when provided."
            )
        return str(evolution_id) if evolution_id else None, errors

    def _is_positive_int(self, value: Any) -> bool:
        if isinstance(value, bool):
            return False
        try:
            parsed = int(value)
        except (TypeError, ValueError):
            return False
        return parsed >= 1


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate configured event, evolution, and metric references."
    )
    parser.add_argument(
        "--config",
        default="config/config.yaml",
        help="Path to canonical config file.",
    )
    parser.add_argument(
        "--event-id",
        default=None,
        help="Optional event_id to validate as the selected event.",
    )
    return parser.parse_args(argv)


def evaluate_config(config_path: Path, event_id: Optional[str] = None) -> EvaluationResult:
    from template_generator import TemplateGenerationRequest, load_generation_inputs

    loaded = load_generation_inputs(
        TemplateGenerationRequest(config_path=config_path.resolve(), event_id=event_id)
    )
    return Evaluator(
        metrics_doc=loaded.metrics_doc,
        evolutions_doc=loaded.evolutions_doc,
        events_doc=loaded.events_doc,
    ).evaluate(event_id=event_id)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    result = evaluate_config(Path(args.config), event_id=args.event_id)
    print(result.error_message())
    return 0 if result.is_valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
