from __future__ import annotations

from collections import Counter
from typing import Optional

from objects import (
    AssessmentConfig,
    EventDefinition,
    EventEvolutionInstance,
    EvolutionDefinition,
    WorkbookGenerationPlan,
)


class EventResolver:
    def __init__(self, config: AssessmentConfig) -> None:
        self.config = config

    def resolve(self, event_id: Optional[str] = None) -> WorkbookGenerationPlan:
        event = self._selected_event(event_id)
        if event is None:
            return self._all_evolutions_plan()

        instances = []
        valid_references = [
            reference for reference in event.evolution_refs if reference.evolution_id
        ]
        occurrence_counts = Counter(
            reference.evolution_id for reference in valid_references
        )
        occurrence_indexes = Counter()  # type: Counter[str]
        for reference in valid_references:
            occurrence_indexes[reference.evolution_id] += 1
            occurrence_index = occurrence_indexes[reference.evolution_id]
            occurrence_count = occurrence_counts[reference.evolution_id]
            source = self.config.evolutions_by_id.get(
                reference.evolution_id,
                EvolutionDefinition.placeholder(reference.evolution_id),
            )
            evolution = source.with_event_display(
                display_name=reference.display_name,
                occurrence_index=occurrence_index,
                occurrence_count=1,
            )
            instance_id = self._instance_id(
                event=event,
                evolution=evolution,
                occurrence_index=occurrence_index,
                occurrence_count=occurrence_count,
            )
            instances.append(
                EventEvolutionInstance(
                    event_id=event.event_id,
                    event_name=event.event_name,
                    instance_id=instance_id,
                    evolution=evolution,
                    event_occurrence_index=occurrence_index,
                    event_occurrence_count=occurrence_count,
                )
            )

        return WorkbookGenerationPlan(
            event_id=event.event_id,
            event_name=event.event_name,
            instances=tuple(instances),
        )

    def _selected_event(self, event_id: Optional[str]) -> Optional[EventDefinition]:
        if not self.config.events_by_id:
            return None
        if event_id is None:
            return next(iter(self.config.events_by_id.values()))
        event = self.config.events_by_id.get(str(event_id))
        if event is None:
            raise ValueError("Requested event_id '{0}' was not found".format(event_id))
        return event

    def _all_evolutions_plan(self) -> WorkbookGenerationPlan:
        instances = []
        for evolution in self.config.evolutions_by_id.values():
            instances.append(
                EventEvolutionInstance(
                    event_id=None,
                    event_name=None,
                    instance_id=evolution.evolution_id,
                    evolution=evolution,
                    event_occurrence_index=1,
                    event_occurrence_count=1,
                )
            )
        return WorkbookGenerationPlan(
            event_id=None,
            event_name=None,
            instances=tuple(instances),
        )

    def _instance_id(
        self,
        *,
        event: EventDefinition,
        evolution: EvolutionDefinition,
        occurrence_index: int,
        occurrence_count: int,
    ) -> str:
        if occurrence_count <= 1:
            return "{0}__{1}".format(event.event_id, evolution.evolution_id)
        return "{0}__{1}__{2}".format(
            event.event_id,
            evolution.evolution_id,
            occurrence_index,
        )
