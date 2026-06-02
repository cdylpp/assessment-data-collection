# TODO

- [ ] Add support for Event definitions. Events are a collection of evolutions. Users can add new events by adding to the list and defining the structure.
- [ ] Add support for multiple evolutions in a event. For example, if an event occurs twice (denoted by occurrence) then update the display name to append the occurrence number.
- [ ] Add feature that validates an Event. On failure, give a descriptive error message that explains what is missing and where the error originated from. Define the Evaluator class in a separate file for modularity.

## Backlog

- [ ] Define formal config schemas for config, metrics, evolutions, events, and master mappings. Blocked on schema/modeling dependencies.
- [ ] Introduce explicit strict and permissive generation modes. Blocked on schema/modeling dependencies.
