# Safety Model

DataOps Commander is intentionally designed so that an LLM cannot directly mutate arbitrary data.

## Trust boundaries

### LLM

Trusted for:

- semantic interpretation
- mapping a natural-language request to a structured action
- selecting grounded schema objects from the provided catalog

Not trusted for:

- filesystem paths
- raw SQL execution
- bypassing action limits
- deciding whether deterministic safety checks should be ignored

### Deterministic runtime

Owns:

- allowed workspace resolution
- operation validation
- preview generation
- human confirmation
- file locking
- snapshots
- SQL parameterization
- transaction boundaries
- atomic file replacement
- post-write verification
- audit persistence

## Stale preview protection

A write plan is tied to the source version inspected when the preview was created.

If the source changes before confirmation, execution is blocked instead of applying an old plan to new data.

## SQL safety

The SQLite interface supports structured select, update and delete plans.

It deliberately does not expose arbitrary SQL generation.

Write operations require filter conditions, and schema-destructive statements such as `DROP`, `TRUNCATE` or arbitrary `ALTER` are not exposed to the model.

## Undo safety

Undo is snapshot-based.

Before restoration, the runtime verifies that the current source still matches the result of the operation being undone. If the file changed externally afterward, automatic rollback is rejected.

This is designed to protect newer edits from being overwritten by an old snapshot.
