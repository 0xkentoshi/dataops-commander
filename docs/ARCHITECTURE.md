# Architecture

DataOps Commander separates language understanding from data execution.

## Main request path

```text
Telegram message
   ↓
AI Command Interpreter
   ↓
CommandEnvelope
   ├─ mode
   ├─ ordered tasks
   ├─ source hint
   └─ execution options
   ↓
Source Catalog + Real Schema
   ↓
Excel Planner / SQLite Planner
   ↓
Structured operation plan
   ↓
Deterministic guards
   ├─ read → execute immediately
   └─ write → preview + confirmation
                    ↓
              source version check
                    ↓
                 snapshot
                    ↓
                 executor
                    ↓
                verification
                    ↓
                 audit DB
```

## AI command interpreter

`app/intent.py` handles the semantic interpretation of the whole user message.

It can produce data tasks or operational modes such as:

- data
- undo
- repeat last operation
- send last processed file
- rename the last file
- update an already pending operation
- confirm / cancel a pending operation

Natural-language semantics belong to the LLM. Python validates the structured result.

## Source grounding

The selected source and its actual schema are passed into the planning stage.

This prevents the model from treating invented column names, sheets or tables as if they existed.

## Excel path

The Excel planner receives workbook metadata and returns a structured `ExcelOperationPlan`.

The deterministic Excel executor owns:

- filtering
- updates
- row deletion
- column operations
- clear-values logic
- deduplication
- snapshot creation
- temporary-copy validation
- atomic replacement
- post-write verification

## SQLite path

The SQL planner returns a structured `SqlOperationPlan`.

The model never returns executable raw SQL.

The runtime turns structured filters and assignments into parameterized SQL inside a transaction.

`UPDATE` and `DELETE` without filters are rejected by the plan schema and execution layer.

## Multi-task execution

For a write batch, DataOps uses a staging copy.

Each planned step runs in order against the result of the previous step. Only after the whole sequence succeeds and verifies is the result committed back to the requested destination.

This makes commands such as:

```text
rename a column → update that renamed column → deduplicate
```

possible without planning every step against stale original metadata.

## Audit and recovery

The application persists operation status and audit events in its own SQLite database.

Write operations retain enough information to support a guarded undo. The rollback layer checks the current source hash before restoring an older snapshot so it cannot silently overwrite later external edits.
