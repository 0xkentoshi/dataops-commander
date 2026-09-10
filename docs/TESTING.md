# Testing

The public `v1.0.0` repository contains **34 automated regression tests** using Python's built-in `unittest`.

## Run

```powershell
python -m unittest discover -s tests -v
```

## Covered areas

The suite includes regression coverage for:

- schema-aware AI command interpretation
- preservation of multiple tasks from one user message
- removal of Python semantic regex routing
- Excel batch dependency chaining
- SQLite batch dependency chaining
- source catalog isolation
- read-only Excel operations
- in-place Excel writes
- source-change rejection between preview and execution
- workbook structure discovery
- malformed worksheet dimension handling
- merged-cell write guards
- SQLite transactions
- filtered update / delete requirements
- dashboard state and navigation
- natural-language confirmation behavior
- output naming / copy options
- operation history paths
- snapshot restoration
- undo protection after an external change
- audit database connection lifecycle

## Philosophy

When live QA exposes a reproducible failure, the desired workflow is:

```text
reproduce
→ add or update regression coverage
→ fix
→ run the full suite
```

The tests do not attempt to prove that an LLM will interpret every possible phrase perfectly. They verify that the deterministic system keeps known execution and state-management failures from returning.
