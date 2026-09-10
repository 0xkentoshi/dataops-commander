# Changelog

## [1.0.1] - 2026-09-10

### Changed

- converted the Telegram/admin/runtime interface to English while preserving multilingual natural-language input
- improved user-facing planning failures so internal Pydantic/validation traces stay in audit instead of Telegram
- hardened Excel UPDATE structured-plan normalization
- hardened filtered SQLite SELECT planning with catalog-bound filter recovery
- hardened SQLite UPDATE planning with catalog-bound assignment recovery
- added sequential recovery when a malformed SQL write loses both its filter and assignment
- expanded the regression suite from 34 to 50 tests
- added six autoplaying live demo GIFs to the README

### Notes

`v1.0.1` is the post-live-QA portfolio release. It keeps the same execution surface as v1.0.0 while improving interface consistency and planner resilience.

## [1.0.0] - 2026-09-10

### Added

- AI-native natural-language command interpreter
- active-source and schema-aware planning
- Excel read / update / delete / structural cleanup operations
- SQLite select / filtered update / filtered delete
- ordered multi-task execution
- persistent Telegram file dashboard
- local and Telegram-uploaded source catalog
- preview and explicit confirmation for writes
- snapshot creation and guarded undo
- source-version checks before execution
- optional copy instead of changing the original
- result-file naming and sending
- repeat-last-operation workflow
- audit history and health checks
- deterministic SQL safety controls
- atomic Excel write path and verification
- 34 automated regression tests

### Notes

`v1.0.0` is the portfolio baseline focused on Excel and SQLite operations.
CSV / TSV execution adapters, Google Sheets and additional database backends remain future work.
