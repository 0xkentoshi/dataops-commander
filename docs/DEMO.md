# Portfolio Demo Scenario

A short live demo should show that DataOps Commander can understand a compound request, preview it, safely execute it, and undo it.

## Recommended sequence

1. Open the persistent file dashboard.
2. Select `price.xlsx`.
3. Show the detected workbook schema.
4. Send one compound request:

```text
rename Товар to Название, set Периферия stock to 77,
then remove duplicates by Артикул
```

5. Show the multi-step preview.
6. Confirm.
7. Show the completion summary / snapshot.
8. Open or download the resulting file.
9. Trigger Undo and show the restored source.
10. Switch to `warehouse.sqlite3` and run a safe filtered query.

This demonstrates:

- natural-language interpretation
- schema grounding
- multi-step planning
- human approval
- real execution
- audit / snapshot safety
- Excel + SQLite support
