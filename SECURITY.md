# Security

## Never commit

- `.env`
- Telegram bot tokens
- private user allowlists
- audit databases
- snapshots
- processed result files
- Telegram-uploaded customer files
- local workspace data

The repository contains a safe `.env.example` with empty secret fields.

## Data access

Use a dedicated `DATA_WORKSPACE_DIR`.

Do not point DataOps Commander at a disk root or an uncontrolled folder containing unrelated sensitive files.

## Human confirmation

Keep write confirmation enabled. Review the generated preview before approving a destructive operation.

## Backups

Snapshots are a safety mechanism, not a replacement for normal data backups.

For valuable production data, use independent backups in addition to DataOps rollback snapshots.
