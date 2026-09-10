from __future__ import annotations

import json
import shutil
import sqlite3
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = PROJECT_ROOT / "workspace_files"
EXCEL_TEMPLATE = PROJECT_ROOT / "demo_data" / "price.xlsx"


def create_excel_examples() -> None:
    if not EXCEL_TEMPLATE.is_file():
        raise FileNotFoundError(
            "Не найден demo_data/price.xlsx. Сначала распакуйте весь архив v3."
        )
    targets = [
        WORKSPACE / "price.xlsx",
        WORKSPACE / "reports" / "sales_report.xlsx",
    ]
    for target in targets:
        if target.exists():
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(EXCEL_TEMPLATE, target)


def create_sqlite_example() -> None:
    database = WORKSPACE / "warehouse.sqlite3"
    if database.exists():
        return
    connection = sqlite3.connect(database)
    try:
        connection.execute(
            """
            CREATE TABLE products (
                id INTEGER PRIMARY KEY,
                article TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                category TEXT NOT NULL,
                stock INTEGER NOT NULL,
                price REAL NOT NULL,
                status TEXT NOT NULL
            )
            """
        )
        connection.executemany(
            """
            INSERT INTO products(article, name, category, stock, price, status)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            [
                ("KB-001", "Клавиатура", "Периферия", 12, 3490, "active"),
                ("MS-002", "Мышь", "Периферия", 4, 1890, "active"),
                ("MN-003", "Монитор 24", "Мониторы", 7, 17990, "active"),
                ("HD-004", "Наушники", "Аудио", 0, 5290, "out_of_stock"),
                ("CM-005", "Веб-камера", "Периферия", 9, 4190, "active"),
                ("MC-006", "Микрофон", "Аудио", 2, 8990, "active"),
            ],
        )
        connection.execute(
            """
            CREATE TABLE invoices (
                id INTEGER PRIMARY KEY,
                customer TEXT NOT NULL,
                amount REAL NOT NULL,
                due_date TEXT NOT NULL,
                paid INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        connection.executemany(
            """
            INSERT INTO invoices(customer, amount, due_date, paid)
            VALUES (?, ?, ?, ?)
            """,
            [
                ("Альфа", 42000, "2026-08-01", 0),
                ("Бета", 18500, "2026-09-20", 0),
                ("Гамма", 77000, "2026-07-15", 1),
                ("Дельта", 9300, "2026-08-10", 0),
                ("Омега", 56000, "2026-10-01", 0),
            ],
        )
        connection.commit()
    finally:
        connection.close()


def write_text_examples() -> None:
    examples = {
        WORKSPACE / "customers.json": json.dumps(
            [
                {"id": 1, "name": "Альфа", "segment": "B2B"},
                {"id": 2, "name": "Бета", "segment": "Retail"},
            ],
            ensure_ascii=False,
            indent=2,
        ),
        WORKSPACE / "queries" / "readonly_examples.sql": (
            "-- Демо SQL-скрипт: отображается в каталоге, но не исполняется LLM.\n"
            "SELECT id, article, name, stock FROM products WHERE stock < 5;\n"
        ),
        WORKSPACE / "notes.txt": (
            "Этот файл нужен, чтобы показать определение неподдерживаемых форматов.\n"
        ),
    }
    for path, content in examples.items():
        if path.exists():
            continue
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")


def main() -> None:
    WORKSPACE.mkdir(parents=True, exist_ok=True)
    create_excel_examples()
    create_sqlite_example()
    write_text_examples()
    files = sorted(path.relative_to(WORKSPACE) for path in WORKSPACE.rglob("*") if path.is_file())
    print(f"Demo workspace готов: {WORKSPACE}")
    for path in files:
        print(f"- {path}")


if __name__ == "__main__":
    main()
