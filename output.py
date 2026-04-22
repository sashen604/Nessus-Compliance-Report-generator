"""Output helpers for CLI tables and exports."""
from __future__ import annotations

from typing import Iterable, List, Dict, Optional
import csv
import json

try:
    from tabulate import tabulate
except ImportError:  # pragma: no cover
    tabulate = None


def render_table(title: str, rows: List[Dict], headers: Optional[List[str]] = None) -> str:
    if not rows:
        return f"{title}\n(No data)"

    if headers is None:
        headers = list(rows[0].keys())

    table_data = [[row.get(header, "") for header in headers] for row in rows]

    if tabulate:
        table = tabulate(table_data, headers=headers, tablefmt="github")
    else:
        # Minimal fallback
        header_line = "\t".join(headers)
        body_lines = ["\t".join(str(value) for value in row) for row in table_data]
        table = "\n".join([header_line] + body_lines)

    return f"{title}\n{table}"


def print_table(title: str, rows: List[Dict], headers: Optional[List[str]] = None) -> None:
    print(render_table(title, rows, headers=headers))
    print()


def export_csv(path: str, rows: List[Dict]) -> None:
    if not rows:
        return

    fieldnames = sorted({key for row in rows for key in row.keys()})
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def export_json(path: str, rows: List[Dict]) -> None:
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(rows, handle, indent=2, ensure_ascii=False)
