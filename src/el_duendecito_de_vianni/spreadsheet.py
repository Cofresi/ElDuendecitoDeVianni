from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pandas as pd


def _format_value(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value).strip()


def read_employees(path: str | Path) -> list[dict[str, str]]:
    file_path = Path(path)
    engine = "xlrd" if file_path.suffix.lower() == ".xls" else "openpyxl"
    frame = pd.read_excel(file_path, dtype=object, engine=engine)
    frame = frame.dropna(how="all")
    employees: list[dict[str, str]] = []
    for _, row in frame.iterrows():
        employees.append({str(column): _format_value(row[column]) for column in frame.columns})
    return employees
