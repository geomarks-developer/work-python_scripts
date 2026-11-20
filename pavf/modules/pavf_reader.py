"""
pavf_reader.py

Reads PAVF source files and converts them into normalized PavfRecord objects.

Supported sources:
    - .xlsx / .xlsm Excel workbooks
    - .csv files
    - a directory containing supported files

The parser is deliberately header-driven so it can tolerate reasonable column
name variations without hard-coding one workbook layout.
"""
from __future__ import annotations

import csv
import re
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

from .models import PavfRecord


class PavfReader:
    SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".csv"}

    # Canonical PavfRecord field -> possible source header labels.
    HEADER_ALIASES = {
        "part_number": {
            "part number", "part no", "part #", "part", "partnumber",
            "part_number", "material", "material number",
        },
        "problem_date": {
            "problem date", "date", "pavf date", "issue date", "problem_date",
            "date of problem",
        },
        "box_qty": {
            "box qty", "box quantity", "boxes", "qty boxes", "box_qty",
            "number of boxes", "no of boxes",
        },
        "po_number": {
            "po", "po number", "po #", "purchase order", "purchase order number",
            "po_number",
        },
    }

    def __init__(self, source: str | Path):
        self.source = Path(source)
        self._model_fields = self._get_model_fields()

    def read(self) -> list[PavfRecord]:
        """Read a file or directory and return normalized PavfRecord objects."""
        files = self._source_files()
        records: list[PavfRecord] = []

        for path in files:
            suffix = path.suffix.lower()
            if suffix in {".xlsx", ".xlsm"}:
                records.extend(self._read_excel(path))
            elif suffix == ".csv":
                records.extend(self._read_csv(path))

        return records

    def _source_files(self) -> list[Path]:
        if not self.source.exists():
            raise FileNotFoundError(f"PAVF source not found: {self.source}")

        if self.source.is_file():
            if self.source.suffix.lower() not in self.SUPPORTED_SUFFIXES:
                raise ValueError(f"Unsupported PAVF file type: {self.source.suffix}")
            return [self.source]

        files = sorted(
            p for p in self.source.iterdir()
            if p.is_file()
            and p.suffix.lower() in self.SUPPORTED_SUFFIXES
            and not p.name.startswith("~$")
        )
        if not files:
            raise ValueError(f"No supported PAVF files found in {self.source}")
        return files

    def _read_excel(self, path: Path) -> list[PavfRecord]:
        wb = load_workbook(path, data_only=True, read_only=True)
        records: list[PavfRecord] = []
        try:
            for ws in wb.worksheets:
                header_row, field_columns = self._find_header(ws)
                if header_row is None:
                    continue

                for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                    record = self._row_to_record(
                        row=row,
                        field_columns=field_columns,
                        source_file=path,
                        source_sheet=ws.title,
                    )
                    if record is not None:
                        records.append(record)
        finally:
            wb.close()
        return records

    def _find_header(self, ws) -> tuple[int | None, dict[str, int]]:
        """Find the first likely header row within the leading workbook rows."""
        scan_limit = min(ws.max_row or 1, 40)
        best_row: int | None = None
        best_map: dict[str, int] = {}
        best_score = -1

        for row_number, row in enumerate(
            ws.iter_rows(min_row=1, max_row=scan_limit, values_only=True),
            start=1,
        ):
            mapping = self._map_headers(row)
            required_hits = sum(
                field in mapping for field in ("part_number", "problem_date")
            )
            score = len(mapping) + required_hits * 3
            if score > best_score:
                best_row, best_map, best_score = row_number, mapping, score

        # A part number is essential for the existing matcher. Require it.
        if "part_number" not in best_map:
            return None, {}
        return best_row, best_map

    def _read_csv(self, path: Path) -> list[PavfRecord]:
        records: list[PavfRecord] = []
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.reader(handle)
            rows = list(reader)

        if not rows:
            return records

        best_index: int | None = None
        best_map: dict[str, int] = {}
        best_score = -1
        for index, row in enumerate(rows[:40]):
            mapping = self._map_headers(row)
            required_hits = sum(
                field in mapping for field in ("part_number", "problem_date")
            )
            score = len(mapping) + required_hits * 3
            if score > best_score:
                best_index, best_map, best_score = index, mapping, score

        if best_index is None or "part_number" not in best_map:
            return records

        for row in rows[best_index + 1:]:
            record = self._row_to_record(
                row=row,
                field_columns=best_map,
                source_file=path,
                source_sheet="CSV",
            )
            if record is not None:
                records.append(record)
        return records

    def _map_headers(self, row: Iterable[Any]) -> dict[str, int]:
        normalized_aliases = {
            field: {self._normalize_header(a) for a in aliases}
            for field, aliases in self.HEADER_ALIASES.items()
        }
        mapping: dict[str, int] = {}

        for index, value in enumerate(row):
            header = self._normalize_header(value)
            if not header:
                continue

            for field, aliases in normalized_aliases.items():
                if field not in self._model_fields:
                    continue
                if header in aliases:
                    mapping.setdefault(field, index)
                    break

            # Also accept exact model field names automatically.
            for field in self._model_fields:
                if header == self._normalize_header(field):
                    mapping.setdefault(field, index)

        return mapping

    def _row_to_record(
        self,
        row: Iterable[Any],
        field_columns: dict[str, int],
        source_file: Path,
        source_sheet: str,
    ) -> PavfRecord | None:
        values = list(row)
        data: dict[str, Any] = {}

        for field, column_index in field_columns.items():
            value = values[column_index] if column_index < len(values) else None
            data[field] = self._coerce(field, value)

        part = self.normalize_part(data.get("part_number"))
        if not part:
            return None
        data["part_number"] = part

        # Optional source/audit fields are populated only if they actually exist
        # in the project's PavfRecord dataclass.
        optional_audit = {
            "source_file": str(source_file),
            "source_sheet": source_sheet,
        }
        for key, value in optional_audit.items():
            if key in self._model_fields:
                data[key] = value

        allowed = {key: val for key, val in data.items() if key in self._model_fields}
        try:
            return PavfRecord(**allowed)
        except TypeError as exc:
            raise TypeError(
                f"Could not create PavfRecord from {source_file.name} / "
                f"{source_sheet}. Parsed fields: {sorted(allowed)}. "
                "Check models.py for required fields and add header aliases if needed."
            ) from exc

    def _get_model_fields(self) -> set[str]:
        if is_dataclass(PavfRecord):
            return {field.name for field in fields(PavfRecord)}
        annotations = getattr(PavfRecord, "__annotations__", {})
        if annotations:
            return set(annotations)
        return {"part_number", "problem_date", "box_qty", "po_number"}

    @staticmethod
    def _normalize_header(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip().lower()
        text = text.replace("#", " number ")
        text = re.sub(r"[_\-/]+", " ", text)
        text = re.sub(r"[^a-z0-9 ]+", "", text)
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def normalize_part(value: Any) -> str:
        if value is None:
            return ""
        return (
            str(value)
            .upper()
            .replace("-", "")
            .replace("_", "")
            .replace(" ", "")
            .replace("'", "")
            .strip()
        )

    @staticmethod
    def _coerce(field: str, value: Any) -> Any:
        if value in (None, ""):
            return None

        if field == "problem_date":
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            text = str(value).strip()
            for fmt in (
                "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y",
                "%m-%d-%Y", "%m-%d-%y",
            ):
                try:
                    return datetime.strptime(text, fmt).date()
                except ValueError:
                    continue
            return value

        if field == "box_qty":
            if isinstance(value, (int, float)):
                return value
            text = str(value).replace(",", "").strip()
            try:
                number = float(text)
                return int(number) if number.is_integer() else number
            except ValueError:
                return value

        if field in {"part_number", "po_number"}:
            # Prevent Excel numeric identifiers from becoming '12345.0'.
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value).strip()

        return value


def read_pavf_records(source: str | Path) -> list[PavfRecord]:
    """Convenience loader compatible with PAVFReconciliation(pavf_loader=...)."""
    return PavfReader(source).read()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Parse PAVF source records.")
    parser.add_argument("source", help="PAVF workbook, CSV, or folder")
    args = parser.parse_args()

    parsed = read_pavf_records(args.source)
    print(f"PAVF records found: {len(parsed)}")
    for item in parsed[:10]:
        print(item)
