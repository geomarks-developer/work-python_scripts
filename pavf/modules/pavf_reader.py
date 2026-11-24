"""Read PAVF Excel/CSV files into normalized PavfRecord objects."""
from __future__ import annotations

import csv
import re
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .models import PavfRecord

import warnings

warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported*"
)

def normalize_part_number(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper().strip())


def extract_part_number_from_filename(path: str | Path) -> str | None:
    stem = Path(path).stem.upper()
    match = re.search(r"PAVF\s+FORM\s+(S[A-Z0-9]+(?:[-_][A-Z0-9]+)*)", stem)
    return normalize_part_number(match.group(1)) if match else None


class PavfReader:
    SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".csv"}

    def __init__(self, source: str | Path):
        self.source = Path(source)
        self._model_fields = self._get_model_fields()

    def read(self) -> list[PavfRecord]:
        files = self._source_files()
        candidates: list[tuple[PavfRecord, bool, int]] = []

        #print("\n================================")
        #print(f"PAVF SOURCE FILES: {len(files)}")
        #print("================================")

        for order, path in enumerate(files):
        #    print(f"\nLoading {path.name}")
            if path.suffix.lower() in {".xlsx", ".xlsm"}:
                record = self._read_excel_form(path)
                if record is not None:
                    is_revised = "REVISED" in path.stem.upper()
                    candidates.append((record, is_revised, order))
            elif path.suffix.lower() == ".csv":
                for record in self._read_csv(path):
                    candidates.append((record, False, order))

        # Keep one form per part/date. A file containing "Revised" wins.
        selected: dict[tuple[str, Any], tuple[PavfRecord, bool, int]] = {}
        for record, is_revised, order in candidates:
            key = (
                normalize_part_number(getattr(record, "part_number", None)),
                self._date_only(getattr(record, "problem_date", None)),
            )
            current = selected.get(key)
            if current is None or (is_revised and not current[1]) or (is_revised == current[1] and order > current[2]):
                selected[key] = (record, is_revised, order)

        records = [item[0] for item in selected.values()]
        return records

    def _source_files(self) -> list[Path]:
        if not self.source.exists():
            raise FileNotFoundError(f"PAVF source not found: {self.source}")
        if self.source.is_file():
            if self.source.suffix.lower() not in self.SUPPORTED_SUFFIXES:
                raise ValueError(f"Unsupported PAVF file type: {self.source.suffix}")
            return [self.source]
        files = sorted(
            path for path in self.source.iterdir()
            if path.is_file()
            and path.suffix.lower() in self.SUPPORTED_SUFFIXES
            and not path.name.startswith("~$")
        )
        if not files:
            raise ValueError(f"No supported PAVF files found in {self.source}")
        return files

    def _read_excel_form(self, path: Path) -> PavfRecord | None:
        workbook = load_workbook(path, data_only=True, read_only=False)
        try:
            sheet_name = next(
                (name for name in workbook.sheetnames if name.strip().upper() == "PAVF"),
                None,
            )
            if sheet_name is None:
                print(f"Skipped - PAVF sheet not found: {path.name}")
                return None

            worksheet = workbook[sheet_name]

            # Fixed cells in the PAVF form.
            cell_part = worksheet["C10"].value
            part_number = normalize_part_number(cell_part) or extract_part_number_from_filename(path)
            problem_date = self._coerce_date(worksheet["C12"].value)
            box_qty = self._coerce_number(worksheet["D20"].value)
            po_number = self._coerce_identifier(worksheet["H20"].value)

            if not part_number:
                print(f"Skipped - part number not found: {path.name}")
                return None

            data: dict[str, Any] = {
                "part_number": part_number,
                "problem_date": problem_date,
                "box_qty": box_qty,
                "po_number": po_number,
                "source_file": str(path),
                "source_sheet": worksheet.title,
            }
            return self._make_record(data, path.name)
        finally:
            workbook.close()

    def _read_csv(self, path: Path) -> list[PavfRecord]:
        records: list[PavfRecord] = []
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                normalized = {self._normalize_header(key): value for key, value in row.items()}
                data = {
                    "part_number": normalize_part_number(
                        normalized.get("part number") or normalized.get("partnumber") or normalized.get("part")
                    ),
                    "problem_date": self._coerce_date(
                        normalized.get("problem date") or normalized.get("date")
                    ),
                    "box_qty": self._coerce_number(
                        normalized.get("box qty") or normalized.get("box quantity") or normalized.get("boxes")
                    ),
                    "po_number": self._coerce_identifier(
                        normalized.get("po number") or normalized.get("po")
                    ),
                    "source_file": str(path),
                    "source_sheet": "CSV",
                }
                if data["part_number"]:
                    records.append(self._make_record(data, path.name))
        return records

    def _make_record(self, data: dict[str, Any], source_name: str) -> PavfRecord:
        allowed = {key: value for key, value in data.items() if key in self._model_fields}
        try:
            return PavfRecord(**allowed)
        except TypeError as exc:
            raise TypeError(
                f"Could not create PavfRecord from {source_name}. "
                f"Fields supplied: {sorted(allowed)}."
            ) from exc

    def _get_model_fields(self) -> set[str]:
        if is_dataclass(PavfRecord):
            return {field.name for field in fields(PavfRecord)}
        annotations = getattr(PavfRecord, "__annotations__", {})
        return set(annotations) if annotations else {
            "part_number", "problem_date", "box_qty", "po_number"
        }

    @staticmethod
    def _normalize_header(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip().lower().replace("_", " ")
        return re.sub(r"\s+", " ", text)

    @staticmethod
    def _date_only(value: Any) -> Any:
        if isinstance(value, datetime):
            return value.date()
        return value

    @staticmethod
    def _coerce_date(value: Any) -> date | None:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if value in (None, ""):
            return None
        text = str(value).strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    @staticmethod
    def _coerce_number(value: Any) -> int | float | None:
        if value in (None, ""):
            return None
        if isinstance(value, (int, float)):
            number = float(value)
        else:
            try:
                number = float(str(value).replace(",", "").strip())
            except ValueError:
                return None
        return int(number) if number.is_integer() else number

    @staticmethod
    def _coerce_identifier(value: Any) -> str | None:
        if value in (None, ""):
            return None
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()


def read_pavf_records(source: str | Path) -> list[PavfRecord]:
    return PavfReader(source).read()
