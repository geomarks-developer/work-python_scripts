"""
pavf_reconciliation.py

End-to-end orchestration for the PAVF reconciliation project.

Pipeline:
    BreakdownReader
        -> PAVF records
        -> Email approval evidence
        -> ReconciliationMatcher
        -> ReconciliationReport

The existing project modules have evolved independently, so this orchestrator
uses small adapter functions. That keeps file loading separate from matching
and reporting and makes it easy to plug in the exact PAVF/email readers.
"""
from __future__ import annotations

import argparse
import json
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Iterable
from .breakdown_reader import BreakdownReader
from .pavf_reader import read_pavf_records
from .email_evidence_reader import read_email_evidence
from .reconciliation_matcher import ReconciliationMatcher
from .reconciliation_report import create_reconciliation_report

try:
    from .models import BreakdownRecord, EmailEvidence, PavfRecord
except ImportError as exc:
    raise ImportError(
        "Could not import BreakdownRecord, EmailEvidence, and PavfRecord "
        "from models.py. Keep pavf_reconciliation.py in the PAVF project folder."
    ) from exc


RecordLoader = Callable[[Path], Iterable[Any]]


class PAVFReconciliation:
    """Coordinate input loading, reconciliation, and Excel reporting."""

    def __init__(
        self,
        breakdown_path: str | Path,
        pavf_path: str | Path,
        email_evidence_path: str | Path,
        output_path: str | Path = "PAVF_Reconciliation_Report.xlsx",
        *,
        pavf_loader: RecordLoader | None = None,
        email_loader: RecordLoader | None = None,
    ) -> None:
        self.breakdown_path = Path(breakdown_path)
        self.pavf_path = Path(pavf_path)
        self.email_evidence_path = Path(email_evidence_path)
        self.output_path = Path(output_path)
        self.pavf_loader = pavf_loader
        self.email_loader = email_loader

    def run(self) -> Path:
        """Run the complete reconciliation and return the report path."""
        self._validate_inputs()

        print("[1/5] Reading reimbursement breakdown...")
        breakdown_records = self.load_breakdown_records()
        print(f"      {len(breakdown_records)} breakdown records")

        print("[2/5] Reading PAVF records...")
        pavf_records = self.load_pavf_records()
        print(f"      {len(pavf_records)} PAVF records")

        print("[3/5] Reading approval evidence...")
        email_evidence = self.load_email_evidence()
        print(f"      {len(email_evidence)} evidence records")

        print("[4/5] Reconciling records...")
        matcher = ReconciliationMatcher()
        results = matcher.reconcile(
            breakdown_records=breakdown_records,
            pavf_records=pavf_records,
            email_evidence=email_evidence,
        )
        self._print_status_summary(results)

        print("[5/5] Creating Excel report...")
        report_path = create_reconciliation_report(
            results,
            self.output_path,
        )
        print(f"      Report created: {report_path.resolve()}")
        return report_path

    def load_breakdown_records(self) -> list[BreakdownRecord]:
        """Read breakdown workbook dictionaries and adapt them to the model."""
        rows = BreakdownReader(self.breakdown_path).read()
        return [self._construct(BreakdownRecord, row) for row in rows]

    def load_pavf_records(self) -> list[PavfRecord]:
        """
        Load PAVF records.

        A project-specific loader may be supplied. Otherwise JSON is supported
        directly. This avoids making assumptions about the layout of the PAVF
        source workbook while the project reader is still evolving.
        """
        if self.pavf_loader is not None:
            rows = list(self.pavf_loader(self.pavf_path))
        else:
            rows = self._load_json_records(
                self.pavf_path,
                label="PAVF",
            )
        return [self._construct(PavfRecord, row) for row in rows]

    def load_email_evidence(self) -> list[EmailEvidence]:
        """
        Load normalized email evidence.

        A project-specific loader may be supplied. Otherwise JSON is supported
        directly. email_classifier.py classifies a message; a reader/parser can
        be connected here to turn mailbox/export files into EmailEvidence.
        """
        if self.email_loader is not None:
            rows = list(self.email_loader(self.email_evidence_path))
        else:
            rows = self._load_json_records(
                self.email_evidence_path,
                label="email evidence",
            )
        return [self._construct(EmailEvidence, row) for row in rows]

    def _validate_inputs(self) -> None:
        missing = [
            p
            for p in (
                self.breakdown_path,
                self.pavf_path,
                self.email_evidence_path,
            )
            if not p.exists()
        ]
        if missing:
            values = "\n".join(f"  - {path}" for path in missing)
            raise FileNotFoundError(f"Missing input file(s):\n{values}")

    @staticmethod
    def _load_json_records(path: Path, label: str) -> list[dict[str, Any]]:
        if path.suffix.lower() != ".json":
            raise ValueError(
                f"No {label} loader was supplied for {path.name}. "
                "Pass a project-specific loader to PAVFReconciliation, or "
                "provide normalized .json input."
            )

        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)

        if isinstance(payload, dict):
            # Accept a common wrapper such as {"records": [...]}.
            payload = payload.get("records", payload.get("data"))

        if not isinstance(payload, list):
            raise ValueError(
                f"{path.name} must contain a JSON list or a records/data list."
            )

        if not all(isinstance(item, dict) for item in payload):
            raise ValueError(f"Every record in {path.name} must be a JSON object.")

        return payload

    @classmethod
    def _construct(cls, model_type: type, value: Any) -> Any:
        """Convert dictionaries into project dataclasses without guessing fields."""
        if isinstance(value, model_type):
            return value
        if not isinstance(value, dict):
            raise TypeError(
                f"Expected {model_type.__name__} or dict, got {type(value).__name__}."
            )

        data = dict(value)
        data = cls._coerce_known_values(data)

        if is_dataclass(model_type):
            allowed = {f.name for f in fields(model_type)}
            data = {key: val for key, val in data.items() if key in allowed}

        try:
            return model_type(**data)
        except TypeError as exc:
            raise TypeError(
                f"Could not create {model_type.__name__} from fields "
                f"{sorted(data.keys())}. Check models.py and the input adapter."
            ) from exc

    @staticmethod
    def _coerce_known_values(data: dict[str, Any]) -> dict[str, Any]:
        """Convert ISO date text commonly found in normalized JSON exports."""
        converted = dict(data)

        date_fields = {"date", "problem_date"}
        datetime_fields = {"email_date", "approval_date"}

        for key in date_fields:
            value = converted.get(key)
            if isinstance(value, str) and value.strip():
                try:
                    converted[key] = date.fromisoformat(value.strip())
                except ValueError:
                    pass

        for key in datetime_fields:
            value = converted.get(key)
            if isinstance(value, str) and value.strip():
                text = value.strip().replace("Z", "+00:00")
                try:
                    converted[key] = datetime.fromisoformat(text)
                except ValueError:
                    pass

        return converted

    @staticmethod
    def _enum_text(value: Any) -> str:
        if isinstance(value, Enum):
            return str(value.value)
        return str(value or "")

    @classmethod
    def _print_status_summary(cls, results: Iterable[Any]) -> None:
        counts: dict[str, int] = {}
        for result in results:
            status = cls._enum_text(getattr(result, "status", "UNKNOWN"))
            counts[status] = counts.get(status, 0) + 1

        for status, count in sorted(counts.items()):
            print(f"      {status}: {count}")


def run_reconciliation(
    breakdown_path: str | Path,
    pavf_path: str | Path,
    email_evidence_path: str | Path,
    output_path: str | Path = "PAVF_Reconciliation_Report.xlsx",
    *,
    pavf_loader: RecordLoader | None = None,
    email_loader: RecordLoader | None = None,
) -> Path:
    """Convenience API for calling the orchestrator from another Python module."""
    return PAVFReconciliation(
        breakdown_path=breakdown_path,
        pavf_path=pavf_path,
        email_evidence_path=email_evidence_path,
        output_path=output_path,
        pavf_loader=pavf_loader,
        email_loader=email_loader,
    ).run()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the PAVF reconciliation pipeline and create an Excel report."
    )
    parser.add_argument("--breakdown", required=True, help="Breakdown .xlsx file")
    parser.add_argument(
        "--pavf",
        required=True,
        help="Normalized PAVF .json file unless a custom loader is used",
    )
    parser.add_argument(
        "--emails",
        required=True,
        help="Normalized email-evidence .json file unless a custom loader is used",
    )
    parser.add_argument(
        "--output",
        default="PAVF_Reconciliation_Report.xlsx",
        help="Output Excel report path",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()

    try:
        run_reconciliation(
            breakdown_path=args.breakdown,
            pavf_path=args.pavf,
            email_evidence_path=args.emails,
            output_path=args.output,
        )
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
