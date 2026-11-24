"""Shared data models for pyPAVF."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from pathlib import Path


@dataclass
class ApprovalLedgerRecord:
    shipment_date: date | None
    part_number: str

    box_qty: int | float | None
    po_number: str | None

    approval_status: str

    approval_date: datetime | None

    pavf_file: str

    approval_email: str | None

    approval_phrase: str | None


class ApprovalStatus(str, Enum):
    APPROVED = "APPROVED"
    NOT_FOUND = "NOT_FOUND"


class ReconciliationStatus(str, Enum):
    APPROVED = "APPROVED"
    PAVF_NOT_FOUND = "PAVF_NOT_FOUND"
    APPROVAL_NOT_FOUND = "APPROVAL_NOT_FOUND"
    BOX_QTY_MISMATCH = "BOX_QTY_MISMATCH"


@dataclass
class BreakdownRecord:
    date: date | datetime
    program: str
    part_number: str
    part_qty: float = 0
    box_qty: float = 0
    amount: float = 0


@dataclass
class PavfRecord:
    part_number: str
    problem_date: date | None = None
    box_qty: float | None = None
    po_number: str | None = None
    source_file: str | Path | None = None
    source_sheet: str | None = None


@dataclass
class EmailEvidence:
    email_file: str | Path | None = None
    email_date: datetime | None = None
    related_parts: list[str] = field(default_factory=list)
    approval_status: ApprovalStatus = ApprovalStatus.NOT_FOUND
    approval_phrase: str | None = None
    evidence_text: str | None = None
    subject: str | None = None
    sender: str | None = None
    email_type: str | None = None
    confidence: int | None = None


@dataclass
class ApprovalMatch:
    approval_found: bool
    approval_status: ApprovalStatus
    evidence_file: str | Path | None = None
    evidence_phrase: str | None = None
    evidence_text: str | None = None
    approval_date: datetime | None = None


@dataclass
class ReconciliationResult:
    breakdown: BreakdownRecord
    pavf: PavfRecord | None
    approval: ApprovalMatch | None
    status: ReconciliationStatus
    notes: list[str] = field(default_factory=list)
