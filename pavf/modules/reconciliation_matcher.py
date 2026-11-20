"""
reconciliation_matcher.py

Matches:

    Breakdown Records
           ↓
    PAVF Records
           ↓
    Approval Evidence

Produces:

    ReconciliationResult
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import List

from .models import (
    ApprovalStatus,
    ApprovalMatch,
    BreakdownRecord,
    EmailEvidence,
    PavfRecord,
    ReconciliationResult,
    ReconciliationStatus,
)


# ---------------------------------------------------------
# Configuration
# ---------------------------------------------------------

PART_MATCH_POINTS = 50
DATE_MATCH_POINTS = 25
PO_MATCH_POINTS = 15
BOX_QTY_MATCH_POINTS = 10

MIN_PAVF_MATCH_SCORE = 50

MIN_APPROVAL_MATCH_SCORE = 50


# ---------------------------------------------------------
# Matcher
# ---------------------------------------------------------


class ReconciliationMatcher:

    def reconcile(
        self,
        breakdown_records: List[BreakdownRecord],
        pavf_records: List[PavfRecord],
        email_evidence: List[EmailEvidence],
    ) -> List[ReconciliationResult]:

        results = []

        for breakdown in breakdown_records:

            matched_pavf = self._find_best_pavf(
                breakdown,
                pavf_records
            )

            matched_approval = self._find_best_approval(
                breakdown,
                email_evidence
            )

            result = self._build_result(
                breakdown,
                matched_pavf,
                matched_approval
            )

            results.append(result)

        return results

    # -----------------------------------------------------

    def _find_best_pavf(
        self,
        breakdown: BreakdownRecord,
        pavf_records: List[PavfRecord]
    ) -> PavfRecord | None:

        best_score = 0
        best_record = None

        for pavf in pavf_records:

            # A PAVF is shipment/date-specific. Do not let a strong part or
            # box match attach a PAVF from another reimbursement date.
            if not self._dates_match(breakdown.date, pavf.problem_date):
                continue

            score = self._score_pavf_match(
                breakdown,
                pavf
            )

            if score > best_score:
                best_score = score
                best_record = pavf

        if best_score < MIN_PAVF_MATCH_SCORE:
            return None

        return best_record

    # -----------------------------------------------------

    def _score_pavf_match(
        self,
        breakdown: BreakdownRecord,
        pavf: PavfRecord
    ) -> int:

        score = 0

        if (
            self.normalize_part(
                breakdown.part_number
            )
            ==
            self.normalize_part(
                pavf.part_number
            )
        ):
            score += PART_MATCH_POINTS

        if breakdown.date == pavf.problem_date:
            score += DATE_MATCH_POINTS

        if breakdown.box_qty == pavf.box_qty:
            score += BOX_QTY_MATCH_POINTS

        if pavf.po_number:
            score += PO_MATCH_POINTS

        return score

    # -----------------------------------------------------

    def _find_best_approval(
        self,
        breakdown: BreakdownRecord,
        evidence_list: List[EmailEvidence]
    ) -> ApprovalMatch | None:

        best_score = 0
        best_evidence = None

        normalized_part = self.normalize_part(
            breakdown.part_number
        )

        for evidence in evidence_list:

            score = 0

            for part in evidence.related_parts:

                if (
                    self.normalize_part(part)
                    ==
                    normalized_part
                ):
                    score += 70

            if evidence.email_date:

                if (
                    evidence.email_date.date()
                    ==
                    breakdown.date
                ):
                    score += 20

            if (
                evidence.approval_status
                ==
                ApprovalStatus.APPROVED
            ):
                score += 10

            if score > best_score:

                best_score = score
                best_evidence = evidence

        if (
            best_evidence is None
            or best_score < MIN_APPROVAL_MATCH_SCORE
        ):
            return ApprovalMatch(
                approval_found=False,
                approval_status=ApprovalStatus.NOT_FOUND,
                evidence_file=None,
                evidence_phrase=None,
                evidence_text=None,
                approval_date=None
            )

        return ApprovalMatch(

            approval_found=(
                best_evidence.approval_status
                ==
                ApprovalStatus.APPROVED
            ),

            approval_status=
                best_evidence.approval_status,

            evidence_file=
                best_evidence.email_file,

            evidence_phrase=
                best_evidence.approval_phrase,

            evidence_text=
                best_evidence.evidence_text,

            approval_date=
                best_evidence.email_date
        )

    # -----------------------------------------------------

    def _build_result(
        self,
        breakdown: BreakdownRecord,
        pavf: PavfRecord | None,
        approval: ApprovalMatch | None
    ) -> ReconciliationResult:

        notes = []

        if pavf is None:

            return ReconciliationResult(
                breakdown=breakdown,
                pavf=None,
                approval=approval,
                status=ReconciliationStatus.PAVF_NOT_FOUND,
                notes=[
                    "No matching PAVF found"
                ]
            )

        if (
            pavf.box_qty is not None
            and pavf.box_qty != breakdown.box_qty
        ):

            return ReconciliationResult(
                breakdown=breakdown,
                pavf=pavf,
                approval=approval,
                status=ReconciliationStatus.BOX_QTY_MISMATCH,
                notes=[
                    (
                        f"Breakdown boxes "
                        f"{breakdown.box_qty}"
                    ),
                    (
                        f"PAVF boxes "
                        f"{pavf.box_qty}"
                    )
                ]
            )

        notes.append(
            "PAVF located"
        )

        if approval is None:

            return ReconciliationResult(
                breakdown=breakdown,
                pavf=pavf,
                approval=None,
                status=(
                    ReconciliationStatus
                    .APPROVAL_NOT_FOUND
                ),
                notes=notes
            )

        if not approval.approval_found:

            return ReconciliationResult(
                breakdown=breakdown,
                pavf=pavf,
                approval=approval,
                status=(
                    ReconciliationStatus
                    .APPROVAL_NOT_FOUND
                ),
                notes=notes
            )

        notes.append(
            "Approval evidence located"
        )

        return ReconciliationResult(
            breakdown=breakdown,
            pavf=pavf,
            approval=approval,
            status=ReconciliationStatus.APPROVED,
            notes=notes
        )

    # -----------------------------------------------------

    @staticmethod
    def _dates_match(left, right) -> bool:
        if left is None or right is None:
            return False
        if isinstance(left, datetime):
            left = left.date()
        if isinstance(right, datetime):
            right = right.date()
        return left == right

    @staticmethod
    def normalize_part(
        value: str | None
    ) -> str:

        if not value:
            return ""

        return (
            value.upper()
            .replace("-", "")
            .replace("_", "")
            .replace(" ", "")
            .replace("'", "")
            .strip()
        )