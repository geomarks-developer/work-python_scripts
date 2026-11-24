"""Match breakdown records to PAVF forms and thread-based approval evidence."""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
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

PART_MATCH_POINTS = 50
DATE_MATCH_POINTS = 25
BOX_QTY_MATCH_POINTS = 20
PO_MATCH_POINTS = 5
MIN_PAVF_MATCH_SCORE = PART_MATCH_POINTS + DATE_MATCH_POINTS


class ReconciliationMatcher:
    def reconcile(
        self,
        breakdown_records: List[BreakdownRecord],
        pavf_records: List[PavfRecord],
        email_evidence: List[EmailEvidence],
    ) -> List[ReconciliationResult]:
        results: list[ReconciliationResult] = []

        for breakdown in breakdown_records:
            pavf = self._find_best_pavf(breakdown, pavf_records)

            # Approval evidence is evaluated only after an exact transaction
            # has been established by part/group plus PAVF problem date.
            approval = (
                self._find_thread_approval(breakdown, pavf, email_evidence)
                if pavf is not None
                else self._not_found_approval()
            )

            results.append(self._build_result(breakdown, pavf, approval))

        return results

    # -----------------------------------------------------
    # PAVF matching
    # -----------------------------------------------------
    def _find_best_pavf(
        self,
        breakdown: BreakdownRecord,
        pavf_records: List[PavfRecord],
    ) -> PavfRecord | None:
        best_score = 0
        best_record: PavfRecord | None = None

        print(f"\nSearching PAVF for {breakdown.part_number}")
        print(f"Normalized: {self.normalize_part(breakdown.part_number)}")

        for pavf in pavf_records:
            # Both conditions are mandatory. This prevents a 9/11 PAVF from
            # matching every monthly breakdown row for the same part group.
            if not self._parts_match(breakdown.part_number, pavf.part_number):
                continue
            if not self._dates_match(breakdown.date, pavf.problem_date):
                continue

            score = self._score_pavf_match(breakdown, pavf)
            print(
                f"Candidate: {pavf.part_number} "
                f"Date={pavf.problem_date} Score={score}"
            )

            if score > best_score:
                best_score = score
                best_record = pavf

        print(f"Best Score={best_score}")
        return best_record if best_score >= MIN_PAVF_MATCH_SCORE else None

    def _score_pavf_match(
        self,
        breakdown: BreakdownRecord,
        pavf: PavfRecord,
    ) -> int:
        score = 0
        if self._parts_match(breakdown.part_number, pavf.part_number):
            score += PART_MATCH_POINTS
        if self._dates_match(breakdown.date, pavf.problem_date):
            score += DATE_MATCH_POINTS
        if self._numbers_match(breakdown.box_qty, pavf.box_qty):
            score += BOX_QTY_MATCH_POINTS
        if getattr(pavf, "po_number", None):
            score += PO_MATCH_POINTS
        return score

    # -----------------------------------------------------
    # Thread-based approval matching
    # -----------------------------------------------------
    def _find_thread_approval(
        self,
        breakdown: BreakdownRecord,
        pavf: PavfRecord,
        evidence_list: List[EmailEvidence],
    ) -> ApprovalMatch:
        """
        Find approval through the email conversation rather than a fixed date
        window. A thread qualifies when its request/reply content identifies the
        PAVF part. The approval can arrive later, including after a weekend.
        """
        thread_groups: dict[str, list[EmailEvidence]] = {}
        for evidence in evidence_list:
            thread_groups.setdefault(self._thread_key(evidence), []).append(evidence)

        qualifying_threads: list[list[EmailEvidence]] = []

        for thread_items in thread_groups.values():
            thread_parts: list[str] = []
            for item in thread_items:
                thread_parts.extend(list(getattr(item, "related_parts", []) or []))
                thread_parts.extend(self._parts_from_evidence_text(item))

            if any(
                self._parts_match(pavf.part_number, part)
                or self._parts_match(breakdown.part_number, part)
                for part in thread_parts
            ):
                qualifying_threads.append(thread_items)

        # Prefer explicit approved evidence from a qualifying request/reply
        # thread. Email date is deliberately not used as a rejection rule.
        approved_candidates: list[EmailEvidence] = []
        for thread_items in qualifying_threads:
            for evidence in thread_items:
                if evidence.approval_status == ApprovalStatus.APPROVED:
                    approved_candidates.append(evidence)

        if not approved_candidates:
            return self._not_found_approval()

        # Prefer the latest dated approval when dates are available.
        best_evidence = max(
            approved_candidates,
            key=lambda item: self._date_sort_value(
                getattr(item, "email_date", None)
            ),
        )

        return ApprovalMatch(
            approval_found=True,
            approval_status=ApprovalStatus.APPROVED,
            evidence_file=best_evidence.email_file,
            evidence_phrase=best_evidence.approval_phrase,
            evidence_text=best_evidence.evidence_text,
            approval_date=best_evidence.email_date,
        )

    @classmethod
    def _thread_key(cls, evidence: EmailEvidence) -> str:
        """Normalize request/reply filenames into a shared conversation key."""
        raw_name = str(getattr(evidence, "email_file", "") or "")
        name = Path(raw_name).stem.upper()

        # Remove repeated mail prefixes such as RE_, RE-, FW:, FWD_.
        previous = None
        while name != previous:
            previous = name
            name = re.sub(r"^(?:RE|FW|FWD)[\s_:\-]+", "", name).strip()

        # Treat slash and underscore date punctuation consistently.
        return re.sub(r"[^A-Z0-9]+", "", name)

    @classmethod
    def _parts_from_evidence_text(cls, evidence: EmailEvidence) -> list[str]:
        """Capture S-prefixed part numbers from the reply and quoted request."""
        chunks = [
            getattr(evidence, "evidence_text", None),
            getattr(evidence, "approval_phrase", None),
        ]
        text = "\n".join(str(chunk) for chunk in chunks if chunk)
        return re.findall(r"\bS[A-Z0-9]{7,20}(?:[_/\-][A-Z0-9]{1,20})*\b", text.upper())

    @staticmethod
    def _not_found_approval() -> ApprovalMatch:
        return ApprovalMatch(
            approval_found=False,
            approval_status=ApprovalStatus.NOT_FOUND,
            evidence_file=None,
            evidence_phrase=None,
            evidence_text=None,
            approval_date=None,
        )

    # -----------------------------------------------------
    # Result construction
    # -----------------------------------------------------
    def _build_result(
        self,
        breakdown: BreakdownRecord,
        pavf: PavfRecord | None,
        approval: ApprovalMatch | None,
    ) -> ReconciliationResult:
        if pavf is None:
            return ReconciliationResult(
                breakdown=breakdown,
                pavf=None,
                approval=approval,
                status=ReconciliationStatus.PAVF_NOT_FOUND,
                notes=["No matching PAVF found for this part and date"],
            )

        if pavf.box_qty is not None and not self._numbers_match(
            pavf.box_qty,
            breakdown.box_qty,
        ):
            return ReconciliationResult(
                breakdown=breakdown,
                pavf=pavf,
                approval=approval,
                status=ReconciliationStatus.BOX_QTY_MISMATCH,
                notes=[
                    f"Breakdown boxes {breakdown.box_qty}",
                    f"PAVF boxes {pavf.box_qty}",
                ],
            )

        notes = ["PAVF located", "Box quantity matches"]

        if approval is None or not approval.approval_found:
            return ReconciliationResult(
                breakdown=breakdown,
                pavf=pavf,
                approval=approval,
                status=ReconciliationStatus.APPROVAL_NOT_FOUND,
                notes=notes + ["No explicit approval found in the matching email thread"],
            )

        return ReconciliationResult(
            breakdown=breakdown,
            pavf=pavf,
            approval=approval,
            status=ReconciliationStatus.APPROVED,
            notes=notes + ["Approval evidence located in matching email thread"],
        )

    # -----------------------------------------------------
    # Normalization helpers
    # -----------------------------------------------------
    @classmethod
    def _parts_match(cls, left: str | None, right: str | None) -> bool:
        left_norm = cls.normalize_part(left)
        right_norm = cls.normalize_part(right)

        if not left_norm or not right_norm:
            return False
        if left_norm == right_norm:
            return True

        return bool(cls._group_members(left).intersection(cls._group_members(right)))

    @classmethod
    def _group_members(cls, value: str | None) -> set[str]:
        if not value:
            return set()

        raw = str(value).upper().strip().replace(" ", "")
        pieces = [piece for piece in re.split(r"[_/]+", raw) if piece]
        if not pieces:
            return set()

        first = cls.normalize_part(pieces[0])
        members = {first} if first else set()

        for suffix in pieces[1:]:
            clean = cls.normalize_part(suffix)
            if not clean:
                continue
            if clean.startswith("S") and len(clean) >= 8:
                members.add(clean)
            elif len(clean) < len(first):
                members.add(first[:-len(clean)] + clean)
            else:
                members.add(clean)

        return members

    @staticmethod
    def normalize_part(value: str | None) -> str:
        if not value:
            return ""
        return re.sub(r"[^A-Z0-9]", "", str(value).upper())

    @staticmethod
    def _numbers_match(left, right) -> bool:
        try:
            return float(left) == float(right)
        except (TypeError, ValueError):
            return left == right

    @classmethod
    def _dates_match(cls, left, right) -> bool:
        left_date = cls._as_date(left)
        right_date = cls._as_date(right)
        return (
            left_date is not None
            and right_date is not None
            and left_date == right_date
        )

    @staticmethod
    def _as_date(value) -> date | None:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value

        text = str(value).strip()
        if not text:
            return None

        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%m-%d-%Y",
            "%m-%d-%y",
        ):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    @classmethod
    def _date_sort_value(cls, value) -> datetime:
        parsed = cls._as_date(value)
        return datetime.combine(parsed, datetime.min.time()) if parsed else datetime.min
