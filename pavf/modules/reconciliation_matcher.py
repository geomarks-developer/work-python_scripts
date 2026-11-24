"""Match breakdown rows to dated PAVFs and transaction-specific email approvals."""
from __future__ import annotations
import re
from datetime import date, datetime
from pathlib import Path
from typing import List
from .models import ApprovalStatus, ApprovalMatch, BreakdownRecord, EmailEvidence, PavfRecord, ReconciliationResult, ReconciliationStatus

class ReconciliationMatcher:
    def reconcile(self, breakdown_records: List[BreakdownRecord], pavf_records: List[PavfRecord], email_evidence: List[EmailEvidence]) -> List[ReconciliationResult]:
        results = []
        for breakdown in breakdown_records:
            pavf = self._find_best_pavf(breakdown, pavf_records)
            approval = self._find_thread_approval(breakdown, pavf, email_evidence) if pavf else self._not_found_approval()
            results.append(self._build_result(breakdown, pavf, approval))
        return results

    def _find_best_pavf(self, breakdown, pavf_records):
        candidates = [p for p in pavf_records if self._parts_match(breakdown.part_number, p.part_number) and self._dates_match(breakdown.date, p.problem_date)]
        if not candidates:
            return None
        return max(candidates, key=lambda p: (self._numbers_match(breakdown.box_qty, p.box_qty), bool(getattr(p, "po_number", None))))

    def _find_thread_approval(self, breakdown, pavf, evidence_list):
        transaction_date = self._as_date(pavf.problem_date)
        groups = {}
        for evidence in evidence_list:
            groups.setdefault(self._thread_key(evidence), []).append(evidence)

        approved = []
        for items in groups.values():
            parts = []
            dates = set()
            for item in items:
                parts.extend(list(getattr(item, "related_parts", []) or []))
                parts.extend(self._parts_from_evidence_text(item))
                dates.update(self._thread_transaction_dates(item, transaction_date))

            # part match is wrong#
            part_matches = any(
                self._parts_match(pavf.part_number, part)
                or self._parts_match(breakdown.part_number, part)
                for part in parts
            )

            date_matches = (
                transaction_date is not None
                and transaction_date in dates
            )

            approved_in_thread = [
                item
                for item in items
                if item.approval_status == ApprovalStatus.APPROVED
            ]

            #
            # If we have valid part numbers in the thread,
            # require both part and date matches.
            #
            if part_matches and date_matches:
                approved.extend(approved_in_thread)
                continue

            #
            # If the thread contains an approval and date match
            # but no matching part information, still accept it.
            #
            if approved_in_thread and date_matches:
                approved.extend(approved_in_thread)

        if not approved:
            return self._not_found_approval()
        best = max(approved, key=lambda item: self._date_sort_value(getattr(item, "email_date", None)))
        return ApprovalMatch(approval_found=True, approval_status=ApprovalStatus.APPROVED, evidence_file=best.email_file, evidence_phrase=best.approval_phrase, evidence_text=best.evidence_text, approval_date=best.email_date)

    @classmethod
    def _thread_transaction_dates(cls, evidence, reference_date):
        if reference_date is None:
            return set()
        values = [str(getattr(evidence, "subject", "") or ""), Path(str(getattr(evidence, "email_file", "") or "")).stem]
        found = set()
        for value in values:
            for match in re.finditer(r"(?<!\d)(1[0-2]|0?[1-9])[/_-](3[01]|[12]\d|0?[1-9])(?:[/_-](20\d{2}))?(?!\d)", value):
                year = int(match.group(3)) if match.group(3) else reference_date.year
                try: found.add(date(year, int(match.group(1)), int(match.group(2))))
                except ValueError: pass
        return found

    @staticmethod
    def _thread_key(evidence):
        name = Path(str(getattr(evidence, "email_file", "") or "")).stem.upper()
        old = None
        while old != name:
            old = name
            name = re.sub(r"^(?:RE|FW|FWD)[\s_:\-]+", "", name).strip()
        return re.sub(r"[^A-Z0-9]+", "", name)

    @staticmethod
    def _parts_from_evidence_text(evidence):
        text = "\n".join(str(x) for x in (getattr(evidence, "evidence_text", None), getattr(evidence, "approval_phrase", None)) if x)
        return re.findall(r"\bS\d{7,}(?:[_/\-]\d+)*\b", text.upper())

    @staticmethod
    def _not_found_approval():
        return ApprovalMatch(approval_found=False, approval_status=ApprovalStatus.NOT_FOUND, evidence_file=None, evidence_phrase=None, evidence_text=None, approval_date=None)

    def _build_result(self, breakdown, pavf, approval):
        if pavf is None:
            return ReconciliationResult(breakdown=breakdown, pavf=None, approval=approval, status=ReconciliationStatus.PAVF_NOT_FOUND, notes=["No matching PAVF found for this part and date"])
        if pavf.box_qty is not None and not self._numbers_match(pavf.box_qty, breakdown.box_qty):
            return ReconciliationResult(breakdown=breakdown, pavf=pavf, approval=approval, status=ReconciliationStatus.BOX_QTY_MISMATCH, notes=[f"Breakdown boxes {breakdown.box_qty}", f"PAVF boxes {pavf.box_qty}"])
        notes = ["PAVF located", "Box quantity matches"]
        if not approval or not approval.approval_found:
            return ReconciliationResult(breakdown=breakdown, pavf=pavf, approval=approval, status=ReconciliationStatus.APPROVAL_NOT_FOUND, notes=notes + ["No explicit approval found in the matching transaction thread"])
        return ReconciliationResult(breakdown=breakdown, pavf=pavf, approval=approval, status=ReconciliationStatus.APPROVED, notes=notes + ["Approval evidence located in matching transaction thread"])

    @classmethod
    def _parts_match(cls, left, right):
        if cls.normalize_part(left) == cls.normalize_part(right) and cls.normalize_part(left): return True
        return bool(cls._group_members(left) & cls._group_members(right))

    @classmethod
    def _group_members(cls, value):
        if not value: return set()
        pieces = [p for p in re.split(r"[_/]+", str(value).upper().replace(" ", "")) if p]
        if not pieces: return set()
        first = cls.normalize_part(pieces[0]); members = {first} if first else set()
        for suffix in pieces[1:]:
            clean = cls.normalize_part(suffix)
            if clean.startswith("S") and len(clean) >= 8: members.add(clean)
            elif clean and len(clean) < len(first): members.add(first[:-len(clean)] + clean)
            elif clean: members.add(clean)
        return members

    @staticmethod
    def normalize_part(value): return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    @staticmethod
    def _numbers_match(left, right):
        try: return float(left) == float(right)
        except (TypeError, ValueError): return left == right
    @classmethod
    def _dates_match(cls, left, right):
        a, b = cls._as_date(left), cls._as_date(right)
        return a is not None and b is not None and a == b
    @staticmethod
    def _as_date(value):
        if isinstance(value, datetime): return value.date()
        if isinstance(value, date): return value
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y"):
            try: return datetime.strptime(str(value).strip(), fmt).date()
            except (ValueError, AttributeError): pass
        return None
    @classmethod
    def _date_sort_value(cls, value):
        parsed = cls._as_date(value)
        return datetime.combine(parsed, datetime.min.time()) if parsed else datetime.min
