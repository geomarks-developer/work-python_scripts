from .models import (
    ApprovalLedgerRecord,
    ApprovalStatus
)

from .reconciliation_matcher import (
    ReconciliationMatcher
)


class ApprovalLedgerBuilder:

    def build(
        self,
        pavf_records,
        email_evidence,
    ):
        matcher = ReconciliationMatcher()

        ledger = []

        for pavf in pavf_records:

            approval = matcher._find_thread_approval(
                breakdown=pavf,
                pavf=pavf,
                evidence_list=email_evidence,
            )

            ledger.append(
                ApprovalLedgerRecord(
                    shipment_date=pavf.problem_date,
                    part_number=pavf.part_number,

                    box_qty=pavf.box_qty,
                    po_number=pavf.po_number,

                    approval_status=approval.approval_status.value,

                    approval_date=approval.approval_date,

                    pavf_file=pavf.source_file,

                    approval_email=approval.evidence_file,

                    approval_phrase=approval.evidence_phrase,
                )
            )

        return ledger