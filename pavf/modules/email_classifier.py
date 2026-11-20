"""
email_classifier.py

Classifies PAVF email traffic into
business categories.

Author: George Marks Project
"""

from __future__ import annotations

import re

from dataclasses import dataclass
from enum import Enum


class EmailType(str, Enum):

    SUBMISSION = "SUBMISSION"

    REQUEST_APPROVAL = "REQUEST_APPROVAL"

    APPROVED = "APPROVED"

    GUIDANCE = "GUIDANCE"

    REVISION = "REVISION"

    OTHER = "OTHER"


@dataclass
class EmailClassification:

    email_type: EmailType

    confidence: int

    evidence: str | None

    part_numbers: list[str]

    approval_found: bool


class EmailClassifier:

    PART_PATTERN = re.compile(
        r"\bS\d+(?:[-_]\d+)?\b",
        re.IGNORECASE
    )

    APPROVAL_PATTERNS = [

        r"this is approved",

        r"approved for",

        r"alternate packaging is approved",

        r"alt pack .* approved",

        r"\bis approved\b",

        r"\bapproved\b"
    ]

    REQUEST_PATTERNS = [

        r"looking for approval",

        r"good to ship",

        r"please let me know if we are good",

        r"can we ship",

        r"ship boxes tomorrow",

        r"approval"
    ]

    GUIDANCE_PATTERNS = [

        r"must be approved",

        r"prior to shipment",

        r"approval to build into cardboard",

        r"newest version of the pavf",

        r"going forward"
    ]

    REVISION_PATTERNS = [

        r"revised pavf",

        r"revised reports",

        r"i.ve revised",

        r"sending.*revised",

        r"actual shipment"
    ]

    SUBMISSION_PATTERNS = [

        r"attached is the pavf",

        r"attached is the pavf report",

        r"attached is the report",

        r"parts that were shipped in boxes",

        r"parts that will be shipped in boxes"
    ]

    def classify(
        self,
        subject: str,
        sender: str,
        body: str,
    ) -> EmailClassification:

        text = (
            f"{subject}\n{body}"
        ).lower()

        parts = self._extract_parts(
            subject + "\n" + body
        )

        #
        # IMPORTANT:
        # Check GUIDANCE BEFORE APPROVED
        #
        # Example:
        #
        # "must be approved prior to shipment"
        #
        # is policy guidance, NOT approval.
        #
        guidance = self._find_pattern(
            text,
            self.GUIDANCE_PATTERNS
        )

        if guidance:

            return EmailClassification(

                email_type=
                    EmailType.GUIDANCE,

                confidence=95,

                evidence=guidance,

                part_numbers=parts,

                approval_found=False
            )

        approved = self._find_pattern(
            text,
            self.APPROVAL_PATTERNS
        )

        if approved:

            return EmailClassification(

                email_type=
                    EmailType.APPROVED,

                confidence=100,

                evidence=approved,

                part_numbers=parts,

                approval_found=True
            )

        revision = self._find_pattern(
            text,
            self.REVISION_PATTERNS
        )

        if revision:

            return EmailClassification(

                email_type=
                    EmailType.REVISION,

                confidence=90,

                evidence=revision,

                part_numbers=parts,

                approval_found=False
            )

        request = self._find_pattern(
            text,
            self.REQUEST_PATTERNS
        )

        if request:

            return EmailClassification(

                email_type=
                    EmailType.REQUEST_APPROVAL,

                confidence=90,

                evidence=request,

                part_numbers=parts,

                approval_found=False
            )

        submission = self._find_pattern(
            text,
            self.SUBMISSION_PATTERNS
        )

        if submission:

            return EmailClassification(

                email_type=
                    EmailType.SUBMISSION,

                confidence=85,

                evidence=submission,

                part_numbers=parts,

                approval_found=False
            )

        return EmailClassification(

            email_type=EmailType.OTHER,

            confidence=50,

            evidence=None,

            part_numbers=parts,

            approval_found=False
        )

    def _find_pattern(
        self,
        text: str,
        patterns: list[str]
    ) -> str | None:

        for pattern in patterns:

            match = re.search(
                pattern,
                text,
                flags=re.IGNORECASE
            )

            if match:
                return match.group(0)

        return None

    def _extract_parts(
        self,
        text: str
    ) -> list[str]:

        found = self.PART_PATTERN.findall(
            text
        )

        cleaned = []

        for part in found:

            value = (
                part.upper()
                .replace("-", "_")
            )

            if value not in cleaned:

                cleaned.append(value)

        return cleaned