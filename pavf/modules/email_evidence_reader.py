"""
email_evidence_reader.py

Reads saved PAVF email evidence and converts messages into EmailEvidence
objects using email_classifier.EmailClassifier.

Supported input:
    - .eml files
    - .txt files
    - a folder containing .eml / .txt files

Outlook .msg is intentionally not guessed at here; it requires a separate
parser/dependency and can be added once a real sample is available.
"""
from __future__ import annotations

from dataclasses import fields, is_dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from enum import Enum
from pathlib import Path
import re
from typing import Any, Iterable

from .email_classifier import EmailClassifier, EmailType
from .models import ApprovalStatus, EmailEvidence


class EmailEvidenceReader:
    SUPPORTED_SUFFIXES = {".eml", ".txt"}

    def __init__(self, source: str | Path):
        self.source = Path(source)
        self.classifier = EmailClassifier()
        self._model_fields = self._get_model_fields(EmailEvidence)

    def read(self) -> list[EmailEvidence]:
        """Read all supported messages and return normalized evidence records."""
        evidence: list[EmailEvidence] = []
        for path in self._source_files():
            message = self._read_message(path)
            classification = self.classifier.classify(
                subject=message["subject"],
                sender=message["sender"],
                body=message["body"],
            )
            evidence.append(
                self._build_evidence(path, message, classification)
            )
        return evidence

    def _source_files(self) -> list[Path]:
        if not self.source.exists():
            raise FileNotFoundError(f"Email evidence source not found: {self.source}")

        if self.source.is_file():
            if self.source.suffix.lower() not in self.SUPPORTED_SUFFIXES:
                raise ValueError(
                    f"Unsupported email evidence file type: {self.source.suffix}"
                )
            return [self.source]

        files = sorted(
            path for path in self.source.rglob("*")
            if path.is_file()
            and path.suffix.lower() in self.SUPPORTED_SUFFIXES
            and not path.name.startswith("~$")
        )
        if not files:
            raise ValueError(f"No .eml or .txt evidence files found in {self.source}")
        return files

    def _read_message(self, path: Path) -> dict[str, Any]:
        if path.suffix.lower() == ".eml":
            return self._read_eml(path)
        return self._read_txt(path)

    def _read_eml(self, path: Path) -> dict[str, Any]:
        with path.open("rb") as handle:
            msg = BytesParser(policy=policy.default).parse(handle)

        subject = self._clean_header(msg.get("Subject"))
        sender = self._clean_header(msg.get("From"))
        sent_at = self._parse_email_date(msg.get("Date"))
        body = self._extract_body(msg)

        return {
            "subject": subject,
            "sender": sender,
            "email_date": sent_at,
            "body": body,
        }

    def _read_txt(self, path: Path) -> dict[str, Any]:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        subject = self._extract_text_header(text, "subject")
        sender = self._extract_text_header(text, "from")
        date_text = self._extract_text_header(text, "date")
        sent_at = self._parse_email_date(date_text)

        # Treat everything as body as well. The classifier benefits from having
        # any subject/header text available even when the export is informal.
        return {
            "subject": subject,
            "sender": sender,
            "email_date": sent_at,
            "body": text,
        }

    def _build_evidence(self, path: Path, message: dict[str, Any], classification: Any) -> EmailEvidence:
        approval_status = self._approval_status(classification.email_type)
        evidence_text = self._evidence_context(
            message.get("body", ""),
            classification.evidence,
        )

        candidates = {
            "email_file": str(path),
            "email_date": message.get("email_date"),
            "subject": message.get("subject"),
            "sender": message.get("sender"),
            "related_parts": list(classification.part_numbers),
            "approval_status": approval_status,
            "approval_phrase": classification.evidence,
            "evidence_text": evidence_text,
            "email_type": self._enum_value(classification.email_type),
            "confidence": classification.confidence,
        }

        data = {
            key: value
            for key, value in candidates.items()
            if key in self._model_fields
        }

        try:
            return EmailEvidence(**data)
        except TypeError as exc:
            raise TypeError(
                f"Could not create EmailEvidence from {path.name}. "
                f"Parsed fields: {sorted(data)}. Check models.py for required fields."
            ) from exc

    @staticmethod
    def _approval_status(email_type: EmailType) -> ApprovalStatus:
        if email_type == EmailType.APPROVED:
            for name in ("APPROVED",):
                if hasattr(ApprovalStatus, name):
                    return getattr(ApprovalStatus, name)

        # Prefer explicit non-approved states if the project model has one;
        # otherwise NOT_FOUND matches the current reconciliation matcher.
        for name in ("NOT_FOUND", "PENDING", "NOT_APPROVED"):
            if hasattr(ApprovalStatus, name):
                return getattr(ApprovalStatus, name)

        raise AttributeError(
            "ApprovalStatus must define APPROVED plus a non-approved state "
            "such as NOT_FOUND."
        )

    @staticmethod
    def _extract_body(msg: Any) -> str:
        if msg.is_multipart():
            plain_parts: list[str] = []
            html_parts: list[str] = []
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    continue
                content_type = part.get_content_type()
                if content_type not in {"text/plain", "text/html"}:
                    continue
                try:
                    content = part.get_content()
                except Exception:
                    continue
                if not isinstance(content, str):
                    continue
                if content_type == "text/plain":
                    plain_parts.append(content)
                else:
                    html_parts.append(content)
            if plain_parts:
                return "\n".join(plain_parts).strip()
            return EmailEvidenceReader._html_to_text("\n".join(html_parts))

        try:
            content = msg.get_content()
        except Exception:
            content = ""
        if msg.get_content_type() == "text/html":
            return EmailEvidenceReader._html_to_text(str(content))
        return str(content or "").strip()

    @staticmethod
    def _html_to_text(html: str) -> str:
        text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"(?i)</p\s*>", "\n", text)
        text = re.sub(r"<[^>]+>", " ", text)
        return re.sub(r"[ \t]+", " ", text).strip()

    @staticmethod
    def _parse_email_date(value: Any) -> datetime | None:
        if not value:
            return None
        if isinstance(value, datetime):
            return value
        text = str(value).strip()
        try:
            return parsedate_to_datetime(text)
        except (TypeError, ValueError, OverflowError):
            pass
        text = text.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(text)
        except ValueError:
            return None

    @staticmethod
    def _extract_text_header(text: str, header: str) -> str:
        match = re.search(
            rf"(?im)^\s*{re.escape(header)}\s*:\s*(.+?)\s*$",
            text,
        )
        return match.group(1).strip() if match else ""

    @staticmethod
    def _clean_header(value: Any) -> str:
        return str(value or "").strip()

    @staticmethod
    def _evidence_context(body: str, phrase: str | None, window: int = 220) -> str | None:
        if not phrase:
            return None
        match = re.search(re.escape(phrase), body, flags=re.IGNORECASE)
        if not match:
            return phrase
        start = max(0, match.start() - window)
        end = min(len(body), match.end() + window)
        context = body[start:end]
        return re.sub(r"\s+", " ", context).strip()

    @staticmethod
    def _get_model_fields(model_type: type) -> set[str]:
        if is_dataclass(model_type):
            return {field.name for field in fields(model_type)}
        annotations = getattr(model_type, "__annotations__", {})
        return set(annotations) if annotations else {
            "email_file", "email_date", "related_parts", "approval_status",
            "approval_phrase", "evidence_text",
        }

    @staticmethod
    def _enum_value(value: Any) -> str:
        if isinstance(value, Enum):
            return str(value.value)
        return str(value or "")


def read_email_evidence(source: str | Path) -> list[EmailEvidence]:
    """Convenience loader compatible with PAVFReconciliation(email_loader=...)."""
    return EmailEvidenceReader(source).read()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Parse saved PAVF email evidence.")
    parser.add_argument("source", help=".eml/.txt file or folder")
    args = parser.parse_args()

    records = read_email_evidence(args.source)
    print(f"Email evidence records found: {len(records)}")
    for record in records[:10]:
        print(record)
