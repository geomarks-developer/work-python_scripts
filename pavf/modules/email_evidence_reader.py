"""Read saved PAVF email evidence into EmailEvidence objects."""
from __future__ import annotations
from dataclasses import fields, is_dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from enum import Enum
from pathlib import Path
import re
from typing import Any
from .email_classifier import EmailClassifier, EmailType
from .models import ApprovalStatus, EmailEvidence

class EmailEvidenceReader:
    SUPPORTED_SUFFIXES = {".eml", ".txt"}
    # These describe a requirement/process, not a granted approval.
    NON_GRANT_PATTERNS = (
        r"\bmust\s+be\s+approved\b",
        r"\bneeds?\s+to\s+be\s+approved\b",
        r"\brequires?\s+approval\b",
        r"\bprior\s+approval\b",
        r"\bbefore\s+giving\b.{0,40}\bapproval\b",
        r"\bapproval\s+(?:is\s+)?required\b",
        r"\blooking\s+for\s+approval\b",
        r"\brequest(?:ing)?\s+approval\b",
        r"\bplease\s+(?:let\s+me\s+know\s+if|approve)\b",
    )
    # Affirmative phrases that actually grant permission.
    GRANT_PATTERNS = (
        r"\bthis\s+is\s+approved\b",
        r"\bis\s+approved\b",
        r"\bare\s+approved\b",
        r"\bhas\s+been\s+approved\b",
        r"\bhave\s+been\s+approved\b",
        r"\bapproved\s+for\b",
        r"\bapproved\s+to\b",
        r"\byou\s+(?:are|['’]re)\s+approved\b",
        r"\bwe\s+approve\b",
        r"\bapproval\s+(?:is\s+)?granted\b",
        r"^\s*approved[.!]?\s*$",
        r"\bgood\s+to\s+ship\b",
        r"\bok(?:ay)?\s+to\s+ship\b",
    )

    def __init__(self, source: str | Path):
        self.source = Path(source)
        self.classifier = EmailClassifier()
        self._model_fields = self._get_model_fields(EmailEvidence)

    def read(self) -> list[EmailEvidence]:
        evidence = []
        for path in self._source_files():
            message = self._read_message(path)
            classification = self.classifier.classify(
                subject=message["subject"], sender=message["sender"], body=message["body"]
            )
            evidence.append(self._build_evidence(path, message, classification))
        return evidence

    def _source_files(self):
        if not self.source.exists():
            raise FileNotFoundError(f"Email evidence source not found: {self.source}")
        if self.source.is_file():
            if self.source.suffix.lower() not in self.SUPPORTED_SUFFIXES:
                raise ValueError(f"Unsupported email evidence file type: {self.source.suffix}")
            return [self.source]
        files = sorted(p for p in self.source.rglob("*") if p.is_file() and p.suffix.lower() in self.SUPPORTED_SUFFIXES and not p.name.startswith("~$"))
        if not files:
            raise ValueError(f"No .eml or .txt evidence files found in {self.source}")
        return files

    def _read_message(self, path):
        return self._read_eml(path) if path.suffix.lower() == ".eml" else self._read_txt(path)

    def _read_eml(self, path):
        with path.open("rb") as handle:
            msg = BytesParser(policy=policy.default).parse(handle)
        return {"subject": self._clean_header(msg.get("Subject")), "sender": self._clean_header(msg.get("From")), "email_date": self._parse_email_date(msg.get("Date")), "body": self._extract_body(msg)}

    def _read_txt(self, path):
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        return {"subject": self._extract_text_header(text, "subject"), "sender": self._extract_text_header(text, "from"), "email_date": self._parse_email_date(self._extract_text_header(text, "date")), "body": text}

    def _build_evidence(self, path, message, classification):
        body = message.get("body", "")

        granted_phrase = self._find_granted_approval(body)

        approval_status = self._approval_status(
            classification.email_type,
            granted_phrase
        )

        phrase = (
            granted_phrase
            if approval_status == ApprovalStatus.APPROVED
            else None
        )

        evidence_text = self._evidence_context(body, phrase)

        candidates = {
            "email_file": str(path),
            "email_date": message.get("email_date"),
            "subject": message.get("subject"),
            "sender": message.get("sender"),
            "related_parts": list(classification.part_numbers),
            "approval_status": approval_status,
            "approval_phrase": phrase,
            "evidence_text": evidence_text,
            "email_type": self._enum_value(classification.email_type),
            "confidence": classification.confidence,
        }

        data = {
            k: v
            for k, v in candidates.items()
            if k in self._model_fields
        }

        try:
            return EmailEvidence(**data)
        except TypeError as exc:
            raise TypeError(
                f"Could not create EmailEvidence from {path.name}. "
                f"Parsed fields: {sorted(data)}."
            ) from exc

        



    @classmethod
    def _find_granted_approval(cls, body: str) -> str | None:
        if not body:
            return None

        segments = re.split(r"[\r\n]+|(?<=[.!?])\s+", body)

        for segment in segments:
            text = re.sub(r"\s+", " ", segment).strip()

            if not text:
                continue

            if any(re.search(p, text, re.I) for p in cls.NON_GRANT_PATTERNS):
                continue

            for pattern in cls.GRANT_PATTERNS:
                match = re.search(pattern, text, re.I)

                if match:
                    return text

        return None

    @staticmethod
    def _approval_status(email_type, granted_phrase):
        # Explicit affirmative approval language wins.
        if granted_phrase:
            return ApprovalStatus.APPROVED

        for name in (
            "NOT_FOUND",
            "PENDING",
            "NOT_APPROVED",
        ):
            if hasattr(ApprovalStatus, name):
                return getattr(ApprovalStatus, name)

        raise AttributeError(
            "ApprovalStatus must define APPROVED plus "
            "a non-approved state."
        )

    @staticmethod
    def _extract_body(msg):
        if msg.is_multipart():
            plain, html = [], []
            for part in msg.walk():
                if part.get_content_disposition() == "attachment": continue
                if part.get_content_type() not in {"text/plain", "text/html"}: continue
                try: content = part.get_content()
                except Exception: continue
                if not isinstance(content, str): continue
                (plain if part.get_content_type() == "text/plain" else html).append(content)
            if plain: return "\n".join(plain).strip()
            return EmailEvidenceReader._html_to_text("\n".join(html))
        try: content = msg.get_content()
        except Exception: content = ""
        return EmailEvidenceReader._html_to_text(str(content)) if msg.get_content_type() == "text/html" else str(content or "").strip()

    @staticmethod
    def _html_to_text(html):
        text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"(?i)</p\s*>", "\n", text)
        text = re.sub(r"<[^>]+>", " ", text)
        return re.sub(r"[ \t]+", " ", text).strip()

    @staticmethod
    def _parse_email_date(value):
        if not value: return None
        if isinstance(value, datetime): return value
        try: return parsedate_to_datetime(str(value).strip())
        except (TypeError, ValueError, OverflowError): pass
        try: return datetime.fromisoformat(str(value).strip().replace("Z", "+00:00"))
        except ValueError: return None

    @staticmethod
    def _extract_text_header(text, header):
        m = re.search(rf"(?im)^\s*{re.escape(header)}\s*:\s*(.+?)\s*$", text)
        return m.group(1).strip() if m else ""

    @staticmethod
    def _clean_header(value): return str(value or "").strip()

    @staticmethod
    def _evidence_context(body, phrase, window=220):
        if not phrase: return None
        m = re.search(re.escape(phrase), body, re.I)
        if not m: return phrase
        return re.sub(r"\s+", " ", body[max(0,m.start()-window):min(len(body),m.end()+window)]).strip()

    @staticmethod
    def _get_model_fields(model_type):
        if is_dataclass(model_type): return {f.name for f in fields(model_type)}
        annotations = getattr(model_type, "__annotations__", {})
        return set(annotations) if annotations else {"email_file", "email_date", "related_parts", "approval_status", "approval_phrase", "evidence_text"}

    @staticmethod
    def _enum_value(value): return str(value.value) if isinstance(value, Enum) else str(value or "")

def read_email_evidence(source: str | Path) -> list[EmailEvidence]:
    return EmailEvidenceReader(source).read()
