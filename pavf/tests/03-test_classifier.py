# 03-test_classifier.py

from pathlib import Path
from email import policy
from email.parser import BytesParser

from modules.email_classifier import (
    EmailClassifier
)

classifier = EmailClassifier()

for file in Path("data").glob("*.eml"):

    with open(file, "rb") as f:

        msg = BytesParser(
           policy=policy.default
        ).parse(f)

    body = ""

    body_part = msg.get_body(
       preferencelist=("plain",)
    )

    if body_part:
        body = body_part.get_content()

    result = classifier.classify(
        msg["Subject"] or "",
        msg["From"] or "",
        body
    )

    print()
    print("=" * 80)
    print(file.name)

    print(
        result.email_type,
        result.evidence
    )