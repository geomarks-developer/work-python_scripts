# 02-test_emails.py

from pathlib import Path
from email import policy
from email.parser import BytesParser

folder = Path("data")

for file in folder.glob("*.eml"):

    print("\n" + "=" * 80)
    print(file.name)

    with open(file, "rb") as f:
        msg = BytesParser(
            policy=policy.default
        ).parse(f)

    print("\nSUBJECT:")
    print(msg["Subject"])

    print("\nFROM:")
    print(msg["From"])

    body = ""

    body_part = msg.get_body(
        preferencelist=("plain",)
    )

    if body_part:
        body = body_part.get_content()

    print("\nBODY:")
    print(body[:1500])

    print("\nATTACHMENTS:")

    if msg.is_multipart():

        for part in msg.iter_attachments():

            print(
                "  ",
                part.get_filename()
            )