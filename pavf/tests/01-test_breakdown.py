# test_breakdown.py

from modules.breakdown_reader import BreakdownReader

reader = BreakdownReader(
    r"data/Wabash Reimbursement Breakdown 2026-09.xlsx"
)

records = reader.read()

print(f"\nRecords Found: {len(records)}\n")

for record in records[:20]:
    print(record)