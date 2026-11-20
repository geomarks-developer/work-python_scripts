"""
breakdown_reader.py

Reads the Wabash Reimbursement Breakdown workbook and converts the
Summary sheet into a normalized list of reimbursement records.

Author: George Marks Project
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Dict, Any

from openpyxl import load_workbook

import re

class BreakdownReader:
    """
    Reads reimbursement breakdown workbook.

    Converts:

        Date | Part Qty | Box Qty | Total $
                     repeating columns

    into

        [
            {
                "date": ...,
                "program": ...,
                "part_qty": ...,
                "box_qty": ...,
                "amount": ...
            }
        ]
    """

    SUMMARY_SHEET = "Summary"

    # Workbook layout as seen in screenshot:
    HEADER_ROW = 6
    DETAIL_HEADER_ROW = 7
    DATA_START_ROW = 8

    def __init__(self, workbook_path: str | Path):
        self.workbook_path = Path(workbook_path)

    def extract_part_number(self, header):

        if not header:
            return ""

        m = re.search(
            r'(S\d+(?:[_-]\d+)*)',
            str(header),
            re.IGNORECASE
        )

        if m:
            return m.group(1)

        return str(header)


    def read(self) -> List[Dict[str, Any]]:
        """
        Return normalized reimbursement records.
        """

        wb = load_workbook(
            self.workbook_path,
            data_only=True
        )

        if self.SUMMARY_SHEET not in wb.sheetnames:
            raise ValueError(
                f"Summary sheet not found in "
                f"{self.workbook_path.name}"
            )

        ws = wb[self.SUMMARY_SHEET]

        program_blocks = self._find_program_blocks(ws)

        records = []

        max_row = ws.max_row

        for row in range(self.DATA_START_ROW, max_row + 1):

            trx_date = ws.cell(row=row, column=2).value

            if trx_date is None:
                continue

            for block in program_blocks:

                part_qty = ws.cell(
                    row=row,
                    column=block["part_col"]
                ).value

                box_qty = ws.cell(
                    row=row,
                    column=block["box_col"]
                ).value

                amount = ws.cell(
                    row=row,
                    column=block["amount_col"]
                ).value

                part_qty = self._to_number(part_qty)
                box_qty = self._to_number(box_qty)
                amount = self._to_number(amount)

                # Skip empty entries
                if (
                    part_qty == 0 and
                    box_qty == 0 and
                    amount == 0
                ):
                    continue

                records.append({
                    "date": trx_date,
                    "program": block["program"],
                    "part_number": self.extract_part_number(block["program"]),
                    "part_qty": part_qty,
                    "box_qty": box_qty,
                    "amount": amount,
                    "part_col": block["part_col"],
                    "box_col": block["box_col"],
                    "amount_col": block["amount_col"]
                })

        wb.close()

        return records

    def _find_program_blocks(self, ws):
        """
        Find repeating program structure.

        Example:

        B = Date

        C,D,E  = S92201310_20
        F,G,H  = S92201330_40
        I,J,K  = S92201400_10

        Returns:

        [
            {
                "program":"S92201310_20",
                "part_col":3,
                "box_col":4,
                "amount_col":5
            }
        ]
        """

        blocks = []

        max_col = ws.max_column

        col = 3

        while col <= max_col:

            program = ws.cell(
                row=self.HEADER_ROW,
                column=col
            ).value

            detail = ws.cell(
                row=self.DETAIL_HEADER_ROW,
                column=col
            ).value

            if (
                program is not None and
                detail == "Part Qty"
            ):
                blocks.append({
                    "program": str(program).strip(),
                    "part_col": col,
                    "box_col": col + 1,
                    "amount_col": col + 2
                })

                col += 3

            else:
                col += 1

        return blocks

    @staticmethod
    def _to_number(value):

        if value is None:
            return 0

        if isinstance(value, (int, float)):
            return value

        try:
            cleaned = (
                str(value)
                .replace("$", "")
                .replace(",", "")
                .strip()
            )

            if cleaned == "":
                return 0

            return float(cleaned)

        except Exception:
            return 0


if __name__ == "__main__":

    reader = BreakdownReader(
        r"Wabash Reimbursement Breakdown 2026-09.xlsx"
    )

    records = reader.read()

    print(f"Records found: {len(records)}")

    for rec in records[:10]:
        print(rec)