"""
reconciliation_report.py

Creates a formatted Excel reconciliation report from a list of
ReconciliationResult objects produced by reconciliation_matcher.py.

Output sheets:
    1. Summary
    2. Reconciliation Detail
    3. Exceptions Review
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Any

from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


class ReconciliationReport:
    """Build an Excel report from reconciliation results."""

    DETAIL_HEADERS = [
        "Breakdown Date",
        "Part Number",
        "Program",
        "Part Qty",
        "Breakdown Box Qty",
        "Amount",
        "PAVF Found",
        "PAVF Problem Date",
        "PAVF Box Qty",
        "PO Number",
        "Approval Status",
        "Approval Date",
        "Evidence File",
        "Evidence Phrase",
        "Reconciliation Status",
        "Notes",
    ]

    EXCEPTION_HEADERS = [
        "Status",
        "Part Number",
        "Breakdown Date",
        "Amount",
        "Evidence / Issue",
        "Reviewer Notes",
    ]

    NAVY = "17365D"
    WHITE = "FFFFFF"
    LIGHT_BLUE = "DDEBF7"
    LIGHT_GREEN = "C6EFCE"
    LIGHT_YELLOW = "FFF2CC"
    LIGHT_ORANGE = "F4B183"
    LIGHT_RED = "F4CCCC"

    APPROVED = "APPROVED"
    PAVF_NOT_FOUND = "PAVF_NOT_FOUND"
    APPROVAL_NOT_FOUND = "APPROVAL_NOT_FOUND"
    BOX_QTY_MISMATCH = "BOX_QTY_MISMATCH"

    def __init__(self, output_path: str | Path):
        self.output_path = Path(output_path)

    def create(self, results: Iterable[Any]) -> Path:
        """Create the reconciliation workbook and return its path."""
        results = list(results)

        self.output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        detail_ws = wb.active
        detail_ws.title = "Reconciliation Detail"

        self._build_detail_sheet(detail_ws, results)

        summary_ws = wb.create_sheet("Summary", 0)
        self._build_summary_sheet(summary_ws, len(results))

        exceptions_ws = wb.create_sheet("Exceptions Review")
        self._build_exceptions_sheet(exceptions_ws, results)

        wb.save(self.output_path)
        return self.output_path

    def _build_detail_sheet(self, ws, results: list[Any]) -> None:
        ws.sheet_view.showGridLines = False
        ws.append(self.DETAIL_HEADERS)
        self._style_header(ws, 1)
        ws.freeze_panes = "A2"

        for result in results:
            ws.append(self._detail_row(result))

        widths = [
            16, 18, 28, 12, 19, 14, 13, 18,
            15, 16, 18, 18, 32, 34, 24, 46,
        ]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        for row in range(2, ws.max_row + 1):
            for col in (1, 8, 12):
                ws.cell(row, col).number_format = "mm/dd/yyyy"
            ws.cell(row, 6).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            for col in (4, 5, 9):
                ws.cell(row, col).number_format = '#,##0;[Red](#,##0);-'
            ws.cell(row, 16).alignment = Alignment(
                vertical="top", wrap_text=True
            )

        if results:
            table_ref = f"A1:P{ws.max_row}"
            table = Table(
                displayName="ReconciliationTable",
                ref=table_ref,
            )
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            ws.add_table(table)

            status_range = f"O2:O{ws.max_row}"
            self._add_status_formatting(ws, status_range)

    def _build_summary_sheet(self, ws, result_count: int) -> None:
        ws.sheet_view.showGridLines = False
        ws.merge_cells("A1:H1")
        ws["A1"] = "PAVF Reconciliation Report"
        ws["A1"].fill = PatternFill("solid", fgColor=self.NAVY)
        ws["A1"].font = Font(color=self.WHITE, bold=True, size=18)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 30

        ws["A3"] = "Reconciliation Summary"
        ws["A3"].font = Font(bold=True, size=14, color=self.NAVY)

        # Detail data always starts at row 2. Keep formulas valid for an empty set.
        last_row = max(result_count + 1, 2)
        status_range = f"'Reconciliation Detail'!O2:O{last_row}"
        amount_range = f"'Reconciliation Detail'!F2:F{last_row}"

        metrics = [
            ("Total Breakdown Records", result_count),
            ("Approved", f'=COUNTIF({status_range},"{self.APPROVED}")'),
            ("PAVF Not Found", f'=COUNTIF({status_range},"{self.PAVF_NOT_FOUND}")'),
            ("Approval Not Found", f'=COUNTIF({status_range},"{self.APPROVAL_NOT_FOUND}")'),
            ("Box Qty Mismatch", f'=COUNTIF({status_range},"{self.BOX_QTY_MISMATCH}")'),
            ("Total Breakdown Amount", f"=SUM({amount_range})"),
        ]

        for row, (label, value) in enumerate(metrics, start=4):
            ws.cell(row, 1, label)
            ws.cell(row, 1).font = Font(bold=True, color="666666")
            ws.cell(row, 2, value)
            ws.cell(row, 2).fill = PatternFill(
                "solid", fgColor=self.LIGHT_BLUE
            )
            ws.cell(row, 2).font = Font(bold=True, size=13)

        ws["B9"].number_format = '$#,##0.00;[Red]($#,##0.00);-'

        ws["A11"] = "Status"
        ws["B11"] = "Count"
        self._style_header(ws, 11, start_col=1, end_col=2)

        statuses = [
            self.APPROVED,
            self.PAVF_NOT_FOUND,
            self.APPROVAL_NOT_FOUND,
            self.BOX_QTY_MISMATCH,
        ]
        for row, status in enumerate(statuses, start=12):
            ws.cell(row, 1, status)
            ws.cell(row, 2, f'=COUNTIF({status_range},A{row})')

        chart = DoughnutChart()
        chart.title = "Reconciliation Status"
        chart.holeSize = 55
        chart.height = 7.5
        chart.width = 11

        data = Reference(ws, min_col=2, min_row=11, max_row=15)
        labels = Reference(ws, min_col=1, min_row=12, max_row=15)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showLeaderLines = True
        ws.add_chart(chart, "D3")

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 18
        for col in ("D", "E", "F", "G", "H"):
            ws.column_dimensions[col].width = 16

    def _build_exceptions_sheet(self, ws, results: list[Any]) -> None:
        ws.sheet_view.showGridLines = False
        ws.merge_cells("A1:F1")
        ws["A1"] = "Exceptions Review Queue"
        ws["A1"].fill = PatternFill("solid", fgColor=self.NAVY)
        ws["A1"].font = Font(color=self.WHITE, bold=True, size=16)
        ws["A1"].alignment = Alignment(horizontal="center")

        for col, header in enumerate(self.EXCEPTION_HEADERS, start=1):
            ws.cell(2, col, header)
        self._style_header(ws, 2)
        ws.freeze_panes = "A3"

        row_num = 3
        for result in results:
            status = self._enum_value(self._get(result, "status"))
            if status == self.APPROVED:
                continue

            breakdown = self._get(result, "breakdown")
            approval = self._get(result, "approval")
            notes = self._notes_text(self._get(result, "notes", []))
            evidence = self._get(approval, "evidence_phrase")
            issue = notes or self._clean(evidence)

            ws.append([
                status,
                self._get(breakdown, "part_number"),
                self._get(breakdown, "date"),
                self._get(breakdown, "amount"),
                issue,
                "",
            ])
            ws.cell(row_num, 3).number_format = "mm/dd/yyyy"
            ws.cell(row_num, 4).number_format = (
                '$#,##0.00;[Red]($#,##0.00);-'
            )
            ws.cell(row_num, 5).alignment = Alignment(wrap_text=True)
            ws.cell(row_num, 6).alignment = Alignment(wrap_text=True)
            row_num += 1

        widths = [24, 18, 16, 14, 48, 48]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        if ws.max_row >= 3:
            table = Table(
                displayName="ExceptionsTable",
                ref=f"A2:F{ws.max_row}",
            )
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showRowStripes=True,
                showFirstColumn=False,
                showLastColumn=False,
            )
            ws.add_table(table)
            self._add_status_formatting(ws, f"A3:A{ws.max_row}")

    def _detail_row(self, result: Any) -> list[Any]:
        breakdown = self._get(result, "breakdown")
        pavf = self._get(result, "pavf")
        approval = self._get(result, "approval")

        approval_status = self._enum_value(
            self._get(approval, "approval_status")
        )
        reconciliation_status = self._enum_value(
            self._get(result, "status")
        )

        return [
            self._excel_value(self._get(breakdown, "date")),
            self._get(breakdown, "part_number"),
            self._get(breakdown, "program"),
            self._get(breakdown, "part_qty"),
            self._get(breakdown, "box_qty"),
            self._get(breakdown, "amount"),
            "YES" if pavf is not None else "NO",
            self._excel_value(self._get(pavf, "problem_date")),
            self._get(pavf, "box_qty"),
            self._get(pavf, "po_number"),
            approval_status,
            self._excel_value(self._get(approval, "approval_date")),
            self._path_text(self._get(approval, "evidence_file")),
            self._get(approval, "evidence_phrase"),
            reconciliation_status,
            self._notes_text(self._get(result, "notes", [])),
        ]

    def _add_status_formatting(self, ws, cell_range: str) -> None:
        top_left = cell_range.split(":", 1)[0]
        rules = [
            (self.APPROVED, self.LIGHT_GREEN),
            (self.PAVF_NOT_FOUND, self.LIGHT_RED),
            (self.APPROVAL_NOT_FOUND, self.LIGHT_YELLOW),
            (self.BOX_QTY_MISMATCH, self.LIGHT_ORANGE),
        ]
        for status, color in rules:
            ws.conditional_formatting.add(
                cell_range,
                FormulaRule(
                    formula=[f'{top_left}="{status}"'],
                    fill=PatternFill("solid", fgColor=color),
                ),
            )

    def _style_header(
        self,
        ws,
        row: int,
        start_col: int = 1,
        end_col: int | None = None,
    ) -> None:
        end_col = end_col or ws.max_column
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row, col)
            cell.fill = PatternFill("solid", fgColor=self.NAVY)
            cell.font = Font(color=self.WHITE, bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
        ws.row_dimensions[row].height = 30

    @staticmethod
    def _get(obj: Any, name: str, default: Any = None) -> Any:
        if obj is None:
            return default
        if isinstance(obj, dict):
            return obj.get(name, default)
        return getattr(obj, name, default)

    @staticmethod
    def _clean(value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @classmethod
    def _enum_value(cls, value: Any) -> str:
        if value is None:
            return ""
        raw = getattr(value, "value", value)
        return cls._clean(raw)

    @staticmethod
    def _path_text(value: Any) -> str:
        if value is None:
            return ""
        return str(value)

    @staticmethod
    def _excel_value(value: Any) -> Any:
        # Excel handles date/datetime natively. If a timezone-aware datetime
        # arrives from parsed email evidence, remove timezone information because
        # Excel cannot store tz-aware datetime values.
        if isinstance(value, datetime) and value.tzinfo is not None:
            return value.replace(tzinfo=None)
        if isinstance(value, (date, datetime)):
            return value
        return value

    @classmethod
    def _notes_text(cls, notes: Any) -> str:
        if notes is None:
            return ""
        if isinstance(notes, (list, tuple, set)):
            return "; ".join(cls._clean(x) for x in notes if x is not None)
        return cls._clean(notes)


def create_reconciliation_report(
    results: Iterable[Any],
    output_path: str | Path = "PAVF_Reconciliation_Report.xlsx",
) -> Path:
    """Convenience function used by the PAVF orchestration layer."""
    return ReconciliationReport(output_path).create(results)
