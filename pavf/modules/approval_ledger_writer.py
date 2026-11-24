import pandas as pd


class ApprovalLedgerWriter:

    def write(
        self,
        records,
        output_file,
    ):

        rows = []

        for r in records:

            rows.append(
                {
                    "Date":
                        r.shipment_date,

                    "Part Number":
                        r.part_number,

                    "Cardboard Boxes":
                        r.box_qty,

                    "PO Number":
                        r.po_number,

                    "Approval":
                        r.approval_status,

                    "Approval Date":
                        r.approval_date,

                    "PAVF File":
                        r.pavf_file,

                    "Approval Email":
                        r.approval_email,
                }
            )

        df = pd.DataFrame(rows)

        df = df.sort_values(
            ["Date", "Part Number"]
        )

        df.to_excel(
            output_file,
            index=False
        )