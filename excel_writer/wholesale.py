from excel_writer.template import ExcelWriter


class WholesaleExcelWriter(ExcelWriter):
    @property
    def output_schema(self) -> list:
        return [
            "PO#",
            "Material",
            "Description",
            "Qty",
            "Unit Price",
            "Create Date",
            "Due Date",
            "GT CRD",
        ]

    @property
    def date_columns(self) -> list:
        return ["Create Date", "Due Date", "GT CRD"]

    @property
    def number_columns(self) -> list:
        return ["Qty", "Unit Price"]
