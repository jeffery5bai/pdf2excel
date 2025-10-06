from excel_writer.template import ExcelWriter


class RetailExcelWriter(ExcelWriter):
    @property
    def col_length_offset(self) -> int:
        return 5

    @property
    def output_schema(self) -> list:
        return [
            "Ship Date",
            "GT Confirmed Ship Date",
            "Kohler PO",
            "Kohler Sales Order#",
            "THD PO#",
            "Kohler SKU",
            "THD SKU",
            "Description",
            "Qty",
            "Unit Price",
            "Ship To",
            "Order Date",
        ]

    @property
    def date_columns(self) -> list:
        return ["Order Date", "Ship Date", "GT Confirmed Ship Date"]

    @property
    def number_columns(self) -> list:
        return ["Qty", "Unit Price"]
