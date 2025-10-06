import re
from datetime import datetime, timedelta
from typing import Dict, List

from pdf_parser.template import POParser


class WholesalePOParser(POParser):
    def parse_po_content(self, text: str, file_type: str = "original") -> List[Dict]:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        PO_ID_POSITION = 2
        CREATE_DATE_POSITION = 18 if file_type == "original" else 19
        INFO_POSITION = -1
        target_pattern = "No./Description"
        for i, line in enumerate(lines):
            if target_pattern in line:
                INFO_POSITION = i + 1 if file_type == "original" else i + 2
                break

        result = {}

        # PO# ("Purchase Order xxxxxx")
        if len(lines) >= PO_ID_POSITION + 1:
            po_match = re.search(r"Purchase Order\s+([A-Z0-9]+)", lines[PO_ID_POSITION])
            if po_match:
                result["PO#"] = po_match.group(1)

        # Material, Description, Qty, Unit Price (use "EACH" as anchor)
        if len(lines) >= INFO_POSITION + 2:
            mdqu_match = re.search(
                r"\d+\s+\w+\s+([A-Z0-9\-]+)\s+(.+?)\s+(\d+)\s+EACH\s+(\d+\.\d+)",
                lines[INFO_POSITION],
            )
            if mdqu_match:
                result["Material"] = mdqu_match.group(1)
                desc_part1 = mdqu_match.group(2).strip()
                desc_part2 = lines[INFO_POSITION + 1]
                result["Description"] = (desc_part1 + " " + desc_part2).strip()
                result["Qty"] = mdqu_match.group(3)
                result["Unit Price"] = mdqu_match.group(4)

        # Create Date (first item in the next line of "Date Terms Ship Via")
        if len(lines) >= CREATE_DATE_POSITION + 1:
            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", lines[CREATE_DATE_POSITION])
            if date_match:
                create_date_str = date_match.group(1)
                result["Create Date"] = datetime.strptime(create_date_str, "%m/%d/%Y")

        # Due Date (right after "Delivery Requested Date")
        due_date_match = re.search(
            r"Delivery Requested Date\s+(\d{2}/\d{2}/\d{4})", text
        )
        if due_date_match:
            due_date_str = due_date_match.group(1)
            result["Due Date"] = datetime.strptime(due_date_str, "%m/%d/%Y")

        # calculate GT CRD
        if "Create Date" in result.keys():
            result["GT CRD"] = result["Create Date"] + timedelta(days=self.gt_crd_days)

        return [result]


def parse_po_text(text: str, file_type: str = "original") -> dict:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    PO_ID_POSITION = 2
    CREATE_DATE_POSITION = 18 if file_type == "original" else 19
    INFO_POSITION = -1
    target_pattern = "No./Description"
    for i, line in enumerate(lines):
        if target_pattern in line:
            INFO_POSITION = i + 1 if file_type == "original" else i + 2
            break

    result = {}

    # PO# ("Purchase Order xxxxxx")
    if len(lines) >= PO_ID_POSITION + 1:
        po_match = re.search(r"Purchase Order\s+([A-Z0-9]+)", lines[PO_ID_POSITION])
        if po_match:
            result["PO#"] = po_match.group(1)

    # Material, Description, Qty, Unit Price (use "EACH" as anchor)
    if len(lines) >= INFO_POSITION + 2:
        mdqu_match = re.search(
            r"\d+\s+\w+\s+([A-Z0-9\-]+)\s+(.+?)\s+(\d+)\s+EACH\s+(\d+\.\d+)",
            lines[INFO_POSITION],
        )
        if mdqu_match:
            result["Material"] = mdqu_match.group(1)
            desc_part1 = mdqu_match.group(2).strip()
            desc_part2 = lines[INFO_POSITION + 1]
            result["Description"] = (desc_part1 + " " + desc_part2).strip()
            result["Qty"] = mdqu_match.group(3)
            result["Unit Price"] = mdqu_match.group(4)

    # Create Date (first item in the next line of "Date Terms Ship Via")
    if len(lines) >= CREATE_DATE_POSITION + 1:
        date_match = re.search(r"(\d{2}/\d{2}/\d{4})", lines[CREATE_DATE_POSITION])
        if date_match:
            create_date_str = date_match.group(1)
            result["Create Date"] = datetime.strptime(create_date_str, "%m/%d/%Y")

    # Due Date (right after "Delivery Requested Date")
    due_date_match = re.search(r"Delivery Requested Date\s+(\d{2}/\d{2}/\d{4})", text)
    if due_date_match:
        due_date_str = due_date_match.group(1)
        result["Due Date"] = datetime.strptime(due_date_str, "%m/%d/%Y")

    # calculate GT CRD
    if "Create Date" in result.keys():
        result["GT CRD"] = result["Create Date"] + timedelta(days=GT_CRD_DAYS)

    return result
