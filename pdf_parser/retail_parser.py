import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional

from pdf_parser.template import POParser

class RetailPOParser(POParser):
    def parse_po_content(self, text: str, words: List[Dict]) -> List[Dict]:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        PO_ID_POSITION = 2
        CREATE_DATE_POSITION = 18

        info_positions = []
        sale_order_positions = []
        customer_po_positions = []
        delivery_requested_date_positions = []

        info_pattern = "No./Description"
        sale_order_pattern = "Kohler Sales Order Number"
        customer_po_pattern = "Customer Purchase Order Number"
        delivery_requested_date_pattern = "Delivery Requested Date"
        for i, line in enumerate(lines):
            if info_pattern in line:
                info_positions.append(i + 1)
            if sale_order_pattern in line:
                sale_order_positions.append(i)
            if customer_po_pattern in line:
                customer_po_positions.append(i)
            if delivery_requested_date_pattern in line:
                delivery_requested_date_positions.append(i)

        assert (
            len(info_positions)
            == len(sale_order_positions)
            == len(customer_po_positions)
            == len(delivery_requested_date_positions)
        ), f"Positions length mismatch: info {len(info_positions)}, sale_order {len(sale_order_positions)}, customer_po {len(customer_po_positions)}, delivery_request_date {len(delivery_request_date_positions)}"

        # Kohler PO ("Purchase Order xxxxxx")
        if len(lines) >= PO_ID_POSITION + 1:
            po_match = re.search(r"Purchase Order\s+([A-Z0-9]+)", lines[PO_ID_POSITION])
            if po_match:
                kohler_po = po_match.group(1)
        # Create Date (first item in the next line of "Date Terms Ship Via")
        if len(lines) >= CREATE_DATE_POSITION + 1:
            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", lines[CREATE_DATE_POSITION])
            if date_match:
                create_date_str = date_match.group(1)
                create_date = datetime.strptime(create_date_str, "%m/%d/%Y")

        ship_to = extract_ship_to_first_line(words, anchor_keyword="PNA", debug=False)

        results = []
        for idx in range(len(info_positions)):
            item = {"Kohler PO": kohler_po, "Order Date": create_date, "Ship To": ship_to, "THD SKU": ""}
            info_pos = info_positions[idx]
            sale_order_pos = sale_order_positions[idx]
            customer_po_pos = customer_po_positions[idx]
            delivery_requested_date_pos = delivery_requested_date_positions[idx]

            # Material, Description, Qty, Unit Price (use "EACH" as anchor)
            if len(lines) >= info_pos + 2:
                mdqu_match = re.search(
                    r"\d+\s+\w+\s+([A-Z0-9\-]+)\s+(.+?)\s+(\d+)\s+EACH\s+(\d+\.\d+)",
                    lines[info_pos],
                )
                if mdqu_match:
                    item["Kohler SKU"] = mdqu_match.group(1)
                    desc_part1 = mdqu_match.group(2).strip()
                    desc_part2 = lines[info_pos + 1]
                    item["Description"] = (desc_part1 + " " + desc_part2).strip()
                    item["Qty"] = mdqu_match.group(3)
                    item["Unit Price"] = mdqu_match.group(4)

            # Kohler Sales Order# (right after "Kohler Sales Order Number")
            if len(lines) >= sale_order_pos + 1:
                so_match = re.search(
                    r"KOHLER SALES ORDER\s+([A-Z0-9\-]+)", lines[sale_order_pos]
                )
                if so_match:
                    item["Kohler Sales Order#"] = so_match.group(1)

            # THD PO# (right after "Customer Purchase Order Number")
            if len(lines) >= customer_po_pos + 1:
                po_match = re.search(
                    r"CUSTOMER PO\s+([A-Z0-9\-]+)",
                    lines[customer_po_pos],
                )
                if po_match:
                    item["THD PO#"] = po_match.group(1)

            # Ship Date (right after "Delivery Requested Date")
            if len(lines) >= delivery_requested_date_pos + 1:
                ship_date_match = re.search(
                    r"Delivery Requested Date\s+(\d{2}/\d{2}/\d{4})",
                    lines[delivery_requested_date_pos],
                )
                if ship_date_match:
                    ship_date_str = ship_date_match.group(1)
                    item["Ship Date"] = datetime.strptime(ship_date_str, "%m/%d/%Y")

            # calculate GT Confirmed Ship Date
            if "Order Date" in item.keys():
                item["GT Confirmed Ship Date"] = item["Order Date"] + timedelta(
                    days=self.gt_crd_days
                )

            results.append(item)

        return results


    def extract_ship_to_first_line(
        self,
        words: List[Dict],
        anchor_keyword: str = "PNA",
        line_tol: float = 3.0,  # 同一行判定容差 (points)
        debug: bool = False,
    ) -> Optional[str]:
        """
        從 word-level 座標資料擷取 Ship-To 那一行 (去掉開頭 'S')，回傳像:
        "THD DI DFC #6707 - LUCKEY"

        Args:
            words: list of dict, 每個 dict 至少要有 keys: 'text','x0','x1','top','bottom'
            anchor_keyword: 用來定位 Bill-To anchor 的 keyword（預設 'PNA'）
            line_tol: 判定屬於同一水平列的 top 差容差
            gap_threshold: 將同一列分成多個 cluster 的水平 gap 閾值
            debug: 若 True 會回傳更多中間資訊（print 出來）
        Returns:
            Ship-To 第一行字串 (不含 leading 'S')，若找不到回傳 None
        """
        CLUSTER_WIDTH_TOL = 150
        CLUSTER_HEIGHT_TOL = 25

        if not words:
            return None

        # normalize text and ensure numeric fields exist
        normalized = []
        for w in words:
            if "text" not in w or "x0" not in w or "x1" not in w or "top" not in w:
                continue
            # create a shallow copy and normalized text
            ww = dict(w)
            ww["text"] = ww["text"].strip()
            normalized.append(ww)
        if not normalized:
            return None

        # 1) 找到所有包含 anchor_keyword 的 candidates（忽略大小寫）
        anchors = [w for w in normalized if anchor_keyword.upper() in w["text"].upper()]
        if not anchors:
            if debug:
                print("No anchor found for:", anchor_keyword)
            return None

        # 選擇「所在同一行包含最少 words」那個 anchor（比較穩定）
        def line_count_for(w):
            return sum(1 for t in normalized if abs(t["top"] - w["top"]) <= line_tol)

        anchor = min(anchors, key=line_count_for)

        if debug:
            print("Chosen anchor:", anchor)

        # 2) 取 anchor 所在行的所有 words (同一行)
        anchor_top = anchor["top"]
        anchor_line = [w for w in normalized if abs(w["top"] - anchor_top) <= line_tol]
        if not anchor_line:
            if debug:
                print("No same-line words found for anchor_top:", anchor_top)
            return None
        if len(anchor_line) < 2:
            if debug:
                print("Not enough words in the same line:", anchor_line)
            return None
        # 製作一個 psuedo word 代表第二欄
        psuedo_word = anchor_line[0].copy()
        psuedo_word["text"] = "SECOND_COLUMN"
        psuedo_word["x0"] = (anchor_line[0]["x0"] + anchor_line[1]["x0"]) / 2 - 10
        psuedo_word["x1"] = (anchor_line[0]["x1"] + anchor_line[1]["x1"]) / 2
        anchor_line.insert(1, psuedo_word)
        # 排序（由左到右）
        anchor_line.sort(key=lambda w: w["x0"])

        # 3) 依水平 gap 分群成 clusters（每個 cluster 代表同一欄）
        # 寬度 130，高度 25
        clusters = []
        for anchor in anchor_line:
            current = [anchor]
            current.extend(
                [
                    w
                    for w in normalized
                    if 0 <= w["top"] - anchor["top"] <= CLUSTER_HEIGHT_TOL
                    and 0 <= w["x1"] - anchor["x0"] <= CLUSTER_WIDTH_TOL
                ]
            )
            clusters.append(current)

        if debug:
            print("clusters count:", len(clusters))
            for i, c in enumerate(clusters):
                print(
                    f" cluster {i}: "
                    + " | ".join([f"{t['text']}@{int(t['x0'])}" for t in c])
                )

        # 4) 定位 Ship-To cluster 和 line
        ship_cluster = clusters[1]
        target_word = [w for w in ship_cluster if w["text"] == "S"][0]
        target_line_words = sorted(
            [w for w in ship_cluster if abs(w["top"] - target_word["top"]) <= line_tol],
            key=lambda w: (w["top"], w["x0"]),
        )
        target_line = " ".join(w["text"] for w in target_line_words).strip()

        # 5) 試著去掉行首的 'S' 或 'S,' 之類
        # 例如 "S THD DI DFC #6707 - LUCKEY" -> "THD DI DFC #6707 - LUCKEY"
        target_line = re.sub(r"^\s*S[\s,:-]*", "", target_line, flags=re.IGNORECASE).strip()

        if debug:
            print("Extracted target_line:", target_line)

        return target_line




def parse_retail_po_text(text: str, words: List[Dict]) -> List[Dict]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    PO_ID_POSITION = 2
    CREATE_DATE_POSITION = 18

    info_positions = []
    sale_order_positions = []
    customer_po_positions = []
    delivery_requested_date_positions = []

    info_pattern = "No./Description"
    sale_order_pattern = "Kohler Sales Order Number"
    customer_po_pattern = "Customer Purchase Order Number"
    delivery_requested_date_pattern = "Delivery Requested Date"
    for i, line in enumerate(lines):
        if info_pattern in line:
            info_positions.append(i + 1)
        if sale_order_pattern in line:
            sale_order_positions.append(i)
        if customer_po_pattern in line:
            customer_po_positions.append(i)
        if delivery_requested_date_pattern in line:
            delivery_requested_date_positions.append(i)

    assert (
        len(info_positions)
        == len(sale_order_positions)
        == len(customer_po_positions)
        == len(delivery_requested_date_positions)
    ), f"Positions length mismatch: info {len(info_positions)}, sale_order {len(sale_order_positions)}, customer_po {len(customer_po_positions)}, delivery_request_date {len(delivery_request_date_positions)}"

    # Kohler PO ("Purchase Order xxxxxx")
    if len(lines) >= PO_ID_POSITION + 1:
        po_match = re.search(r"Purchase Order\s+([A-Z0-9]+)", lines[PO_ID_POSITION])
        if po_match:
            kohler_po = po_match.group(1)
    # Create Date (first item in the next line of "Date Terms Ship Via")
    if len(lines) >= CREATE_DATE_POSITION + 1:
        date_match = re.search(r"(\d{2}/\d{2}/\d{4})", lines[CREATE_DATE_POSITION])
        if date_match:
            create_date_str = date_match.group(1)
            create_date = datetime.strptime(create_date_str, "%m/%d/%Y")

    ship_to = extract_ship_to_first_line(words, anchor_keyword="PNA", debug=False)

    results = []
    for idx in range(len(info_positions)):
        item = {"Kohler PO": kohler_po, "Order Date": create_date, "Ship To": ship_to}
        info_pos = info_positions[idx]
        sale_order_pos = sale_order_positions[idx]
        customer_po_pos = customer_po_positions[idx]
        delivery_requested_date_pos = delivery_requested_date_positions[idx]

        # Material, Description, Qty, Unit Price (use "EACH" as anchor)
        if len(lines) >= info_pos + 2:
            mdqu_match = re.search(
                r"\d+\s+\w+\s+([A-Z0-9\-]+)\s+(.+?)\s+(\d+)\s+EACH\s+(\d+\.\d+)",
                lines[info_pos],
            )
            if mdqu_match:
                item["Kohler SKU"] = mdqu_match.group(1)
                desc_part1 = mdqu_match.group(2).strip()
                desc_part2 = lines[info_pos + 1]
                item["Description"] = (desc_part1 + " " + desc_part2).strip()
                item["Qty"] = mdqu_match.group(3)
                item["Unit Price"] = mdqu_match.group(4)

        # Kohler Sales Order# (right after "Kohler Sales Order Number")
        if len(lines) >= sale_order_pos + 1:
            so_match = re.search(
                r"KOHLER SALES ORDER\s+([A-Z0-9\-]+)", lines[sale_order_pos]
            )
            if so_match:
                item["Kohler Sales Order#"] = so_match.group(1)

        # THD PO# (right after "Customer Purchase Order Number")
        if len(lines) >= customer_po_pos + 1:
            po_match = re.search(
                r"CUSTOMER PO\s+([A-Z0-9\-]+)",
                lines[customer_po_pos],
            )
            if po_match:
                item["THD PO#"] = po_match.group(1)

        # Ship Date (right after "Delivery Requested Date")
        if len(lines) >= delivery_requested_date_pos + 1:
            ship_date_match = re.search(
                r"Delivery Requested Date\s+(\d{2}/\d{2}/\d{4})",
                lines[delivery_requested_date_pos],
            )
            if ship_date_match:
                ship_date_str = ship_date_match.group(1)
                item["Ship Date"] = datetime.strptime(ship_date_str, "%m/%d/%Y")

        # calculate GT Confirmed Ship Date
        if "Order Date" in item.keys():
            item["GT Confirmed Ship Date"] = item["Order Date"] + timedelta(
                days=GT_CRD_DAYS
            )

        results.append(item)

    return results


def extract_ship_to_first_line(
    words: List[Dict],
    anchor_keyword: str = "PNA",
    line_tol: float = 3.0,  # 同一行判定容差 (points)
    debug: bool = False,
) -> Optional[str]:
    """
    從 word-level 座標資料擷取 Ship-To 那一行 (去掉開頭 'S')，回傳像:
      "THD DI DFC #6707 - LUCKEY"

    Args:
        words: list of dict, 每個 dict 至少要有 keys: 'text','x0','x1','top','bottom'
        anchor_keyword: 用來定位 Bill-To anchor 的 keyword（預設 'PNA'）
        line_tol: 判定屬於同一水平列的 top 差容差
        gap_threshold: 將同一列分成多個 cluster 的水平 gap 閾值
        debug: 若 True 會回傳更多中間資訊（print 出來）
    Returns:
        Ship-To 第一行字串 (不含 leading 'S')，若找不到回傳 None
    """
    CLUSTER_WIDTH_TOL = 150
    CLUSTER_HEIGHT_TOL = 25

    if not words:
        return None

    # normalize text and ensure numeric fields exist
    normalized = []
    for w in words:
        if "text" not in w or "x0" not in w or "x1" not in w or "top" not in w:
            continue
        # create a shallow copy and normalized text
        ww = dict(w)
        ww["text"] = ww["text"].strip()
        normalized.append(ww)
    if not normalized:
        return None

    # 1) 找到所有包含 anchor_keyword 的 candidates（忽略大小寫）
    anchors = [w for w in normalized if anchor_keyword.upper() in w["text"].upper()]
    if not anchors:
        if debug:
            print("No anchor found for:", anchor_keyword)
        return None

    # 選擇「所在同一行包含最少 words」那個 anchor（比較穩定）
    def line_count_for(w):
        return sum(1 for t in normalized if abs(t["top"] - w["top"]) <= line_tol)

    anchor = min(anchors, key=line_count_for)

    if debug:
        print("Chosen anchor:", anchor)

    # 2) 取 anchor 所在行的所有 words (同一行)
    anchor_top = anchor["top"]
    anchor_line = [w for w in normalized if abs(w["top"] - anchor_top) <= line_tol]
    if not anchor_line:
        if debug:
            print("No same-line words found for anchor_top:", anchor_top)
        return None
    if len(anchor_line) < 2:
        if debug:
            print("Not enough words in the same line:", anchor_line)
        return None
    # 製作一個 psuedo word 代表第二欄
    psuedo_word = anchor_line[0].copy()
    psuedo_word["text"] = "SECOND_COLUMN"
    psuedo_word["x0"] = (anchor_line[0]["x0"] + anchor_line[1]["x0"]) / 2 - 10
    psuedo_word["x1"] = (anchor_line[0]["x1"] + anchor_line[1]["x1"]) / 2
    anchor_line.insert(1, psuedo_word)
    # 排序（由左到右）
    anchor_line.sort(key=lambda w: w["x0"])

    # 3) 依水平 gap 分群成 clusters（每個 cluster 代表同一欄）
    # 寬度 130，高度 25
    clusters = []
    for anchor in anchor_line:
        current = [anchor]
        current.extend(
            [
                w
                for w in normalized
                if 0 <= w["top"] - anchor["top"] <= CLUSTER_HEIGHT_TOL
                and 0 <= w["x1"] - anchor["x0"] <= CLUSTER_WIDTH_TOL
            ]
        )
        clusters.append(current)

    if debug:
        print("clusters count:", len(clusters))
        for i, c in enumerate(clusters):
            print(
                f" cluster {i}: "
                + " | ".join([f"{t['text']}@{int(t['x0'])}" for t in c])
            )

    # 4) 定位 Ship-To cluster 和 line
    ship_cluster = clusters[1]
    target_word = [w for w in ship_cluster if w["text"] == "S"][0]
    target_line_words = sorted(
        [w for w in ship_cluster if abs(w["top"] - target_word["top"]) <= line_tol],
        key=lambda w: (w["top"], w["x0"]),
    )
    target_line = " ".join(w["text"] for w in target_line_words).strip()

    # 5) 試著去掉行首的 'S' 或 'S,' 之類
    # 例如 "S THD DI DFC #6707 - LUCKEY" -> "THD DI DFC #6707 - LUCKEY"
    target_line = re.sub(r"^\s*S[\s,:-]*", "", target_line, flags=re.IGNORECASE).strip()

    if debug:
        print("Extracted target_line:", target_line)

    return target_line
