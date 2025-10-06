from typing import List, Dict

class POParser:
    def __init__(self):
        pass

    @property
    def gt_crd_days(self) -> int:
        return 70

    def parse_po_content(self, text: str, file_type: str = "original") -> List[Dict]:
        raise NotImplementedError("Subclasses should implement this method")