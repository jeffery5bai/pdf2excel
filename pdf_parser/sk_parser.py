from pdf_parser.wholesale_parser import WholesalePOParser


class SKPOParser(WholesalePOParser):
    def __init__(self):
        super().__init__()
        self._gt_crd_days = None

    def set_gt_crd_days(self, text: str):
        if "SPLASH" in text:
            self._gt_crd_days = 45
        else:
            self._gt_crd_days = 60

    @property
    def gt_crd_days(self) -> int:
        if self._gt_crd_days is None:
            raise ValueError(
                "gt_crd_days is not set. Please call set_gt_crd_days() first."
            )
        return self._gt_crd_days
