class Package:
    def __init__(
        self, excelReciverString: str, excel_row: int, excel_column: int
    ) -> None:
        self.excelReciverString: str = excelReciverString
        self.excel_row: int = excel_row
        self.excel_column: int = excel_column
        self.recipientName: str = None
        self.recipientNameAddtional: str = None
        self.address1: str = None
        self.address2: str = None
        self.address3: str = None
        self.country: str = None
        self.postalCode: str = None
        self.city: str = None
        self.state: str = None
        self.phoneNumber: str = None
        self.email: str = None
        self.weight: float = None
        self.service: str = None
        self.referenceNumbers: list[str] = []
        self.packageCount: str = None

    def __str__(self) -> str:
        return f"Empfaenger: '{self.recipientName}'"
