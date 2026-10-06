from pydantic import BaseModel, Field
from typing import List, Optional

class TpBankContractItem(BaseModel):
    contractNumber: str = ""
    effectiveDate: Optional[str] = None
    expiryDate: Optional[str] = None
    status: Optional[str] = None
    year: Optional[int] = None
    period: Optional[str] = None
    hrbp: Optional[str] = None
    note: Optional[str] = None

class TpBankInfo(BaseModel):
    estimatedContractCount: Optional[int] = None
    hrContractCount: Optional[int] = None
    contracts: List[TpBankContractItem] = Field(default_factory=list)

class BmkHrInfo(BaseModel):
    contractCount: int = 0
    idCardCount: int = 0
    liquidationCount: int = 0
    taxCommitmentCount: int = 0

class BmkSystemInfo(BaseModel):
    contractCount: int = 0
    idCardCount: int = 0
    liquidationCount: int = 0
    taxCommitmentCount: int = 0

class ReconciliationResultDetail(BaseModel):
    contract: Optional[str] = None       # "success" | "failed" | "warn" | None
    idCard: Optional[str] = None         # "success" | "failed" | None
    liquidation: Optional[str] = None    # "success" | "failed" | None
    taxCommitment: Optional[str] = None # "success" | "failed" | None

class ReconciliationRecordBase(BaseModel):
    employeeCode: str
    fullName: str = ""
    idNumber: Optional[str] = None  # Số CCCD
    createdSource: Optional[str] = "bmk_system"  # Nguồn tạo: bmk_system, bmk_hr, tpbank
    isBmkSystemExist: bool = False  # Đánh dấu mã NV có tồn tại ở bmk_ctv_collaborators hay không
    isSynced: bool = False  # Flag sync đánh dấu nhân viên đã thực hiện đối soát hay chưa
    departmentLevel1: Optional[str] = None
    position: Optional[str] = None
    employmentStatus: Optional[str] = None
    onboardDate: Optional[str] = None
    offboardDate: Optional[str] = None
    contractExpiryDate: Optional[str] = None  # Ngày hợp đồng đến hạn
    tpbankInfo: TpBankInfo = Field(default_factory=TpBankInfo)
    bmkHrInfo: BmkHrInfo = Field(default_factory=BmkHrInfo)
    bmkSystemInfo: BmkSystemInfo = Field(default_factory=BmkSystemInfo)
    result: Optional[ReconciliationResultDetail] = None
    reconciliationStatus: Optional[str] = "pending"


class ReconciliationRecordCreate(ReconciliationRecordBase):
    pass

class ReconciliationRecordUpdate(BaseModel):
    fullName: Optional[str] = None
    idNumber: Optional[str] = None
    createdSource: Optional[str] = None
    isSynced: Optional[bool] = None
    departmentLevel1: Optional[str] = None
    position: Optional[str] = None
    employmentStatus: Optional[str] = None
    onboardDate: Optional[str] = None
    offboardDate: Optional[str] = None
    contractExpiryDate: Optional[str] = None
    tpbankInfo: Optional[TpBankInfo] = None
    bmkHrInfo: Optional[BmkHrInfo] = None
    bmkSystemInfo: Optional[BmkSystemInfo] = None
    result: Optional[ReconciliationResultDetail] = None
    reconciliationStatus: Optional[str] = None

class ReconciliationRecordResponse(ReconciliationRecordBase):
    id: str
    createdAt: str
    updatedAt: str

    model_config = {
        "json_schema_extra": {
            "example": {
                "id": "65a1f0c2e4b0f0a1b2c3d4e5",
                "employeeCode": "45114",
                "fullName": "Nguyễn Văn A",
                "idNumber": "001204014772",
                "createdSource": "bmk_system",
                "departmentLevel1": "Khối Khách hàng Cá nhân",
                "position": "Cộng tác viên bán hàng",
                "employmentStatus": "Hiện diện",
                "onboardDate": "2024-01-15",
                "offboardDate": None,
                "contractExpiryDate": "2024-12-31",
                "tpbankInfo": {
                    "contracts": [
                        {
                            "contractNumber": "HD-2024-001",
                            "effectiveDate": "2024-01-15",
                            "expiryDate": "2024-12-31",
                            "status": "Hiệu lực",
                            "year": 2024,
                            "period": "Kỳ 1",
                            "hrbp": "HRBP.FOS",
                            "note": "Hợp đồng năm 2024"
                        }
                    ]
                },
                "bmkHrInfo": {
                    "contractCount": 1,
                    "idCardCount": 1,
                    "liquidationCount": 0,
                    "taxCommitmentCount": 1
                },
                "bmkSystemInfo": {
                    "contractCount": 1,
                    "idCardCount": 1,
                    "liquidationCount": 0,
                    "taxCommitmentCount": 1
                },
                "reconciliationStatus": "matched",
                "createdAt": "2026-09-21T00:00:00Z",
                "updatedAt": "2026-09-21T00:00:00Z"
            }
        }
    }

class ReconciliationListResponse(BaseModel):
    items: List[ReconciliationRecordResponse]
    total: int
    page: int
    pageSize: int
    totalPages: int

class ImportHrTpBankResult(BaseModel):
    status: str
    message: str
    totalProcessed: int
    createdCount: int
    updatedCount: int


class ReconciliationHistoryStats(BaseModel):
    totalSuccess: int = 0  # Cả 3 SL HĐ, BBTL, CCCD đều khớp
    totalWarnBank: int = 0  # SL HĐ lệch với bank
    totalMismatchContract: int = 0  # Lệch SL HĐ
    totalMismatchIdCard: int = 0  # Lệch CCCD
    totalMismatchLiquidation: int = 0  # Lệch BBTL


class ReconciliationResultFileInfo(BaseModel):
    filename: str
    s3Key: str
    s3Bucket: str
    fileSize: Optional[int] = 0


class ReconciliationHistoryResponse(BaseModel):
    id: str
    filename: str
    s3Key: str
    s3Bucket: str
    fileSize: Optional[int] = 0
    uploadedBy: str
    username: str
    totalRows: int = 0
    successRows: int = 0
    failedRows: int = 0
    status: str = "success"  # "success" | "failed"
    message: str = ""
    stats: ReconciliationHistoryStats = Field(default_factory=ReconciliationHistoryStats)
    resultFile: Optional[ReconciliationResultFileInfo] = None
    createdAt: str


class ReconciliationHistoryListResponse(BaseModel):
    items: List[ReconciliationHistoryResponse]
    total: int
    page: int
    pageSize: int
    totalPages: int


class ExpiringMovementItem(BaseModel):
    date: str
    label: str
    count: int

class ExpiringBarChartData(BaseModel):
    days: int
    totalExpiring: int
    items: List[ExpiringMovementItem]

class ExpiringPieChartData(BaseModel):
    active: int
    resigned: int
    total: int

class ExpiringContractsStatsResponse(BaseModel):
    pieChart: ExpiringPieChartData
    barChart: ExpiringBarChartData

