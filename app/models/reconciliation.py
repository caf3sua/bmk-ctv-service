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

class ReconciliationRecordBase(BaseModel):
    employeeCode: str
    fullName: str = ""
    idNumber: Optional[str] = None  # Số CCCD
    createdSource: Optional[str] = "bmk_system"  # Nguồn tạo: bmk_system, bmk_hr, tpbank
    departmentLevel1: Optional[str] = None
    position: Optional[str] = None
    employmentStatus: Optional[str] = None
    onboardDate: Optional[str] = None
    offboardDate: Optional[str] = None
    tpbankInfo: TpBankInfo = Field(default_factory=TpBankInfo)
    bmkHrInfo: BmkHrInfo = Field(default_factory=BmkHrInfo)
    bmkSystemInfo: BmkSystemInfo = Field(default_factory=BmkSystemInfo)
    reconciliationStatus: Optional[str] = "pending"

class ReconciliationRecordCreate(ReconciliationRecordBase):
    pass

class ReconciliationRecordUpdate(BaseModel):
    fullName: Optional[str] = None
    idNumber: Optional[str] = None
    createdSource: Optional[str] = None
    departmentLevel1: Optional[str] = None
    position: Optional[str] = None
    employmentStatus: Optional[str] = None
    onboardDate: Optional[str] = None
    offboardDate: Optional[str] = None
    tpbankInfo: Optional[TpBankInfo] = None
    bmkHrInfo: Optional[BmkHrInfo] = None
    bmkSystemInfo: Optional[BmkSystemInfo] = None
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
