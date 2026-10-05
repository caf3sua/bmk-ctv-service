from pydantic import BaseModel, Field
from typing import List, Optional, Dict

class ServiceContractPeriod(BaseModel):
    startDate: Optional[str] = None
    endDate: Optional[str] = None

class ChecklistFileItem(BaseModel):
    name: str = ""
    updatedDate: Optional[str] = None

class CccdChecklist(BaseModel):
    checked: bool = False
    file: Optional[ChecklistFileItem] = None

class CktChecklist(BaseModel):
    checked: bool = False
    file: Optional[ChecklistFileItem] = None

class HddvChecklist(BaseModel):
    # Một cộng tác viên có thể có nhiều hợp đồng dịch vụ theo thời gian (gia hạn, ký lại...).
    # Luôn phải giữ tối thiểu 1 phần tử (có thể rỗng ngày) để còn chỗ nhập liệu.
    contract_date: List[ServiceContractPeriod] = Field(
        default_factory=lambda: [ServiceContractPeriod()], min_length=1
    )
    files: List[ChecklistFileItem] = Field(default_factory=list)

class BbtlChecklist(BaseModel):
    date: Optional[str] = None
    file: Optional[ChecklistFileItem] = None

class Checklist(BaseModel):
    cccd: CccdChecklist = Field(default_factory=CccdChecklist)
    ckt: CktChecklist = Field(default_factory=CktChecklist)
    hddv: HddvChecklist = Field(default_factory=HddvChecklist)
    bbtl: BbtlChecklist = Field(default_factory=BbtlChecklist)

class HandoverInfoItem(BaseModel):
    handoverDate: str
    handoverPerson: Optional[str] = ""
    createdAt: Optional[str] = None

class CollaboratorBase(BaseModel):
    employeeCode: str
    fullName: str = ""
    taxCode: str = ""
    dob: Optional[str] = None
    idNumber: str = ""
    email: str = ""
    phone: str = ""
    address: str = ""
    noted: Optional[str] = ""
    handoverInfo: List[HandoverInfoItem] = Field(default_factory=list)
    checklist: Checklist = Field(default_factory=Checklist)

class CollaboratorCreate(CollaboratorBase):
    pass

class CollaboratorUpdate(CollaboratorBase):
    pass

class CollaboratorResponse(CollaboratorBase):
    createdAt: str
    updatedAt: str

    model_config = {
        "json_schema_extra": {
            "example": {
                "employeeCode": "CTV001",
                "fullName": "Đỗ Xuân Yến",
                "taxCode": "8786335019",
                "dob": "1991-03-10",
                "idNumber": "079090000001",
                "email": "do.xuan.yen1@example.com",
                "phone": "0910000137",
                "address": "188 Trần Phú, TP. Nha Trang, Khánh Hòa",
                "noted": "Thông tin lưu ý",
                "handoverInfo": [
                    {
                        "handoverDate": "2024-12-15",
                        "handoverPerson": "Nguyễn Văn B",
                        "createdAt": "2024-12-15T02:00:00.000Z"
                    }
                ],
                "checklist": {
                    "cccd": {"checked": True, "file": None},
                    "ckt": {"checked": False, "file": None},
                    "hddv": {
                        "contract_date": [{"startDate": None, "endDate": None}],
                        "files": []
                    },
                    "bbtl": {"date": None, "file": None}
                },
                "createdAt": "2024-12-15T02:00:00.000Z",
                "updatedAt": "2024-12-15T02:00:00.000Z"
            }
        }
    }
