from pydantic import BaseModel, Field
from typing import List, Optional

class ServiceContractPeriod(BaseModel):
    startDate: Optional[str] = None
    endDate: Optional[str] = None

class Checklist(BaseModel):
    submittedIdCard: bool = False
    # Một cộng tác viên có thể có nhiều hợp đồng dịch vụ theo thời gian (gia hạn, ký lại...).
    # Luôn phải giữ tối thiểu 1 phần tử (có thể rỗng ngày) để còn chỗ nhập liệu.
    serviceContracts: List[ServiceContractPeriod] = Field(
        default_factory=lambda: [ServiceContractPeriod()], min_length=1
    )
    submittedTaxCommitment: bool = False
    liquidationDate: Optional[str] = None
    submittedCV: bool = False
    submittedResidenceInfo: bool = False
    submittedDegree: bool = False

class CollaboratorBase(BaseModel):
    employeeCode: str
    fullName: str = ""
    taxCode: str = ""
    dob: Optional[str] = None
    idNumber: str = ""
    email: str = ""
    phone: str = ""
    address: str = ""
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
                "checklist": {
                    "submittedIdCard": True,
                    "serviceContracts": [{"startDate": None, "endDate": None}],
                    "submittedTaxCommitment": False,
                    "liquidationDate": None,
                    "submittedCV": True,
                    "submittedResidenceInfo": True,
                    "submittedDegree": True
                },
                "createdAt": "2024-12-15T02:00:00.000Z",
                "updatedAt": "2024-12-15T02:00:00.000Z"
            }
        }
    }
