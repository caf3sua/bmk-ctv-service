from pydantic import BaseModel, Field
from typing import List, Optional

class BankAccountInfo(BaseModel):
    accountName: Optional[str] = ""
    accountNumber: Optional[str] = ""
    bankName: Optional[str] = ""

class ExpenseDetailItem(BaseModel):
    id: Optional[str] = None
    employeeCode: str
    fullName: str = ""
    taxCode: Optional[str] = ""
    bankAccount: BankAccountInfo = Field(default_factory=BankAccountInfo)
    taxableIncome: float = 0.0
    personalIncomeTax: float = 0.0
    netIncome: float = 0.0
    workLocation: Optional[str] = ""
    paymentDate: str
    paymentContent: Optional[str] = ""
    onboardDate: Optional[str] = None
    offboardDate: Optional[str] = None
    period: str
    createdAt: Optional[str] = None
    updatedAt: Optional[str] = None
    createdBy: Optional[str] = None
    updatedBy: Optional[str] = None

class CollaboratorExpenseSummary(BaseModel):
    id: Optional[str] = None
    employeeCode: str
    fullName: str = ""
    taxCode: Optional[str] = ""
    totalPaymentCount: int = 0
    latestPaymentDate: Optional[str] = None
    latestPeriod: Optional[str] = None
    createdAt: Optional[str] = None
    updatedAt: Optional[str] = None
    createdBy: Optional[str] = None
    updatedBy: Optional[str] = None

class MonthlyCollaboratorCountStat(BaseModel):
    period: str
    label: str
    count: int

class MonthlyPaymentAmountStat(BaseModel):
    period: str
    label: str
    totalNetIncome: float = 0.0
    totalTaxableIncome: float = 0.0
    totalPersonalIncomeTax: float = 0.0

class ExpenseStatsResponse(BaseModel):
    totalPaidAmount: float = 0.0
    totalTaxableAmount: float = 0.0
    totalTaxAmount: float = 0.0
    totalPaymentTransactions: int = 0
    totalCollaborators: int = 0
    latestPeriod: Optional[str] = None
    collaboratorCountsByMonth: List[MonthlyCollaboratorCountStat] = Field(default_factory=list)
    paymentAmountsByMonth: List[MonthlyPaymentAmountStat] = Field(default_factory=list)

class ExpenseImportResponse(BaseModel):
    success: bool
    totalRows: int
    importedDetails: int
    updatedDetails: int
    newReconciliationCreated: int
    collaboratorsCount: int
    periodsAffected: List[str] = Field(default_factory=list)
    message: str
