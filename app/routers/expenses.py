import re
from datetime import datetime, date, timezone
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import List, Optional

from dateutil.relativedelta import relativedelta
from fastapi import APIRouter, Depends, Form, HTTPException, Query, UploadFile, File, status
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pymongo import InsertOne, UpdateOne

from app.core.database import get_db
from app.core.logging import get_logger
from app.core.security import get_current_user
from app.models.expense import (
    ExpenseStatsResponse,
    ExpenseImportResponse,
    MonthlyCollaboratorCountStat,
    MonthlyPaymentAmountStat,
)

logger = get_logger(__name__)

router = APIRouter(prefix="/api/expenses", tags=["Collaborator Expenses"])

COLLABORATOR_EXPENSES_COLLECTION = "bmk_ctv_collaborator_expenses"
EXPENSE_DETAILS_COLLECTION = "bmk_ctv_expense_details"
RECONCILIATION_COLLECTION = "bmk_ctv_reconciliations"

TEMPLATE_PATH = Path(__file__).resolve().parent.parent.parent / "templates" / "import_thanh_toan_phi_CTV.xlsx"
EXPORT_TEMPLATE_PATH = Path(__file__).resolve().parent.parent.parent / "templates" / "export_thanh_toan_phi_CTV.xlsx"
MONTHLY_EXPORT_TEMPLATE_PATH = Path(__file__).resolve().parent.parent.parent / "templates" / "export_bao_cao_chi_phi_theo_thang.xlsx"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _actor_name(current_user: dict) -> str:
    return current_user.get("name") or current_user.get("fullName") or current_user.get("username", "System")


def parse_employee_code(val) -> Optional[str]:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ("-", "none", "null"):
        return None
    if s.endswith(".0"):
        s = s[:-2]
    if s.isdigit() and len(s) < 5:
        s = s.zfill(5)
    return s


def parse_str_val(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("none", "null", "-"):
        return ""
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def parse_float_val(val) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "").replace(" ", "").replace("₫", "").replace("VND", "").replace("vnd", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def parse_date_str(val) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() in ("-", "none", "null"):
        return None
    if " " in s:
        s = s.split(" ")[0].strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def calculate_period(payment_date_str: str) -> str:
    """Calculate period 'YYYYMM' from payment date 'YYYY-MM-DD'."""
    try:
        parts = payment_date_str.split("-")
        if len(parts) >= 2:
            return f"{parts[0]}{parts[1]}"
    except Exception:
        pass
    return "UNKNOWN"


def format_period_label(period_str: str) -> str:
    """Format 'YYYYMM' to 'MM/YYYY'."""
    if len(period_str) == 6 and period_str.isdigit():
        return f"{period_str[4:]}/{period_str[:4]}"
    return period_str


def format_display_date(date_str: Optional[str]) -> str:
    """Format 'YYYY-MM-DD' to 'DD/MM/YYYY'."""
    if not date_str:
        return ""
    s = str(date_str).strip()
    if not s or s.lower() in ("none", "null", "-"):
        return ""
    if " " in s:
        s = s.split(" ")[0].strip()
    try:
        parts = s.split("-")
        if len(parts) == 3 and len(parts[0]) == 4:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
    except Exception:
        pass
    return s


def format_row_ranges(row_numbers: List[int]) -> str:
    """Gom các số dòng liên tiếp thành dạng range, ví dụ: '10-12, 15, 517-1235'."""
    if not row_numbers:
        return ""
    sorted_rows = sorted(set(row_numbers))
    ranges = []
    start = sorted_rows[0]
    end = start
    for r in sorted_rows[1:]:
        if r == end + 1:
            end = r
        else:
            ranges.append(f"{start}-{end}" if start != end else f"{start}")
            start = end = r
    ranges.append(f"{start}-{end}" if start != end else f"{start}")
    if len(ranges) <= 5:
        return ", ".join(ranges)
    return ", ".join(ranges[:5]) + f", ... (tổng {len(sorted_rows)} dòng)"


def add_six_months(date_str: Optional[str]) -> Optional[str]:
    if not date_str:
        return None
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d").date()
        exp_dt = dt + relativedelta(months=6)
        return exp_dt.strftime("%Y-%m-%d")
    except Exception:
        return None


@router.get("/template")
async def download_template(current_user: dict = Depends(get_current_user)):
    """Tải file mẫu Excel import thanh toán phí CTV."""
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File mẫu import_thanh_toan_phi_CTV.xlsx không tồn tại trên hệ thống",
        )
    return FileResponse(
        path=str(TEMPLATE_PATH),
        filename="import_thanh_toan_phi_CTV.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.get("/periods")
async def list_periods(
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy danh sách các kỳ thanh toán duy nhất (sorted mới nhất -> cũ nhất)."""
    periods = await db[EXPENSE_DETAILS_COLLECTION].distinct("period")
    periods = [p for p in periods if p and p != "UNKNOWN"]
    periods.sort(reverse=True)
    return {"periods": periods}


@router.get("/locations")
async def list_locations(
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy danh sách các nơi làm việc duy nhất."""
    locations = await db[EXPENSE_DETAILS_COLLECTION].distinct("workLocation")
    locations = [loc for loc in locations if loc and loc.strip()]
    locations.sort()
    return {"locations": locations}


@router.get("/stats", response_model=ExpenseStatsResponse)
async def get_expense_stats(
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy thống kê tổng quan và dữ liệu cho 2 biểu đồ:
    1. Số lượng CTV được thanh toán theo từng tháng
    2. Số tiền BMK thanh toán theo từng tháng
    """
    total_collabs = await db[COLLABORATOR_EXPENSES_COLLECTION].count_documents({})
    total_tx = await db[EXPENSE_DETAILS_COLLECTION].count_documents({})

    pipeline_summary = [
        {
            "$group": {
                "_id": None,
                "totalPaidAmount": {"$sum": "$netIncome"},
                "totalTaxableAmount": {"$sum": "$taxableIncome"},
                "totalTaxAmount": {"$sum": "$personalIncomeTax"},
            }
        }
    ]
    summary_res = await db[EXPENSE_DETAILS_COLLECTION].aggregate(pipeline_summary).to_list(1)
    total_paid = summary_res[0]["totalPaidAmount"] if summary_res else 0.0
    total_taxable = summary_res[0]["totalTaxableAmount"] if summary_res else 0.0
    total_tax = summary_res[0]["totalTaxAmount"] if summary_res else 0.0

    # Biểu đồ 1: Số CTV duy nhất được thanh toán theo từng tháng
    pipeline_monthly_collabs = [
        {
            "$group": {
                "_id": {"period": "$period", "employeeCode": "$employeeCode"}
            }
        },
        {
            "$group": {
                "_id": "$_id.period",
                "count": {"$sum": 1}
            }
        },
        {"$sort": {"_id": 1}}
    ]
    monthly_collabs_res = await db[EXPENSE_DETAILS_COLLECTION].aggregate(pipeline_monthly_collabs).to_list(100)
    collab_counts_by_month: List[MonthlyCollaboratorCountStat] = []
    for item in monthly_collabs_res:
        p = item["_id"]
        if p and p != "UNKNOWN":
            collab_counts_by_month.append(
                MonthlyCollaboratorCountStat(
                    period=p,
                    label=format_period_label(p),
                    count=item["count"],
                )
            )

    # Biểu đồ 2: Số tiền thanh toán theo từng tháng (Net income, Taxable, Tax)
    pipeline_monthly_amount = [
        {
            "$group": {
                "_id": "$period",
                "totalNetIncome": {"$sum": "$netIncome"},
                "totalTaxableIncome": {"$sum": "$taxableIncome"},
                "totalPersonalIncomeTax": {"$sum": "$personalIncomeTax"},
            }
        },
        {"$sort": {"_id": 1}}
    ]
    monthly_amount_res = await db[EXPENSE_DETAILS_COLLECTION].aggregate(pipeline_monthly_amount).to_list(100)
    payment_amounts_by_month: List[MonthlyPaymentAmountStat] = []
    for item in monthly_amount_res:
        p = item["_id"]
        if p and p != "UNKNOWN":
            payment_amounts_by_month.append(
                MonthlyPaymentAmountStat(
                    period=p,
                    label=format_period_label(p),
                    totalNetIncome=item.get("totalNetIncome", 0.0),
                    totalTaxableIncome=item.get("totalTaxableIncome", 0.0),
                    totalPersonalIncomeTax=item.get("totalPersonalIncomeTax", 0.0),
                )
            )

    latest_period = None
    if collab_counts_by_month:
        latest_period = collab_counts_by_month[-1].period

    return ExpenseStatsResponse(
        totalPaidAmount=round(total_paid, 2),
        totalTaxableAmount=round(total_taxable, 2),
        totalTaxAmount=round(total_tax, 2),
        totalPaymentTransactions=total_tx,
        totalCollaborators=total_collabs,
        latestPeriod=latest_period,
        collaboratorCountsByMonth=collab_counts_by_month,
        paymentAmountsByMonth=payment_amounts_by_month,
    )


@router.get("/details")
async def list_expense_details(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    search: Optional[str] = Query(None),
    period: Optional[str] = Query(None),
    work_location: Optional[str] = Query(None),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy danh sách chi tiết chi phí thanh toán cho CTV."""
    query = {}
    if period and period.strip():
        query["period"] = period.strip()
    if work_location and work_location.strip():
        query["workLocation"] = work_location.strip()
    if search and search.strip():
        s = search.strip()
        reg = {"$regex": re.escape(s), "$options": "i"}
        query["$or"] = [
            {"employeeCode": reg},
            {"fullName": reg},
            {"taxCode": reg},
            {"bankAccount.accountNumber": reg},
            {"workLocation": reg},
        ]

    total = await db[EXPENSE_DETAILS_COLLECTION].count_documents(query)
    skip = (page - 1) * page_size
    cursor = (
        db[EXPENSE_DETAILS_COLLECTION]
        .find(query)
        .sort([("paymentDate", -1), ("employeeCode", 1)])
        .skip(skip)
        .limit(page_size)
    )

    items = []
    async for doc in cursor:
        doc["id"] = str(doc.get("_id"))
        doc.pop("_id", None)
        items.append(doc)

    total_pages = max(1, (total + page_size - 1) // page_size)
    return {
        "items": items,
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
    }


@router.get("/collaborators")
async def list_expense_collaborators(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    search: Optional[str] = Query(None),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy danh sách tổng hợp CTV chi phí (collection chính)."""
    query = {}
    if search and search.strip():
        s = search.strip()
        reg = {"$regex": re.escape(s), "$options": "i"}
        query["$or"] = [
            {"employeeCode": reg},
            {"fullName": reg},
            {"taxCode": reg},
        ]

    total = await db[COLLABORATOR_EXPENSES_COLLECTION].count_documents(query)
    skip = (page - 1) * page_size
    cursor = (
        db[COLLABORATOR_EXPENSES_COLLECTION]
        .find(query)
        .sort([("totalPaymentCount", -1), ("employeeCode", 1)])
        .skip(skip)
        .limit(page_size)
    )

    items = []
    async for doc in cursor:
        doc["id"] = str(doc.get("_id"))
        doc.pop("_id", None)
        items.append(doc)

    total_pages = max(1, (total + page_size - 1) // page_size)
    return {
        "items": items,
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
    }


@router.post("/import", response_model=ExpenseImportResponse)
async def import_expenses(
    file: UploadFile = File(...),
    period: str = Form(..., description="Kỳ chi phí (YYYY-MM hoặc YYYYMM) do người dùng chọn"),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Import dữ liệu thanh toán phí CTV từ file Excel:
    1. Kiểm tra định dạng kỳ và file Excel.
    2. Trước khi xử lý, xóa toàn bộ dữ liệu chi tiết của kỳ đã chọn trong bmk_ctv_expense_details.
    3. Thêm mới bản ghi vào bmk_ctv_expense_details với period = kỳ người dùng đã chọn ở dialog.
    4. Tạo mới record vào bmk_ctv_reconciliations nếu CTV chưa có trong danh sách đối soát.
    5. Cập nhật thông tin và tính lại chính xác totalPaymentCount trong bmk_ctv_collaborator_expenses.
    """
    if not period or not isinstance(period, str):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Vui lòng chọn kỳ chi phí (tháng/năm) cần import",
        )
    target_period = period.strip().replace("-", "")
    if len(target_period) != 6 or not target_period.isdigit():
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Kỳ chi phí không hợp lệ: '{period}'. Định dạng hợp lệ: YYYY-MM hoặc YYYYMM (ví dụ: 2026-10 hoặc 202610).",
        )

    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Chỉ hỗ trợ file Excel (.xlsx, .xlsm)",
        )

    content = await file.read()
    try:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Không đọc được file Excel: {str(e)}",
        )

    sheet_name = "Gui BMK" if "Gui BMK" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File Excel không chứa dữ liệu",
        )

    # Tìm dòng tiêu đề chính
    header_row_idx = None
    header = []
    for idx, r in enumerate(rows):
        if r and any(
            cell and any(k in str(cell).upper() for k in ["MNV", "MÃ NV", "MÃ NHÂN VIÊN", "MÃ CTV", "EMPLOYEE CODE"])
            for cell in r
        ):
            header_row_idx = idx
            header = [str(c).strip() if c is not None else "" for c in r]
            break

    if header_row_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không tìm thấy dòng tiêu đề chứa cột 'MNV' trong file Excel",
        )

    header_upper = [h.upper() for h in header]

    def find_col(candidates: List[str]) -> Optional[int]:
        for candidate in candidates:
            cand_u = candidate.upper()
            for idx, h in enumerate(header_upper):
                if cand_u in h:
                    return idx
        return None

    code_idx = find_col(["MNV", "MÃ NV", "MÃ NHÂN VIÊN", "MÃ CTV", "MA NV"])
    name_idx = find_col(["HỌ VÀ TÊN", "HỌ TÊN", "HO VA TEN", "FULL NAME", "NAME"])
    tax_idx = find_col(["MÃ SỐ THUẾ", "MA SO THUE", "MST", "TAX CODE"])
    acc_name_idx = find_col(["TÊN TÀI KHOẢN", "TEN TAI KHOAN", "TÊN TK", "ACCOUNT NAME"])
    acc_num_idx = find_col(["SỐ TÀI KHOẢN", "SO TAI KHOAN", "SỐ TK", "ACCOUNT NUMBER"])
    bank_idx = find_col(["NGÂN HÀNG", "NGAN HANG", "BANK"])
    taxable_idx = find_col(["TỔNG TN", "TONG TN", "THU NHẬP CHỊU THUẾ", "TN CHỊU THUẾ"])
    tax_tncn_idx = find_col(["THUẾ TNCN", "THUE TNCN", "TNCN"])
    net_idx = find_col(["THỰC LĨNH", "THUC LINH", "NET"])
    work_loc_idx = find_col(["NƠI LÀM VIỆC", "NOI LAM VIEC", "LOCATION"])
    pay_date_idx = find_col(["NGÀY BMK TT", "NGAY BMK TT", "NGÀY TT", "NGAY TT", "NGÀY THANH TOÁN"])
    pay_content_idx = find_col(["NỘI DUNG TT", "NOI DUNG TT", "NỘI DUNG", "NOI DUNG"])
    onboard_idx = find_col(["NGÀY GIA NHẬP", "NGAY GIA NHAP", "GIA NHẬP"])
    offboard_idx = find_col(["NGÀY CHẤM DỨT", "NGAY CHAM DUT", "CHẤM DỨT"])

    if code_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File thiếu cột 'MNV' (Mã cộng tác viên)",
        )

    actor = _actor_name(current_user)
    now_iso = _now()

    # 3. Trước khi xử lý, xóa hết dữ liệu của collection theo period = tháng user chọn (để thực hiện lại import từ đầu)
    old_period_collab_codes = set(
        await db[EXPENSE_DETAILS_COLLECTION].distinct("employeeCode", {"period": target_period})
    )
    delete_res = await db[EXPENSE_DETAILS_COLLECTION].delete_many({"period": target_period})
    deleted_details_count = delete_res.deleted_count
    logger.info(
        f"Đã xóa {deleted_details_count} bản ghi cũ thuộc kỳ {target_period} trong {EXPENSE_DETAILS_COLLECTION} trước khi import."
    )

    # 1. Thu thập mã CTV sẵn có trong bmk_ctv_reconciliations
    existing_reconcil_codes = set()
    async for doc in db[RECONCILIATION_COLLECTION].find({}, {"employeeCode": 1}):
        c = doc.get("employeeCode")
        if c:
            existing_reconcil_codes.add(str(c).strip())

    # 2. Thu thập thông tin CTV sẵn có trong bmk_ctv_collaborator_expenses (collection chính)
    existing_collabs = {}
    async for doc in db[COLLABORATOR_EXPENSES_COLLECTION].find(
        {},
        {"employeeCode": 1, "fullName": 1, "taxCode": 1, "onboardDate": 1, "offboardDate": 1}
    ):
        c = doc.get("employeeCode")
        if c:
            code_str = str(c).strip()
            existing_collabs[code_str] = {
                "fullName": doc.get("fullName", ""),
                "taxCode": doc.get("taxCode", ""),
                "onboardDate": doc.get("onboardDate"),
                "offboardDate": doc.get("offboardDate"),
                "is_new": False,
                "update_onboard": False,
                "update_offboard": False,
                "latestPaymentDate": None,
                "latestPeriod": None,
            }

    warnings_list: List[str] = []
    warnings_seen: set = set()
    missing_date_rows: List[int] = []
    missing_date_net_total: float = 0.0
    collab_tx_counter: dict = {}

    reconcil_ops = []
    detail_ops = []

    affected_collab_codes = set()
    periods_affected = {target_period}
    new_reconcil_created = 0
    total_valid_rows = 0

    for r_idx, row in enumerate(rows[header_row_idx + 1 :]):
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        raw_code = row[code_idx] if code_idx < len(row) else None
        employee_code = parse_employee_code(raw_code)
        if not employee_code:
            continue

        excel_row_num = header_row_idx + 1 + r_idx + 1  # Số dòng thực tế (1-indexed) trong file Excel

        full_name = parse_str_val(row[name_idx]) if (name_idx is not None and name_idx < len(row)) else ""
        tax_code = parse_str_val(row[tax_idx]) if (tax_idx is not None and tax_idx < len(row)) else ""
        acc_name = parse_str_val(row[acc_name_idx]) if (acc_name_idx is not None and acc_name_idx < len(row)) else ""
        acc_num = parse_str_val(row[acc_num_idx]) if (acc_num_idx is not None and acc_num_idx < len(row)) else ""
        bank_name = parse_str_val(row[bank_idx]) if (bank_idx is not None and bank_idx < len(row)) else ""
        taxable_income = parse_float_val(row[taxable_idx]) if (taxable_idx is not None and taxable_idx < len(row)) else 0.0
        tax_tncn = parse_float_val(row[tax_tncn_idx]) if (tax_tncn_idx is not None and tax_tncn_idx < len(row)) else 0.0
        net_income = parse_float_val(row[net_idx]) if (net_idx is not None and net_idx < len(row)) else 0.0
        work_location = parse_str_val(row[work_loc_idx]) if (work_loc_idx is not None and work_loc_idx < len(row)) else ""
        payment_date = parse_date_str(row[pay_date_idx]) if (pay_date_idx is not None and pay_date_idx < len(row)) else None
        payment_content = parse_str_val(row[pay_content_idx]) if (pay_content_idx is not None and pay_content_idx < len(row)) else ""
        onboard_date = parse_date_str(row[onboard_idx]) if (onboard_idx is not None and onboard_idx < len(row)) else None
        offboard_date = parse_date_str(row[offboard_idx]) if (offboard_idx is not None and offboard_idx < len(row)) else None

        if not payment_date:
            # Ghi nhận dòng thiếu ngày BMK TT để gom cảnh báo chung
            missing_date_rows.append(excel_row_num)
            missing_date_net_total += net_income
            continue

        # 2. Thay vì tính period từ ngày BMK thanh toán, gán bằng giá trị target_period do user chọn ở dialog
        period = target_period
        affected_collab_codes.add(employee_code)
        total_valid_rows += 1

        # 1. Kiểm tra tồn tại trong bmk_ctv_reconciliations
        if employee_code not in existing_reconcil_codes:
            contract_expiry_date = add_six_months(onboard_date)
            new_reconcil_doc = {
                "_id": employee_code,
                "employeeCode": employee_code,
                "fullName": full_name or "",
                "idNumber": tax_code or None,
                "contractOnboardDate": onboard_date,
                "contractExpiryDate": contract_expiry_date,
                "isCalculateExpired": True if contract_expiry_date else None,
                "createdSource": "HTR BMK",
                "isBmkSystemExist": False,
                "isSynced": False,
                "departmentLevel1": None,
                "position": None,
                "employmentStatus": "Hiệu lực",
                "onboardDate": onboard_date,
                "offboardDate": offboard_date,
                "tpbankInfo": {"contracts": []},
                "bmkHrInfo": {
                    "contractCount": 0,
                    "idCardCount": 0,
                    "liquidationCount": 0,
                    "taxCommitmentCount": 0,
                },
                "bmkSystemInfo": {
                    "contractCount": 0,
                    "idCardCount": 0,
                    "liquidationCount": 0,
                    "taxCommitmentCount": 0,
                },
                "result": None,
                "reconciliationStatus": "pending",
                "createdAt": now_iso,
                "updatedAt": now_iso,
            }
            reconcil_ops.append(
                UpdateOne(
                    {"employeeCode": employee_code},
                    {"$setOnInsert": new_reconcil_doc},
                    upsert=True,
                )
            )
            existing_reconcil_codes.add(employee_code)
            new_reconcil_created += 1

        # 2. Chuẩn bị insert vào bmk_ctv_expense_details (đã được xóa sạch kỳ này từ trước)
        tx_seq_key = (employee_code, tax_code, payment_date)
        collab_tx_counter[tx_seq_key] = collab_tx_counter.get(tx_seq_key, 0) + 1
        tx_index = collab_tx_counter[tx_seq_key]

        detail_doc = {
            "employeeCode": employee_code,
            "fullName": full_name,
            "taxCode": tax_code,
            "txIndex": tx_index,
            "bankAccount": {
                "accountName": acc_name,
                "accountNumber": acc_num,
                "bankName": bank_name,
            },
            "taxableIncome": taxable_income,
            "personalIncomeTax": tax_tncn,
            "netIncome": net_income,
            "workLocation": work_location,
            "paymentDate": payment_date,
            "paymentContent": payment_content,
            "period": target_period,
            "createdAt": now_iso,
            "createdBy": actor,
            "updatedAt": now_iso,
            "updatedBy": actor,
        }
        detail_ops.append(InsertOne(detail_doc))

        # 3. Quản lý thông tin CTV ở collection chính (bmk_ctv_collaborator_expenses):
        # - Mỗi CTV chỉ có 1 onboardDate và 1 offboardDate lưu ở collection chính.
        # - Chỉ update từ NONE, nếu đang có giá trị mà file khác giá trị thì cảnh báo.
        if employee_code not in existing_collabs:
            existing_collabs[employee_code] = {
                "fullName": full_name,
                "taxCode": tax_code,
                "onboardDate": onboard_date,
                "offboardDate": offboard_date,
                "is_new": True,
                "update_onboard": False,
                "update_offboard": False,
                "latestPaymentDate": payment_date,
                "latestPeriod": target_period,
            }
        else:
            collab_info = existing_collabs[employee_code]
            if full_name and not collab_info.get("fullName"):
                collab_info["fullName"] = full_name
            if tax_code and not collab_info.get("taxCode"):
                collab_info["taxCode"] = tax_code

            if not collab_info.get("latestPaymentDate") or (payment_date and payment_date > collab_info["latestPaymentDate"]):
                collab_info["latestPaymentDate"] = payment_date
                collab_info["latestPeriod"] = target_period

            # Kiểm tra onboardDate (chỉ update từ NONE, nếu có giá trị khác thì cảnh báo)
            curr_onboard = collab_info.get("onboardDate")
            if onboard_date:
                if not curr_onboard:
                    collab_info["onboardDate"] = onboard_date
                    collab_info["update_onboard"] = True
                elif curr_onboard != onboard_date:
                    warn_key = (employee_code, "onboard", curr_onboard, onboard_date)
                    if warn_key not in warnings_seen:
                        warnings_seen.add(warn_key)
                        warnings_list.append(
                            f"Dòng {excel_row_num}: CTV {employee_code} ({full_name or collab_info.get('fullName', '')}) - "
                            f"Ngày gia nhập trong hệ thống ({format_display_date(curr_onboard)}) "
                            f"khác ngày trong file ({format_display_date(onboard_date)})"
                        )

            # Kiểm tra offboardDate (chỉ update từ NONE, nếu có giá trị khác thì cảnh báo)
            curr_offboard = collab_info.get("offboardDate")
            if offboard_date:
                if not curr_offboard:
                    collab_info["offboardDate"] = offboard_date
                    collab_info["update_offboard"] = True
                elif curr_offboard != offboard_date:
                    warn_key = (employee_code, "offboard", curr_offboard, offboard_date)
                    if warn_key not in warnings_seen:
                        warnings_seen.add(warn_key)
                        warnings_list.append(
                            f"Dòng {excel_row_num}: CTV {employee_code} ({full_name or collab_info.get('fullName', '')}) - "
                            f"Ngày chấm dứt trong hệ thống ({format_display_date(curr_offboard)}) "
                            f"khác ngày trong file ({format_display_date(offboard_date)})"
                        )

    # Thực thi bulk operations vào MongoDB
    if reconcil_ops:
        await db[RECONCILIATION_COLLECTION].bulk_write(reconcil_ops, ordered=False)

    imported_details = 0
    updated_details = 0
    if detail_ops:
        res = await db[EXPENSE_DETAILS_COLLECTION].bulk_write(detail_ops, ordered=False)
        imported_details = res.inserted_count
        updated_details = 0

    # Chuẩn bị bulk operations cho collection chính (bmk_ctv_collaborator_expenses)
    collab_ops = []
    for code in affected_collab_codes:
        info = existing_collabs.get(code)
        if not info:
            continue
        if info.get("is_new"):
            new_collab_doc = {
                "_id": code,
                "employeeCode": code,
                "fullName": info.get("fullName", ""),
                "taxCode": info.get("taxCode", ""),
                "onboardDate": info.get("onboardDate"),
                "offboardDate": info.get("offboardDate"),
                "totalPaymentCount": 0,
                "latestPaymentDate": info.get("latestPaymentDate"),
                "latestPeriod": info.get("latestPeriod"),
                "createdAt": now_iso,
                "createdBy": actor,
                "updatedAt": now_iso,
                "updatedBy": actor,
            }
            collab_ops.append(
                UpdateOne(
                    {"employeeCode": code},
                    {"$setOnInsert": new_collab_doc},
                    upsert=True,
                )
            )
        else:
            set_fields = {
                "updatedAt": now_iso,
                "updatedBy": actor,
            }
            if info.get("fullName"):
                set_fields["fullName"] = info["fullName"]
            if info.get("taxCode"):
                set_fields["taxCode"] = info["taxCode"]
            if info.get("latestPaymentDate"):
                set_fields["latestPaymentDate"] = info["latestPaymentDate"]
            if info.get("latestPeriod"):
                set_fields["latestPeriod"] = info["latestPeriod"]
            if info.get("update_onboard") and info.get("onboardDate"):
                set_fields["onboardDate"] = info["onboardDate"]
            if info.get("update_offboard") and info.get("offboardDate"):
                set_fields["offboardDate"] = info["offboardDate"]

            collab_ops.append(
                UpdateOne(
                    {"employeeCode": code},
                    {"$set": set_fields},
                )
            )

    if collab_ops:
        await db[COLLABORATOR_EXPENSES_COLLECTION].bulk_write(collab_ops, ordered=False)

    # 4. Gộp tất cả mã CTV có liên quan (bao gồm CTV trong file mới và CTV có dữ liệu cũ thuộc kỳ vừa bị xóa)
    all_affected_collab_codes = affected_collab_codes.union(old_period_collab_codes)

    # Tính toán và cập nhật lại chính xác totalPaymentCount cho tất cả CTV có liên quan
    if all_affected_collab_codes:
        pipeline_counts = [
            {"$match": {"employeeCode": {"$in": list(all_affected_collab_codes)}}},
            {"$group": {"_id": "$employeeCode", "count": {"$sum": 1}}},
        ]
        counts_res = await db[EXPENSE_DETAILS_COLLECTION].aggregate(pipeline_counts).to_list(len(all_affected_collab_codes) + 50)
        found_counts = {item["_id"]: item["count"] for item in counts_res}
        count_ops = [
            UpdateOne({"employeeCode": code}, {"$set": {"totalPaymentCount": found_counts.get(code, 0)}})
            for code in all_affected_collab_codes
        ]
        if count_ops:
            await db[COLLABORATOR_EXPENSES_COLLECTION].bulk_write(count_ops, ordered=False)

        # Cập nhật latestPaymentDate và latestPeriod chuẩn xác sau khi import / xóa
        pipeline_latest = [
            {"$match": {"employeeCode": {"$in": list(all_affected_collab_codes)}}},
            {"$sort": {"paymentDate": -1, "period": -1}},
            {
                "$group": {
                    "_id": "$employeeCode",
                    "latestPaymentDate": {"$first": "$paymentDate"},
                    "latestPeriod": {"$first": "$period"},
                }
            },
        ]
        latest_res = await db[EXPENSE_DETAILS_COLLECTION].aggregate(pipeline_latest).to_list(len(all_affected_collab_codes) + 50)
        latest_map = {item["_id"]: item for item in latest_res}
        latest_ops = [
            UpdateOne(
                {"employeeCode": code},
                {"$set": {
                    "latestPaymentDate": latest_map[code].get("latestPaymentDate") if code in latest_map else None,
                    "latestPeriod": latest_map[code].get("latestPeriod") if code in latest_map else None,
                }}
            )
            for code in all_affected_collab_codes
        ]
        if latest_ops:
            await db[COLLABORATOR_EXPENSES_COLLECTION].bulk_write(latest_ops, ordered=False)

    # Gom các dòng thiếu ngày BMK TT thành 1 cảnh báo duy nhất
    if missing_date_rows:
        range_str = format_row_ranges(missing_date_rows)
        warn_missing = (
            f"Phát hiện {len(missing_date_rows)} dòng thiếu 'NGÀY BMK TT' (Dòng {range_str}) "
            f"với tổng thực lĩnh {missing_date_net_total:,.0f} ₫. "
            f"Các dòng này đã bị bỏ qua do không có ngày thanh toán để xác định ngày chuyển tiền."
        )
        warnings_list.insert(0, warn_missing)

    target_period_label = format_period_label(target_period)
    logger.info(
        f"Import chi phí CTV kỳ {target_period_label} hoàn tất bởi {actor}: {total_valid_rows} dòng xử lý, "
        f"{imported_details} bản ghi mới (xóa {deleted_details_count} bản ghi cũ), "
        f"{new_reconcil_created} CTV mới tạo vào bmk_ctv_reconciliations, {len(warnings_list)} cảnh báo."
    )

    msg = (
        f"Import thành công kỳ {target_period_label}: {total_valid_rows} dòng xử lý "
        f"({imported_details} mới"
        + (f", xóa {deleted_details_count} bản ghi cũ" if deleted_details_count > 0 else "")
        + f"), thêm {new_reconcil_created} CTV vào đối soát."
    )
    if missing_date_rows:
        msg += f" Cảnh báo: Bỏ qua {len(missing_date_rows)} dòng thiếu 'NGÀY BMK TT' (Dòng {format_row_ranges(missing_date_rows)})."
    if warnings_list:
        msg += f" Có {len(warnings_list)} cảnh báo/thông báo cần lưu ý."

    return ExpenseImportResponse(
        success=True,
        totalRows=total_valid_rows,
        importedDetails=imported_details,
        updatedDetails=updated_details,
        newReconciliationCreated=new_reconcil_created,
        collaboratorsCount=len(affected_collab_codes),
        periodsAffected=sorted(list(periods_affected), reverse=True),
        warnings=warnings_list,
        message=msg,
    )


@router.get("/export")
async def export_expenses(
    period: Optional[str] = Query(None),
    work_location: Optional[str] = Query(None),
    search: Optional[str] = Query(None),
    mode: str = Query("details"),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Xuất file Excel dữ liệu chi phí thanh toán CTV.
    - mode == 'details': Xuất danh sách chi tiết các đợt thanh toán.
    - mode == 'collaborators': Xuất tổng hợp CTV trên từng row bằng template export_thanh_toan_phi_CTV.xlsx,
      các kỳ thanh toán xếp giảm dần theo ngày thanh toán (#1 là lần gần nhất).
    """
    if not isinstance(search, str):
        search = None
    if not isinstance(period, str):
        period = None
    if not isinstance(work_location, str):
        work_location = None
    if not isinstance(mode, str) or mode not in ["details", "collaborators"]:
        mode = "details"

    if mode == "details":
        wb = Workbook()
        ws = wb.active
        ws.title = "Chi tiết thanh toán"
        headers = [
            "STT", "Mã CTV", "Họ và tên", "Mã số thuế", "Tên tài khoản", "Số tài khoản",
            "Ngân hàng", "Tổng TN chịu thuế", "Thuế TNCN", "Thực lĩnh", "Nơi làm việc",
            "Ngày BMK TT", "Nội dung TT", "Ngày gia nhập", "Ngày chấm dứt", "Kỳ thanh toán"
        ]
        ws.append(headers)

        header_font = Font(name="Tahoma", size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        data_font = Font(name="Tahoma", size=10)
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center

        query = {}
        if period and period.strip():
            query["period"] = period.strip()
        if work_location and work_location.strip():
            query["workLocation"] = work_location.strip()
        if search and search.strip():
            s = search.strip()
            reg = {"$regex": re.escape(s), "$options": "i"}
            query["$or"] = [
                {"employeeCode": reg},
                {"fullName": reg},
                {"taxCode": reg},
            ]

        cursor = db[EXPENSE_DETAILS_COLLECTION].find(query).sort([("paymentDate", -1), ("employeeCode", 1)])
        row_idx = 2
        async for doc in cursor:
            bank = doc.get("bankAccount") or {}
            row_data = [
                row_idx - 1,
                doc.get("employeeCode", ""),
                doc.get("fullName", ""),
                doc.get("taxCode", ""),
                bank.get("accountName", ""),
                bank.get("accountNumber", ""),
                bank.get("bankName", ""),
                doc.get("taxableIncome", 0.0),
                doc.get("personalIncomeTax", 0.0),
                doc.get("netIncome", 0.0),
                doc.get("workLocation", ""),
                format_display_date(doc.get("paymentDate")),
                doc.get("paymentContent", ""),
                format_display_date(doc.get("onboardDate")),
                format_display_date(doc.get("offboardDate")),
                doc.get("period", ""),
            ]
            ws.append(row_data)

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.border = thin_border
                if col_idx in (1, 2, 4, 6, 12, 14, 15, 16):
                    cell.alignment = align_center
                elif col_idx in (8, 9, 10):
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = align_left

            row_idx += 1

        filename = f"chi_tiet_chi_phi_ctv_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    else:
        # Sử dụng template export_thanh_toan_phi_CTV.xlsx theo yêu cầu
        if not EXPORT_TEMPLATE_PATH.exists():
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File mẫu export_thanh_toan_phi_CTV.xlsx không tồn tại trên hệ thống",
            )

        from openpyxl.utils import get_column_letter

        wb = load_workbook(str(EXPORT_TEMPLATE_PATH))
        sheet_name = "Gui BMK" if "Gui BMK" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        # Xóa các dòng placeholder cũ (từ dòng 10 trở đi)
        if ws.max_row >= 10:
            ws.delete_rows(10, ws.max_row - 9)

        query = {}
        if search and search.strip():
            s = search.strip()
            reg = {"$regex": re.escape(s), "$options": "i"}
            query["$or"] = [
                {"employeeCode": reg},
                {"fullName": reg},
                {"taxCode": reg},
            ]

        collabs = (
            await db[COLLABORATOR_EXPENSES_COLLECTION]
            .find(query)
            .sort([("totalPaymentCount", -1), ("employeeCode", 1)])
            .to_list(10000)
        )
        collab_codes = [c.get("employeeCode") for c in collabs if c.get("employeeCode")]

        # Lấy tất cả chi tiết thanh toán của các CTV này, sắp xếp giảm dần theo ngày thanh toán
        details_cursor = db[EXPENSE_DETAILS_COLLECTION].find(
            {"employeeCode": {"$in": collab_codes}}
        ).sort([("paymentDate", -1), ("_id", -1)])
        all_details = await details_cursor.to_list(100000)

        payments_by_code = defaultdict(list)
        for d in all_details:
            c_code = d.get("employeeCode")
            if c_code:
                payments_by_code[c_code].append(d)

        data_font = Font(name="Times New Roman", size=11)
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        num_format_currency = '#,##0'

        # Xác định số kỳ thanh toán tối đa để mở rộng cột tiêu đề nếu > 3 kỳ (template có sẵn #1, #2, #3)
        max_payments = max([len(payments_by_code.get(c, [])) for c in collab_codes] or [1])
        if max_payments > 3:
            h_font = Font(name="Times New Roman", size=11, bold=True)
            h_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
            for p_i in range(3, max_payments):
                b_col = 9 + p_i * 2
                ws.merge_cells(start_row=8, start_column=b_col, end_row=8, end_column=b_col + 1)
                c8 = ws.cell(row=8, column=b_col, value=f"#{p_i + 1}")
                c8.font = h_font
                c8.alignment = align_center
                c8.fill = h_fill
                c8.border = thin_border
                ws.cell(row=8, column=b_col + 1).border = thin_border

                c9_net = ws.cell(row=9, column=b_col, value="THỰC LĨNH")
                c9_net.font = h_font
                c9_net.alignment = align_center
                c9_net.fill = h_fill
                c9_net.border = thin_border

                c9_date = ws.cell(row=9, column=b_col + 1, value="NGÀY BMK TT")
                c9_date.font = h_font
                c9_date.alignment = align_center
                c9_date.fill = h_fill
                c9_date.border = thin_border

                col_letter_net = get_column_letter(b_col)
                col_letter_date = get_column_letter(b_col + 1)
                ws.column_dimensions[col_letter_net].width = 17.1
                ws.column_dimensions[col_letter_date].width = 15.5

        total_cols = max(14, 8 + max_payments * 2)

        # Mở rộng merge tiêu đề bảng ở dòng 6
        if total_cols > 9:
            try:
                ws.unmerge_cells('A6:I6')
            except Exception:
                pass
            ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=total_cols)

        for idx, collab in enumerate(collabs, start=1):
            r_num = 9 + idx
            ws.row_dimensions[r_num].height = 20.0
            code = collab.get("employeeCode", "")
            full_name = collab.get("fullName", "")
            tax_code = collab.get("taxCode", "")
            onboard_date_str = format_display_date(collab.get("onboardDate"))
            offboard_date_str = format_display_date(collab.get("offboardDate"))
            ctv_payments = payments_by_code.get(code, [])
            total_payments_count = len(ctv_payments) or collab.get("totalPaymentCount", 0)
            total_net_income = sum(p.get("netIncome", 0.0) for p in ctv_payments)

            # Cột thông tin cơ bản:
            # Col 1: STT | Col 2: MNV | Col 3: HỌ VÀ TÊN | Col 4: MÃ SỐ THUẾ | Col 5: Số lần thanh toán
            # Col 6: Ngày gia nhập | Col 7: Ngày chấm dứt | Col 8: Tổng Thực lĩnh
            c_stt = ws.cell(row=r_num, column=1, value=idx)
            c_code = ws.cell(row=r_num, column=2, value=code)
            c_name = ws.cell(row=r_num, column=3, value=full_name)
            c_tax = ws.cell(row=r_num, column=4, value=tax_code)
            c_count = ws.cell(row=r_num, column=5, value=total_payments_count)
            c_onboard = ws.cell(row=r_num, column=6, value=onboard_date_str)
            c_offboard = ws.cell(row=r_num, column=7, value=offboard_date_str)
            c_tot_net = ws.cell(row=r_num, column=8, value=total_net_income)

            c_stt.alignment = align_center
            c_code.alignment = align_center
            c_name.alignment = align_left
            c_tax.alignment = align_center
            c_count.alignment = align_center
            c_onboard.alignment = align_center
            c_offboard.alignment = align_center
            c_tot_net.alignment = align_right
            c_tot_net.number_format = num_format_currency

            for cell_obj in (c_stt, c_code, c_name, c_tax, c_count, c_onboard, c_offboard, c_tot_net):
                cell_obj.font = data_font
                cell_obj.border = thin_border

            # Điền các kỳ thanh toán giảm dần theo ngày thanh toán (#1 là lần gần nhất)
            for p_idx, p in enumerate(ctv_payments):
                b_col = 9 + p_idx * 2
                c_net = ws.cell(row=r_num, column=b_col, value=p.get("netIncome", 0.0))
                c_date = ws.cell(row=r_num, column=b_col + 1, value=format_display_date(p.get("paymentDate")))

                c_net.alignment = align_right
                c_net.number_format = num_format_currency
                c_net.font = data_font
                c_net.border = thin_border

                c_date.alignment = align_center
                c_date.font = data_font
                c_date.border = thin_border

            # Viền border cho các cột còn lại nếu CTV có ít hơn số kỳ tối đa
            filled_cols = 8 + len(ctv_payments) * 2
            for empty_col in range(filled_cols + 1, total_cols + 1):
                empty_cell = ws.cell(row=r_num, column=empty_col)
                empty_cell.border = thin_border
                empty_cell.font = data_font

        filename = f"tong_hop_chi_phi_ctv_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@router.get("/export-monthly")
async def export_monthly_expenses(
    year: int = Query(..., description="Năm cần xuất báo cáo (ví dụ: 2026)"),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Xuất file Excel báo cáo chi phí CTV theo từng tháng của năm được chọn,
    sử dụng template export_bao_cao_chi_phi_theo_thang.xlsx.
    Tháng nào CTV không có dữ liệu để trống.
    """
    if not MONTHLY_EXPORT_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="File mẫu export_bao_cao_chi_phi_theo_thang.xlsx không tồn tại trên hệ thống",
        )

    year_str = str(year)
    cursor = db[EXPENSE_DETAILS_COLLECTION].find({
        "$or": [
            {"period": {"$regex": f"^{year_str}"}},
            {"paymentDate": {"$regex": f"^{year_str}"}}
        ]
    })
    details = await cursor.to_list(100000)

    # Tổng hợp theo từng CTV và từng tháng trong năm
    collab_data = {}
    for d in details:
        code = d.get("employeeCode")
        if not code:
            continue
        code = str(code).strip()

        # Xác định tháng trong năm
        p = str(d.get("period") or "").replace("-", "").strip()
        month = None
        if len(p) == 6 and p.startswith(year_str) and p[4:6].isdigit():
            m_val = int(p[4:6])
            if 1 <= m_val <= 12:
                month = m_val
        if month is None:
            pay_date = str(d.get("paymentDate") or "").strip()
            parts = pay_date.split("-")
            if len(parts) >= 2 and parts[0] == year_str and parts[1].isdigit():
                m_val = int(parts[1])
                if 1 <= m_val <= 12:
                    month = m_val

        if month is None:
            continue

        net = float(d.get("netIncome") or 0.0)

        if code not in collab_data:
            collab_data[code] = {
                "employeeCode": code,
                "fullName": d.get("fullName") or "",
                "taxCode": d.get("taxCode") or "",
                "monthly": defaultdict(float),
                "count": 0,
            }
        collab_data[code]["monthly"][month] += net
        collab_data[code]["count"] += 1
        if not collab_data[code]["fullName"] and d.get("fullName"):
            collab_data[code]["fullName"] = d.get("fullName")
        if not collab_data[code]["taxCode"] and d.get("taxCode"):
            collab_data[code]["taxCode"] = d.get("taxCode")

    # Lấy thông tin ngày gia nhập / ngày chấm dứt từ COLLABORATOR_EXPENSES_COLLECTION và RECONCILIATION_COLLECTION
    collab_codes = list(collab_data.keys())
    existing_collabs = {}
    if collab_codes:
        async for c_doc in db[COLLABORATOR_EXPENSES_COLLECTION].find(
            {"employeeCode": {"$in": collab_codes}},
            {"employeeCode": 1, "fullName": 1, "taxCode": 1, "onboardDate": 1, "offboardDate": 1}
        ):
            c_code = c_doc.get("employeeCode")
            if c_code:
                existing_collabs[str(c_code).strip()] = c_doc

        async for r_doc in db[RECONCILIATION_COLLECTION].find(
            {"employeeCode": {"$in": collab_codes}},
            {"employeeCode": 1, "fullName": 1, "idNumber": 1, "onboardDate": 1, "offboardDate": 1}
        ):
            c_code = r_doc.get("employeeCode")
            if c_code:
                c_str = str(c_code).strip()
                if c_str not in existing_collabs:
                    existing_collabs[c_str] = r_doc
                else:
                    if not existing_collabs[c_str].get("onboardDate") and r_doc.get("onboardDate"):
                        existing_collabs[c_str]["onboardDate"] = r_doc.get("onboardDate")
                    if not existing_collabs[c_str].get("offboardDate") and r_doc.get("offboardDate"):
                        existing_collabs[c_str]["offboardDate"] = r_doc.get("offboardDate")

    # Sắp xếp danh sách CTV theo mã (số trước, chữ sau)
    def sort_key(c):
        return (0, int(c)) if c.isdigit() else (1, c)

    sorted_codes = sorted(collab_codes, key=sort_key)

    # Mở template Excel
    wb = load_workbook(str(MONTHLY_EXPORT_TEMPLATE_PATH))
    sheet_name = "Gui BMK" if "Gui BMK" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    # Cập nhật tiêu đề báo cáo
    title_cell = ws.cell(row=6, column=1)
    title_cell.value = f"BẢNG THANH TOÁN PHÍ DỊCH VỤ CỘNG TÁC VIÊN NĂM {year}"

    # Xóa các dòng dữ liệu cũ nếu có (từ dòng 10 trở đi)
    if ws.max_row >= 10:
        ws.delete_rows(10, ws.max_row - 9)

    # Style chuẩn
    data_font = Font(name="Times New Roman", size=11)
    bold_font = Font(name="Times New Roman", size=11, bold=True)
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    double_bottom_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    num_format_currency = "#,##0"

    # Mapping cột tháng trong template:
    # Col 9 (I): Tháng 1
    # Col 10 (J): Tháng 2
    # Col 11 (K): Tháng 3
    # Col 12 (L): Để trống (duplicate Tháng 3 trong template)
    # Col 13 (M): Tháng 4
    # Col 14 (N): Tháng 5
    # Col 15 (O): Tháng 6
    # Col 16 (P): Tháng 7
    # Col 17 (Q): Tháng 8
    # Col 18 (R): Tháng 9
    # Col 19 (S): Tháng 10
    # Col 20 (T): Tháng 11
    # Col 21 (U): Tháng 12
    month_col_map = {
        1: 9,
        2: 10,
        3: 11,
        4: 13,
        5: 14,
        6: 15,
        7: 16,
        8: 17,
        9: 18,
        10: 19,
        11: 20,
        12: 21,
    }

    current_row = 10
    total_year_all = 0.0
    total_count_all = 0
    monthly_totals_all = defaultdict(float)

    for idx, code in enumerate(sorted_codes, start=1):
        item = collab_data[code]
        info = existing_collabs.get(code, {})

        full_name = info.get("fullName") or item.get("fullName", "")
        tax_code = info.get("taxCode") or info.get("idNumber") or item.get("taxCode", "")
        onboard_date = info.get("onboardDate")
        offboard_date = info.get("offboardDate")
        tx_count = item.get("count", 0)

        # Tính tổng thực lĩnh cả năm của CTV này
        c_monthly = item["monthly"]
        c_yearly_total = sum(c_monthly.values())

        total_year_all += c_yearly_total
        total_count_all += tx_count

        # Ghi các cột cơ bản
        ws.cell(row=current_row, column=1, value=idx)
        ws.cell(row=current_row, column=2, value=code)
        ws.cell(row=current_row, column=3, value=full_name)
        ws.cell(row=current_row, column=4, value=tax_code)
        ws.cell(row=current_row, column=5, value=tx_count)
        ws.cell(row=current_row, column=6, value=format_display_date(onboard_date))
        ws.cell(row=current_row, column=7, value=format_display_date(offboard_date))
        ws.cell(row=current_row, column=8, value=c_yearly_total)

        # Ghi các cột tháng (tháng nào không có để trống)
        for m in range(1, 13):
            col_target = month_col_map[m]
            m_val = c_monthly.get(m)
            if m_val and m_val > 0:
                ws.cell(row=current_row, column=col_target, value=m_val)
                monthly_totals_all[m] += m_val
            else:
                ws.cell(row=current_row, column=col_target, value=None)

        # Cột L (Col 12) để trống theo đúng template
        ws.cell(row=current_row, column=12, value=None)

        # Format styles cho từng cell
        for c_idx in range(1, 22):
            cell = ws.cell(row=current_row, column=c_idx)
            cell.font = data_font
            cell.border = thin_border
            if c_idx in (1, 2, 4, 5, 6, 7):
                cell.alignment = align_center
            elif c_idx == 3:
                cell.alignment = align_left
            else:
                cell.alignment = align_right
                cell.number_format = num_format_currency

        current_row += 1

    # Thêm dòng TỔNG CỘNG ở cuối bảng
    ws.cell(row=current_row, column=1, value="")
    ws.cell(row=current_row, column=2, value="")
    ws.cell(row=current_row, column=3, value="TỔNG CỘNG")
    ws.cell(row=current_row, column=4, value="")
    ws.cell(row=current_row, column=5, value=total_count_all)
    ws.cell(row=current_row, column=6, value="")
    ws.cell(row=current_row, column=7, value="")
    ws.cell(row=current_row, column=8, value=total_year_all)

    for m in range(1, 13):
        col_target = month_col_map[m]
        m_tot = monthly_totals_all.get(m, 0.0)
        ws.cell(row=current_row, column=col_target, value=m_tot if m_tot > 0 else None)

    ws.cell(row=current_row, column=12, value=None)

    for c_idx in range(1, 22):
        cell = ws.cell(row=current_row, column=c_idx)
        cell.font = bold_font
        cell.border = double_bottom_border
        cell.fill = total_fill
        if c_idx in (1, 2, 4, 6, 7):
            cell.alignment = align_center
        elif c_idx in (3, 5):
            cell.alignment = align_left if c_idx == 3 else align_center
        else:
            cell.alignment = align_right
            cell.number_format = num_format_currency

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    wb.close()

    filename = f"bao_cao_chi_phi_ctv_theo_thang_{year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    headers_resp = {
        "Content-Disposition": f"attachment; filename=\"{filename}\"",
        "Access-Control-Expose-Headers": "Content-Disposition",
    }
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers_resp,
    )
