import os
from datetime import date, datetime, timezone
from io import BytesIO
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, status
from fastapi.responses import FileResponse, StreamingResponse
from urllib.parse import quote
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from typing import List, Optional
from app.core.activity_log import record_activity
from app.core.database import get_db
from app.core.logging import get_logger
from pydantic import BaseModel
from app.core.security import get_current_user
from app.models.collaborator import CollaboratorCreate, CollaboratorUpdate, CollaboratorResponse
from app.core.s3 import upload_to_s3, get_s3_client
from app.core.config import settings


router = APIRouter(prefix="/api/collaborators", tags=["Collaborators"])
logger = get_logger(__name__)

COLLECTION = "bmk_ctv_collaborators"

SERVICE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE_PATH = os.path.join(SERVICE_ROOT, "templates", "collaborator_checklist_template.xlsx")
CONTRACT_REPORT_TEMPLATE_PATH = os.path.join(SERVICE_ROOT, "templates", "mau_bao_cao_hop_dong_ctv.xlsx")

# Template có 3 dòng tiêu đề (row 2-4, do merge cells) rồi mới tới dữ liệu (row 5 trở đi).
HEADER_START_ROW = 2
DATA_START_ROW = 5

EMPLOYEE_CODE_LABEL = "Mã nhân viên"
START_DATE_LABEL = "Ngày bắt đầu"
END_DATE_LABEL = "Ngày kết thúc"
LIQUIDATION_DATE_LABEL = "Biên bản thanh lí"

# (Tên cột trong file Excel, tên field key tương ứng trong checklist)
CHECKLIST_COLUMNS = [
    ("CCCD", "cccd"),
    ("Cam kết thuế", "ckt"),
]

# (Tên cột trong file Excel, tên field text tương ứng ở hồ sơ CTV) - dùng khi cần tạo mới CTV từ file import
PROFILE_TEXT_COLUMNS = [
    ("Họ tên", "fullName"),
    ("Mã số thuế", "taxCode"),
    ("Số CCCD", "idNumber"),
    ("Email", "email"),
    ("Số điện thoại", "phone"),
    ("Địa chỉ", "address"),
]
DOB_LABEL = "Ngày sinh"
HANDOVER_DATE_LABEL = "Ngày bàn giao"
NOTED_LABEL = "Ghi chú"

DATE_TEXT_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y"]

_UNSET = object()

def _now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

def _parse_excel_date(value) -> str | None:
    """Chuyển giá trị 1 ô Excel (kiểu Date hoặc Text) thành chuỗi ISO 'YYYY-MM-DD'."""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return from_excel(value).date().isoformat()

    text = str(value).strip()
    if not text or text in ("-", "—"):
        return None
    for fmt in DATE_TEXT_FORMATS:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"không nhận dạng được định dạng ngày '{text}'")

def _format_date_vn(val) -> str:
    if not val:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")
    val_str = str(val).strip()
    if not val_str or val_str in ("-", "—", "None", "null"):
        return ""
    if len(val_str) == 10 and val_str[2] == "/" and val_str[5] == "/":
        return val_str
    if len(val_str) >= 10 and val_str[4] == "-" and val_str[7] == "-":
        parts = val_str[:10].split("-")
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(val_str[:10], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return val_str

def _cell_str(row, idx) -> str:
    if idx is None or idx >= len(row):
        return ""
    value = row[idx]
    return "" if value is None else str(value).strip()

def _to_response(doc: dict) -> dict:
    doc = dict(doc)
    doc["employeeCode"] = doc["_id"]
    if "noted" not in doc or doc["noted"] is None:
        doc["noted"] = ""

    raw_handover_info = doc.get("handoverInfo")
    if not isinstance(raw_handover_info, list):
        raw_handover_info = []

    doc["handoverInfo"] = raw_handover_info
    doc.pop("handoverDate", None)
    doc.pop("handoverPerson", None)
    return doc

def _actor_name(current_user: dict) -> str:
    return current_user.get("name") or current_user.get("username", "")

@router.get("", response_model=List[CollaboratorResponse])
async def list_collaborators(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Fetch all collaborators, sorted by employee code."""
    items = []
    cursor = db[COLLECTION].find({}).sort("_id", 1)
    async for doc in cursor:
        items.append(_to_response(doc))
    return items

@router.get("/template")
async def download_import_template(current_user: dict = Depends(get_current_user)):
    """Download the Excel template used for bulk checklist import."""
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không tìm thấy file mẫu")
    return FileResponse(
        TEMPLATE_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="mau_import_checklist_ctv.xlsx",
    )

@router.get("/export")
async def export_collaborators(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Export all collaborators' checklist status to an Excel file using the template."""
    full_name = _actor_name(current_user)
    if not os.path.exists(TEMPLATE_PATH):
        await record_activity(
            db, action="export_collaborators", result="fail", full_name=full_name,
            username=current_user.get("username", ""),
            message=f"{full_name} đã xuất thất bại danh sách cộng tác viên",
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Không tìm thấy file template mẫu"
        )

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    data_font = Font(name='Aptos Narrow', size=11)

    cursor = db[COLLECTION].find({}).sort("_id", 1)
    row_idx = 5
    stt = 1

    async for doc in cursor:
        checklist = doc.get("checklist") or {}
        
        cccd = checklist.get("cccd") or {}
        ckt = checklist.get("ckt") or {}
        hddv = checklist.get("hddv") or {}
        bbtl = checklist.get("bbtl") or {}

        contracts = hddv.get("contract_date", [])
        start_date = ""
        end_date = ""
        if contracts:
            last_contract = contracts[-1]
            if isinstance(last_contract, dict):
                start_date = last_contract.get("startDate") or ""
                end_date = last_contract.get("endDate") or ""
            else:
                start_date = getattr(last_contract, "startDate", "") or ""
                end_date = getattr(last_contract, "endDate", "") or ""

        handover_info = doc.get("handoverInfo") or []
        sorted_ho = sorted(
            [h for h in handover_info if isinstance(h, dict) and h.get("handoverDate")],
            key=lambda x: str(x.get("handoverDate"))
        )
        latest_ho = sorted_ho[-1] if sorted_ho else {}

        row_values = [
            stt,
            doc.get("_id") or doc.get("employeeCode", ""),
            doc.get("fullName", ""),
            doc.get("taxCode", ""),
            doc.get("dob", ""),
            doc.get("idNumber", ""),
            doc.get("email", ""),
            doc.get("phone", ""),
            doc.get("address", ""),
            start_date,
            end_date,
            "Đã nộp" if cccd.get("checked") else "",
            "Đã nộp" if ckt.get("checked") else "",
            bbtl.get("date") or "",
            latest_ho.get("handoverPerson") or "",
            latest_ho.get("handoverDate") or "",
            doc.get("noted") or "",
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border

            # Alignment formatting
            if col_idx in [1, 2, 4, 5, 6, 8, 10, 11, 12, 13, 14, 16]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        row_idx += 1
        stt += 1

    # Clear remaining rows in the template if they are beyond our data rows
    for r in range(row_idx, ws.max_row + 1):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    exported_count = stt - 1
    await record_activity(
        db, action="export_collaborators", result="success", full_name=full_name,
        username=current_user.get("username", ""),
        message=f"{full_name} đã xuất thành công {exported_count} cộng tác viên",
    )

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=danh_sach_ctv.xlsx"},
    )

@router.get("/export-doisoat")
async def export_collaborators_doisoat(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Export all collaborators for đối soát to an Excel file using the template."""
    full_name = _actor_name(current_user)
    doisoat_template_path = os.path.join(SERVICE_ROOT, "templates", "mau_doisoat_ctv.xlsx")
    if not os.path.exists(doisoat_template_path):
        await record_activity(
            db, action="export_collaborators_doisoat", result="fail", full_name=full_name,
            username=current_user.get("username", ""),
            message=f"{full_name} đã xuất đối soát thất bại danh sách cộng tác viên",
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Không tìm thấy file template mẫu đối soát"
        )

    wb = load_workbook(doisoat_template_path)
    ws = wb.active

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    data_font = Font(name='Aptos Narrow', size=11)

    cursor = db[COLLECTION].find({}).sort("_id", 1)
    row_idx = 5
    stt = 1

    async for doc in cursor:
        checklist = doc.get("checklist") or {}
        
        cccd = checklist.get("cccd") or {}
        ckt = checklist.get("ckt") or {}
        hddv = checklist.get("hddv") or {}
        bbtl = checklist.get("bbtl") or {}

        # Hợp đồng dịch vụ: Lấy ngày bắt đầu và ngày kết thúc của hợp đồng cuối cùng
        contracts = hddv.get("contract_date", [])
        start_date = ""
        end_date = ""
        if contracts:
            last_contract = contracts[-1]
            if isinstance(last_contract, dict):
                start_date = last_contract.get("startDate") or ""
                end_date = last_contract.get("endDate") or ""
            else:
                start_date = getattr(last_contract, "startDate", "") or ""
                end_date = getattr(last_contract, "endDate", "") or ""

        # Bản scan: Hợp đồng dịch vụ có file upload (số lượng file)
        hddv_files = hddv.get("files") or []
        hddv_scan = len(hddv_files) if hddv_files else ""

        # CCCD
        cccd_excel = "X" if cccd.get("checked") else ""
        cccd_scan = "X" if cccd.get("file") else ""

        # Cam kết thuế
        ckt_excel = "X" if ckt.get("checked") else ""
        ckt_scan = "X" if ckt.get("file") else ""

        # Biên bản thanh lí
        bbtl_excel = "X" if bbtl.get("date") else ""
        bbtl_scan = "X" if bbtl.get("file") else ""

        row_values = [
            stt,
            doc.get("_id") or doc.get("employeeCode", ""),
            doc.get("fullName", ""),
            start_date,
            end_date,
            hddv_scan,
            cccd_excel,
            cccd_scan,
            ckt_excel,
            ckt_scan,
            bbtl_excel,
            bbtl_scan
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border

            # Alignment formatting
            if col_idx == 3:  # Họ tên
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        row_idx += 1
        stt += 1

    # Clear remaining rows in the template if they are beyond our data rows
    for r in range(row_idx, ws.max_row + 1):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    exported_count = stt - 1
    await record_activity(
        db, action="export_collaborators_doisoat", result="success", full_name=full_name,
        username=current_user.get("username", ""),
        message=f"{full_name} đã xuất đối soát thành công {exported_count} cộng tác viên",
    )

    filename = f"bmk_ctv_doisoat_hoso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/export-contracts")
async def export_collaborators_contracts(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Export contract report for all collaborators using template mau_bao_cao_hop_dong_ctv.xlsx."""
    full_name = _actor_name(current_user)
    if not os.path.exists(CONTRACT_REPORT_TEMPLATE_PATH):
        await record_activity(
            db, action="export_collaborators_contracts", result="fail", full_name=full_name,
            username=current_user.get("username", ""),
            message=f"{full_name} đã xuất thất bại báo cáo hợp đồng cộng tác viên (không tìm thấy file mẫu)",
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Không tìm thấy file template mẫu báo cáo hợp đồng"
        )

    wb = load_workbook(CONTRACT_REPORT_TEMPLATE_PATH)
    ws = wb["Báo cáo HĐ"] if "Báo cáo HĐ" in wb.sheetnames else wb.active

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    data_font = Font(name='Aptos Narrow', size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    cursor = db[COLLECTION].find({}).sort("_id", 1)
    docs = await cursor.to_list(None)

    # Lấy thông tin hợp đồng cho từng cộng tác viên và xác định số cột hợp đồng tối đa
    parsed_collabs = []
    max_contracts = 0
    for doc in docs:
        checklist = doc.get("checklist") or {}
        hddv = checklist.get("hddv") or {}
        raw_contracts = hddv.get("contract_date") or []
        valid_contracts = []
        for c in raw_contracts:
            if isinstance(c, dict):
                s = c.get("startDate")
                e = c.get("endDate")
            else:
                s = getattr(c, "startDate", None)
                e = getattr(c, "endDate", None)
            if s or e:
                valid_contracts.append({"startDate": s, "endDate": e})
        valid_contracts.sort(key=lambda x: str(x.get("startDate") or x.get("endDate") or ""))
        if len(valid_contracts) > max_contracts:
            max_contracts = len(valid_contracts)
        parsed_collabs.append({
            "employee_code": doc.get("_id") or doc.get("employeeCode", ""),
            "full_name": doc.get("fullName", ""),
            "contracts": valid_contracts,
        })

    num_slots = max(5, max_contracts)

    # Nếu có cộng tác viên có nhiều hơn 5 hợp đồng, mở rộng thêm các cột tiêu đề hợp đồng
    if num_slots > 5:
        header_fill = PatternFill(start_color="EF9263", end_color="EF9263", fill_type="solid")
        header_font = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
        sub_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        for k in range(6, num_slots + 1):
            c_start = 5 + (k - 1) * 2
            c_end = c_start + 1
            ws.merge_cells(start_row=2, start_column=c_start, end_row=2, end_column=c_end)
            h_cell1 = ws.cell(row=2, column=c_start, value=f"Hợp đồng {k}")
            h_cell1.font = header_font
            h_cell1.fill = header_fill
            h_cell1.alignment = center_align
            h_cell1.border = thin_border
            h_cell2 = ws.cell(row=2, column=c_end)
            h_cell2.fill = header_fill
            h_cell2.border = thin_border

            sub1 = ws.cell(row=3, column=c_start, value="Ngày bắt đầu")
            sub1.font = sub_font
            sub1.fill = header_fill
            sub1.alignment = center_align
            sub1.border = thin_border

            sub2 = ws.cell(row=3, column=c_end, value="Ngày kết thúc")
            sub2.font = sub_font
            sub2.fill = header_fill
            sub2.alignment = center_align
            sub2.border = thin_border

    row_idx = 4
    stt = 1
    max_col = 4 + num_slots * 2

    for item in parsed_collabs:
        c_list = item["contracts"]
        contract_count = len(c_list)
        row_values = [
            stt,
            item["employee_code"],
            item["full_name"],
            contract_count,
        ]
        for i in range(num_slots):
            if i < contract_count:
                row_values.append(_format_date_vn(c_list[i].get("startDate")))
                row_values.append(_format_date_vn(c_list[i].get("endDate")))
            else:
                row_values.append(None)
                row_values.append(None)

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = left_align if col_idx == 3 else center_align
            if col_idx >= 5:
                cell.number_format = '@'

        row_idx += 1
        stt += 1

    # Dọn dẹp các dòng mẫu còn dư nếu có
    for r in range(row_idx, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    exported_count = stt - 1
    await record_activity(
        db, action="export_collaborators_contracts", result="success", full_name=full_name,
        username=current_user.get("username", ""),
        message=f"{full_name} đã xuất thành công báo cáo hợp đồng dịch vụ cho {exported_count} cộng tác viên",
    )

    filename = f"bmk_ctv_bao_cao_hop_dong_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.post("/import")
async def import_collaborators(
    file: UploadFile = File(...), db=Depends(get_db), current_user: dict = Depends(get_current_user)
):
    """Bulk-create/update collaborators from an uploaded Excel file (see /template) and log to S3/MongoDB history."""
    username = current_user.get("username", "unknown")
    full_name = _actor_name(current_user)
    logger.info(f"Bắt đầu import '{file.filename}' bởi user='{username}'")

    # Generate unique S3 Key and read contents
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    s3_key = f"excel/{timestamp}_{file.filename}"
    s3_bucket = settings.S3_BUCKET

    content = await file.read()

    # 1. Upload to S3
    try:
        upload_to_s3(content, s3_key)
    except Exception as s3_err:
        error_msg = f"Lỗi upload S3: {str(s3_err)}"
        await db["bmk_ctv_upload_history"].insert_one({
            "filename": file.filename,
            "s3Key": s3_key,
            "s3Bucket": s3_bucket,
            "uploadedBy": full_name,
            "username": username,
            "rowsProcessed": 0,
            "createdCount": 0,
            "updatedCount": 0,
            "status": "fail",
            "message": error_msg,
            "createdAt": _now(),
            "group": "CTV"
        })
        await record_activity(
            db, action="import_collaborators", result="fail", full_name=full_name, username=username,
            message=f"{full_name} đã nhập thất bại 0 cộng tác viên ({error_msg})",
        )
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Không thể lưu trữ file lên S3: {str(s3_err)}"
        )

    # 2. Process logic
    try:
        if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Chỉ hỗ trợ file Excel (.xlsx)")

        try:
            wb = load_workbook(BytesIO(content), data_only=True)
        except Exception:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Không đọc được file Excel, vui lòng dùng đúng file mẫu")

        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < DATA_START_ROW:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="File không có dữ liệu")

        header_row_slices = rows[HEADER_START_ROW - 1 : DATA_START_ROW - 1]
        header = [
            next((str(cell).strip() for cell in reversed(col_cells) if cell not in (None, "")), "")
            for col_cells in zip(*header_row_slices)
        ]

        def find_col_idx(lbl):
            for i, h in enumerate(header):
                if h and h.strip().lower() == lbl.strip().lower():
                    return i
            return None

        code_idx = find_col_idx(EMPLOYEE_CODE_LABEL)
        if code_idx is None:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Không tìm thấy cột 'Mã nhân viên' trong file, vui lòng dùng đúng file mẫu",
            )
        start_date_idx = find_col_idx(START_DATE_LABEL)
        end_date_idx = find_col_idx(END_DATE_LABEL)
        liquidation_idx = find_col_idx(LIQUIDATION_DATE_LABEL)
        dob_idx = find_col_idx(DOB_LABEL)
        handover_date_idx = find_col_idx(HANDOVER_DATE_LABEL)
        noted_idx = find_col_idx(NOTED_LABEL)

        column_indices = {}
        for label, field in CHECKLIST_COLUMNS:
            idx = find_col_idx(label)
            if idx is not None:
                column_indices[field] = idx

        profile_indices = {}
        for label, field in PROFILE_TEXT_COLUMNS:
            idx = find_col_idx(label)
            if idx is not None:
                profile_indices[field] = idx

        updated: List[str] = []
        created: List[str] = []
        date_errors: List[str] = []
        checklist_status_by_emp = {}

        for current_row_num, row in enumerate(rows[DATA_START_ROW - 1 :], start=DATA_START_ROW):
            if row is None or all(cell in (None, "") for cell in row):
                continue

            raw_code = row[code_idx] if code_idx < len(row) else None
            employee_code = str(raw_code).strip() if raw_code is not None else ""
            if not employee_code:
                if not all(cell in (None, "") for cell in row):
                    date_errors.append(f"Dòng {current_row_num}: Thiếu Mã nhân viên")
                continue

            existing = await db[COLLECTION].find_one({"_id": employee_code})

            checklist_values = {}
            if employee_code not in checklist_status_by_emp:
                if existing:
                    db_checklist = existing.get("checklist") or {}
                    checklist_status_by_emp[employee_code] = {
                        field: db_checklist.get(field, {}).get("checked", False) if isinstance(db_checklist.get(field), dict) else False
                        for field, _ in CHECKLIST_COLUMNS
                    }
                else:
                    checklist_status_by_emp[employee_code] = {}

            for field, col_idx in column_indices.items():
                value = row[col_idx] if col_idx < len(row) else None
                is_checked = value is not None and str(value).strip() != ""
                if field in checklist_status_by_emp[employee_code]:
                    checklist_status_by_emp[employee_code][field] = checklist_status_by_emp[employee_code][field] or is_checked
                else:
                    checklist_status_by_emp[employee_code][field] = is_checked
                checklist_values[field] = checklist_status_by_emp[employee_code][field]

            def parse_cell_date(idx):
                if idx is None or idx >= len(row):
                    return _UNSET
                raw = row[idx]
                if raw is None or (isinstance(raw, str) and not raw.strip()):
                    return None
                try:
                    return _parse_excel_date(raw)
                except ValueError as exc:
                    date_errors.append(f"Dòng {current_row_num} (Mã NV: {employee_code}), cột '{header[idx]}': {exc}")
                    return _UNSET

            start_val = parse_cell_date(start_date_idx)
            end_val = parse_cell_date(end_date_idx)
            liquidation_val = parse_cell_date(liquidation_idx)

            if existing:
                updates = {}
                for field, val in checklist_values.items():
                    updates[f"checklist.{field}.checked"] = val

                if start_date_idx is not None or end_date_idx is not None:
                    new_start = start_val if start_val is not _UNSET else None
                    new_end = end_val if end_val is not _UNSET else None

                    # Chỉ xử lý kiểm tra / thêm mới hợp đồng dịch vụ nếu có ít nhất 1 giá trị ngày (không phải cả 2 đều rỗng)
                    if new_start is not None or new_end is not None:
                        hddv = existing.get("checklist", {}).get("hddv") or {}
                        contracts = hddv.get("contract_date") or []

                        def _norm_date(val):
                            if not val:
                                return None
                            s = str(val).strip()
                            return s.split("T")[0] if "T" in s else s if s else None

                        norm_new_start = _norm_date(new_start)
                        norm_new_end = _norm_date(new_end)

                        # Rule: Kiểm tra xem cặp (startDate, endDate) đã tồn tại trong danh sách hợp đồng chưa
                        already_exists = False
                        for c in contracts:
                            if isinstance(c, dict):
                                c_start = c.get("startDate")
                                c_end = c.get("endDate")
                            else:
                                c_start = getattr(c, "startDate", None)
                                c_end = getattr(c, "endDate", None)

                            if _norm_date(c_start) == norm_new_start and _norm_date(c_end) == norm_new_end:
                                already_exists = True
                                break

                        # Nếu chưa tồn tại: bổ sung vào danh sách hợp đồng (nếu đã tồn tại: không thêm mới)
                        if not already_exists:
                            new_contract_item = {"startDate": new_start, "endDate": new_end}
                            # Nếu danh sách hiện tại chỉ có đúng 1 hợp đồng rỗng (chưa có ngày bắt đầu và kết thúc)
                            # thì thay thế phần tử rỗng đó bằng hợp đồng mới
                            if (
                                len(contracts) == 1
                                and isinstance(contracts[0], dict)
                                and not _norm_date(contracts[0].get("startDate"))
                                and not _norm_date(contracts[0].get("endDate"))
                            ):
                                new_contracts = [new_contract_item]
                            else:
                                new_contracts = list(contracts) + [new_contract_item]
                            updates["checklist.hddv.contract_date"] = new_contracts

                if liquidation_val is not _UNSET:
                    bbtl = existing.get("checklist", {}).get("bbtl") or {}
                    if liquidation_val is not None or not bbtl.get("date"):
                        updates["checklist.bbtl.date"] = liquidation_val

                # Cập nhật thông tin hồ sơ nếu có giá trị mới trong dòng hiện tại
                for label, field in PROFILE_TEXT_COLUMNS:
                    if field == "handoverPerson":
                        continue
                    idx = profile_indices.get(field)
                    val = _cell_str(row, idx)
                    if val:
                        updates[field] = val

                if dob_idx is not None:
                    dob_val = parse_cell_date(dob_idx)
                    if dob_val is not _UNSET and dob_val is not None:
                        updates["dob"] = dob_val

                # Xử lý thông tin bàn giao (handoverInfo là list các object {handoverDate, handoverPerson, createdAt})
                handover_person_val = _cell_str(row, find_col_idx("Người bàn giao")) if find_col_idx("Người bàn giao") is not None else ""
                handover_date_val = parse_cell_date(handover_date_idx) if handover_date_idx is not None else _UNSET

                existing_handover_info = list(existing.get("handoverInfo") or [])

                if handover_date_val is not _UNSET and handover_date_val is not None:
                    found_item = None
                    for item in existing_handover_info:
                        if isinstance(item, dict) and item.get("handoverDate") == handover_date_val:
                            found_item = item
                            break

                    if found_item:
                        if handover_person_val:
                            found_item["handoverPerson"] = handover_person_val
                        if not found_item.get("createdAt"):
                            found_item["createdAt"] = _now()
                    else:
                        existing_handover_info.append({
                            "handoverDate": handover_date_val,
                            "handoverPerson": handover_person_val,
                            "createdAt": _now()
                        })

                    updates["handoverInfo"] = existing_handover_info

                if noted_idx is not None:
                    new_noted = _cell_str(row, noted_idx)
                    if new_noted:
                        old_noted = (existing.get("noted") or "").strip()
                        if old_noted:
                            existing_lines = [line.strip() for line in old_noted.splitlines() if line.strip()]
                            if new_noted not in existing_lines:
                                updates["noted"] = f"{old_noted}\n{new_noted}"
                        else:
                            updates["noted"] = new_noted

                updates["updatedAt"] = _now()
                await db[COLLECTION].update_one({"_id": employee_code}, {"$set": updates})
                updated.append(employee_code)
            else:
                now_ts = _now()
                dob_val = parse_cell_date(dob_idx)
                handover_date_val = parse_cell_date(handover_date_idx) if handover_date_idx is not None else _UNSET
                handover_person_val = _cell_str(row, find_col_idx("Người bàn giao")) if find_col_idx("Người bàn giao") is not None else ""
                noted_val = _cell_str(row, noted_idx) if noted_idx is not None else ""

                handover_info_list = []
                if handover_date_val is not _UNSET and handover_date_val is not None:
                    handover_info_list.append({
                        "handoverDate": handover_date_val,
                        "handoverPerson": handover_person_val,
                        "createdAt": now_ts
                    })

                new_doc = {
                    "_id": employee_code,
                    "employeeCode": employee_code,
                    "fullName": _cell_str(row, profile_indices.get("fullName")),
                    "taxCode": _cell_str(row, profile_indices.get("taxCode")),
                    "dob": None if dob_val is _UNSET else dob_val,
                    "idNumber": _cell_str(row, profile_indices.get("idNumber")),
                    "email": _cell_str(row, profile_indices.get("email")),
                    "phone": _cell_str(row, profile_indices.get("phone")),
                    "address": _cell_str(row, profile_indices.get("address")),
                    "noted": noted_val,
                    "handoverInfo": handover_info_list,
                    "checklist": {
                        "cccd": {
                            "checked": checklist_values.get("cccd", False),
                            "file": None
                        },
                        "ckt": {
                            "checked": checklist_values.get("ckt", False),
                            "file": None
                        },
                        "hddv": {
                            "contract_date": [{
                                "startDate": None if start_val is _UNSET else start_val,
                                "endDate": None if end_val is _UNSET else end_val,
                            }],
                            "files": []
                        },
                        "bbtl": {
                            "date": None if liquidation_val is _UNSET else liquidation_val,
                            "file": None
                        }
                    },
                    "createdAt": now_ts,
                    "updatedAt": now_ts,
                }
                await db[COLLECTION].insert_one(new_doc)
                created.append(employee_code)

        total_processed = len(created) + len(updated)
        logger.info(
            f"Import '{file.filename}' hoàn tất bởi user='{username}': "
            f"tạo mới={len(created)}, cập nhật={len(updated)}, lỗi ngày tháng={len(date_errors)}"
        )
        if date_errors:
            logger.warning(f"Import '{file.filename}' có {len(date_errors)} lỗi định dạng ngày: {date_errors}")

        success_msg = f"Nhập thành công {total_processed} cộng tác viên (Tạo mới: {len(created)}, Cập nhật: {len(updated)})"
        if date_errors:
            success_msg += f". Có {len(date_errors)} lỗi dữ liệu:\n" + "\n".join(f"- {err}" for err in date_errors)

        await record_activity(
            db, action="import_collaborators", result="success", full_name=full_name, username=username,
            message=f"{full_name} đã nhập thành công {total_processed} cộng tác viên",
        )

        # Record success history
        await db["bmk_ctv_upload_history"].insert_one({
            "filename": file.filename,
            "s3Key": s3_key,
            "s3Bucket": s3_bucket,
            "uploadedBy": full_name,
            "username": username,
            "rowsProcessed": total_processed,
            "createdCount": len(created),
            "updatedCount": len(updated),
            "status": "success",
            "message": success_msg,
            "createdAt": _now(),
            "group": "CTV"
        })

        return {
            "updatedCount": len(updated),
            "updated": updated,
            "createdCount": len(created),
            "created": created,
            "dateErrorCount": len(date_errors),
            "dateErrors": date_errors,
        }

    except HTTPException as he:
        logger.warning(f"Import thất bại (HTTP {he.status_code}): {he.detail}")
        await db["bmk_ctv_upload_history"].insert_one({
            "filename": file.filename,
            "s3Key": s3_key,
            "s3Bucket": s3_bucket,
            "uploadedBy": full_name,
            "username": username,
            "rowsProcessed": 0,
            "createdCount": 0,
            "updatedCount": 0,
            "status": "fail",
            "message": he.detail,
            "createdAt": _now(),
            "group": "CTV"
        })
        await record_activity(
            db, action="import_collaborators", result="fail", full_name=full_name, username=username,
            message=f"{full_name} đã nhập thất bại 0 cộng tác viên: {he.detail}",
        )
        raise he

    except Exception as exc:
        error_detail = str(exc)
        logger.error(f"Import thất bại (Lỗi không xác định): {error_detail}")
        await db["bmk_ctv_upload_history"].insert_one({
            "filename": file.filename,
            "s3Key": s3_key,
            "s3Bucket": s3_bucket,
            "uploadedBy": full_name,
            "username": username,
            "rowsProcessed": 0,
            "createdCount": 0,
            "updatedCount": 0,
            "status": "fail",
            "message": f"Lỗi hệ thống: {error_detail}",
            "createdAt": _now(),
            "group": "CTV"
        })
        await record_activity(
            db, action="import_collaborators", result="fail", full_name=full_name, username=username,
            message=f"{full_name} đã nhập thất bại 0 cộng tác viên: {error_detail}",
        )
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=f"Không đọc được file hoặc lỗi xử lý dữ liệu: {error_detail}")

@router.get("/{employee_code}", response_model=CollaboratorResponse)
async def get_collaborator(employee_code: str, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Fetch a single collaborator by employee code."""
    doc = await db[COLLECTION].find_one({"_id": employee_code})
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )
    return _to_response(doc)

def _normalize_handover_info(doc: dict) -> dict:
    handover_info = doc.get("handoverInfo") or []
    normalized = []
    now_str = _now()
    for item in handover_info:
        if isinstance(item, dict) and item.get("handoverDate"):
            normalized.append({
                "handoverDate": item["handoverDate"],
                "handoverPerson": item.get("handoverPerson") or "",
                "createdAt": item.get("createdAt") or now_str
            })
        elif hasattr(item, "handoverDate") and getattr(item, "handoverDate"):
            normalized.append({
                "handoverDate": getattr(item, "handoverDate"),
                "handoverPerson": getattr(item, "handoverPerson", "") or "",
                "createdAt": getattr(item, "createdAt", None) or now_str
            })
    doc["handoverInfo"] = normalized
    doc.pop("handoverDate", None)
    doc.pop("handoverPerson", None)
    return doc

@router.post("", response_model=CollaboratorResponse, status_code=status.HTTP_201_CREATED)
async def create_collaborator(payload: CollaboratorCreate, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Create a new collaborator profile."""
    full_name = _actor_name(current_user)
    employee_code = payload.employeeCode.strip()
    if not employee_code:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Mã nhân viên là bắt buộc")

    existing = await db[COLLECTION].find_one({"_id": employee_code})
    if existing:
        await record_activity(
            db, action="create_collaborator", result="fail", full_name=full_name,
            username=current_user.get("username", ""),
            message=f"{full_name} tạo thất bại hồ sơ cho cộng tác viên mã {employee_code}",
            employee_code=employee_code,
        )
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f'Mã nhân viên "{employee_code}" đã tồn tại'
        )

    now = _now()
    doc = payload.model_dump()
    doc = _normalize_handover_info(doc)
    doc["employeeCode"] = employee_code
    doc["_id"] = employee_code
    doc["createdAt"] = now
    doc["updatedAt"] = now
    await db[COLLECTION].insert_one(doc)
    noted_val = (payload.noted or "").strip()
    noted_suffix = f" (Ghi chú: '{noted_val}')" if noted_val else ""
    await record_activity(
        db, action="create_collaborator", result="success", full_name=full_name,
        username=current_user.get("username", ""),
        message=f"{full_name} tạo thành công hồ sơ cho cộng tác viên mã {employee_code}{noted_suffix}",
        employee_code=employee_code,
    )
    return _to_response(doc)

@router.put("/{employee_code}", response_model=CollaboratorResponse)
async def update_collaborator(employee_code: str, payload: CollaboratorUpdate, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Update an existing collaborator profile."""
    full_name = _actor_name(current_user)
    username = current_user.get("username", "")

    existing = await db[COLLECTION].find_one({"_id": employee_code})
    if not existing:
        await record_activity(
            db, action="update_collaborator", result="fail", full_name=full_name, username=username,
            message=f"{full_name} cập nhật thất bại hồ sơ cho cộng tác viên mã {employee_code}",
            employee_code=employee_code,
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )

    new_code = payload.employeeCode.strip() or employee_code
    if new_code != employee_code:
        conflict = await db[COLLECTION].find_one({"_id": new_code})
        if conflict:
            await record_activity(
                db, action="update_collaborator", result="fail", full_name=full_name, username=username,
                message=f"{full_name} cập nhật thất bại hồ sơ cho cộng tác viên mã {employee_code}",
                employee_code=employee_code,
            )
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f'Mã nhân viên "{new_code}" đã tồn tại'
            )

    doc = payload.model_dump()
    doc = _normalize_handover_info(doc)
    doc["employeeCode"] = new_code
    doc["_id"] = new_code
    doc["createdAt"] = existing["createdAt"]
    doc["updatedAt"] = _now()

    # Kiểm tra các thay đổi về thông tin cá nhân để ghi log chi tiết
    personal_fields = {
        "fullName": "Họ tên",
        "taxCode": "Mã số thuế",
        "dob": "Ngày sinh",
        "idNumber": "Số CCCD",
        "email": "Email",
        "phone": "Số điện thoại",
        "address": "Địa chỉ",
        "noted": "Ghi chú/Lưu ý",
    }
    personal_changes = []
    for field_key, field_name in personal_fields.items():
        old_val = existing.get(field_key)
        new_val = doc.get(field_key)
        old_str = (str(old_val).strip()) if old_val is not None else ""
        new_str = (str(new_val).strip()) if new_val is not None else ""
        if old_str != new_str:
            if field_key in ("noted", "handoverPerson", "handoverDate"):
                if not old_str and new_str:
                    personal_changes.append(f"thêm {field_name} '{new_str}'")
                elif old_str and not new_str:
                    personal_changes.append(f"xóa {field_name}")
                else:
                    personal_changes.append(f"{field_name} '{old_str}' -> '{new_str}'")
            else:
                if old_str and new_str:
                    personal_changes.append(f"{field_name} ('{old_str}' -> '{new_str}')")
                elif new_str:
                    personal_changes.append(f"{field_name}: '{new_str}'")
                else:
                    personal_changes.append(f"xóa {field_name}")

    if new_code != employee_code:
        await db[COLLECTION].delete_one({"_id": employee_code})
    await db[COLLECTION].replace_one({"_id": new_code}, doc, upsert=True)

    if personal_changes:
        changes_str = "; ".join(personal_changes)
        log_msg = f"{full_name} đã cập nhật thông tin cá nhân cho cộng tác viên mã {new_code} ({changes_str})"
    else:
        log_msg = f"{full_name} cập nhật thành công hồ sơ cho cộng tác viên mã {new_code}"

    logger.info(f"User '{username}' cập nhật CTV '{new_code}': {personal_changes if personal_changes else 'Không đổi thông tin cá nhân'}")

    await record_activity(
        db, action="update_collaborator", result="success", full_name=full_name, username=username,
        message=log_msg,
        employee_code=new_code,
    )
    return _to_response(doc)

@router.delete("/{employee_code}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_collaborator(employee_code: str, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Delete a collaborator profile by employee code."""
    full_name = _actor_name(current_user)
    username = current_user.get("username", "")

    result = await db[COLLECTION].delete_one({"_id": employee_code})
    if result.deleted_count == 0:
        await record_activity(
            db, action="delete_collaborator", result="fail", full_name=full_name, username=username,
            message=f"{full_name} xóa thất bại hồ sơ cộng tác viên mã {employee_code}",
            employee_code=employee_code,
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )

    await record_activity(
        db, action="delete_collaborator", result="success", full_name=full_name, username=username,
        message=f"{full_name} xóa thành công hồ sơ cộng tác viên mã {employee_code}",
        employee_code=employee_code,
    )
    return None

import re

def _extract_employee_code(filename: str) -> str | None:
    """Extract employee code (e.g. CTV001 or 43028) from filename."""
    base_name = os.path.splitext(filename)[0]
    m = re.search(r'(CTV\d+)', base_name, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    matches = re.findall(r'\d{3,}', base_name)
    if matches:
        return matches[0]
    matches = re.findall(r'\d+', base_name)
    return matches[0] if matches else None

class UploadUrlRequest(BaseModel):
    filename: str
    doc_type: str

class UploadConfirmRequest(BaseModel):
    filename: str
    doc_type: str
    employee_code: str
    s3_key: str

@router.post("/documents/upload-url")
async def get_upload_url(
    req: UploadUrlRequest,
    db = Depends(get_db),
    current_user: dict = Depends(get_current_user)
):
    """Generate a presigned S3 URL for uploading a document."""
    if req.doc_type not in ("idCard", "serviceContract", "taxCommitment", "liquidation"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Loại hồ sơ '{req.doc_type}' không hợp lệ"
        )

    employee_code = _extract_employee_code(req.filename)
    if not employee_code:
        return {
            "filename": req.filename,
            "status": "fail",
            "message": f"Không tìm thấy mã cộng tác viên trong tên file '{req.filename}'",
            "employeeCode": None
        }

    collaborator = await db[COLLECTION].find_one({"_id": employee_code})
    if not collaborator:
        return {
            "filename": req.filename,
            "status": "fail",
            "message": f"Không tìm thấy cộng tác viên với mã '{employee_code}'",
            "employeeCode": employee_code
        }

    timestamp = int(datetime.now().timestamp())
    yyyyMM = datetime.now().strftime("%Y%m")
    s3_key = f"documents/contracts/{yyyyMM}/{employee_code}_{req.doc_type}_{timestamp}_{req.filename}"

    try:
        s3 = get_s3_client()
        upload_url = s3.generate_presigned_url(
            ClientMethod="put_object",
            Params={
                "Bucket": settings.S3_BUCKET,
                "Key": s3_key
            },
            ExpiresIn=3600
        )
        return {
            "filename": req.filename,
            "status": "success",
            "uploadUrl": upload_url,
            "s3Key": s3_key,
            "employeeCode": employee_code
        }
    except Exception as e:
        logger.error(f"Lỗi generate presigned URL cho S3: {str(e)}")
        return {
            "filename": req.filename,
            "status": "fail",
            "message": f"Lỗi tạo đường dẫn tải lên: {str(e)}",
            "employeeCode": employee_code
        }

def _extract_file_key(file_obj) -> str | None:
    """Helper to safely extract the S3 key from either a string or a ChecklistFileItem dict."""
    if not file_obj:
        return None
    if isinstance(file_obj, dict):
        return file_obj.get("name")
    if isinstance(file_obj, str):
        return file_obj
    return None

def _get_document_upload_updates(doc_type: str, s3_key: str) -> dict:
    """Helper to return updates dictionary for a document upload based on doc_type."""
    now_str = _now()
    file_item = {"name": s3_key, "updatedDate": now_str}
    if doc_type == "idCard":
        return {
            "$set": {
                "checklist.cccd.file": file_item,
                "checklist.cccd.checked": True,
                "updatedAt": now_str
            }
        }
    elif doc_type == "taxCommitment":
        return {
            "$set": {
                "checklist.ckt.file": file_item,
                "checklist.ckt.checked": True,
                "updatedAt": now_str
            }
        }
    elif doc_type == "liquidation":
        return {
            "$set": {
                "checklist.bbtl.file": file_item,
                "updatedAt": now_str
            }
        }
    elif doc_type == "serviceContract":
        return {
            "$push": {"checklist.hddv.files": file_item},
            "$set": {"updatedAt": now_str}
        }
    return {}

@router.post("/documents/upload-confirm")
async def confirm_upload(
    req: UploadConfirmRequest,
    db = Depends(get_db),
    current_user: dict = Depends(get_current_user)
):
    """Confirm a successful upload and update the collaborator checklist."""
    if req.doc_type not in ("idCard", "serviceContract", "taxCommitment", "liquidation"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Loại hồ sơ '{req.doc_type}' không hợp lệ"
        )

    collaborator = await db[COLLECTION].find_one({"_id": req.employee_code})
    if not collaborator:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Không tìm thấy cộng tác viên với mã '{req.employee_code}'"
        )

    updates = _get_document_upload_updates(req.doc_type, req.s3_key)
    await db[COLLECTION].update_one({"_id": req.employee_code}, updates)

    full_name = _actor_name(current_user)
    username = current_user.get("username", "")
    await record_activity(
        db, action="upload_collaborator_document", result="success", full_name=full_name, username=username,
        message=f"{full_name} đã upload tài liệu {req.doc_type} cho cộng tác viên {req.employee_code}",
        employee_code=req.employee_code,
    )

    return {
        "filename": req.filename,
        "status": "success",
        "message": f"Tải lên và cập nhật hồ sơ CTV {req.employee_code} thành công",
        "employeeCode": req.employee_code
    }

@router.post("/documents/upload")
async def upload_document(
    file: UploadFile = File(...),
    doc_type: str = Form(...),
    db = Depends(get_db),
    current_user: dict = Depends(get_current_user)
):
    """Upload a document (ID Card, Service Contract, Tax Commitment, Liquidation) for a collaborator."""
    if doc_type not in ("idCard", "serviceContract", "taxCommitment", "liquidation"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Loại hồ sơ '{doc_type}' không hợp lệ"
        )

    employee_code = _extract_employee_code(file.filename)
    if not employee_code:
        return {
            "filename": file.filename,
            "status": "fail",
            "message": f"Không tìm thấy mã cộng tác viên trong tên file '{file.filename}'",
            "employeeCode": None
        }

    collaborator = await db[COLLECTION].find_one({"_id": employee_code})
    if not collaborator:
        return {
            "filename": file.filename,
            "status": "fail",
            "message": f"Không tìm thấy cộng tác viên với mã '{employee_code}'",
            "employeeCode": employee_code
        }

    content = await file.read()
    timestamp = int(datetime.now().timestamp())
    yyyyMM = datetime.now().strftime("%Y%m")
    s3_key = f"documents/contracts/{yyyyMM}/{employee_code}_{doc_type}_{timestamp}_{file.filename}"

    try:
        upload_to_s3(content, s3_key)
    except Exception as e:
        return {
            "filename": file.filename,
            "status": "fail",
            "message": f"Lỗi upload S3: {str(e)}",
            "employeeCode": employee_code
        }

    updates = _get_document_upload_updates(doc_type, s3_key)
    await db[COLLECTION].update_one({"_id": employee_code}, updates)

    full_name = _actor_name(current_user)
    username = current_user.get("username", "")
    await record_activity(
        db, action="upload_collaborator_document", result="success", full_name=full_name, username=username,
        message=f"{full_name} đã upload tài liệu {doc_type} cho cộng tác viên {employee_code}",
        employee_code=employee_code,
    )

    return {
        "filename": file.filename,
        "status": "success",
        "message": f"Tải lên và cập nhật hồ sơ CTV {employee_code} thành công",
        "employeeCode": employee_code
    }

@router.get("/{employee_code}/documents/{doc_type}/download")
async def download_collaborator_document(
    employee_code: str,
    doc_type: str,
    file_key: Optional[str] = None,
    db = Depends(get_db),
    current_user: dict = Depends(get_current_user)
):
    """Download a collaborator document from S3 MinIO."""
    if doc_type not in ("idCard", "serviceContract", "taxCommitment", "liquidation"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Loại hồ sơ '{doc_type}' không hợp lệ"
        )

    collaborator = await db[COLLECTION].find_one({"_id": employee_code})
    if not collaborator:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Không tìm thấy cộng tác viên với mã '{employee_code}'"
        )

    checklist = collaborator.get("checklist", {})
    s3_key = None

    if doc_type == "idCard":
        s3_key = _extract_file_key(checklist.get("cccd", {}).get("file"))
    elif doc_type == "taxCommitment":
        s3_key = _extract_file_key(checklist.get("ckt", {}).get("file"))
    elif doc_type == "liquidation":
        s3_key = _extract_file_key(checklist.get("bbtl", {}).get("file"))
    elif doc_type == "serviceContract":
        files = checklist.get("hddv", {}).get("files") or []
        if file_key:
            for f in files:
                f_name = f.get("name") if isinstance(f, dict) else f
                if f_name == file_key:
                    s3_key = file_key
                    break
            if not s3_key:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Tệp tin không thuộc về cộng tác viên này"
                )
        elif files:
            s3_key = _extract_file_key(files[-1])

    if not s3_key:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Không tìm thấy tệp tin {doc_type} cho cộng tác viên này"
        )

    parts = s3_key.split("_", 3)
    original_filename = parts[-1] if len(parts) > 3 else f"{doc_type}_document"

    try:
        s3 = get_s3_client()
        response = s3.get_object(Bucket=settings.S3_BUCKET, Key=s3_key)

        def iter_chunks():
            for chunk in response["Body"].iter_chunks(chunk_size=1024 * 1024):
                yield chunk

        return StreamingResponse(
            iter_chunks(),
            media_type=response.get("ContentType", "application/octet-stream"),
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{quote(original_filename)}"
            }
        )
    except Exception as e:
        logger.error(f"Lỗi khi tải file từ S3 cho CTV={employee_code}, doc_type={doc_type}: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Không thể tải file từ S3 MinIO"
        )


@router.delete("/{employee_code}/documents/{doc_type}")
async def delete_collaborator_document(
    employee_code: str,
    doc_type: str,
    file_key: Optional[str] = None,
    db = Depends(get_db),
    current_user: dict = Depends(get_current_user)
):
    """Delete an uploaded document (such as service contract) from MongoDB and S3."""
    if doc_type not in ("idCard", "serviceContract", "taxCommitment", "liquidation"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Loại hồ sơ '{doc_type}' không hợp lệ"
        )

    collaborator = await db[COLLECTION].find_one({"_id": employee_code})
    if not collaborator:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Không tìm thấy cộng tác viên với mã '{employee_code}'"
        )

    checklist = collaborator.get("checklist", {})
    target_s3_key = None
    update_op = None

    if doc_type == "serviceContract":
        files = checklist.get("hddv", {}).get("files") or []
        target_item = None
        for f in files:
            f_name = f.get("name") if isinstance(f, dict) else f
            if f_name == file_key:
                target_item = f
                break
        if not file_key or target_item is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Không tìm thấy tệp tin hợp đồng dịch vụ cần xóa"
            )
        target_s3_key = file_key
        pull_target = {"name": file_key} if isinstance(target_item, dict) else file_key
        update_op = {
            "$pull": {"checklist.hddv.files": pull_target},
            "$set": {
                "updatedAt": _now()
            }
        }
    elif doc_type == "idCard":
        current_file = _extract_file_key(checklist.get("cccd", {}).get("file"))
        if not current_file:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không có tệp tin CCCD để xóa")
        target_s3_key = current_file
        update_op = {
            "$set": {"checklist.cccd.file": None, "updatedAt": _now()}
        }
    elif doc_type == "taxCommitment":
        current_file = _extract_file_key(checklist.get("ckt", {}).get("file"))
        if not current_file:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không có tệp tin Cam kết thuế để xóa")
        target_s3_key = current_file
        update_op = {
            "$set": {"checklist.ckt.file": None, "updatedAt": _now()}
        }
    elif doc_type == "liquidation":
        current_file = _extract_file_key(checklist.get("bbtl", {}).get("file"))
        if not current_file:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không có tệp tin Biên bản thanh lý để xóa")
        target_s3_key = current_file
        update_op = {
            "$set": {"checklist.bbtl.file": None, "updatedAt": _now()}
        }

    # Atomic update on MongoDB
    if update_op:
        await db[COLLECTION].update_one({"_id": employee_code}, update_op)

    # Delete object from S3 if key exists
    if target_s3_key:
        try:
            s3 = get_s3_client()
            s3.delete_object(Bucket=settings.S3_BUCKET, Key=target_s3_key)
            logger.info(f"Đã xóa file S3 Key={target_s3_key} cho CTV={employee_code}")
        except Exception as e:
            logger.warning(f"Lỗi khi xóa file S3 Key={target_s3_key}: {str(e)}")

    full_name = _actor_name(current_user)
    username = current_user.get("username", "")
    await record_activity(
        db, action="delete_collaborator_document", result="success", full_name=full_name, username=username,
        message=f"{full_name} đã xóa tài liệu {doc_type} ({target_s3_key}) của cộng tác viên {employee_code}",
        employee_code=employee_code,
    )

    return {
        "status": "success",
        "message": f"Xóa tài liệu {doc_type} thành công",
        "employeeCode": employee_code,
        "fileKey": target_s3_key
    }
