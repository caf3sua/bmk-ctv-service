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

def _cell_str(row, idx) -> str:
    if idx is None or idx >= len(row):
        return ""
    value = row[idx]
    return "" if value is None else str(value).strip()

def _to_response(doc: dict) -> dict:
    doc = dict(doc)
    doc["employeeCode"] = doc["_id"]
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
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border

            # Alignment formatting
            if col_idx in [1, 2, 4, 5, 6, 8, 10, 11, 12, 13, 14]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        row_idx += 1
        stt += 1

    # Clear remaining rows in the template if they are beyond our data rows
    for r in range(row_idx, ws.max_row + 1):
        for c in range(1, 15):
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

        if EMPLOYEE_CODE_LABEL not in header:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Không tìm thấy cột 'Mã nhân viên' trong file, vui lòng dùng đúng file mẫu",
            )
        code_idx = header.index(EMPLOYEE_CODE_LABEL)
        start_date_idx = header.index(START_DATE_LABEL) if START_DATE_LABEL in header else None
        end_date_idx = header.index(END_DATE_LABEL) if END_DATE_LABEL in header else None
        liquidation_idx = header.index(LIQUIDATION_DATE_LABEL) if LIQUIDATION_DATE_LABEL in header else None

        column_indices = {field: header.index(label) for label, field in CHECKLIST_COLUMNS if label in header}
        profile_indices = {field: header.index(label) for label, field in PROFILE_TEXT_COLUMNS if label in header}
        dob_idx = header.index(DOB_LABEL) if DOB_LABEL in header else None

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
                    hddv = existing.get("checklist", {}).get("hddv") or {}
                    contracts = hddv.get("contract_date") or []
                    last_contract = dict(contracts[-1]) if contracts and isinstance(contracts[-1], dict) else {}
                    if start_val is not _UNSET:
                        if start_val is not None or not last_contract.get("startDate"):
                            last_contract["startDate"] = start_val
                    if end_val is not _UNSET:
                        if end_val is not None or not last_contract.get("endDate"):
                            last_contract["endDate"] = end_val
                    updates["checklist.hddv.contract_date"] = (contracts[:-1] if contracts else []) + [last_contract]

                if liquidation_val is not _UNSET:
                    bbtl = existing.get("checklist", {}).get("bbtl") or {}
                    if liquidation_val is not None or not bbtl.get("date"):
                        updates["checklist.bbtl.date"] = liquidation_val

                # Cập nhật thông tin hồ sơ nếu có giá trị mới trong dòng hiện tại
                for label, field in PROFILE_TEXT_COLUMNS:
                    idx = profile_indices.get(field)
                    val = _cell_str(row, idx)
                    if val:
                        updates[field] = val

                if dob_idx is not None:
                    dob_val = parse_cell_date(dob_idx)
                    if dob_val is not _UNSET and dob_val is not None:
                        updates["dob"] = dob_val

                updates["updatedAt"] = _now()
                await db[COLLECTION].update_one({"_id": employee_code}, {"$set": updates})
                updated.append(employee_code)
            else:
                now_ts = _now()
                dob_val = parse_cell_date(dob_idx)
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
        )
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f'Mã nhân viên "{employee_code}" đã tồn tại'
        )

    now = _now()
    doc = payload.model_dump()
    doc["employeeCode"] = employee_code
    doc["_id"] = employee_code
    doc["createdAt"] = now
    doc["updatedAt"] = now
    await db[COLLECTION].insert_one(doc)
    await record_activity(
        db, action="create_collaborator", result="success", full_name=full_name,
        username=current_user.get("username", ""),
        message=f"{full_name} tạo thành công hồ sơ cho cộng tác viên mã {employee_code}",
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
            )
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f'Mã nhân viên "{new_code}" đã tồn tại'
            )

    doc = payload.model_dump()
    doc["employeeCode"] = new_code
    doc["_id"] = new_code
    doc["createdAt"] = existing["createdAt"]
    doc["updatedAt"] = _now()

    if new_code != employee_code:
        await db[COLLECTION].delete_one({"_id": employee_code})
    await db[COLLECTION].replace_one({"_id": new_code}, doc, upsert=True)
    await record_activity(
        db, action="update_collaborator", result="success", full_name=full_name, username=username,
        message=f"{full_name} cập nhật thành công hồ sơ cho cộng tác viên mã {new_code}",
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
        )
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )

    await record_activity(
        db, action="delete_collaborator", result="success", full_name=full_name, username=username,
        message=f"{full_name} xóa thành công hồ sơ cộng tác viên mã {employee_code}",
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
        return matches[-1]
    matches = re.findall(r'\d+', base_name)
    return matches[-1] if matches else None

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

def _get_document_upload_updates(doc_type: str, s3_key: str) -> dict:
    """Helper to return updates dictionary for a document upload based on doc_type."""
    if doc_type == "idCard":
        return {"$set": {"checklist.cccd.file": s3_key, "checklist.cccd.checked": True, "updatedAt": _now()}}
    elif doc_type == "taxCommitment":
        return {"$set": {"checklist.ckt.file": s3_key, "checklist.ckt.checked": True, "updatedAt": _now()}}
    elif doc_type == "liquidation":
        return {"$set": {"checklist.bbtl.file": s3_key, "updatedAt": _now()}}
    elif doc_type == "serviceContract":
        return {
            "$push": {"checklist.hddv.files": s3_key},
            "$set": {"updatedAt": _now()}
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
        s3_key = checklist.get("cccd", {}).get("file")
    elif doc_type == "taxCommitment":
        s3_key = checklist.get("ckt", {}).get("file")
    elif doc_type == "liquidation":
        s3_key = checklist.get("bbtl", {}).get("file")
    elif doc_type == "serviceContract":
        files = checklist.get("hddv", {}).get("files") or []
        if file_key:
            if file_key in files:
                s3_key = file_key
            else:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Tệp tin không thuộc về cộng tác viên này"
                )
        elif files:
            s3_key = files[-1]

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
