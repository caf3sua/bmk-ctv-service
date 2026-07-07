import os
from datetime import datetime, timezone
from io import BytesIO
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, status
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from typing import List
from app.core.database import get_db
from app.core.security import get_current_user
from app.models.collaborator import CollaboratorCreate, CollaboratorUpdate, CollaboratorResponse

router = APIRouter(prefix="/api/collaborators", tags=["Collaborators"])

COLLECTION = "bmk_ctv_collaborators"

SERVICE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE_PATH = os.path.join(SERVICE_ROOT, "templates", "collaborator_checklist_template.xlsx")

# (Tên cột trong file Excel, tên field boolean tương ứng trong checklist)
CHECKLIST_COLUMNS = [
    ("CCCD", "submittedIdCard"),
    ("Cam kết thuế", "submittedTaxCommitment"),
    ("CV", "submittedCV"),
    ("Thông tin cư trú", "submittedResidenceInfo"),
    ("Bằng cấp", "submittedDegree"),
]

def _now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

def _to_response(doc: dict) -> dict:
    doc = dict(doc)
    doc["employeeCode"] = doc["_id"]
    return doc

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
    if not os.path.exists(TEMPLATE_PATH):
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
        checklist = doc.get("checklist")
        if not isinstance(checklist, dict):
            checklist = {}

        contracts = checklist.get("serviceContracts", [])
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
            "Đã nộp" if checklist.get("submittedIdCard") else "",
            "Đã nộp" if checklist.get("submittedTaxCommitment") else "",
            checklist.get("liquidationDate") or "",
            "Đã nộp" if checklist.get("submittedCV") else "",
            "Đã nộp" if checklist.get("submittedResidenceInfo") else "",
            "Đã nộp" if checklist.get("submittedDegree") else "",
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border

            # Alignment formatting
            if col_idx in [1, 2, 4, 5, 6, 8, 10, 11, 12, 13, 14, 15, 16, 17]:
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

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=danh_sach_ctv.xlsx"},
    )

@router.post("/import")
async def import_collaborators(
    file: UploadFile = File(...), db=Depends(get_db), current_user: dict = Depends(get_current_user)
):
    """Bulk-update collaborators' checklist status from an uploaded Excel file (see /template)."""
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Chỉ hỗ trợ file Excel (.xlsx)")

    content = await file.read()
    try:
        wb = load_workbook(BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không đọc được file Excel, vui lòng dùng đúng file mẫu"
        )

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="File không có dữ liệu")

    header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    if "Mã nhân viên" not in header:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không tìm thấy cột 'Mã nhân viên' trong file, vui lòng dùng đúng file mẫu"
        )
    code_idx = header.index("Mã nhân viên")

    column_indices = {field: header.index(label) for label, field in CHECKLIST_COLUMNS if label in header}

    updated: List[str] = []
    not_found: List[str] = []

    for row in rows[1:]:
        if row is None or all(cell in (None, "") for cell in row):
            continue

        raw_code = row[code_idx] if code_idx < len(row) else None
        employee_code = str(raw_code).strip() if raw_code is not None else ""
        if not employee_code:
            continue

        existing = await db[COLLECTION].find_one({"_id": employee_code})
        if not existing:
            not_found.append(employee_code)
            continue

        updates = {}
        for field, col_idx in column_indices.items():
            value = row[col_idx] if col_idx < len(row) else None
            updates[f"checklist.{field}"] = value is not None and str(value).strip() != ""
        updates["updatedAt"] = _now()

        await db[COLLECTION].update_one({"_id": employee_code}, {"$set": updates})
        updated.append(employee_code)

    return {
        "updatedCount": len(updated),
        "updated": updated,
        "notFoundCount": len(not_found),
        "notFound": not_found,
    }

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
    employee_code = payload.employeeCode.strip()
    if not employee_code:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Mã nhân viên là bắt buộc")

    existing = await db[COLLECTION].find_one({"_id": employee_code})
    if existing:
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
    return _to_response(doc)

@router.put("/{employee_code}", response_model=CollaboratorResponse)
async def update_collaborator(employee_code: str, payload: CollaboratorUpdate, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Update an existing collaborator profile."""
    existing = await db[COLLECTION].find_one({"_id": employee_code})
    if not existing:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )

    new_code = payload.employeeCode.strip() or employee_code
    if new_code != employee_code:
        conflict = await db[COLLECTION].find_one({"_id": new_code})
        if conflict:
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
    return _to_response(doc)

@router.delete("/{employee_code}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_collaborator(employee_code: str, db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Delete a collaborator profile by employee code."""
    result = await db[COLLECTION].delete_one({"_id": employee_code})
    if result.deleted_count == 0:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy cộng tác viên "{employee_code}"'
        )
    return None
