import math
import re
from datetime import datetime, timezone, date
from io import BytesIO
from typing import List, Optional
from fastapi import APIRouter, Depends, File, HTTPException, Query, UploadFile, status
from openpyxl import load_workbook
from pymongo import UpdateOne

from app.core.database import get_db
from app.core.security import get_current_user
from app.models.reconciliation import (
    ReconciliationRecordCreate,
    ReconciliationRecordUpdate,
    ReconciliationRecordResponse,
    ReconciliationListResponse,
    ImportHrTpBankResult,
    BmkSystemInfo,
    TpBankInfo,
    BmkHrInfo,
)

router = APIRouter(prefix="/api/reconciliation", tags=["Reconciliation"])

COLLECTION = "bmk_ctv_reconciliations"
COLLABORATORS_COLLECTION = "bmk_ctv_collaborators"

def _now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

def _to_response(doc: dict) -> dict:
    created_source = doc.get("createdSource") or doc.get("source") or "bmk_system"
    tpbank_info = doc.get("tpbankInfo") or {"contracts": []}
    if isinstance(tpbank_info, dict):
        tpbank_info.setdefault("contracts", [])

    return {
        "id": str(doc["_id"]),
        "employeeCode": doc.get("employeeCode", ""),
        "fullName": doc.get("fullName", ""),
        "idNumber": doc.get("idNumber") or doc.get("cccd"),
        "createdSource": created_source,
        "isBmkSystemExist": doc.get("isBmkSystemExist", False),
        "isSynced": doc.get("isSynced", False),
        "result": doc.get("result"),
        "departmentLevel1": doc.get("departmentLevel1"),
        "position": doc.get("position"),
        "employmentStatus": doc.get("employmentStatus"),
        "onboardDate": doc.get("onboardDate"),
        "offboardDate": doc.get("offboardDate"),
        "tpbankInfo": tpbank_info,
        "bmkHrInfo": doc.get("bmkHrInfo") or {
            "contractCount": 0,
            "idCardCount": 0,
            "liquidationCount": 0,
            "taxCommitmentCount": 0,
        },
        "bmkSystemInfo": doc.get("bmkSystemInfo") or {
            "contractCount": 0,
            "idCardCount": 0,
            "liquidationCount": 0,
            "taxCommitmentCount": 0,
        },
        "reconciliationStatus": doc.get("reconciliationStatus", "pending"),
        "createdAt": doc.get("createdAt", _now()),
        "updatedAt": doc.get("updatedAt", _now()),
    }


def _calculate_system_info(collab_doc: dict) -> dict:
    checklist = collab_doc.get("checklist") or {}
    
    # CCCD: 1 nếu cccd.checked = true hoặc cccd.file khác null
    cccd = checklist.get("cccd") or {}
    id_card_count = 1 if (cccd.get("checked") or cccd.get("file")) else 0
    
    # Cam kết thuế: 1 nếu ckt.checked = true hoặc ckt.file khác null
    ckt = checklist.get("ckt") or {}
    tax_count = 1 if (ckt.get("checked") or ckt.get("file")) else 0
    
    # Biên bản thanh lý: 1 nếu bbtl.date != null hoặc bbtl.file khác null
    bbtl = checklist.get("bbtl") or {}
    liquidation_count = 1 if (bbtl.get("date") or bbtl.get("file")) else 0
    
    # Hợp đồng dịch vụ: count từ hddv.contract_date (khởi tạo với startDate & endDate khác null)
    hddv = checklist.get("hddv") or {}
    contract_dates = hddv.get("contract_date") or []
    contract_count = 0
    for d in contract_dates:
        if isinstance(d, dict):
            start_d = d.get("startDate")
            end_d = d.get("endDate")
            if start_d and end_d:
                contract_count += 1

    return {
        "contractCount": contract_count,
        "idCardCount": id_card_count,
        "liquidationCount": liquidation_count,
        "taxCommitmentCount": tax_count,
    }

@router.get("", response_model=ReconciliationListResponse)
async def list_reconciliation_records(
    keyword: Optional[str] = Query(None, description="Tìm theo mã CTV, họ tên, đơn vị"),
    employment_status: Optional[str] = Query(None, description="Lọc theo tình trạng nhân sự"),
    reconciliation_status: Optional[str] = Query(None, description="Lọc theo trạng thái đối soát"),
    created_source: Optional[str] = Query(None, description="Lọc theo nguồn tạo (bmk_system, bmk_hr, tpbank)"),
    is_synced: Optional[str] = Query(None, description="Lọc theo trạng thái đối soát (synced, unsynced)"),
    result_status: Optional[str] = Query(None, description="Lọc theo kết quả đối soát (all, match_all, mismatch_contract, mismatch_idcard, mismatch_liquidation)"),
    page: int = Query(1, ge=1, description="Số trang"),
    page_size: int = Query(20, ge=1, le=200, description="Số dòng mỗi trang"),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    conditions = []
    if keyword and keyword.strip():
        term = keyword.strip()
        conditions.append({
            "$or": [
                {"employeeCode": {"$regex": re.escape(term), "$options": "i"}},
                {"fullName": {"$regex": re.escape(term), "$options": "i"}},
                {"idNumber": {"$regex": re.escape(term), "$options": "i"}},
                {"departmentLevel1": {"$regex": re.escape(term), "$options": "i"}},
                {"position": {"$regex": re.escape(term), "$options": "i"}},
            ]
        })

    if employment_status and employment_status.strip() and employment_status.strip() != "all":
        conditions.append({"employmentStatus": employment_status.strip()})

    if reconciliation_status and reconciliation_status.strip() and reconciliation_status.strip() != "all":
        conditions.append({"reconciliationStatus": reconciliation_status.strip()})

    if is_synced and is_synced.strip() and is_synced.strip() != "all":
        val = is_synced.strip().lower()
        if val in ["synced", "true", "1"]:
            conditions.append({"isSynced": True})
        elif val in ["unsynced", "false", "0"]:
            conditions.append({
                "$or": [
                    {"isSynced": False},
                    {"isSynced": {"$exists": False}},
                ]
            })

    if result_status and result_status.strip() and result_status.strip() != "all":
        r_val = result_status.strip().lower()
        if r_val in ["match_all", "success_all"]:
            conditions.append({
                "result.contract": "success",
                "result.idCard": "success",
                "result.liquidation": "success",
            })
        elif r_val in ["warn_bank_contract", "warn_contract", "warn", "mismatch_bank"]:
            conditions.append({"result.contract": "warn"})
        elif r_val in ["mismatch_contract", "failed_contract"]:
            conditions.append({"result.contract": "failed"})
        elif r_val in ["mismatch_idcard", "failed_idcard"]:
            conditions.append({"result.idCard": "failed"})
        elif r_val in ["mismatch_liquidation", "failed_liquidation"]:
            conditions.append({"result.liquidation": "failed"})

    if created_source and created_source.strip() and created_source.strip() != "all":
        src = created_source.strip()
        if src == "bmk_system":
            conditions.append({
                "$or": [
                    {"createdSource": "bmk_system"},
                    {"createdSource": "BMK System"},
                    {"createdSource": {"$exists": False}},
                ]
            })
        elif src == "bmk_hr":
            conditions.append({
                "$or": [
                    {"createdSource": "bmk_hr"},
                    {"createdSource": "HR BMK"},
                ]
            })
        elif src in ["hr_tpbank", "HR TP Bank", "HR Tp bank"]:
            conditions.append({
                "$or": [
                    {"createdSource": "hr_tpbank"},
                    {"createdSource": "HR TP Bank"},
                    {"createdSource": "HR Tp bank"},
                ]
            })
        elif src == "tpbank":
            conditions.append({
                "$or": [
                    {"createdSource": "tpbank"},
                    {"createdSource": "TP Bank"},
                ]
            })
        else:
            conditions.append({"createdSource": src})

    query = {"$and": conditions} if conditions else {}

    total = await db[COLLECTION].count_documents(query)
    total_pages = max(1, math.ceil(total / page_size))
    skip = (page - 1) * page_size

    cursor = db[COLLECTION].find(query).sort("updatedAt", -1).skip(skip).limit(page_size)
    items = []
    async for doc in cursor:
        items.append(_to_response(doc))

    return {
        "items": items,
        "total": total,
        "page": page,
        "pageSize": page_size,
        "totalPages": total_pages,
    }

@router.get("/{employee_code}", response_model=ReconciliationRecordResponse)
async def get_reconciliation_record(
    employee_code: str,
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Lấy thông tin đối soát của 1 cộng tác viên."""
    doc = await db[COLLECTION].find_one({"employeeCode": employee_code.strip()})
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy thông tin đối soát cho CTV mã "{employee_code}"'
        )
    return _to_response(doc)

@router.post("/sync-system-info")
async def sync_bmk_system_info(
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """
    Đồng bộ thông tin số lượng hồ sơ từ collection bmk_ctv_collaborators
    vào trường bmkSystemInfo, createdSource="BMK System" và isBmkSystemExist=True của collection bmk_ctv_reconciliations.
    """
    collab_cursor = db[COLLABORATORS_COLLECTION].find({})
    operations = []
    total_processed = 0
    synced_codes = set()
    now = _now()

    async for collab in collab_cursor:
        code = collab.get("employeeCode") or str(collab.get("_id"))
        if not code:
            continue
        code = str(code).strip()
        synced_codes.add(code)

        system_info = _calculate_system_info(collab)
        full_name = collab.get("fullName", "")
        id_number = collab.get("idNumber") or collab.get("cccd") or None

        operations.append(
            UpdateOne(
                {"employeeCode": code},
                {
                    "$set": {
                        "idNumber": id_number,
                        "bmkSystemInfo": system_info,
                        "createdSource": "BMK System",
                        "isBmkSystemExist": True,
                        "updatedAt": now,
                    },
                    "$setOnInsert": {
                        "_id": code,
                        "employeeCode": code,
                        "fullName": full_name,
                        "departmentLevel1": None,
                        "position": None,
                        "employmentStatus": None,
                        "onboardDate": None,
                        "offboardDate": None,
                        "tpbankInfo": {"contracts": []},
                        "bmkHrInfo": {
                            "contractCount": 0,
                            "idCardCount": 0,
                            "liquidationCount": 0,
                            "taxCommitmentCount": 0,
                        },
                        "reconciliationStatus": "pending",
                        "createdAt": now,
                    },
                },
                upsert=True,
            )
        )

        total_processed += 1

        if len(operations) >= 500:
            await db[COLLECTION].bulk_write(operations, ordered=False)
            operations = []

    if operations:
        await db[COLLECTION].bulk_write(operations, ordered=False)

    # Đánh dấu isBmkSystemExist = False cho các bản ghi trong bmk_ctv_reconciliations không nằm trong bmk_ctv_collaborators
    if synced_codes:
        await db[COLLECTION].update_many(
            {"employeeCode": {"$nin": list(synced_codes)}},
            {"$set": {"isBmkSystemExist": False}}
        )

    return {
        "status": "success",
        "message": f"Đã đồng bộ thành công {total_processed} cộng tác viên từ hệ thống BMK",
        "totalProcessed": total_processed,
    }



@router.post("/import-hr-bmk")
async def import_hr_bmk_data(
    file: UploadFile = File(...),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """
    Import/Update dữ liệu từ file Excel HR BMK vào collection bmk_ctv_reconciliations.
    """
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

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File Excel không chứa dữ liệu",
        )

    # Tìm dòng tiêu đề chứa 'MNV' hoặc 'MÃ NV'
    header_row_idx = None
    header = []
    for idx, r in enumerate(rows):
        if r and any(cell and ("MNV" in str(cell).upper() or "MÃ NV" in str(cell).upper() or "MÃ NHÂN VIÊN" in str(cell).upper()) for cell in r):
            header_row_idx = idx
            header = [str(c).strip() if c is not None else "" for c in r]
            break

    if header_row_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không tìm thấy dòng tiêu đề chứa cột 'MNV' trong file Excel",
        )

    header_upper = [h.upper() for h in header]

    def find_col_idx(candidates: List[str]) -> Optional[int]:
        for candidate in candidates:
            cand_u = candidate.upper()
            for idx, h in enumerate(header_upper):
                if cand_u in h:
                    return idx
        return None

    code_idx = find_col_idx(["MNV", "MÃ NV", "MÃ NHÂN VIÊN"])
    name_idx = find_col_idx(["HỌ VÀ TÊN", "HỌ TÊN"])
    cccd_idx = find_col_idx(["SỐ CCCD", "CCCD"])
    contract_idx = find_col_idx(["HỢP ĐỒNG", "SL HĐ"])
    
    # Cột CCCD đếm số lượng (thường đứng sau HỢP ĐỒNG)
    id_card_cnt_idx = None
    for idx, h in enumerate(header_upper):
        if "CCCD" in h and idx != cccd_idx:
            id_card_cnt_idx = idx
            break
    if id_card_cnt_idx is None:
        id_card_cnt_idx = cccd_idx

    bbtl_idx = find_col_idx(["BBTL", "BIÊN BẢN THANH LÝ"])
    tax_idx = find_col_idx(["CAM KẾT THUẾ", "CKT"])

    if code_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File thiếu cột 'MNV' (Mã nhân viên)",
        )

    operations = []
    now = _now()

    def parse_int(val) -> int:
        if val is None or val == "":
            return 0
        try:
            return int(float(str(val).strip()))
        except (ValueError, TypeError):
            return 0

    def parse_str(val) -> Optional[str]:
        if val is None:
            return None
        s = str(val).strip()
        if not s or s == "-":
            return None
        return s

    created_count = 0
    updated_count = 0
    total_processed = 0

    # Lấy danh sách các employeeCode đã tồn tại để tránh query 3,000 lần
    existing_codes = set()
    async for doc in db[COLLECTION].find({}, {"employeeCode": 1}):
        code = doc.get("employeeCode")
        if code:
            existing_codes.add(str(code).strip())

    for row in rows[header_row_idx + 1 :]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        raw_code = row[code_idx] if code_idx < len(row) else None
        employee_code = parse_str(raw_code)
        if not employee_code:
            continue

        if employee_code.isdigit() and len(employee_code) < 5:
            employee_code = employee_code.zfill(5)

        full_name = parse_str(row[name_idx] if name_idx is not None and name_idx < len(row) else None) or ""
        id_number = parse_str(row[cccd_idx] if cccd_idx is not None and cccd_idx < len(row) else None)

        contract_cnt = parse_int(row[contract_idx] if contract_idx is not None and contract_idx < len(row) else None)
        id_card_cnt = parse_int(row[id_card_cnt_idx] if id_card_cnt_idx is not None and id_card_cnt_idx < len(row) else None)
        liquidation_cnt = parse_int(row[bbtl_idx] if bbtl_idx is not None and bbtl_idx < len(row) else None)
        tax_cnt = parse_int(row[tax_idx] if tax_idx is not None and tax_idx < len(row) else None)

        bmk_hr_info = {
            "contractCount": contract_cnt,
            "idCardCount": id_card_cnt,
            "liquidationCount": liquidation_cnt,
            "taxCommitmentCount": tax_cnt,
        }

        if employee_code in existing_codes:
            updated_count += 1
        else:
            created_count += 1
            existing_codes.add(employee_code)

        operations.append(
            UpdateOne(
                {"employeeCode": employee_code},
                {
                    "$set": {
                        "idNumber": id_number,
                        "bmkHrInfo": bmk_hr_info,
                        "createdSource": "HR BMK",
                        "updatedAt": now,
                    },
                    "$setOnInsert": {
                        "_id": employee_code,
                        "employeeCode": employee_code,
                        "fullName": full_name,
                        "departmentLevel1": None,
                        "position": None,
                        "employmentStatus": None,
                        "onboardDate": None,
                        "offboardDate": None,
                        "tpbankInfo": {"contracts": []},
                        "bmkSystemInfo": {
                            "contractCount": 0,
                            "idCardCount": 0,
                            "liquidationCount": 0,
                            "taxCommitmentCount": 0,
                        },
                        "reconciliationStatus": "pending",
                        "createdAt": now,
                    },
                },
                upsert=True,
            )
        )

        total_processed += 1

        if len(operations) >= 500:
            await db[COLLECTION].bulk_write(operations, ordered=False)
            operations = []

    if operations:
        await db[COLLECTION].bulk_write(operations, ordered=False)

    return {
        "status": "success",
        "message": f"Đã nhập thành công {total_processed} cộng tác viên từ HR BMK (Tạo mới: {created_count}, Cập nhật: {updated_count})",
        "totalProcessed": total_processed,
        "createdCount": created_count,
        "updatedCount": updated_count,
    }


@router.post("/import-hr-tpbank", response_model=ImportHrTpBankResult)
async def import_hr_tpbank_contracts(
    file: UploadFile = File(...),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """
    Import số lượng Hợp đồng từ file Excel HR TP Bank (template_ctv_hr_tpbank.xlsx).
    Cột mẫu: STT, Mã NV, Họ và tên, Số lượng hợp đồng.
    - Lưu vào TpBankInfo.hrContractCount
    - Nếu CTV chưa tồn tại: tạo mới với employeeCode, fullName, TpBankInfo.hrContractCount, createdSource="HR TP Bank"
    - Nếu CTV đã tồn tại: cập nhật TpBankInfo.hrContractCount
    """
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

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File Excel không chứa dữ liệu",
        )

    # Tìm dòng tiêu đề chứa 'MNV', 'MÃ NV', hoặc 'MÃ NHÂN VIÊN'
    header_row_idx = None
    header = []
    for idx, r in enumerate(rows):
        if r and any(cell and ("MNV" in str(cell).upper() or "MÃ NV" in str(cell).upper() or "MÃ NHÂN VIÊN" in str(cell).upper()) for cell in r):
            header_row_idx = idx
            header = [str(c).strip() if c is not None else "" for c in r]
            break

    if header_row_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không tìm thấy dòng tiêu đề chứa cột 'Mã NV' trong file Excel",
        )

    header_upper = [h.upper() for h in header]

    def find_col_idx(candidates: List[str]) -> Optional[int]:
        for candidate in candidates:
            cand_u = candidate.upper()
            for idx, h in enumerate(header_upper):
                if cand_u in h:
                    return idx
        return None

    code_idx = find_col_idx(["MÃ NV", "MNV", "MÃ NHÂN VIÊN"])
    name_idx = find_col_idx(["HỌ VÀ TÊN", "HỌ TÊN", "TÊN"])
    contract_idx = find_col_idx(["SỐ LƯỢNG HỢP ĐỒNG", "SL HỢP ĐỒNG", "SL HĐ", "HỢP ĐỒNG"])

    if code_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File thiếu cột 'Mã NV'",
        )

    def parse_int(val) -> Optional[int]:
        if val is None or str(val).strip() in ["", "-", "None", "null"]:
            return None
        try:
            return int(float(str(val).strip()))
        except (ValueError, TypeError):
            return None

    def parse_str(val) -> Optional[str]:
        if val is None:
            return None
        s = str(val).strip()
        if not s or s == "-":
            return None
        return s

    operations = []
    now = _now()
    created_count = 0
    updated_count = 0
    total_processed = 0

    existing_docs = {}
    async for doc in db[COLLECTION].find({}, {
        "employeeCode": 1,
        "fullName": 1,
        "tpbankInfo": 1,
        "bmkHrInfo": 1,
        "bmkSystemInfo": 1,
        "employmentStatus": 1,
        "isSynced": 1,
        "result": 1,
    }):
        code = doc.get("employeeCode")
        if code:
            existing_docs[str(code).strip()] = doc

    for row in rows[header_row_idx + 1:]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        raw_code = row[code_idx] if code_idx < len(row) else None
        employee_code = parse_str(raw_code)
        if not employee_code:
            continue

        if employee_code.isdigit() and len(employee_code) < 5:
            employee_code = employee_code.zfill(5)

        full_name = parse_str(row[name_idx] if name_idx is not None and name_idx < len(row) else None) or ""
        contract_cnt = parse_int(row[contract_idx] if contract_idx is not None and contract_idx < len(row) else None)

        total_processed += 1

        if employee_code in existing_docs:
            updated_count += 1
            ex_doc = existing_docs[employee_code]
            update_fields = {
                "tpbankInfo.hrContractCount": contract_cnt,
                "updatedAt": now,
            }
            if ex_doc.get("isSynced") or ex_doc.get("result"):
                est_cnt = (ex_doc.get("tpbankInfo") or {}).get("estimatedContractCount")
                bmk_hr = ex_doc.get("bmkHrInfo") or {}
                bmk_sys = ex_doc.get("bmkSystemInfo") or {}
                emp_status = ex_doc.get("employmentStatus")
                result_obj = _calculate_reconciliation_result(
                    estimated_contract_count=est_cnt,
                    hr_contract_count=contract_cnt,
                    bmk_hr_info=bmk_hr,
                    bmk_system_info=bmk_sys,
                    employment_status=emp_status,
                )
                update_fields["result"] = result_obj
            if full_name:
                operations.append(
                    UpdateOne(
                        {"employeeCode": employee_code, "$or": [{"fullName": ""}, {"fullName": None}, {"fullName": {"$exists": False}}]},
                        {"$set": {"fullName": full_name}}
                    )
                )
            operations.append(
                UpdateOne(
                    {"employeeCode": employee_code},
                    {"$set": update_fields}
                )
            )
        else:
            created_count += 1
            existing_codes.add(employee_code)
            new_doc = {
                "_id": employee_code,
                "employeeCode": employee_code,
                "fullName": full_name,
                "createdSource": "HR TP Bank",
                "departmentLevel1": None,
                "position": None,
                "employmentStatus": None,
                "onboardDate": None,
                "offboardDate": None,
                "tpbankInfo": {
                    "estimatedContractCount": None,
                    "hrContractCount": contract_cnt,
                    "contracts": [],
                },
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
                "reconciliationStatus": "pending",
                "createdAt": now,
                "updatedAt": now,
            }
            operations.append(
                UpdateOne(
                    {"employeeCode": employee_code},
                    {"$setOnInsert": new_doc},
                    upsert=True,
                )
            )

        if len(operations) >= 500:
            await db[COLLECTION].bulk_write(operations, ordered=False)
            operations = []

    if operations:
        await db[COLLECTION].bulk_write(operations, ordered=False)

    return {
        "status": "success",
        "message": f"Đã nhập thành công số lượng HĐ TP Bank: {total_processed} CTV (Tạo mới: {created_count}, Cập nhật: {updated_count})",
        "totalProcessed": total_processed,
        "createdCount": created_count,
        "updatedCount": updated_count,
    }


def _parse_excel_date(val) -> tuple[Optional[str], Optional[datetime]]:
    if val is None:
        return None, None
    if isinstance(val, (datetime, date)):
        if val.year == 1900:
            return None, None
        dt = datetime(val.year, val.month, val.day) if isinstance(val, date) and not isinstance(val, datetime) else val
        return dt.strftime("%Y-%m-%d"), dt
    
    s = str(val).strip()
    if not s or s.upper() in ["N/A", "N/A", "NONE", "-", "NULL"] or "1900" in s:
        return None, None
    
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
        try:
            dt = datetime.strptime(s, fmt)
            if dt.year == 1900:
                return None, None
            return dt.strftime("%Y-%m-%d"), dt
        except ValueError:
            pass
    return None, None


def _calculate_tpbank_contract_count(onboard_dt: Optional[datetime], offboard_dt: Optional[datetime]) -> Optional[int]:
    if not onboard_dt:
        return None
    end_dt = offboard_dt if offboard_dt else datetime.now()
    if end_dt < onboard_dt:
        return 0

    full_months = (end_dt.year - onboard_dt.year) * 12 + (end_dt.month - onboard_dt.month)
    if end_dt.day < onboard_dt.day:
        full_months -= 1

    if full_months < 0:
        full_months = 0

    return (full_months // 6) + 1


def _calculate_reconciliation_result(
    estimated_contract_count: Optional[int],
    hr_contract_count: Optional[int],
    bmk_hr_info: dict,
    bmk_system_info: dict,
    employment_status: Optional[str],
) -> dict:
    sys_contract = (bmk_system_info or {}).get("contractCount", 0) or 0
    hr_contract = (bmk_hr_info or {}).get("contractCount", 0) or 0
    total_bmk_contract = sys_contract + hr_contract

    # 1. Đầu tiên kiểm tra nếu TpBankInfo.estimatedContractCount != TpBankInfo.hrContractCount -> cảnh báo Lệch Bank
    if estimated_contract_count != hr_contract_count:
        contract_res = "warn"
    else:
        # 2. Nếu TpBankInfo.estimatedContractCount == TpBankInfo.hrContractCount thì thực hiện tiếp logic hiện tại
        if estimated_contract_count is not None and estimated_contract_count <= total_bmk_contract:
            contract_res = "success"
        else:
            contract_res = "failed"

    # Đối soát idCard: nếu bmkSystemInfo.idCardCount = 1 hoặc bmkHrInfo.idCardCount = 1 -> success còn lại failed
    sys_idcard = (bmk_system_info or {}).get("idCardCount", 0) or 0
    hr_idcard = (bmk_hr_info or {}).get("idCardCount", 0) or 0
    if sys_idcard == 1 or hr_idcard == 1:
        idcard_res = "success"
    else:
        idcard_res = "failed"

    # Đối soát liquidation: nếu nhân viên đang hiệu lực (ngày nghỉ việc null) thì success, nhân viên nghỉ việc thì bmkSystemInfo.liquidationCount = 1 hoặc bmkHrInfo.liquidationCount = 1 thì success còn lại failed
    sys_liq = (bmk_system_info or {}).get("liquidationCount", 0) or 0
    hr_liq = (bmk_hr_info or {}).get("liquidationCount", 0) or 0
    if employment_status == "Nghỉ việc":
        if sys_liq == 1 or hr_liq == 1:
            liq_res = "success"
        else:
            liq_res = "failed"
    else:
        liq_res = "success"

    sys_tax = (bmk_system_info or {}).get("taxCommitmentCount", 0) or 0
    hr_tax = (bmk_hr_info or {}).get("taxCommitmentCount", 0) or 0
    tax_res = "success" if (sys_tax == 1 or hr_tax == 1) else "failed"

    return {
        "contract": contract_res,
        "idCard": idcard_res,
        "liquidation": liq_res,
        "taxCommitment": tax_res,
    }


@router.post("/reconcile-tpbank")
async def reconcile_tpbank_data(
    file: UploadFile = File(...),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """
    Upload file Đối soát TP Bank, thực hiện đối soát và cập nhật collection bmk_ctv_reconciliations.
    1. Tại thời điểm upload update tất cả row có isSynced = False
    2. Đọc MNV, HỌ VÀ TÊN, NGÀY GIA NHẬP, NGÀY THÔI VIỆC
    3. Nếu chưa có -> tạo mới với createdSource="TP Bank", isSynced=True
    4. Nếu đã có -> update NGÀY GIA NHẬP, NGÀY THÔI VIỆC, isSynced=True
    5. Tính số HĐ dự tính lưu vào tpbankInfo.estimatedContractCount
    """
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Chỉ hỗ trợ file Excel (.xlsx, .xlsm)",
        )

    # 1. Update tất cả bản ghi hiện tại thành isSynced = False và result = None
    await db[COLLECTION].update_many({}, {"$set": {"isSynced": False, "result": None}})

    content = await file.read()
    try:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Không đọc được file Excel: {str(e)}",
        )

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File Excel không chứa dữ liệu",
        )

    header_row_idx = None
    header = []
    for idx, r in enumerate(rows):
        if r and any(cell and ("MNV" in str(cell).upper() or "MÃ NV" in str(cell).upper() or "MÃ NHÂN VIÊN" in str(cell).upper()) for cell in r):
            header_row_idx = idx
            header = [str(c).strip() if c is not None else "" for c in r]
            break

    if header_row_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không tìm thấy dòng tiêu đề chứa cột 'MNV' trong file Excel",
        )

    header_upper = [h.upper() for h in header]

    def find_col_idx(candidates: List[str]) -> Optional[int]:
        for candidate in candidates:
            cand_u = candidate.upper()
            for idx, h in enumerate(header_upper):
                if cand_u in h:
                    return idx
        return None

    code_idx = find_col_idx(["MNV", "MÃ NV", "MÃ NHÂN VIÊN"])
    name_idx = find_col_idx(["HỌ VÀ TÊN", "HỌ TÊN"])
    onboard_idx = find_col_idx(["NGÀY GIA NHẬP", "NGÀY VÀO"])
    offboard_idx = find_col_idx(["NGÀY THÔI VIỆC", "NGÀY NGHỈ"])

    if code_idx is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File thiếu cột 'MNV' (Mã nhân viên)",
        )

    existing_docs = {}
    async for doc in db[COLLECTION].find({}, {
        "employeeCode": 1,
        "fullName": 1,
        "tpbankInfo": 1,
        "bmkHrInfo": 1,
        "bmkSystemInfo": 1,
    }):
        code = doc.get("employeeCode")
        if code:
            existing_docs[str(code).strip()] = doc

    operations = []
    now = _now()
    created_count = 0
    updated_count = 0
    total_processed = 0

    for row in rows[header_row_idx + 1:]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        raw_code = row[code_idx] if code_idx < len(row) else None
        if raw_code is None:
            continue
        employee_code = str(raw_code).strip()
        if not employee_code or employee_code.upper() in ["MNV", "MÃ NV"]:
            continue

        if employee_code.isdigit() and len(employee_code) < 5:
            employee_code = employee_code.zfill(5)

        full_name = str(row[name_idx]).strip() if name_idx is not None and name_idx < len(row) and row[name_idx] is not None else ""
        raw_onboard = row[onboard_idx] if onboard_idx is not None and onboard_idx < len(row) else None
        raw_offboard = row[offboard_idx] if offboard_idx is not None and offboard_idx < len(row) else None

        onboard_str, onboard_dt = _parse_excel_date(raw_onboard)
        offboard_str, offboard_dt = _parse_excel_date(raw_offboard)

        # quy tắc: Nếu ngày gia nhập trước ngày 1/4/2025 (dd/MM/yyyy) thì Ngày gia nhập = 1/4/2025
        if onboard_dt and onboard_dt < datetime(2025, 4, 1):
            onboard_dt = datetime(2025, 4, 1)
            onboard_str = "2025-04-01"

        # quy tắc: nếu nhân viên có giá trị ở cột NGÀY THÔI VIỆC là null hoặc N/a hoặc ngày 1/0/1900 (4 số cuối 1900)
        # thì ngày thôi việc là null và đang làm việc employmentStatus = "Hiệu lực" còn lại "Nghỉ việc"
        if offboard_dt is None:
            offboard_str = None
            employment_status = "Hiệu lực"
        else:
            employment_status = "Nghỉ việc"

        contract_count = _calculate_tpbank_contract_count(onboard_dt, offboard_dt)

        if employee_code in existing_docs:
            updated_count += 1
            existing_doc = existing_docs[employee_code]
            tpbank_info = existing_doc.get("tpbankInfo") or {}
            contracts = tpbank_info.get("contracts") or []
            bmk_hr_info = existing_doc.get("bmkHrInfo") or {}
            bmk_system_info = existing_doc.get("bmkSystemInfo") or {}

            result_obj = _calculate_reconciliation_result(
                estimated_contract_count=contract_count,
                hr_contract_count=tpbank_info.get("hrContractCount"),
                bmk_hr_info=bmk_hr_info,
                bmk_system_info=bmk_system_info,
                employment_status=employment_status,
            )

            new_tpbank_info = {
                "estimatedContractCount": contract_count,
                "hrContractCount": tpbank_info.get("hrContractCount"),
                "contracts": contracts,
            }
            update_set = {
                "onboardDate": onboard_str,
                "offboardDate": offboard_str,
                "employmentStatus": employment_status,
                "isSynced": True,
                "tpbankInfo": new_tpbank_info,
                "result": result_obj,
                "updatedAt": now,
            }
            if full_name and not existing_doc.get("fullName"):
                update_set["fullName"] = full_name

            operations.append(
                UpdateOne(
                    {"employeeCode": employee_code},
                    {"$set": update_set}
                )
            )
        else:
            created_count += 1
            existing_docs[employee_code] = {"employeeCode": employee_code, "fullName": full_name}
            new_tpbank_info = {
                "estimatedContractCount": contract_count,
                "hrContractCount": None,
                "contracts": [],
            }
            empty_info = {
                "contractCount": 0,
                "idCardCount": 0,
                "liquidationCount": 0,
                "taxCommitmentCount": 0,
            }
            result_obj = _calculate_reconciliation_result(
                estimated_contract_count=contract_count,
                hr_contract_count=None,
                bmk_hr_info=empty_info,
                bmk_system_info=empty_info,
                employment_status=employment_status,
            )
            operations.append(
                UpdateOne(
                    {"employeeCode": employee_code},
                    {
                        "$set": {
                            "onboardDate": onboard_str,
                            "offboardDate": offboard_str,
                            "employmentStatus": employment_status,
                            "isSynced": True,
                            "tpbankInfo": new_tpbank_info,
                            "result": result_obj,
                            "updatedAt": now,
                        },
                        "$setOnInsert": {
                            "_id": employee_code,
                            "employeeCode": employee_code,
                            "fullName": full_name,
                            "idNumber": None,
                            "createdSource": "TP Bank",
                            "isBmkSystemExist": False,
                            "departmentLevel1": None,
                            "position": None,
                            "bmkHrInfo": empty_info,
                            "bmkSystemInfo": empty_info,
                            "reconciliationStatus": "pending",
                            "createdAt": now,
                        },
                    },
                    upsert=True,
                )
            )

        total_processed += 1

        if len(operations) >= 500:
            await db[COLLECTION].bulk_write(operations, ordered=False)
            operations = []

    if operations:
        await db[COLLECTION].bulk_write(operations, ordered=False)

    return {
        "status": "success",
        "message": f"Đối soát thành công {total_processed} bản ghi từ file TP Bank (Tạo mới: {created_count}, Cập nhật: {updated_count})",
        "totalProcessed": total_processed,
        "createdCount": created_count,
        "updatedCount": updated_count,
    }




