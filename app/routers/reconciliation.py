import math
import re
from datetime import datetime, timezone
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query, status
from app.core.database import get_db
from app.core.security import get_current_user
from app.models.reconciliation import (
    ReconciliationRecordCreate,
    ReconciliationRecordUpdate,
    ReconciliationRecordResponse,
    ReconciliationListResponse,
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
    return {
        "id": str(doc["_id"]),
        "employeeCode": doc.get("employeeCode", ""),
        "fullName": doc.get("fullName", ""),
        "idNumber": doc.get("idNumber") or doc.get("cccd"),
        "createdSource": created_source,
        "departmentLevel1": doc.get("departmentLevel1"),
        "position": doc.get("position"),
        "employmentStatus": doc.get("employmentStatus"),
        "onboardDate": doc.get("onboardDate"),
        "offboardDate": doc.get("offboardDate"),
        "tpbankInfo": doc.get("tpbankInfo") or {"contracts": []},
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
    
    # CCCD
    cccd = checklist.get("cccd") or {}
    id_card_count = 1 if (cccd.get("checked") or cccd.get("file")) else 0
    
    # Cam kết thuế
    ckt = checklist.get("ckt") or {}
    tax_count = 1 if (ckt.get("checked") or ckt.get("file")) else 0
    
    # Biên bản thanh lý
    bbtl = checklist.get("bbtl") or {}
    liquidation_count = 1 if (bbtl.get("date") or bbtl.get("file")) else 0
    
    # Hợp đồng dịch vụ
    hddv = checklist.get("hddv") or {}
    contract_dates = hddv.get("contract_date") or []
    files = hddv.get("files") or []
    
    valid_dates = [d for d in contract_dates if (isinstance(d, dict) and (d.get("startDate") or d.get("endDate")))]
    contract_count = max(len(valid_dates), len(files))
    if contract_count == 0 and len(contract_dates) > 0:
        # Nếu có mảng contract_date nhưng chưa điền ngày
        contract_count = 0

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

    if created_source and created_source.strip() and created_source.strip() != "all":
        src = created_source.strip()
        if src == "bmk_system":
            conditions.append({
                "$or": [
                    {"createdSource": "bmk_system"},
                    {"createdSource": {"$exists": False}},
                ]
            })
        else:
            conditions.append({"createdSource": src})

    query = {"$and": conditions} if conditions else {}

    total = await db[COLLECTION].count_documents(query)
    total_pages = max(1, math.ceil(total / page_size))
    skip = (page - 1) * page_size

    cursor = db[COLLECTION].find(query).sort("employeeCode", 1).skip(skip).limit(page_size)
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

from pymongo import UpdateOne

@router.post("/sync-system-info")
async def sync_bmk_system_info(
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """
    Đồng bộ thông tin số lượng hồ sơ từ collection bmk_ctv_collaborators
    vào trường bmkSystemInfo của collection bmk_ctv_reconciliations.
    Nếu CTV chưa có trong bảng đối soát, tự động tạo mới bản ghi với thông tin cơ bản.
    """
    collab_cursor = db[COLLABORATORS_COLLECTION].find({})
    operations = []
    total_processed = 0
    now = _now()

    async for collab in collab_cursor:
        code = collab.get("employeeCode") or str(collab.get("_id"))
        if not code:
            continue

        system_info = _calculate_system_info(collab)
        full_name = collab.get("fullName", "")
        id_number = collab.get("idNumber") or collab.get("cccd") or None

        operations.append(
            UpdateOne(
                {"employeeCode": code},
                {
                    "$set": {
                        "bmkSystemInfo": system_info,
                        "idNumber": id_number,
                        "updatedAt": now,
                    },
                    "$setOnInsert": {
                        "fullName": full_name,
                        "createdSource": "bmk_system",
                        "departmentLevel1": None,
                        "position": None,
                        "employmentStatus": "Hiện diện",
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

        if len(operations) >= 500:
            await db[COLLECTION].bulk_write(operations, ordered=False)
            total_processed += len(operations)
            operations = []

    if operations:
        await db[COLLECTION].bulk_write(operations, ordered=False)
        total_processed += len(operations)

    return {
        "status": "success",
        "message": f"Đã đồng bộ {total_processed} cộng tác viên từ hệ thống BMK",
        "totalProcessed": total_processed,
    }

