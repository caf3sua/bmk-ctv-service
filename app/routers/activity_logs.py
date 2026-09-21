import re
from fastapi import APIRouter, Depends, Query
from typing import List, Optional
from app.core.database import get_db
from app.core.security import get_current_user
from app.models.activity_log import ActivityLogResponse

router = APIRouter(prefix="/api/activity-logs", tags=["Activity Logs"])

COLLECTION = "bmk_ctv_activity_logs"

def _to_response(doc: dict) -> dict:
    return {
        "id": str(doc["_id"]),
        "action": doc.get("action"),
        "result": doc.get("result"),
        "fullName": doc.get("fullName", ""),
        "username": doc.get("username", ""),
        "message": doc.get("message", ""),
        "employeeCode": doc.get("employeeCode"),
        "createdAt": doc.get("createdAt", ""),
    }

@router.get("", response_model=List[ActivityLogResponse])
async def list_activity_logs(
    employee_code: Optional[str] = Query(None, description="Lọc log theo mã cộng tác viên"),
    employeeCode: Optional[str] = Query(None, description="Lọc log theo mã cộng tác viên (alias)"),
    db=Depends(get_db),
    current_user: dict = Depends(get_current_user),
):
    """Fetch system activity logs, newest first (any authenticated user)."""
    items = []
    target_code = (employee_code or employeeCode)
    query = {}
    if target_code:
        code = target_code.strip()
        query = {
            "$or": [
                {"employeeCode": code},
                {"employeeCode": {"$regex": f"^{re.escape(code)}$", "$options": "i"}},
                {
                    "employeeCode": None,
                    "message": {"$regex": rf"(?i)\b{re.escape(code)}\b"}
                }
            ]
        }

    cursor = db[COLLECTION].find(query).sort("createdAt", -1).limit(1000)
    async for doc in cursor:
        items.append(_to_response(doc))
    return items


