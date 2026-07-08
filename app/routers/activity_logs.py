from fastapi import APIRouter, Depends
from typing import List
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
        "createdAt": doc.get("createdAt", ""),
    }

@router.get("", response_model=List[ActivityLogResponse])
async def list_activity_logs(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Fetch system activity logs, newest first (any authenticated user)."""
    items = []
    cursor = db[COLLECTION].find({}).sort("createdAt", -1).limit(1000)
    async for doc in cursor:
        items.append(_to_response(doc))
    return items
