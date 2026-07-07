from datetime import datetime, timezone
from fastapi import APIRouter, Depends, HTTPException, status
from typing import List
from app.core.database import get_db
from app.core.security import get_current_user
from app.models.collaborator import CollaboratorCreate, CollaboratorUpdate, CollaboratorResponse

router = APIRouter(prefix="/api/collaborators", tags=["Collaborators"])

COLLECTION = "bmk_ctv_collaborators"

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
