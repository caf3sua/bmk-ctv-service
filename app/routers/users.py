from datetime import datetime, timezone
from fastapi import APIRouter, Depends, HTTPException, status
from typing import List
from app.core.database import get_db
from app.core.security import get_current_admin_user, hash_password
from app.models.user import UserCreate, UserUpdate, UserResponse

router = APIRouter(prefix="/api/users", tags=["Users"])

COLLECTION = "bmk_ctv_users"

def _now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

def _to_response(doc: dict) -> dict:
    doc = dict(doc)
    doc.pop("hashedPassword", None)
    return doc

@router.get("", response_model=List[UserResponse])
async def list_users(db=Depends(get_db), current_user: dict = Depends(get_current_admin_user)):
    """Fetch all admin/staff accounts, sorted by username."""
    items = []
    cursor = db[COLLECTION].find({}).sort("username", 1)
    async for doc in cursor:
        items.append(_to_response(doc))
    return items

@router.get("/{username}", response_model=UserResponse)
async def get_user(username: str, db=Depends(get_db), current_user: dict = Depends(get_current_admin_user)):
    """Fetch a single account by username."""
    doc = await db[COLLECTION].find_one({"username": username})
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy người dùng "{username}"'
        )
    return _to_response(doc)

@router.post("", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
async def create_user(payload: UserCreate, db=Depends(get_db), current_user: dict = Depends(get_current_admin_user)):
    """Create a new admin/staff account."""
    username = payload.username.strip().lower()
    if not username:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Tên đăng nhập là bắt buộc")

    existing = await db[COLLECTION].find_one({"username": username})
    if existing:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f'Tên đăng nhập "{username}" đã tồn tại'
        )

    email = payload.email.lower()
    email_conflict = await db[COLLECTION].find_one({"email": email})
    if email_conflict:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f'Email "{email}" đã được sử dụng'
        )

    now = _now()
    doc = payload.model_dump(exclude={"password"})
    doc["username"] = username
    doc["email"] = email
    doc["_id"] = username
    doc["hashedPassword"] = hash_password(payload.password)
    doc["createdAt"] = now
    doc["updatedAt"] = now
    await db[COLLECTION].insert_one(doc)
    return _to_response(doc)

@router.put("/{username}", response_model=UserResponse)
async def update_user(username: str, payload: UserUpdate, db=Depends(get_db), current_user: dict = Depends(get_current_admin_user)):
    """Update an existing account's profile, role, active status or password."""
    existing = await db[COLLECTION].find_one({"username": username})
    if not existing:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy người dùng "{username}"'
        )

    updates = payload.model_dump(exclude_unset=True, exclude={"password"})

    if "email" in updates:
        updates["email"] = updates["email"].lower()
        email_conflict = await db[COLLECTION].find_one({"email": updates["email"], "username": {"$ne": username}})
        if email_conflict:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f'Email "{updates["email"]}" đã được sử dụng'
            )

    is_last_admin_change = existing.get("role") == "admin" and (
        updates.get("role", "admin") != "admin" or updates.get("active", True) is False
    )
    if is_last_admin_change:
        remaining_admins = await db[COLLECTION].count_documents(
            {"role": "admin", "active": True, "username": {"$ne": username}}
        )
        if remaining_admins == 0:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Không thể hạ quyền hoặc khóa quản trị viên cuối cùng"
            )

    if payload.password:
        updates["hashedPassword"] = hash_password(payload.password)

    updates["updatedAt"] = _now()
    await db[COLLECTION].update_one({"username": username}, {"$set": updates})
    doc = await db[COLLECTION].find_one({"username": username})
    return _to_response(doc)

@router.delete("/{username}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_user(username: str, db=Depends(get_db), current_user: dict = Depends(get_current_admin_user)):
    """Delete an account by username."""
    if username == current_user.get("username"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Không thể tự xóa tài khoản của chính mình"
        )

    existing = await db[COLLECTION].find_one({"username": username})
    if not existing:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f'Không tìm thấy người dùng "{username}"'
        )

    if existing.get("role") == "admin":
        remaining_admins = await db[COLLECTION].count_documents(
            {"role": "admin", "username": {"$ne": username}}
        )
        if remaining_admins == 0:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Không thể xóa quản trị viên cuối cùng"
            )

    await db[COLLECTION].delete_one({"username": username})
    return None
