"""Ghi nhật ký hoạt động hệ thống (login, logout, CRUD/import/export cộng tác viên)."""
from datetime import datetime, timezone
from app.core.logging import get_logger

logger = get_logger(__name__)

COLLECTION = "bmk_ctv_activity_logs"

async def record_activity(db, *, action: str, result: str, full_name: str, username: str, message: str) -> None:
    """Lưu 1 dòng nhật ký hệ thống. Không bao giờ raise lỗi ra ngoài để tránh làm hỏng thao tác chính."""
    doc = {
        "action": action,
        "result": result,
        "fullName": full_name,
        "username": username,
        "message": message,
        "createdAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    try:
        await db[COLLECTION].insert_one(doc)
    except Exception:
        logger.exception(f"Không ghi được nhật ký hệ thống cho action='{action}'")
