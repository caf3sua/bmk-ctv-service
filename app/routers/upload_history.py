from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import StreamingResponse
from typing import List
from urllib.parse import quote
from bson import ObjectId
from app.core.database import get_db
from app.core.security import get_current_user
from app.core.s3 import get_s3_client
from app.core.config import settings
from app.core.logging import get_logger
from app.models.upload_history import UploadHistoryResponse

router = APIRouter(prefix="/api/upload-history", tags=["Upload History"])
logger = get_logger(__name__)

COLLECTION = "bmk_ctv_upload_history"

def _to_response(doc: dict) -> dict:
    return {
        "id": str(doc["_id"]),
        "filename": doc.get("filename", ""),
        "s3Key": doc.get("s3Key", ""),
        "s3Bucket": doc.get("s3Bucket", ""),
        "uploadedBy": doc.get("uploadedBy", ""),
        "username": doc.get("username", ""),
        "rowsProcessed": doc.get("rowsProcessed", 0),
        "createdCount": doc.get("createdCount", 0),
        "updatedCount": doc.get("updatedCount", 0),
        "status": doc.get("status", ""),
        "message": doc.get("message", ""),
        "createdAt": doc.get("createdAt", ""),
        "group": doc.get("group", "CTV"),
    }

@router.get("", response_model=List[UploadHistoryResponse])
async def list_upload_history(db=Depends(get_db), current_user: dict = Depends(get_current_user)):
    """Fetch upload history log, newest first."""
    items = []
    cursor = db[COLLECTION].find({}).sort("createdAt", -1).limit(1000)
    async for doc in cursor:
        items.append(_to_response(doc))
    return items

@router.get("/{history_id}/download")
async def download_uploaded_file(
    history_id: str, db=Depends(get_db), current_user: dict = Depends(get_current_user)
):
    """Stream an uploaded file from MinIO S3."""
    if not ObjectId.is_valid(history_id):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Mã lịch sử upload không hợp lệ")

    history = await db[COLLECTION].find_one({"_id": ObjectId(history_id)})
    if not history:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không tìm thấy lịch sử upload")

    s3_key = history.get("s3Key")
    s3_bucket = history.get("s3Bucket", settings.S3_BUCKET)
    filename = history.get("filename", "download.xlsx")

    if not s3_key:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Không có thông tin lưu trữ S3 cho file này")

    try:
        s3 = get_s3_client()
        response = s3.get_object(Bucket=s3_bucket, Key=s3_key)

        def iter_chunks():
            for chunk in response["Body"].iter_chunks(chunk_size=1024 * 1024):
                yield chunk

        return StreamingResponse(
            iter_chunks(),
            media_type=response.get("ContentType", "application/octet-stream"),
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"
            }
        )
    except Exception as e:
        logger.error(f"Lỗi khi tải file từ S3 cho ID={history_id}: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Không thể tải file từ S3 MinIO"
        )
