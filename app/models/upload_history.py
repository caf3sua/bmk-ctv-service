from pydantic import BaseModel

class UploadHistoryResponse(BaseModel):
    id: str
    filename: str
    s3Key: str
    s3Bucket: str
    uploadedBy: str
    username: str
    rowsProcessed: int
    createdCount: int
    updatedCount: int
    status: str  # "success" | "fail"
    message: str
    createdAt: str
    group: str = "CTV"

    model_config = {
        "json_schema_extra": {
            "example": {
                "id": "65a1f0c2e4b0f0a1b2c3d4e5",
                "filename": "mau_import_checklist_ctv.xlsx",
                "s3Key": "excel/20260721_113244_mau_import_checklist_ctv.xlsx",
                "s3Bucket": "intranet",
                "uploadedBy": "Quản trị viên",
                "username": "admin",
                "rowsProcessed": 5,
                "createdCount": 3,
                "updatedCount": 2,
                "status": "success",
                "message": "Đã nhập thành công: tạo mới 3, cập nhật 2. Lỗi ngày tháng: 0",
                "createdAt": "2026-07-21T04:32:44Z",
            }
        }
    }
