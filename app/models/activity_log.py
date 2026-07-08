from pydantic import BaseModel
from typing import Literal

ActivityAction = Literal[
    "login",
    "logout",
    "create_collaborator",
    "update_collaborator",
    "delete_collaborator",
    "import_collaborators",
    "export_collaborators",
]
ActivityResult = Literal["success", "error", "fail"]

class ActivityLogResponse(BaseModel):
    id: str
    action: ActivityAction
    result: ActivityResult
    fullName: str
    username: str
    message: str
    createdAt: str

    model_config = {
        "json_schema_extra": {
            "example": {
                "id": "65a1f0c2e4b0f0a1b2c3d4e5",
                "action": "login",
                "result": "success",
                "fullName": "Quản trị viên",
                "username": "admin",
                "message": "Quản trị viên đăng nhập hệ thống thành công",
                "createdAt": "2026-01-15T02:00:00.000Z",
            }
        }
    }
