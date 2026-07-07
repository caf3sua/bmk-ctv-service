from pydantic import BaseModel, EmailStr, Field
from typing import Literal, Optional

UserRole = Literal['admin', 'staff']

class UserBase(BaseModel):
    username: str
    name: str = ""
    email: EmailStr
    role: UserRole = 'staff'
    active: bool = True

class UserCreate(UserBase):
    password: str = Field(min_length=6)

class UserUpdate(BaseModel):
    name: Optional[str] = None
    email: Optional[EmailStr] = None
    role: Optional[UserRole] = None
    active: Optional[bool] = None
    password: Optional[str] = Field(default=None, min_length=6)

class UserResponse(UserBase):
    createdAt: str
    updatedAt: str

    model_config = {
        "json_schema_extra": {
            "example": {
                "username": "admin",
                "name": "Quản trị viên",
                "email": "admin@bmk.vn",
                "role": "admin",
                "active": True,
                "createdAt": "2026-01-15T02:00:00.000Z",
                "updatedAt": "2026-01-15T02:00:00.000Z"
            }
        }
    }

class UserLogin(BaseModel):
    username: str
    password: str

class AuthUser(BaseModel):
    username: str
    name: str
    email: EmailStr
    role: UserRole

class TokenResponse(BaseModel):
    token: str
    user: AuthUser
