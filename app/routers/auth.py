from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
from app.core.database import get_db
from app.core.security import verify_password, create_access_token, verify_google_token
from app.models.user import UserLogin, TokenResponse

router = APIRouter(prefix="/api/auth", tags=["Authentication"])

class GoogleLoginRequest(BaseModel):
    token: str

def _issue_token_response(user: dict) -> dict:
    token = create_access_token({"username": user["username"]})
    return {
        "token": token,
        "user": {
            "username": user["username"],
            "name": user.get("name", user["username"]),
            "email": user.get("email", ""),
            "role": user.get("role", "staff"),
        },
    }

@router.post("/login", response_model=TokenResponse)
async def login(credentials: UserLogin, db=Depends(get_db)):
    """Authenticate a user by username/password and return a signed access token."""
    username = credentials.username.strip().lower()
    user = await db["bmk_ctv_users"].find_one({"username": username})

    if not user or not verify_password(credentials.password, user.get("hashedPassword", "")):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Sai tên đăng nhập hoặc mật khẩu"
        )

    if not user.get("active", True):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa hoặc ngừng hoạt động"
        )

    return _issue_token_response(user)

@router.post("/google", response_model=TokenResponse)
async def login_google(credentials: GoogleLoginRequest, db=Depends(get_db)):
    """Authenticate via a Google ID token and return a signed access token."""
    id_info = verify_google_token(credentials.token)
    email = id_info.get("email")
    if not email:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Token Google không chứa địa chỉ email."
        )

    user = await db["bmk_ctv_users"].find_one({"email": email.lower()})
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Tài khoản Google này chưa được cấp quyền truy cập hệ thống."
        )

    if not user.get("active", True):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa hoặc ngừng hoạt động"
        )

    return _issue_token_response(user)
