from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
from app.core.activity_log import record_activity
from app.core.database import get_db
from app.core.logging import get_logger
from app.core.security import verify_password, create_access_token, verify_google_token, get_current_user
from app.models.user import UserLogin, TokenResponse

router = APIRouter(prefix="/api/auth", tags=["Authentication"])
logger = get_logger(__name__)

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
        full_name = user.get("name") or username if user else username
        logger.warning(f"Login thất bại (sai thông tin đăng nhập): username='{username}'")
        await record_activity(
            db, action="login", result="fail", full_name=full_name, username=username,
            message=f"{full_name} đăng nhập hệ thống thất bại",
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Sai tên đăng nhập hoặc mật khẩu"
        )

    full_name = user.get("name") or username

    if not user.get("active", True):
        logger.warning(f"Login thất bại (tài khoản bị khóa): username='{username}'")
        await record_activity(
            db, action="login", result="fail", full_name=full_name, username=username,
            message=f"{full_name} đăng nhập hệ thống thất bại",
        )
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa hoặc ngừng hoạt động"
        )

    logger.info(f"Login thành công: username='{username}'")
    await record_activity(
        db, action="login", result="success", full_name=full_name, username=username,
        message=f"{full_name} đăng nhập hệ thống thành công",
    )
    return _issue_token_response(user)

@router.post("/logout", status_code=status.HTTP_204_NO_CONTENT)
async def logout(current_user: dict = Depends(get_current_user), db=Depends(get_db)):
    """Record a logout event for the current user (JWT is stateless, this is audit-only)."""
    full_name = current_user.get("name") or current_user.get("username")
    await record_activity(
        db, action="logout", result="success", full_name=full_name, username=current_user.get("username"),
        message=f"{full_name} đăng xuất hệ thống thành công",
    )
    return None

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
        logger.warning(f"Login Google thất bại (chưa được cấp quyền): email='{email}'")
        await record_activity(
            db, action="login", result="fail", full_name=email, username=email,
            message=f"{email} đăng nhập hệ thống thất bại",
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Tài khoản Google này chưa được cấp quyền truy cập hệ thống."
        )

    full_name = user.get("name") or user.get("username")

    if not user.get("active", True):
        logger.warning(f"Login Google thất bại (tài khoản bị khóa): email='{email}'")
        await record_activity(
            db, action="login", result="fail", full_name=full_name, username=user.get("username"),
            message=f"{full_name} đăng nhập hệ thống thất bại",
        )
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa hoặc ngừng hoạt động"
        )

    logger.info(f"Login Google thành công: email='{email}'")
    await record_activity(
        db, action="login", result="success", full_name=full_name, username=user.get("username"),
        message=f"{full_name} đăng nhập hệ thống thành công",
    )
    return _issue_token_response(user)
