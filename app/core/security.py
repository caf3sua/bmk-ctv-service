import hashlib
import secrets
import json
import base64
import hmac
import time
from typing import Optional
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from app.core.config import settings
from app.core.database import get_db

# OAuth2 Scheme definition
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)

# --- Password Hashing (PBKDF2-SHA256) ---
def hash_password(password: str) -> str:
    """Hash a password using PBKDF2-SHA256 with 600,000 iterations and a random salt."""
    salt = secrets.token_hex(16)
    iterations = 600000
    hash_bytes = hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt.encode('utf-8'),
        iterations
    )
    hash_hex = hash_bytes.hex()
    return f"pbkdf2:sha256:{iterations}${salt}${hash_hex}"

def verify_password(password: str, hashed_password: str) -> bool:
    """Verify a plain password against the stored PBKDF2-SHA256 hash."""
    try:
        parts = hashed_password.split('$')
        if len(parts) != 3:
            return False
        algo_info, salt, hash_hex = parts
        _, subalgo, iterations_str = algo_info.split(':')
        iterations = int(iterations_str)

        test_hash_bytes = hashlib.pbkdf2_hmac(
            subalgo,
            password.encode('utf-8'),
            salt.encode('utf-8'),
            iterations
        )
        return hmac.compare_digest(test_hash_bytes.hex(), hash_hex)
    except Exception:
        return False

# --- HMAC-SHA256 Signed JSON Tokens (JWT Compliant) ---
def base64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode('utf-8').rstrip('=')

def base64url_decode(data: str) -> bytes:
    padding = '=' * (4 - len(data) % 4)
    return base64.urlsafe_b64decode(data + padding)

def create_access_token(data: dict, expires_delta_minutes: Optional[int] = None) -> str:
    """Generate an HMAC-SHA256 signed JWT-like access token."""
    payload = data.copy()
    if expires_delta_minutes:
        expire = int(time.time()) + (expires_delta_minutes * 60)
    else:
        expire = int(time.time()) + (settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60)

    payload['exp'] = expire

    header = {"alg": "HS256", "typ": "JWT"}

    header_b64 = base64url_encode(json.dumps(header).encode('utf-8'))
    payload_b64 = base64url_encode(json.dumps(payload).encode('utf-8'))

    signature_input = f"{header_b64}.{payload_b64}".encode('utf-8')
    key = settings.SECRET_KEY.encode('utf-8')
    signature = hmac.new(key, signature_input, hashlib.sha256).digest()
    signature_b64 = base64url_encode(signature)

    return f"{header_b64}.{payload_b64}.{signature_b64}"

def verify_access_token(token: str) -> Optional[dict]:
    """Verify the signature and expiration of an access token."""
    try:
        parts = token.split('.')
        if len(parts) != 3:
            return None

        header_b64, payload_b64, signature_b64 = parts

        signature_input = f"{header_b64}.{payload_b64}".encode('utf-8')
        key = settings.SECRET_KEY.encode('utf-8')
        expected_signature = hmac.new(key, signature_input, hashlib.sha256).digest()
        expected_signature_b64 = base64url_encode(expected_signature)

        if not hmac.compare_digest(signature_b64, expected_signature_b64):
            return None

        payload_bytes = base64url_decode(payload_b64)
        payload = json.loads(payload_bytes.decode('utf-8'))

        exp = payload.get('exp')
        if exp and int(time.time()) > exp:
            return None  # Expired

        return payload
    except Exception:
        return None

# --- FastAPI Authentication Dependencies ---
async def get_current_user(token: Optional[str] = Depends(oauth2_scheme), db=Depends(get_db)):
    """FastAPI dependency to retrieve the current logged-in user from the database."""
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Không thể xác thực phiên đăng nhập",
        headers={"WWW-Authenticate": "Bearer"},
    )

    if not token:
        raise credentials_exception

    payload = verify_access_token(token)
    if payload is None:
        raise credentials_exception

    username: str = payload.get("username")
    if username is None:
        raise credentials_exception

    user = await db["bmk_ctv_users"].find_one({"username": username})
    if user is None:
        raise credentials_exception

    if not user.get("active", True):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Tài khoản đã bị khóa hoặc ngừng hoạt động"
        )

    return user

async def get_current_admin_user(current_user: dict = Depends(get_current_user)):
    """FastAPI dependency that additionally requires the 'admin' role."""
    if current_user.get("role") != "admin":
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Bạn không có quyền thực hiện thao tác này"
        )
    return current_user

# --- Google Social Login (ID token verification) ---
def verify_google_token(token: str) -> dict:
    """Verify a Google ID token against Google's tokeninfo endpoint and check the audience."""
    import urllib.request
    import urllib.error

    if not settings.GOOGLE_CLIENT_ID:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Chưa cấu hình Google Client ID trên hệ thống backend."
        )

    try:
        url = f"https://oauth2.googleapis.com/tokeninfo?id_token={token}"
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode('utf-8'))

        if "error" in data or "error_description" in data:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail=data.get("error_description", "Token Google không hợp lệ.")
            )

        if data.get("aud") != settings.GOOGLE_CLIENT_ID:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Token Google không khớp Client ID hệ thống."
            )

        if str(data.get("email_verified")).lower() != "true":
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Email Google chưa được xác minh."
            )

        return data
    except HTTPException:
        raise
    except urllib.error.HTTPError as e:
        try:
            error_data = json.loads(e.read().decode('utf-8'))
            error_detail = error_data.get("error_description", "Xác thực Token Google thất bại.")
        except Exception:
            error_detail = "Xác thực Token Google thất bại."
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=error_detail)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Lỗi xác thực Google: {str(e)}"
        )
