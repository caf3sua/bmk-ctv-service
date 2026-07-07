# BMK CTV Management API

Backend API cho hệ thống quản lý hồ sơ cộng tác viên (CTV) của BMK, xây dựng bằng **FastAPI** và **MongoDB** (driver bất đồng bộ **Motor**).

## Cấu trúc thư mục

```text
bmk-ctv-service/
├── app/
│   ├── core/
│   │   ├── config.py        # Cấu hình đọc từ biến môi trường (Pydantic Settings)
│   │   ├── database.py      # Kết nối MongoDB (Motor client)
│   │   └── security.py      # Băm mật khẩu, sinh/xác thực access token, xác thực Google ID token
│   ├── models/
│   │   ├── collaborator.py  # Schema hồ sơ cộng tác viên
│   │   └── user.py          # Schema tài khoản quản trị (đăng nhập + CRUD)
│   ├── routers/
│   │   ├── auth.py          # Endpoint đăng nhập
│   │   ├── collaborators.py # Endpoint CRUD hồ sơ cộng tác viên
│   │   └── users.py         # Endpoint CRUD tài khoản quản trị hệ thống
│   └── main.py               # Điểm khởi tạo FastAPI & cấu hình CORS
├── seed_data/
│   └── collaborators.json    # Dữ liệu mẫu để seed
├── .env                       # Biến môi trường cục bộ (không commit)
├── .env.example                # Mẫu biến môi trường
├── requirements.txt
├── seed.py                     # Script seed dữ liệu mẫu
└── README.md
```

## Cài đặt

### 1. Yêu cầu
- Python 3.10+
- MongoDB (Atlas cluster hoặc local)

### 2. Tạo virtual environment
```bash
python -m venv .venv
```

### 3. Kích hoạt virtual environment
- Windows (PowerShell):
  ```powershell
  .\.venv\Scripts\Activate.ps1
  ```
- macOS/Linux:
  ```bash
  source .venv/bin/activate
  ```

### 4. Cài dependencies
```bash
pip install -r requirements.txt
```

### 5. Cấu hình môi trường
File `.env` đã có sẵn tại thư mục gốc (copy từ `.env.example` nếu cần tạo lại):
```env
MONGODB_URI=mongodb+srv://<username>:<password>@<cluster-url>/?appName=bmk_ctv
DB_NAME=bmk_ctv
PORT=8000
HOST=0.0.0.0
SECRET_KEY=change-me-to-a-random-secret
ACCESS_TOKEN_EXPIRE_MINUTES=1440
CORS_ORIGINS=http://localhost:5173,http://127.0.0.1:5173
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
```

`GOOGLE_CLIENT_SECRET` được lưu để dự phòng cho sau này, luồng đăng nhập Google hiện tại (ID token từ Google Identity Services phía frontend) chỉ cần `GOOGLE_CLIENT_ID` để xác thực `aud` của token qua endpoint `tokeninfo` của Google — không cần trao đổi client secret ở backend.

## Seed dữ liệu

Nạp dữ liệu mẫu (30 cộng tác viên + 2 tài khoản `admin` / `123456` (role `admin`) và `staff` / `123456` (role `staff`)):
```bash
python seed.py
```
Lệnh này sẽ xóa sạch các collection `bmk_ctv_collaborators`, `bmk_ctv_users` hiện có rồi seed lại.

## Chạy server

```bash
uvicorn app.main:app --reload --port 8000
```

- API Base URL: `http://localhost:8000`
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## API Endpoints

### Authentication
- `POST /api/auth/login` — Đăng nhập bằng `username` / `password`, trả về `token` và thông tin `user`.
- `POST /api/auth/google` — Đăng nhập bằng Google ID token (`{ "token": "<id_token>" }` lấy từ Google Identity Services phía frontend). Backend xác thực token qua Google, tìm tài khoản theo `email` trong `bmk_ctv_users` — chỉ những email đã được admin tạo sẵn trong module Users mới đăng nhập được, tài khoản Google lạ sẽ bị từ chối.

### Collaborators (yêu cầu header `Authorization: Bearer <token>`)
- `GET /api/collaborators` — Danh sách tất cả cộng tác viên.
- `GET /api/collaborators/{employeeCode}` — Chi tiết một cộng tác viên.
- `POST /api/collaborators` — Tạo mới hồ sơ cộng tác viên.
- `PUT /api/collaborators/{employeeCode}` — Cập nhật hồ sơ.
- `DELETE /api/collaborators/{employeeCode}` — Xóa hồ sơ.

`employeeCode` được dùng làm khóa chính (`_id`) trong MongoDB.

### Users — quản trị người dùng hệ thống (yêu cầu role `admin`)
- `GET /api/users` — Danh sách tài khoản (username, name, email, role, active).
- `GET /api/users/{username}` — Chi tiết một tài khoản.
- `POST /api/users` — Tạo mới tài khoản (yêu cầu `password` tối thiểu 6 ký tự).
- `PUT /api/users/{username}` — Cập nhật thông tin, đổi role, khóa/mở khóa (`active`), hoặc đổi mật khẩu.
- `DELETE /api/users/{username}` — Xóa tài khoản.

Ràng buộc an toàn: không thể tự xóa tài khoản đang đăng nhập, không thể xóa/hạ quyền/khóa quản trị viên (`role=admin`, `active=true`) cuối cùng trong hệ thống. Tài khoản bị khóa (`active=false`) sẽ không thể đăng nhập hoặc dùng token hiện có để gọi API.
