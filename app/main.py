import time
import uuid
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from app.core.config import settings
from app.core.database import connect_db, close_db
from app.core.logging import get_logger
from app.routers import activity_logs, auth, collaborators, users

logger = get_logger(__name__)

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup actions
    logger.info("Starting up BMK CTV Management API")
    await connect_db()
    yield
    # Shutdown actions
    logger.info("Shutting down BMK CTV Management API")
    await close_db()

app = FastAPI(
    title="BMK CTV Management API",
    description="Backend API for managing collaborator (CTV) profiles, backed by MongoDB",
    version="1.0.0",
    lifespan=lifespan
)

@app.middleware("http")
async def log_requests(request: Request, call_next):
    request_id = str(uuid.uuid4())
    start = time.time()
    try:
        response = await call_next(request)
    except Exception:
        logger.exception(
            f"Unhandled exception for {request.method} {request.url.path}",
            extra={"request_id": request_id, "method": request.method, "path": request.url.path},
        )
        raise
    duration_ms = (time.time() - start) * 1000
    logger.info(
        f"{request.method} {request.url.path} {response.status_code} {duration_ms:.2f}ms",
        extra={
            "request_id": request_id,
            "method": request.method,
            "path": request.url.path,
            "status_code": response.status_code,
            "duration_ms": round(duration_ms, 2),
        },
    )
    response.headers["X-Request-ID"] = request_id
    return response

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list,
    # Cho phép mọi origin *.vercel.app (production + preview deployments) và localhost khi dev.
    allow_origin_regex=r"^(https://.*\.vercel\.app|http://(localhost|127\.0\.0\.1):\d+)$",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register routers
app.include_router(auth.router)
app.include_router(collaborators.router)
app.include_router(users.router)
app.include_router(activity_logs.router)

@app.get("/", tags=["Root"])
async def root():
    return {
        "status": "online",
        "service": "BMK CTV Management API",
        "version": "1.0.0",
        "docs_url": "/docs"
    }
