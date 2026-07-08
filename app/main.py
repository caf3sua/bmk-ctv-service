from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from app.core.config import settings
from app.core.database import connect_db, close_db
from app.routers import auth, collaborators, users

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup actions
    await connect_db()
    yield
    # Shutdown actions
    await close_db()

app = FastAPI(
    title="BMK CTV Management API",
    description="Backend API for managing collaborator (CTV) profiles, backed by MongoDB",
    version="1.0.0",
    lifespan=lifespan
)

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

@app.get("/", tags=["Root"])
async def root():
    return {
        "status": "online",
        "service": "BMK CTV Management API",
        "version": "1.0.0",
        "docs_url": "/docs"
    }
