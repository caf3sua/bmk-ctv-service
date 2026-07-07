import logging
from motor.motor_asyncio import AsyncIOMotorClient
from app.core.config import settings

logger = logging.getLogger("bmk-ctv-service")

class Database:
    client: AsyncIOMotorClient = None
    db = None

db_instance = Database()

async def connect_db():
    """Establish connection to MongoDB cluster."""
    logger.info("Connecting to MongoDB...")
    uri_lower = settings.MONGODB_URI.lower()
    if "tls=true" in uri_lower or "ssl=true" in uri_lower or uri_lower.startswith("mongodb+srv://"):
        import certifi
        db_instance.client = AsyncIOMotorClient(settings.MONGODB_URI, tlsCAFile=certifi.where())
    else:
        db_instance.client = AsyncIOMotorClient(settings.MONGODB_URI)
    db_instance.db = db_instance.client[settings.DB_NAME]
    logger.info("Connected to MongoDB successfully!")

async def close_db():
    """Close MongoDB connection."""
    if db_instance.client:
        logger.info("Closing MongoDB connection...")
        db_instance.client.close()
        logger.info("Closed MongoDB connection.")

def get_db():
    """Dependency injection helper to get database instance."""
    return db_instance.db
