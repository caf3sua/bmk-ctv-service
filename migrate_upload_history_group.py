"""One-off migration: add group = 'CTV' to all existing upload history logs.
Safe to re-run.
"""
import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "bmk_ctv")

if not MONGODB_URI:
    print("Error: MONGODB_URI is not set in environment variables.")
    exit(1)

def migrate():
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_upload_history"]

    # Find documents where group doesn't exist
    query = {"group": {"$exists": False}}
    cursor = collection.find(query)
    
    migrated = 0
    for doc in cursor:
        collection.update_one(
            {"_id": doc["_id"]},
            {"$set": {"group": "CTV"}}
        )
        migrated += 1

    print(f"Migrated {migrated} upload history document(s) to group 'CTV'.")
    client.close()

if __name__ == "__main__":
    migrate()
