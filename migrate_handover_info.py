"""One-off migration: populate handoverInfo list for existing collaborators that have
legacy handoverDate or handoverPerson fields.

Safe to re-run: only updates documents where handoverInfo is missing or empty.
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
    collection = db["bmk_ctv_collaborators"]

    cursor = collection.find({
        "$or": [
            {"handoverInfo": {"$exists": False}},
            {"handoverInfo": {"$eq": []}},
            {"handoverInfo": None}
        ]
    })
    migrated = 0
    for doc in cursor:
        handover_date = doc.get("handoverDate")
        handover_person = doc.get("handoverPerson") or ""
        created_at = doc.get("createdAt")
        
        handover_info = []
        if handover_date:
            handover_info.append({
                "handoverDate": handover_date,
                "handoverPerson": handover_person,
                "createdAt": created_at
            })

        collection.update_one(
            {"_id": doc["_id"]},
            {
                "$set": {"handoverInfo": handover_info}
            }
        )
        migrated += 1

    print(f"Migrated {migrated} collaborator document(s) with handoverInfo.")
    client.close()

if __name__ == "__main__":
    migrate()
