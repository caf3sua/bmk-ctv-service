"""One-off migration: convert checklist.serviceContract (single object) into
checklist.serviceContracts (array with at least 1 item), to support collaborators
having multiple service contract periods over time.

Safe to re-run: only touches documents that still have the old field.
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

    cursor = collection.find({"checklist.serviceContract": {"$exists": True}})
    migrated = 0
    for doc in cursor:
        old = doc.get("checklist", {}).get("serviceContract") or {}
        new_value = [{"startDate": old.get("startDate"), "endDate": old.get("endDate")}]
        collection.update_one(
            {"_id": doc["_id"]},
            {
                "$set": {"checklist.serviceContracts": new_value},
                "$unset": {"checklist.serviceContract": ""},
            },
        )
        migrated += 1

    print(f"Migrated {migrated} collaborator document(s).")
    client.close()

if __name__ == "__main__":
    migrate()
