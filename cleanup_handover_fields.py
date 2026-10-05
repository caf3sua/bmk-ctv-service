"""One-off cleanup script: remove outer handoverDate and handoverPerson fields from all documents
in bmk_ctv_collaborators collection since handoverInfo is now used.
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

def cleanup():
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_collaborators"]

    result = collection.update_many(
        {"$or": [{"handoverDate": {"$exists": True}}, {"handoverPerson": {"$exists": True}}]},
        {"$unset": {"handoverDate": "", "handoverPerson": ""}}
    )
    print(f"Cleaned up {result.modified_count} collaborator document(s) by removing root handoverDate and handoverPerson fields.")
    client.close()

if __name__ == "__main__":
    cleanup()
