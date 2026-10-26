"""One-off cleanup script: remove contractExpiryDate, contractOnboardDate, and isCalculateExpired fields
from all documents in bmk_ctv_reconciliations collection.
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
    collection = db["bmk_ctv_reconciliations"]

    result = collection.update_many(
        {
            "$or": [
                {"contractExpiryDate": {"$exists": True}},
                {"contractOnboardDate": {"$exists": True}},
                {"isCalculateExpired": {"$exists": True}},
            ]
        },
        {
            "$unset": {
                "contractExpiryDate": "",
                "contractOnboardDate": "",
                "isCalculateExpired": "",
            }
        }
    )
    print(f"Cleaned up {result.modified_count} reconciliation document(s) by removing contractExpiryDate, contractOnboardDate, and isCalculateExpired fields.")
    client.close()

if __name__ == "__main__":
    cleanup()
