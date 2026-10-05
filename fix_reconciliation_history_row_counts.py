"""One-off fix script: update bmk_ctv_reconciliation_history collection to fix totalRows and failedRows
for past runs where empty trailing Excel rows were incorrectly counted as failed rows.
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

def fix_history():
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_reconciliation_history"]

    cursor = collection.find({"failedRows": {"$gt": 0}})
    fixed = 0
    for doc in cursor:
        success = doc.get("successRows", 0)
        total = doc.get("totalRows", 0)
        failed = doc.get("failedRows", 0)
        
        # If total rows was artificially inflated due to empty rows
        if failed > 0 and total > success:
            new_total = success
            new_failed = 0
            collection.update_one(
                {"_id": doc["_id"]},
                {"$set": {"totalRows": new_total, "failedRows": new_failed}}
            )
            fixed += 1

    print(f"Fixed {fixed} reconciliation history document(s).")
    client.close()

if __name__ == "__main__":
    fix_history()
