import asyncio
import os
import sys
import re

# Add project root to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from app.core.database import connect_db, close_db, db_instance

COLLECTION = "bmk_ctv_activity_logs"
TARGET_ACTIONS = [
    "create_collaborator",
    "update_collaborator",
    "delete_collaborator",
    "upload_collaborator_document",
    "delete_collaborator_document",
]

def extract_employee_code_from_message(message: str) -> str | None:
    if not message:
        return None
    match = re.search(r'cộng tác viên (?:mã\s+)?(\S+)', message, re.IGNORECASE)
    if match:
        code = match.group(1).strip()
        # Clean trailing punctuation if any
        code = code.rstrip('.,;:)(')
        return code if code else None
    return None

async def run_migration():
    print("Connecting to MongoDB...")
    await connect_db()
    db = db_instance.db

    query = {"action": {"$in": TARGET_ACTIONS}}
    cursor = db[COLLECTION].find(query)

    total_scanned = 0
    updated_count = 0
    skipped_count = 0

    async for doc in cursor:
        total_scanned += 1
        doc_id = doc["_id"]
        current_code = doc.get("employeeCode")
        message = doc.get("message", "")

        extracted_code = extract_employee_code_from_message(message)

        if extracted_code and current_code != extracted_code:
            await db[COLLECTION].update_one(
                {"_id": doc_id},
                {"$set": {"employeeCode": extracted_code}}
            )
            updated_count += 1
            print(f"[UPDATED] Log ID {doc_id} (action: {doc.get('action')}) -> employeeCode: '{extracted_code}'")
        else:
            skipped_count += 1

    print("\n================ MIGRATION SUMMARY ================")
    print(f"Total target logs scanned : {total_scanned}")
    print(f"Total updated logs        : {updated_count}")
    print(f"Total skipped logs        : {skipped_count}")
    print("===================================================\n")

    await close_db()

if __name__ == "__main__":
    asyncio.run(run_migration())
