import os
import re
from datetime import datetime, timezone
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "bmk_ctv")

if not MONGODB_URI:
    print("Error: MONGODB_URI is not set in environment variables.")
    exit(1)

def _parse_timestamp_from_key(file_key: str) -> str:
    """Try to extract timestamp from s3_key, or return current UTC time."""
    parts = file_key.split("_")
    if len(parts) >= 4 and re.match(r"^\d{9,12}$", parts[2]):
        try:
            ts = int(parts[2])
            return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat().replace("+00:00", "Z")
        except Exception:
            pass
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

def run_migration():
    print(f"Connecting to database '{DB_NAME}'...")
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_collaborators"]

    cursor = collection.find({})
    total_docs = 0
    updated_docs = 0

    for doc in cursor:
        total_docs += 1
        checklist = doc.get("checklist") or {}
        updates = {}
        unsets = {}

        # 1. CCCD file
        cccd = checklist.get("cccd") or {}
        cccd_file = cccd.get("file")
        if isinstance(cccd_file, str) and cccd_file.strip():
            updated_date = cccd.get("uploadedDate") or _parse_timestamp_from_key(cccd_file)
            updates["checklist.cccd.file"] = {"name": cccd_file, "updatedDate": updated_date}
        elif cccd_file is not None and not isinstance(cccd_file, dict):
            updates["checklist.cccd.file"] = None

        # 2. CKT file
        ckt = checklist.get("ckt") or {}
        ckt_file = ckt.get("file")
        if isinstance(ckt_file, str) and ckt_file.strip():
            updated_date = ckt.get("uploadedDate") or _parse_timestamp_from_key(ckt_file)
            updates["checklist.ckt.file"] = {"name": ckt_file, "updatedDate": updated_date}
        elif ckt_file is not None and not isinstance(ckt_file, dict):
            updates["checklist.ckt.file"] = None

        # 3. BBTL file
        bbtl = checklist.get("bbtl") or {}
        bbtl_file = bbtl.get("file")
        if isinstance(bbtl_file, str) and bbtl_file.strip():
            updated_date = bbtl.get("uploadedDate") or _parse_timestamp_from_key(bbtl_file)
            updates["checklist.bbtl.file"] = {"name": bbtl_file, "updatedDate": updated_date}
        elif bbtl_file is not None and not isinstance(bbtl_file, dict):
            updates["checklist.bbtl.file"] = None

        # 4. HDDV files
        hddv = checklist.get("hddv") or {}
        hddv_files = hddv.get("files") or []
        uploaded_dates_map = hddv.get("uploadedDates") or {}
        new_files = []
        files_modified = False

        for f in hddv_files:
            if isinstance(f, str):
                files_modified = True
                u_date = uploaded_dates_map.get(f) or _parse_timestamp_from_key(f)
                new_files.append({"name": f, "updatedDate": u_date})
            elif isinstance(f, dict):
                new_files.append(f)

        if files_modified:
            updates["checklist.hddv.files"] = new_files

        # 5. Clean up redundant old fields
        if "uploadedDate" in cccd:
            unsets["checklist.cccd.uploadedDate"] = ""
        if "uploadedDate" in ckt:
            unsets["checklist.ckt.uploadedDate"] = ""
        if "uploadedDate" in bbtl:
            unsets["checklist.bbtl.uploadedDate"] = ""
        if "uploadedDate" in hddv:
            unsets["checklist.hddv.uploadedDate"] = ""
        if "uploadedDates" in hddv:
            unsets["checklist.hddv.uploadedDates"] = ""

        # Perform atomic update if needed
        update_op = {}
        if updates:
            update_op["$set"] = updates
        if unsets:
            update_op["$unset"] = unsets

        if update_op:
            collection.update_one({"_id": doc["_id"]}, update_op)
            updated_docs += 1

    print(f"Migration completed! Checked {total_docs} records, updated {updated_docs} records.")
    client.close()

if __name__ == "__main__":
    run_migration()
