import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "bmk_ctv")

if not MONGODB_URI:
    print("Error: MONGODB_URI is not set in environment variables.")
    exit(1)

def run_migration():
    print(f"Connecting to database '{DB_NAME}'...")
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_collaborators"]

    print("Checking database records...")
    cursor = collection.find({})
    migrated_count = 0
    already_migrated = 0

    for doc in cursor:
        checklist = doc.get("checklist") or {}

        # If cccd exists as a dictionary, it's already migrated
        if "cccd" in checklist and isinstance(checklist["cccd"], dict):
            already_migrated += 1
            continue

        # Extract old values safely
        submitted_id_card = checklist.get("submittedIdCard", False)
        submitted_tax_commitment = checklist.get("submittedTaxCommitment", False)
        service_contracts = checklist.get("serviceContracts", [{"startDate": None, "endDate": None}])
        liquidation_date = checklist.get("liquidationDate")

        id_card_file = checklist.get("idCardFile")
        tax_commitment_file = checklist.get("taxCommitmentFile")
        service_contract_file = checklist.get("serviceContractFile")
        liquidation_file = checklist.get("liquidationFile")

        # Build new structured checklist
        new_checklist = {
            "cccd": {
                "checked": submitted_id_card,
                "file": id_card_file
            },
            "ckt": {
                "checked": submitted_tax_commitment,
                "file": tax_commitment_file
            },
            "hddv": {
                "contract_date": service_contracts,
                "files": [service_contract_file] if service_contract_file else []
            },
            "bbtl": {
                "date": liquidation_date,
                "file": liquidation_file
            }
        }

        # Update the document with new checklist and clean up top-level old file fields if they existed outside checklist
        collection.update_one(
            {"_id": doc["_id"]},
            {"$set": {"checklist": new_checklist}}
        )
        migrated_count += 1

    print(f"Migration finished. Migrated: {migrated_count} records, Already Migrated: {already_migrated} records.")
    client.close()

if __name__ == "__main__":
    run_migration()
