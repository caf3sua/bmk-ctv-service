"""Script reset calContractExpiryDate về None trong collection bmk_ctv_reconciliations."""
import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "bmk_ctv")

if not MONGODB_URI:
    print("Error: MONGODB_URI is not set in environment variables.")
    exit(1)


def reset():
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]
    collection = db["bmk_ctv_reconciliations"]

    res = collection.update_many({}, {"$set": {"calContractExpiryDate": None}})
    print(f"Successfully reset calContractExpiryDate to None for {res.modified_count} documents.")
    client.close()


if __name__ == "__main__":
    reset()
