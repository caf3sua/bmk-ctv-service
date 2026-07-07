import os
import json
from datetime import datetime, timezone
from pymongo import MongoClient
from dotenv import load_dotenv

from app.core.security import hash_password

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
DB_NAME = os.getenv("DB_NAME", "bmk_ctv")

if not MONGODB_URI:
    print("Error: MONGODB_URI is not set in environment variables.")
    exit(1)

SEED_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "seed_data")

with open(os.path.join(SEED_DATA_DIR, "collaborators.json"), encoding="utf-8") as f:
    collaborators = json.load(f)

for item in collaborators:
    item["_id"] = item["employeeCode"]

_now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

users = [
    {
        "_id": "admin",
        "username": "admin",
        "name": "Quản trị viên",
        "email": "admin@bmk.vn",
        "role": "admin",
        "active": True,
        "hashedPassword": hash_password("123456"),
        "createdAt": _now,
        "updatedAt": _now,
    },
    {
        "_id": "staff",
        "username": "staff",
        "name": "Nhân viên",
        "email": "staff@bmk.vn",
        "role": "staff",
        "active": True,
        "hashedPassword": hash_password("123456"),
        "createdAt": _now,
        "updatedAt": _now,
    }
]

def seed_db():
    print(f"Connecting to database '{DB_NAME}'...")
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]

    print("Seeding collaborators...")
    db["bmk_ctv_collaborators"].delete_many({})
    db["bmk_ctv_collaborators"].insert_many(collaborators)
    print(f"Successfully seeded {len(collaborators)} collaborators.")

    print("Seeding users...")
    db["bmk_ctv_users"].drop()
    db["bmk_ctv_users"].create_index("username", unique=True)
    db["bmk_ctv_users"].create_index("email", unique=True)
    db["bmk_ctv_users"].insert_many(users)
    print(f"Successfully seeded {len(users)} users.")

    print("\nDatabase seeding completed successfully!")
    client.close()

if __name__ == "__main__":
    seed_db()
