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
    },
    {
        "_id": "trangvpccpd",
        "username": "trangvpccpd",
        "name": "Nguyễn Thị Thu Trang",
        "email": "thutrangvpccpd@gmail.com",
        "role": "admin",
        "active": True,
        "hashedPassword": "pbkdf2:sha256:600000$f07dafe436ef09078375c98cef8fe54a$41b27df7dca4143680b2eca9a53d01932373ee3b313826906f2b9c56d590fede",
        "createdAt": "2026-07-26T11:02:25.551592Z",
        "updatedAt": "2026-07-26T11:02:25.551592Z",
    },
    {
        "_id": "hoainamtin2",
        "username": "hoainamtin2",
        "name": "Nguyễn Hoài Nam",
        "email": "hoainamtin2@gmail.com",
        "role": "admin",
        "active": True,
        "hashedPassword": "pbkdf2:sha256:600000$142f81c3a7887f2c508f51a196d75839$501942367f60a491a59479b5880056b0674b9df1fe40a3170356d41c1882fd73",
        "createdAt": "2026-07-26T11:02:57.756822Z",
        "updatedAt": "2026-07-26T11:02:57.756822Z",
    }
]

def seed_db():
    print(f"Connecting to database '{DB_NAME}'...")
    client = MongoClient(MONGODB_URI)
    db = client[DB_NAME]

    print("Seeding collaborators (using upsert)...")
    inserted_colls = 0
    updated_colls = 0
    for item in collaborators:
        res = db["bmk_ctv_collaborators"].replace_one(
            {"_id": item["_id"]},
            item,
            upsert=True
        )
        if res.matched_count > 0:
            updated_colls += 1
        else:
            inserted_colls += 1
    print(f"Successfully seeded collaborators: {inserted_colls} inserted, {updated_colls} updated.")

    print("Seeding users (using upsert)...")
    db["bmk_ctv_users"].create_index("username", unique=True)
    db["bmk_ctv_users"].create_index("email", unique=True)
    
    inserted_users = 0
    updated_users = 0
    for u in users:
        res = db["bmk_ctv_users"].replace_one(
            {"_id": u["_id"]},
            u,
            upsert=True
        )
        if res.matched_count > 0:
            updated_users += 1
        else:
            inserted_users += 1
    print(f"Successfully seeded users: {inserted_users} inserted, {updated_users} updated.")

    print("\nDatabase seeding completed successfully (non-destructive)!")
    client.close()

if __name__ == "__main__":
    seed_db()
