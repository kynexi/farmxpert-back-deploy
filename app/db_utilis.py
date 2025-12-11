from pymongo import MongoClient
import os

DATABASE_URL = os.environ.get("DATABASE_URL")
print(f"DEBUG: DATABASE_URL is set: {bool(DATABASE_URL)}")
print(f"DEBUG: DATABASE_URL value: {DATABASE_URL[:50] if DATABASE_URL else 'NOT SET'}...")

if not DATABASE_URL:
    raise ValueError("DATABASE_URL environment variable not found in Vercel")

client = MongoClient(DATABASE_URL)
db = client["farmxpert"]