from pymongo import MongoClient
import os

client = MongoClient(os.environ["DATABASE_URL"])
db = client["farmxpert"]
