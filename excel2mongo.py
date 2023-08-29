from pymongo import MongoClient
import json
import pandas as pd

def excel2mongo(excel_path, mongo_uri, db_name, collection_name):
    df = pd.read_excel(excel_path, engine='openpyxl')

    client = MongoClient(mongo_uri)
    db = client[db_name]
    collection = db[collection_name]

    records = df.to_dict(orient='records')
    collection.insert_many(records)
    # for record in records:
    #     collection.update_one({'ID':record['ID']}, {'$set':record}, upsert=True)
    print(f"Inserted {len(records)} records into {db_name}.{collection_name}")

    for doc in collection.find():
        print(doc)

    client.close()

if __name__ == "__main__":
    excel_path = 'test.xlsx'
    mongo_uri = 'mongodb://localhost:27017/'
    db_name = 'testdb'
    collection_name = 'test'
    excel2mongo(excel_path, mongo_uri, db_name, collection_name)