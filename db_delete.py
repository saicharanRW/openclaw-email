import sqlite3
import shutil

original_db = "document.db"
new_db = "document_index_clean2.db"

# Copy DB so original stays safe
shutil.copy(original_db, new_db)

conn = sqlite3.connect(new_db)
cursor = conn.cursor()

print("Connected to database:", new_db)

# Find documents starting with Aadan or traffic
cursor.execute("""
SELECT id, file_name
FROM documents
WHERE file_name LIKE 'Aadan%'
   OR file_name LIKE 'traffic%'
""")

docs = cursor.fetchall()

print(f"Found {len(docs)} documents to delete")

document_ids = []

for doc_id, name in docs:
    print(name)
    document_ids.append(doc_id)

# Delete related chunks first
if document_ids:
    placeholders = ",".join("?" * len(document_ids))

    cursor.execute(f"""
    DELETE FROM document_chunks
    WHERE document_id IN ({placeholders})
    """, document_ids)

    cursor.execute(f"""
    DELETE FROM documents
    WHERE id IN ({placeholders})
    """, document_ids)

conn.commit()

print("Deleted Aadan and traffic files successfully")

# Verify remaining records
cursor.execute("SELECT COUNT(*) FROM documents")
print("Remaining documents:", cursor.fetchone()[0])

cursor.execute("SELECT COUNT(*) FROM document_chunks")
print("Remaining chunks:", cursor.fetchone()[0])

conn.close()

print("Cleaned DB saved as:", new_db)