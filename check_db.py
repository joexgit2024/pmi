#!/usr/bin/env python3
import sqlite3
import os

def check_database():
    db_path = 'instance/pmi.db'
    
    if not os.path.exists(db_path):
        print("❌ Database file not found!")
        return
    
    print(f"✅ Database file exists at: {db_path}")
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    
    print(f"\n📊 Database tables ({len(tables)} total):")
    for table in tables:
        table_name = table[0]
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
        count = cursor.fetchone()[0]
        print(f"  - {table_name}: {count} records")
    
    # Check file_uploads specifically
    print("\n📁 Recent file uploads:")
    cursor.execute("""
        SELECT id, filename, file_type, status, upload_date 
        FROM file_uploads 
        ORDER BY upload_date DESC 
        LIMIT 5
    """)
    uploads = cursor.fetchall()
    
    if uploads:
        for upload in uploads:
            print(f"  ID: {upload[0]}, File: {upload[1]}, Type: {upload[2]}, Status: {upload[3]}, Date: {upload[4]}")
    else:
        print("  No file uploads found")
    
    conn.close()
    print("\n✅ Database check complete!")

if __name__ == "__main__":
    check_database()