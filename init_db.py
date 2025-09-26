#!/usr/bin/env python3
"""Initialize the database with all tables."""

import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

from app import create_app, db

def init_db():
    """Initialize the database with all tables."""
    app = create_app()
    
    with app.app_context():
        print("Creating database tables...")
        db.create_all()
        print("Database tables created successfully!")
        
        # Check what tables were created
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        tables = inspector.get_table_names()
        print(f"Created tables: {tables}")

if __name__ == '__main__':
    init_db()