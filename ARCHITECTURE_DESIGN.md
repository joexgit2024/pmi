"""
PMI Web Application - Enterprise Architecture Design
==================================================

Project Structure:
/
├── app.py                          # Application entry point
├── config.py                       # Configuration management
├── requirements.txt               # Dependencies
├── instance/                      # Instance-specific files
│   ├── config.py                 # Local configuration
│   └── pmi.db                    # SQLite database
├── app/                          # Main application package
│   ├── __init__.py               # Flask app factory
│   ├── models/                   # Database models
│   │   ├── __init__.py
│   │   ├── registration.py       # Registration data models
│   │   ├── charity.py           # Charity data models
│   │   ├── matching.py          # Matching results models
│   │   ├── email_tracking.py    # Email tracking models
│   │   └── file_upload.py       # File upload tracking
│   ├── routes/                   # Route handlers (Blueprints)
│   │   ├── __init__.py
│   │   ├── main.py              # Dashboard & main pages
│   │   ├── upload.py            # File upload handling
│   │   ├── matching.py          # Matching operations
│   │   ├── email.py             # Email operations
│   │   └── api.py               # JSON API endpoints
│   ├── services/                 # Business logic layer
│   │   ├── __init__.py
│   │   ├── file_service.py      # File processing service
│   │   ├── matching_service.py   # Matching algorithm service
│   │   ├── email_service.py     # Email generation service
│   │   └── data_service.py      # Data analysis service
│   ├── utils/                    # Utility functions
│   │   ├── __init__.py
│   │   ├── file_utils.py        # File handling utilities
│   │   ├── validation.py        # Data validation
│   │   ├── excel_parser.py      # Excel file processing
│   │   └── background_tasks.py   # Background job utilities
│   ├── templates/               # Jinja2 templates
│   │   ├── base.html            # Base template
│   │   ├── index.html           # Dashboard
│   │   ├── upload/              # Upload pages
│   │   │   ├── files.html       # File upload page
│   │   │   └── progress.html    # Upload progress
│   │   ├── matching/            # Matching pages
│   │   │   ├── dashboard.html   # Matching dashboard
│   │   │   ├── results.html     # Results display
│   │   │   └── download.html    # Download results
│   │   ├── email/               # Email pages
│   │   │   ├── dashboard.html   # Email dashboard
│   │   │   ├── generate.html    # Generate emails
│   │   │   └── history.html     # Email history
│   │   └── components/          # Reusable components
│   │       ├── navbar.html      # Navigation
│   │       ├── alerts.html      # Alert messages
│   │       └── progress.html    # Progress bars
│   └── static/                  # Static assets
│       ├── css/
│       │   ├── main.css         # Custom styles
│       │   └── upload.css       # Upload-specific styles
│       ├── js/
│       │   ├── main.js          # Main JavaScript
│       │   ├── upload.js        # Upload handling
│       │   ├── matching.js      # Matching operations
│       │   └── email.js         # Email operations
│       └── uploads/             # Temporary file storage
└── migrations/                   # Database migrations
    └── versions/

Key Features:
============

1. DASHBOARD (/)
   - Overview of system status
   - Quick stats (total registrations, emails sent, etc.)
   - Recent activity feed
   - Quick action buttons

2. FILE UPLOAD (/upload)
   - Drag & drop interface for Excel files
   - File validation and preview
   - Progress indicators
   - File history and management

3. MATCHING OPERATIONS (/matching)
   - Run matching algorithms
   - View matching results
   - Download Excel reports
   - Matching history

4. EMAIL MANAGEMENT (/email)
   - Generate new emails
   - View email status
   - Download email drafts
   - Email sending history

5. API ENDPOINTS (/api)
   - RESTful API for all operations
   - JSON responses for AJAX calls
   - Background job status
   - Data export endpoints

Database Schema:
===============

1. file_uploads
   - id, filename, file_type, upload_date, file_path, status, user_id

2. registrations
   - id, file_upload_id, first_name, last_name, email, pmi_id, 
     linkedin_url, experience, areas_of_interest, created_at

3. charities
   - id, file_upload_id, organization, initiative, description,
     requirements, priority_level, complexity, created_at

4. matching_results
   - id, registration_id, charity_id, match_score, linkedin_quality,
     created_at, batch_id

5. email_tracking
   - id, registration_id, email_address, sent_date, batch_id,
     draft_file_path, status

6. background_jobs
   - id, job_type, status, progress, result, created_at, completed_at

Security & Best Practices:
=========================

1. Input Validation
   - File type validation
   - Size limits
   - Content validation

2. Error Handling
   - Comprehensive error logging
   - User-friendly error messages
   - Graceful degradation

3. Performance
   - Background job processing
   - Database indexing
   - File caching

4. Scalability
   - Modular architecture
   - Service layer separation
   - Easy database migration

5. Maintainability
   - Clear separation of concerns
   - Comprehensive logging
   - Unit test structure
"""