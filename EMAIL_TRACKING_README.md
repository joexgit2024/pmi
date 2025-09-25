# PMI Email Tracking & Management System

## Overview

This system manages acknowledgment emails for PMI PMDoS registrations with automatic tracking to prevent duplicates and enable incremental processing of new registrations.

## 🎯 Problem Solved

**Before:** Manual tracking, risk of duplicate emails, difficulty handling incremental registrations

**After:** Automatic tracking, duplicate prevention, seamless handling of new registrations

## 📁 File Structure

```
PMI/
├── email_drafts/                          # Original 29 sent emails
│   ├── 01_Maria_Mainhardt_email_draft.txt
│   ├── 02_Neha_Purohit_email_draft.txt
│   └── ... (29 files total)
│
├── new_email_drafts/                      # New incremental emails
│   └── 2025-09-26/                        # Date-based folders
│       ├── 30_Suresh_Reddy_email_draft.txt
│       ├── 31_Julia_Wright_email_draft.txt
│       ├── 32_Emma_Voss_email_draft.txt
│       ├── 33_Lindsey_Norris_email_draft.txt
│       ├── 34_Hadi_Ahmadi_email_draft.txt
│       └── NEW_EMAILS_SUMMARY.md
│
├── email_tracking.json                    # Master tracking database
├── email_tracking_system.py              # Core tracking system
├── enhanced_email_generator.py           # Enhanced email generator
├── email_manager.py                      # Management dashboard
└── dynamic_file_loader.py                # Auto file detection
```

## 🚀 Quick Start

### Check Status
```bash
python email_manager.py status
```

### Generate Emails for New Registrations
```bash
python email_manager.py generate
```

### Generate Detailed Report
```bash
python email_manager.py report
```

## 📊 Current Status (as of 2025-09-26)

- **Total Emails Sent:** 34
- **Original Batch:** 29 emails (in `email_drafts/`)
- **New Batch:** 5 emails (in `new_email_drafts/2025-09-26/`)
- **Registration File:** Auto-detected latest file with 33 registrations
- **New Registrations:** 0 (all current registrations handled)

## 🔄 Workflow for Future Registrations

1. **New registrations arrive** → Upload new Excel file to `input/` folder
2. **Check status** → `python email_manager.py status`
3. **Generate emails** → `python email_manager.py generate`
4. **Review drafts** → Check `new_email_drafts/[DATE]/` folder  
5. **Send emails** → System automatically tracks as sent

## 📋 Key Features

### ✅ Automatic Tracking
- Tracks all sent emails with timestamps
- Prevents duplicate email generation
- Maintains persistent record across runs

### ✅ Incremental Processing
- Identifies only NEW registrations
- Generates drafts only for unprocessed registrations
- Organizes new drafts in date-based folders

### ✅ Dynamic File Detection
- Automatically finds latest registration files
- Works with any filename format/convention
- Ignores temporary Excel files

### ✅ Comprehensive Reporting
- Real-time status dashboard
- Detailed batch tracking
- Summary reports for each batch

## 📁 Email Tracking Database

The `email_tracking.json` file maintains:

```json
{
  "metadata": {
    "total_emails_sent": 34,
    "batches": [
      {
        "batch_id": "initial_batch_29", 
        "date": "2025-09-26",
        "count": 29,
        "folder": "email_drafts"
      },
      {
        "batch_id": "batch_20250926_01",
        "date": "2025-09-26", 
        "count": 5,
        "folder": "new_email_drafts/2025-09-26"
      }
    ]
  },
  "sent_emails": {
    "email@domain.com": {
      "name": "Person Name",
      "sent_date": "2025-09-26",
      "batch_id": "batch_20250926_01",
      "draft_file": "path/to/email/draft.txt"
    }
  }
}
```

## 🔧 Advanced Usage

### Initialize Tracking (First Time Only)
```python
from email_tracking_system import EmailTracker
tracker = EmailTracker()
tracker.initialize_from_existing_drafts()
```

### Manual Email Generation
```python
from email_tracking_system import create_incremental_email_drafts
create_incremental_email_drafts()
```

### Check for New Registrations
```python
from email_tracking_system import EmailTracker
tracker = EmailTracker()
# Load latest data and check
new_registrations = tracker.identify_new_registrations(df)
print(f"New registrations: {len(new_registrations)}")
```

## 🎯 Benefits

1. **No Duplicate Emails** - System tracks all sent emails
2. **Incremental Processing** - Only processes new registrations  
3. **Organized Storage** - Date-based folders for new batches
4. **Comprehensive Tracking** - Full audit trail of all emails
5. **Future-Proof** - Handles any naming convention for input files
6. **Easy Management** - Simple command-line interface

## 📈 Scalability

The system is designed to handle:
- ✅ Multiple batches per day
- ✅ Any number of new registrations
- ✅ Different file naming conventions
- ✅ Long-term tracking across months/years
- ✅ Easy reporting and auditing

## 🆘 Troubleshooting

### No New Emails Generated
- Check if all registrations already have tracking records
- Verify latest Excel file is in `input/` folder
- Run `python email_manager.py status` to see current state

### Tracking Data Missing
- Run the system once to auto-initialize from existing drafts
- Check if `email_tracking.json` exists and is readable

### File Permission Issues
- Close Excel files before running the system
- Ensure write permissions to create new folders

---

## 📞 Summary

This system successfully solved your challenge by:

✅ **Identifying the 5 new registrations** (Suresh Reddy, Julia Wright, Emma Voss, Lindsey Norris, Hadi Ahmadi)

✅ **Creating separate folder structure** for new emails (`new_email_drafts/2025-09-26/`)

✅ **Maintaining tracking records** of all 34 sent emails (29 original + 5 new)

✅ **Future-proofing** for any additional registrations with automatic detection

The system is now ready for ongoing use and will seamlessly handle future registration batches!