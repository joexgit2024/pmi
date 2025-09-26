# PMDoS 2025 Selection Notification System

## 📧 Complete Email Notification System for Selection Results

This system handles personalized email notifications for all three PMDoS 2025 selection outcomes:
- **Selected and Matched** - Full participants with project assignments
- **Selected as Backup** - Reserve participants with induction access
- **Not Selected** - Respectful notifications with future opportunities

---

## 🚀 Quick Start

### 1. Generate All Notification Emails
```bash
python generate_selection_notification_emails.py
```

### 2. Update Selected Participants with Project Details (Optional)
```bash
python update_project_assignments.py
```

### 3. Send Emails
- Navigate to each folder in `selection_notifications/`
- Copy-paste individual drafts into Outlook
- Use tracking checklist in `SELECTION_SUMMARY.md`

---

## 📁 System Structure

```
selection_notifications/
├── selected_and_matched/           # 33 congratulatory emails
│   ├── 01_Maria_Mainhardt_notification.txt
│   ├── 02_Neha_Purohit_notification.txt
│   └── ... (all selected participants)
├── selected_as_backup/             # 33 backup status emails  
│   ├── 01_Maria_Mainhardt_notification.txt
│   └── ... (all backup participants)
├── not_selected/                   # 33 respectful notifications
│   ├── 01_Maria_Mainhardt_notification.txt
│   └── ... (all non-selected applicants)
├── SELECTION_SUMMARY.md            # Complete tracking and instructions
├── selected_and_matched_template.txt
├── selected_as_backup_template.txt
└── not_selected_template.txt
```

---

## 📊 Input Files Required

Place these Excel files in the `input/` directory:
- `selected and matched.xlsx` - Full participants
- `seected as backup.xlsx` - Backup participants  
- `not selected.xlsx` - Non-selected applicants

Each file should contain columns:
- First Name
- Last Name  
- Preferred Email Address

---

## 📧 Email Templates Overview

### 1. Selected and Matched Template
**Subject:** "PMDoS 2025: Congratulations - You're In! Next Steps Inside"

**Key Elements:**
- 🎉 Congratulations message
- 📅 Mandatory induction: October 15, 2025 (Google Meet)
- 🎯 Project assignment details (filled by update script)
- 📋 Commitment requirements and next steps

### 2. Selected as Backup Template  
**Subject:** "PMDoS 2025: Selected as Backup - Important Next Steps"

**Key Elements:**
- ⭐ Backup status explanation
- 📅 Induction invitation: October 15, 2025 (Google Meet)  
- 🔄 Activation process (Oct 16-25, 2025)
- 🎯 Priority for future events

### 3. Not Selected Template
**Subject:** "PMDoS 2025: Selection Update and Future Opportunities"

**Key Elements:**
- 💝 Respectful and empathetic notification
- 📊 Context about competitive selection process
- 🔄 Future opportunities with PMI Sydney Chapter
- 🤝 Alternative ways to contribute to community

---

## 🗓️ Key Dates and Timeline

### October 2025
- **October 1:** Confirmation deadline for selected participants
- **October 15:** Mandatory induction session (7:00-8:30 PM AEDT via Google Meet)
- **October 16-25:** Backup activation period

### November 2025  
- **November 6:** PMDoS 2025 main event (9:00 AM - 5:00 PM)

---

## ⚡ Usage Instructions

### For Each Email Draft:

1. **Open** individual `.txt` file from appropriate folder
2. **Copy** entire content (Ctrl+A, Ctrl+C)  
3. **Log into** pmdos_professionals@pmisydney.org Outlook
4. **Create new email** and paste content
5. **Add subject line** (see templates above)
6. **Send** the email
7. **Check** the tracking box in `SELECTION_SUMMARY.md`

### Google Meet Setup for Induction:

1. **Create** Google Meet for October 15, 2025, 7:00-8:30 PM AEDT
2. **Prepare** meeting materials:
   - Project briefings for selected participants
   - Backup activation process explanation
   - Event logistics and requirements
3. **Send** meeting link to Selected and Backup groups 24 hours before
4. **Record** session for participants who cannot attend live

---

## 🔧 Advanced Features

### Project Assignment Integration

If you've run the matching analysis, use the update script to automatically fill project details:

```bash
# Run matching first (if not done)
python run_complete_analysis.py

# Then update selected emails with project assignments  
python update_project_assignments.py
```

This will replace placeholders with:
- Specific charity organization names
- Project initiative details
- Team partner information

### Customization

**Email Templates:** Edit template files to modify standard content
**Tracking:** Use `SELECTION_SUMMARY.md` for comprehensive progress tracking
**Automation:** Scripts handle personalization automatically

---

## 📋 Checklist for Organizers

### Pre-Sending (October 1-5)
- [ ] Verify all three Excel files contain correct participant lists
- [ ] Run email generation script
- [ ] Review sample emails from each category
- [ ] Update selected participants with project assignments
- [ ] Create Google Meet for October 15 induction

### Sending Phase (October 5-10)  
- [ ] Send all "Selected and Matched" emails first
- [ ] Send all "Selected as Backup" emails  
- [ ] Send all "Not Selected" emails last
- [ ] Track sent emails using SELECTION_SUMMARY.md checklist

### Post-Sending (October 10-15)
- [ ] Monitor confirmation responses from selected participants
- [ ] Send Google Meet link 24 hours before induction
- [ ] Prepare backup activation process
- [ ] Ready project materials for induction session

---

## 📞 Support and Troubleshooting

### Common Issues:
- **Missing project details:** Run `update_project_assignments.py` after matching analysis
- **Email formatting:** Check individual `.txt` files for proper personalization
- **File not found:** Ensure Excel files are in `input/` directory with correct names

### Contact:
- **Technical Issues:** Check script output for error messages
- **Content Questions:** Review template files for messaging
- **Process Support:** Use `SELECTION_SUMMARY.md` for step-by-step guidance

---

## 🎯 Success Metrics

- **99 personalized emails** generated across all three categories
- **Professional communication** maintaining PMI Sydney Chapter standards
- **Clear next steps** for each participant group
- **Comprehensive tracking** system for organizer efficiency

This system ensures respectful, professional, and complete communication with all PMDoS 2025 applicants regardless of their selection outcome.