# PMI System - Quick Instructions

## 🎯 Which Python File to Run?

### **Need PMP-Charity Matching?**
```bash
python run_complete_analysis.py
```
**What you get:**
- Excel file with charity assignments
- LinkedIn analysis report
- Matching summary CSV

---

### **Need to Send Acknowledgment Emails?**
```bash
python email_manager.py generate
```
**What you get:**
- Email drafts for NEW registrations only
- Personalized emails with names and addresses
- Automatic tracking to prevent duplicates

---

## 📋 Other Useful Commands

### Check Email Status
```bash
python email_manager.py status
```
Shows who has been sent emails already and identifies new registrations.

### Check System Status
```bash
python email_manager.py
```
Quick overview of current situation.

---

## 🔄 Typical Workflow

**When new Excel registration file arrives:**

1. **Upload file** to `input/` folder
2. **Generate matching:** `python run_complete_analysis.py`
3. **Send emails:** `python email_manager.py generate`
4. **Done!**

---

## 📁 Output Locations

- **Matching results:** Root folder (Excel files)
- **Email drafts:** `new_email_drafts/[date]/` folder
- **Original emails:** `email_drafts/` folder (don't touch)

---

## ❓ Quick Troubleshooting

**Problem: No new emails generated**
- Check: `python email_manager.py status` 
- Likely: All current registrations already have emails

**Problem: Matching fails**
- Check: Excel file is in `input/` folder
- Check: File is not open in Excel

---

## 🎯 Summary

**2 Main Commands:**
1. `python run_complete_analysis.py` → **Charity matching**
2. `python email_manager.py generate` → **New emails**

That's it! 🚀