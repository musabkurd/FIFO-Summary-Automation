# 🏭 FIFO Inventory Tracker

![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346?logo=microsoft-excel) ![Status](https://img.shields.io/badge/Status-Active-success)

**Automated FIFO expiry tracking system - One-click generation of expiry reports split by Regional Sales Manager.**

---

## 📊 What It Does

Converts raw inventory data into:
- ✅ Color-coded expiry reports (Red/Yellow/Green alerts)
- ✅ Individual reports per Regional Sales Manager
- ✅ Automatic email drafts for distribution
- ✅ Risk categorization (Expired → 3 months)

**Speed:** <30 seconds for complete processing

---

## 🚀 Quick Start

### Step 1: Generate Reports
1. Open your inventory file
2. Run: `FIFO_Auto_Summary_Split.bas`
3. Output: Master report + individual RSM files

### Step 2: Send Emails
1. Open email template
2. Run: `FIFO_Auto_Email.bas`
3. Output: Email drafts in Outlook

---

## 📁 Files

- `FIFO_Auto_Summary_Split.bas` - Creates categorized reports
- `FIFO_Auto_Email.bas` - Generates email notifications
- Template files included for structure

---

## 🎨 Risk Categories

| Color | Criteria | Priority |
|-------|----------|----------|
| 🔴 Red | Expired | URGENT |
| 🟠 Orange | < 1 month | HIGH |
| 🟡 Yellow | 1-2 months | MEDIUM |
| 🟢 Green | 2-3 months | LOW |

---

## 🔧 Requirements

- Excel 2016 or later
- Outlook (for email automation)
- Enable macros in Excel

---

## 📝 License

MIT License - Free to use and modify

---

⭐ **Star if useful!**
