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

### Setup (One-time)

**Import VBA Code:**
1. Download the 2 VBA files from this repo
2. Open your Excel file
3. Press `Alt + F11` (opens VBA Editor)
4. Go to **File** → **Import File...**
5. Select `#1 FIFO Auto Summary, Split VBA.vba`
6. Repeat for `#2 FIFO Auto Email VBA.txt`

### Step 1: Generate Reports
1. Open your inventory file (use template provided)
2. If running from `PERSONAL.XLSB` or the Quick Access Toolbar, make sure the FIFO workbook is the **active workbook**
3. Press `Alt + F8` → Select `FIFO_ULTIMATE_OneClick`
4. Click **Run**
5. Output: Master report + individual RSM files created

### Step 2: Send Emails
1. Open `Email Template FIFO.xlsx`
2. Press `Alt + F8` → Select email macro
3. Click **Run**
4. Output: Email drafts in Outlook (ready to review & send)

---

## 📁 Files in This Repo

| File | Purpose |
|------|---------|
| `#1 FIFO Auto Summary, Split VBA.vba` | Main automation - Import to Excel |
| `#2 FIFO Auto Email VBA.txt` | Email generator - Import to Excel |
| `Email Template FIFO [template].xlsx` | Template for emails |
| `FIFO_Expiry_Report [TEMPLATE].xlsx` | Input data template |
| `SAP Code VBA-0 [TEMPLATE].xlsx` | SAP lookup reference |

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
- **Enable macros:** File → Options → Trust Center → Enable all macros

---

## 💡 Troubleshooting

**Problem:** Macro not found  
**Solution:** Make sure you imported the VBA files (Step 1 in Setup)

**Problem:** Could not find sheet named `Total`  
**Solution:** Activate the correct FIFO workbook first. The active workbook must contain the `Total` sheet.

**Problem:** Security warning  
**Solution:** Enable macros in Excel settings

**Problem:** Email drafts not created  
**Solution:** Open Outlook before running the email macro

---

## 📝 License

MIT License - Free to use and modify

---

⭐ **Star if useful!**
## PERSONAL.XLSB Quick Setup

Exactly. That workflow is correct.

Your clean version is:

1. Unhide `PERSONAL.XLSB`
2. Press `Alt + F11`
3. Paste the VBA into a module in `PERSONAL.XLSB`
4. Replace `ThisWorkbook` with `ActiveWorkbook` where needed
5. Save
6. Hide `PERSONAL.XLSB` again
7. In Excel Options > Quick Access Toolbar
8. Change command list to `Macros`
9. Add the macro you want
10. Click `Modify...` to change the name and icon
11. Done

That is the best setup for your daily Excel automation across all files.

