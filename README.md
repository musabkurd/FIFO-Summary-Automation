# 🏭 FIFO Auto Summary & Split

![Status](https://img.shields.io/badge/Status-Production-success)
![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346?logo=microsoft-excel)
![Speed](https://img.shields.io/badge/Speed-27s-blue)
![Scale](https://img.shields.io/badge/Scale-500%2B_Products-orange)

**One-click FIFO expiry tracking system that segments 500+ products by expiry risk and auto-generates RSM-specific reports in <30 seconds.**

---

## 📈 Impact

| Metric | Before | After | Improvement |
|--------|--------|-------|-------------|
| **Time** | 2 days | 27 seconds | **99.8% faster** |
| **Accuracy** | ~85% | 100% | **Zero errors** |
| **Scale** | Manual | 500+ products | **Fully automated** |
| **Reports** | 1 master | 12+ RSM files | **Auto-segmented** |

---

## 🎯 Problem Solved

### Manual Process Pain Points:
- ❌ Opening 42+ warehouse files individually
- ❌ Manually calculating days to expiry (high error rate)
- ❌ Categorizing items into 4 risk tiers by hand
- ❌ Creating separate reports per Regional Sales Manager
- ❌ Dealing with mixed Arabic/English data
- ❌ **Result:** 2 days per cycle + frequent calculation errors

---

## ✨ Solution

### **Two-Stage Automation Pipeline**

```
┌─────────────────┐
│  Raw SAP Data   │
│  (500+ items)   │
└────────┬────────┘
         │
         ▼
┌─────────────────────────┐
│  Stage 1: Categorizer   │
│  • Calculate expiry     │
│  • Risk categorization  │
│  • Color-coded sheets   │
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────┐
│  Stage 2: RSM Splitter  │
│  • Filter by manager    │
│  • Generate 12+ files   │
│  • Rank top items       │
└─────────────────────────┘
         │
         ▼
    📊 Reports Ready
```

### **Stage 1: Master Report Generator**
Processes raw inventory data:
- ✅ Reads 19-column "Total" sheet
- ✅ Loads SAP distributor codes
- ✅ Calculates `DaysToExpiry = ExpiryDate - Today`
- ✅ **Auto-categorizes into 4 risk tiers:**

| Tier | Color | Criteria | Action |
|------|-------|----------|--------|
| 🔴 **Expired** | Red | Past expiry | URGENT: Remove from stock |
| 🟠 **< 1 Month** | Light Red | ≤30 days | HIGH: Immediate promotions |
| 🟡 **< 2 Months** | Orange | 31-60 days | MEDIUM: Plan sales push |
| 🟢 **< 3 Months** | Yellow | 61-90 days | LOW: Monitor closely |

**Output:** `FIFO_Expiry_Report_[dd-mmm-yyyy].xlsx`

---

### **Stage 2: RSM Splitter**
Creates personalized reports:
- ✅ Filters by RSM Name (Column 17)
- ✅ Creates timestamped folder
- ✅ Per-RSM file contains:
  - All 5 category sheets (filtered)
  - Summary with top distributors ranked by quantity
  - Recalculated totals
  - Professional formatting

**Output:** `FIFO_Per_RSM_[timestamp]/` with 12+ individual files

---

## 🚀 One-Click Execution

```vba
Sub FIFO_ULTIMATE_OneClick()
    ' Runs both stages sequentially
    ' Total execution: ~27 seconds
End Sub
```

**That's it.** No parameters, no configuration needed.

---

## 🏗️ Technical Architecture

### **Input Requirements**

| File | Purpose | Critical Columns |
|------|---------|------------------|
| **Main Workbook** | Raw inventory | "Total" sheet with 19 columns |
| **SAP Lookup** | `SAP Code VBA-0.xlsx` | Maps distributor → SAP codes |

### **Data Flow Diagram**

```
📥 Total Sheet (raw inventory)
    │
    ├─► [Col 14] ExpiryDate → Calculate DaysToExpiry
    ├─► [Col 17] RSM Name → Split by manager
    ├─► [Col 19] Category → Risk tier assignment
    └─► [Col 13] ItemQty → Sum totals
    │
    ▼
📊 Master Report (5 sheets, color-coded)
    │
    ├─► Filter by unique RSM
    └─► Generate individual files
    │
    ▼
📁 FIFO_Per_RSM_[timestamp]/
    ├─► RSM_Ahmad_31-Dec-2025.xlsx
    ├─► RSM_Karwan_31-Dec-2025.xlsx
    └─► ... (12+ files)
```

### **Key Features**

| Feature | Implementation | Benefit |
|---------|---------------|---------|
| **Auto-detection** | No hardcoded paths | Works in any folder |
| **Unicode support** | `ChrW()` for Arabic | Mixed language data |
| **Performance** | Dictionary lookups O(1) | Fast SAP matching |
| **Ranking** | Bubble sort algorithm | Top distributors auto-sorted |
| **Error handling** | Try-catch all operations | Never crashes |

---

## ⚡ Performance

**Real Production Benchmark:**

```
Input:  500+ products × 42 warehouses = 21,000+ rows
Output: 1 master report + 12 RSM files
Time:   27.2 seconds
```

**Optimization Techniques:**
```vba
Application.ScreenUpdating = False      ' Skip UI updates
Application.Calculation = xlCalculationManual  ' Defer formulas
Application.DisplayAlerts = False       ' No popup dialogs
```

---

## 📂 Output Structure

```
📁 Project Root/
│
├── 📄 FIFO_Expiry_Report_31-Dec-2025.xlsx  ← Master report
│
└── 📁 FIFO_Per_RSM_31-12-2025_14-30-45/    ← Timestamped folder
    ├── 📊 FIFO_Report_RSM_Ahmad_31-Dec-2025.xlsx
    ├── 📊 FIFO_Report_RSM_Karwan_31-Dec-2025.xlsx
    ├── 📊 FIFO_Report_RSM_Saman_31-Dec-2025.xlsx
    ├── ... (12+ RSM files)
    └── 📊 Summary per RSM.xlsx  ← Top distributors ranked
```

---

## 🎮 Usage

### **Quick Start**

1. **Prepare data:**
   ```
   Paste raw SAP inventory into "Total" sheet
   ```

2. **Run automation:**
   ```vba
   FIFO_ULTIMATE_OneClick()
   ```

3. **Check output:**
   ```
   Open FIFO_Per_RSM_[timestamp]/ folder
   ```

### **Example Execution**

```
▶ Running FIFO automation...
  ✓ Reading 523 products from Total sheet
  ✓ Loading SAP codes (42 distributors)
  ✓ Calculating expiry dates...
  ✓ Categorizing: 12 expired, 45 <1mo, 89 <2mo, 134 <3mo
  ✓ Creating master report: FIFO_Expiry_Report_31-Dec-2025.xlsx
  ✓ Splitting by RSM (12 managers detected)
  ✓ Generating individual files...
  ✓ Creating summary rankings...
  
✅ Complete! (27.2 seconds)
📁 Output: FIFO_Per_RSM_31-12-2025_14-30-45/
```

---

## 🛡️ Error Handling

| Scenario | Behavior |
|----------|----------|
| Missing "Total" sheet | Alert + graceful exit |
| SAP file not found | Proceeds without codes |
| Invalid dates | Defaults to `daysRemaining = 999` |
| Master report failed | Clear error message |
| Non-matching RSM | Auto-filtered out |

---

## 🔧 Tech Stack

**Core Technologies:**
- **Excel VBA** (2016+)
- **FileSystemObject** - Folder/file operations
- **Scripting.Dictionary** - O(1) lookups
- **Unicode handling** - Arabic text support
- **Dynamic arrays** - Sorting & ranking

**Algorithms:**
- Date calculations (`DateValue`)
- Dictionary-based deduplication
- Bubble sort for top-N ranking
- Memory-efficient bulk operations

---

## 🏢 Business Context

| Attribute | Value |
|-----------|-------|
| **Company** | Karwanchi, Kurdistan (Erbil) |
| **Department** | Stock & Credit Control |
| **Users** | 12+ Regional Sales Managers |
| **Frequency** | Daily during high-volume cycles |
| **Coverage** | 42 warehouses, 500+ SKUs |
| **Impact** | Prevents waste + distributor complaints |

---

## 📊 KPIs Tracked

- 🔢 **Product Count** - Unique SKUs per category
- 📦 **Total Quantity** - Sum of units at risk
- 🏢 **Warehouse Distribution** - Items per location
- ⏰ **Expiry Timeline** - Days remaining per item
- 👤 **RSM Load** - Products assigned per manager

---

## 📝 Version History

| Version | Date | Changes |
|---------|------|---------|
| **v2.0** | Jan 2025 | Production release (auto-email + RSM split) |
| **v1.0** | Dec 2024 | Initial FIFO categorization |

---

## 🤝 Contributing

Currently internal tool. For questions:
- **Owner:** Musab - Stock & Credit Controller
- **Location:** Karwanchi, Kurdistan (Erbil)
- **Support:** Mohammed IT (technical issues)

---

## 📄 License

Proprietary - Internal use at Karwanchi

---

## 🎯 Future Enhancements

- [ ] Power BI dashboard integration
- [ ] Email auto-send (currently creates drafts)
- [ ] Mobile app for RSM field access
- [ ] Predictive expiry forecasting (ML)
- [ ] Real-time SAP API integration

---

**⭐ Star this repo if it helps your inventory management!**

---

> **Note:** This README describes the production system. The template file contains headers only—actual deployment requires SAP data connection.

**Status:** ✅ Active (Daily use since Jan 2025) | 🔧 Zero manual intervention required
