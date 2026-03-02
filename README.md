# 📊 EDVIRON Revenue Analytics Dashboard  
## Data Analyst Assignment - Revenue, Commission & Settlement Analytics  

![Excel](https://img.shields.io/badge/Microsoft_Excel-2019+-blue?logo=microsoft-excel)
![VBA](https://img.shields.io/badge/VBA-Macros_Enabled-green?logo=vba)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Deadline](https://img.shields.io/badge/Deadline-2_Mar_2026_10AM-red)

---

## 📌 Table of Contents
- [Project Overview](#project-overview)
- [Business Context](#business-context)
- [Key Features](#key-features)
- [File Structure](#file-structure)
- [Setup & Execution Guide](#setup--execution-guide)
- [Dashboard Walkthrough](#dashboard-walkthrough)
- [Revenue Logic](#revenue-logic)
- [Assumptions](#assumptions)
- [Testing Checklist](#testing-checklist)
- [Submission Details](#submission-details)
- [Troubleshooting](#troubleshooting)

---

## 🌐 Project Overview
This solution delivers a **fully automated revenue analytics system** for Edviron's payment processing platform. Built entirely in Excel with VBA automation, it transforms raw transaction data into actionable business intelligence through:
- Comprehensive data cleaning & normalization
- Accurate revenue calculations across pricing layers
- Interactive partner/gateway/payment analytics
- Real-time settlement exposure monitoring
- Professional executive dashboard

✅ **100% Requirement Coverage**  
✅ **Audit-Friendly Data Flow**  
✅ **Production-Ready VBA Implementation**  

---

## 💼 Business Context
Edviron operates as a payment processing intermediary between:
```
Schools (Merchants) → ERP Partners → Edviron Platform → Payment Gateways
```

### Revenue Layers:
| Layer | Description | Example |
|-------|-------------|---------|
| **Merchant Pricing** | School pays to ERP | ₹2 flat or 2.5% |
| **Partner Pricing** | ERP pays to Edviron | ₹2 flat |
| **Edviron Buying** | Edviron pays to Gateway | ₹0 flat |

### Settlement Flow:
1. Gateway settles **full transaction amount** to Edviron
2. Edviron pays **ERP Revenue** to partner
3. Edviron retains **Edviron Net Revenue**

---

## 🚀 Key Features Implemented

### ✅ Data Processing
- Automatic pricing format standardization (Flat/% detection)
- Data quality flags for missing values & inconsistencies
- INR conversion for all pricing layers
- Negative margin exception tracking

### ✅ Revenue Calculations
- ERP Revenue = Merchant Pricing - Partner Pricing
- Edviron Net Revenue = Partner Pricing - Edviron Buying
- Edviron Gross Revenue = ERP Revenue + Edviron Net Revenue
- Settlement metrics with pending exposure tracking

### ✅ Analytics & Reporting
- 5 Pivot-based analytical reports:
  - Daily/Weekly/Monthly Revenue Summary
  - Partner (ERP) Performance Dashboard
  - Gateway & Payment Method Analysis
  - Pending vs Settled Exposure Report
  - Data Quality Exception Report
- Dynamic KPI calculations with filters

### ✅ Interactive Dashboard
- **8 Real-time KPI Tiles**: Transactions, GMV, Revenues, Exposure, Users
- **4 Interactive Charts**: Revenue Split, Time Trend, Gateway Contribution, Payment Mix
- **4 Dynamic Filters**: Date Range, Partner, Gateway, Payment Method
- **Refresh Button**: One-click recalculation of all metrics
- **Slicer Integration**: Connected to all pivot tables

### ✅ VBA Automation
- Single-click execution (`RunCompleteAnalysis`)
- Modular, commented code structure
- Error handling with user-friendly messages
- No hard-coded references
- Documentation generator

---

## 📁 File Structure
```
EDVIRON_Revenue_Analytics/
│
├── EDVIRON_Analytics_YourName.xlsm       # Main Excel file (MACROS ENABLED)
│
├── Documentation/
│   ├── README.md                         # This file
│   ├── Revenue_Logic.pdf                 # Detailed calculation guide
│   └── Dashboard_Walkthrough.pdf         # UI/UX guide with screenshots
│
├── Source_Data/
│   └── Raw_Transactions.xlsx             # Original dataset (reference)
│
└── Submission/
    ├── Submission_Checklist.pdf
    └── Screen_Recording_Link.txt
```

---

## ⚙️ Setup & Execution Guide

### Prerequisites
- Microsoft Excel 2019+ or Microsoft 365
- Macros enabled in Excel Security Settings
- Developer tab visible in Excel ribbon

### Step-by-Step Execution
1. **Download & Open**
   - Save `EDVIRON_Analytics_YourName.xlsm` to your computer
   - Open file → Click "Enable Macros" when prompted

2. **Run Complete Analysis**
   - Press `ALT + F8`
   - Select `RunCompleteAnalysis`
   - Click **RUN**
   - Wait 60-90 seconds for completion

3. **Verify Execution**
   - Confirm 6 sheets created:
     - `Raw_Data` (original data)
     - `Clean_Data` (standardized data)
     - `Calc_Model` (revenue calculations)
     - `Reports` (pivot tables)
     - `Dashboard` (interactive interface)
     - `Assumptions_Notes` (documentation)
   - Check popup confirmation message

4. **Use Dashboard**
   - Navigate to `Dashboard` sheet
   - Use slicers to filter data
   - Click "Refresh Dashboard" button after changes
   - Explore all KPIs and charts

> 💡 **Pro Tip**: Press `ALT + F11` to view/edit VBA code. All modules are thoroughly commented.

---

## 📈 Dashboard Walkthrough

### Top Control Panel
| Element | Function |
|---------|----------|
| **Date Range Slicer** | Filter transactions by date period |
| **Partner Slicer** | Analyze specific ERP partners |
| **Gateway Slicer** | Compare payment gateway performance |
| **Payment Method Slicer** | Segment by UPI/Card/NetBanking |
| **Refresh Button** | Recalculate all metrics after filtering |

### KPI Tiles (Real-time)
| Metric | Calculation |
|--------|-------------|
| Total Transactions | `COUNTA(Calc_Model!A:A)-1` |
| Total GMV | `SUM(Calc_Model!G:G)` |
| ERP Revenue | `SUM(Calc_Model!AD:AD)` |
| Edviron Net Revenue | `SUM(Calc_Model!AE:AE)` |
| Edviron Gross Revenue | `SUM(Calc_Model!AF:AF)` |
| Pending Exposure | `SUM(Calc_Model!AI:AI)` |
| Unique Users | `SUMPRODUCT(1/COUNTIFS(...))` |
| Payment Frequency | Transactions ÷ Unique Users |

### Interactive Charts
1. **Revenue Split Pie Chart**: ERP vs Edviron revenue distribution
2. **Time Trend Line Chart**: Daily revenue patterns (responsive to date filter)
3. **Gateway Contribution Bar Chart**: GMV by payment gateway
4. **Payment Method Mix Donut Chart**: Transaction volume distribution

---

## 💰 Revenue Logic

### Pricing Conversion to INR
```excel
=IF(Pricing_Type="FLAT", Rate, Transaction_Amount × Rate ÷ 100)
```

### Core Revenue Formulas
| Metric | Formula | Location |
|--------|---------|----------|
| **Merchant Pricing INR** | `IF(FLAT, Rate, Amount×Rate/100)` | Calc_Model!AA |
| **Partner Pricing INR** | `IF(FLAT, Rate, Amount×Rate/100)` | Calc_Model!AB |
| **Edviron Buying INR** | `IF(FLAT, Rate, Amount×Rate/100)` | Calc_Model!AC |
| **ERP Revenue** | `Merchant_INR - Partner_INR` | Calc_Model!AD |
| **Edviron Net Revenue** | `Partner_INR - Edviron_INR` | Calc_Model!AE |
| **Edviron Gross Revenue** | `ERP_Revenue + Edviron_Net` | Calc_Model!AF |
| **Pending Exposure** | `IF(Status=PENDING, Amount, 0)` | Calc_Model!AI |
| **ERP Payable Outstanding** | `IF(Pending, ERP_Revenue, 0)` | Calc_Model!AJ |

### Settlement Logic
- **Amount Payable to ERP**: Uses `ERP_Commission` field if available, else computed `ERP_Revenue`
- **Amount Retained by Edviron**: `Edviron_Net_Revenue`
- **Gateway Fees**: Estimated at 1% of transaction amount (documented assumption)

---

## 📝 Assumptions Documented

### Critical Assumptions
1. **Pricing Conversion**: Percentage values applied to Transaction Amount (not Order Amount)
2. **Missing Values**: 
   - Empty pricing fields treated as ₹0
   - Missing partners/gateways labeled "Unknown" with flags
3. **Date Handling**: Date extracted from "Date & Time" column; no timezone conversion
4. **Unique Users**: Calculated via COUNTIFS approximation (exact distinct count requires Power Pivot)
5. **ERP Commission Priority**: When available, ERP Commission field overrides computed ERP Revenue
6. **Pending Exposure**: Defined as transactions with Status="PENDING" (Capture Status not used for exposure calc)
7. **Gateway Fees**: Estimated at 1% where not explicitly provided

### Data Limitations Addressed
- Handled mixed pricing formats (2, "2%", "2.5 %", 0.025)
- Created data quality flags for audit trail
- Implemented negative margin exception tracking
- Added validation for missing critical fields

> 📌 Full documentation available in `Assumptions_Notes` sheet within workbook

---

## ✅ Testing Checklist

### Pre-Submission Validation
- [ ] All 6 required sheets exist and contain data
- [ ] VBA runs without errors (`ALT+F8` → `RunCompleteAnalysis`)
- [ ] Revenue calculations verified with manual spot-checks:
  - Transaction #2: ₹15,002.36 × 2.5% = ₹375.06 Merchant Pricing
  - ERP Revenue = ₹375.06 - ₹2.00 = ₹373.06
- [ ] Slicers dynamically update all KPIs and charts
- [ ] "Refresh Dashboard" button recalculates all metrics
- [ ] Exception report shows negative margins (if any)
- [ ] File saved as `.xlsm` (NOT .xlsx)
- [ ] VBA project NOT password protected
- [ ] All formulas use relative references (no hard-coded values)

### Critical Test Scenarios
| Test | Steps | Expected Result |
|------|-------|-----------------|
| **Filter by Partner** | Select "SBNM College" in Partner slicer | All KPIs/charts show only SBNM data |
| **Date Range Filter** | Select Jan 2025 in Date slicer | Time trend chart shows Jan data only |
| **Pending Exposure** | Filter Status="PENDING" | Pending Exposure KPI > 0 |
| **Negative Margin** | Check Exception Report | Flags transactions where Merchant < Partner pricing |
| **Refresh Function** | Change filter → Click Refresh | All metrics update instantly |

---

## 📤 Submission Details

### Required Deliverables
1. **Excel File**: `EDVIRON_Analytics_YourName.xlsm` (MACROS ENABLED)
2. **Documentation**: 
   - `Assumptions_Notes` sheet within workbook
   - This README file
3. **Submission Form**: 
   - Complete Google Form: https://docs.google.com/forms/d/e/1FAIpQLSebP65ps136E1rhZiws_WPyLI3qNOnigCNljQh-H9lzFeB83w/viewform
   - Upload `.xlsm` file before deadline

### Critical Submission Requirements
⚠️ **MUST BE .XLSM FORMAT** (macros enabled)  
⚠️ **NO PASSWORD PROTECTION** on VBA project  
⚠️ **DEADLINE**: March 2, 2026 at 10:00 AM IST  
⚠️ **FILE SIZE**: Under 25MB (remove unnecessary sample data if needed)  
⚠️ **ACCESS**: Ensure Google Drive links have "Anyone with link can view" permissions

---

## 🛠️ Troubleshooting

### Common Issues & Solutions
| Issue | Solution |
|-------|----------|
| **"Macros Disabled" warning** | File → Options → Trust Center → Macro Settings → Enable all macros |
| **Runtime Error 9** | Sheet missing → Run `SetupWorkbookStructure` subroutine first |
| **Slicers not connected** | Right-click slicer → Report Connections → Check all pivot tables |
| **KPIs show #VALUE!** | Check Calc_Model sheet has data → Run `CalculateRevenues` |
| **Charts not updating** | Click "Refresh Dashboard" button → Verify slicer connections |
| **VBA project locked** | Tools → VBAProject Properties → Protection tab → Uncheck "Lock project" |

### Emergency Recovery Steps
1. Save current workbook as backup
2. Close Excel completely
3. Reopen workbook → Enable macros
4. Press `ALT+F11` → Run `RunCompleteAnalysis` again
5. If persistent issues: Delete all sheets except Raw_Data → Re-run analysis

---

## 🌟 Bonus Features Implemented
- ✅ **Automated Exception Reporting**: Tracks negative margins & data issues
- ✅ **Data Quality Dashboard**: Visual flags for missing values
- ✅ **Professional Formatting**: Consistent INR formatting (₹1,25,000)
- ✅ **Audit Trail**: Clear data flow from Raw_Data → Clean_Data → Calc_Model
- ✅ **User Documentation**: Comprehensive Assumptions_Notes sheet
- ✅ **Error Handling**: User-friendly messages for all VBA operations

---
