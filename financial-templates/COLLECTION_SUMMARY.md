# Financial & Accounting Excel Templates Collection

## 📦 Collection Overview

A comprehensive collection of professional financial Excel templates with pre-built formulas, VBA macros, and automation capabilities.

---

## 📊 Excel Templates (5 files)

| Template | Size | Description |
|----------|------|-------------|
| **General_Ledger.xlsx** | 5.6K | Transaction tracking with running balance, debit/credit validation |
| **Balance_Sheet.xlsx** | 5.8K | Assets, liabilities, equity with automatic calculations |
| **Income_Statement.xlsx** | 5.8K | Revenue, expenses, profit/loss with margin analysis |
| **Annual_Budget.xlsx** | 6.5K | 12-month budget planning with variance analysis |
| **Financial_Ratios.xlsx** | 5.8K | 15+ key financial ratios with formulas |

**Features:**
- ✅ Pre-built formulas for automatic calculations
- ✅ Professional formatting and styling
- ✅ Sample data for reference
- ✅ Currency and percentage formatting
- ✅ Ready to use immediately

---

## 📚 Documentation (3 files)

| Document | Size | Description |
|----------|------|-------------|
| **QUICK_START.md** | 5.9K | Get started guide with common tasks |
| **TEMPLATES_GUIDE.md** | 25K | Complete formulas and VBA code reference |
| **VBA_MACROS.txt** | 24K | 50+ ready-to-use VBA macros |

---

## 🤖 Automation Tools

### Python Script
- **Financial_Template_Helper.py** (22K) - Generate templates programmatically

### VBA Macros Included
- General Ledger automation (ID generation, validation, export)
- Balance Sheet calculations and verification
- Income Statement generation
- Budget variance analysis
- Financial ratio calculations
- Invoice generation
- Loan amortization schedules
- Utility functions (protect, backup, export)

---

## 🎯 Key Features

### Formulas
- Running balance calculations
- Financial ratios (liquidity, profitability, efficiency, solvency)
- Variance analysis (budget vs actual)
- Loan amortization (PMT, IPMT, PPMT)
- Depreciation calculations (SLN, DB, DDB, SYD)
- Conditional formatting for visual indicators

### VBA Macros
- Auto-generate transaction IDs
- Validate data entry
- Calculate running balances
- Export to CSV/PDF
- Create comparative reports
- Generate charts
- Protect formulas
- Create backups

### Professional Formatting
- Currency formatting
- Percentage formatting
- Conditional formatting (profit/loss indicators)
- Color-coded sections
- Professional headers and footers

---

## 📋 Template Details

### 1. General Ledger
**Columns:** Transaction ID, Date, Account Code, Description, Debit, Credit, Balance, Category, Reference, Status

**Formulas:**
- Running balance: `=Previous_Balance + Debit - Credit`
- Total debits/credits: `=SUM(Debit_Column)`, `=SUM(Credit_Column)`
- Balance check: `=IF(Total_Debits=Total_Credits, "Balanced", "Out of Balance")`

**Macros:**
- Generate transaction IDs
- Validate debit/credit entry
- Calculate running balance
- Export to CSV
- Filter by date range

---

### 2. Balance Sheet
**Sections:**
- Assets (Current & Non-Current)
- Liabilities (Current & Non-Current)
- Equity (Capital, Retained Earnings, Current Year)

**Formulas:**
- Total Assets: `=Current_Assets + Non_Current_Assets`
- Working Capital: `=Current_Assets - Current_Liabilities`
- Current Ratio: `=Current_Assets / Current_Liabilities`
- Debt-to-Equity: `=Total_Liabilities / Total_Equity`

**Macros:**
- Calculate balance sheet
- Verify balance
- Create comparative balance sheet

---

### 3. Income Statement
**Sections:**
- Revenue (Sales, Other)
- Cost of Goods Sold
- Gross Profit
- Operating Expenses
- Operating Income
- Other Income/Expense
- Net Income

**Formulas:**
- Net Sales: `=Sales - Returns - Discounts`
- Gross Profit: `=Revenue - COGS`
- Operating Income: `=Gross_Profit - Operating_Expenses`
- Net Income: `=Income_Before_Tax - Income_Tax`
- Gross Margin %: `=Gross_Profit / Net_Sales`
- Net Profit Margin %: `=Net_Income / Net_Sales`

**Macros:**
- Generate income statement
- Compare actual vs budget
- Calculate profit margins

---

### 4. Annual Budget
**Structure:**
- 12-month grid (Jan-Dec)
- Revenue categories
- Expense categories
- Net income calculation
- Total and average columns

**Formulas:**
- Monthly total: `=SUM(Jan:Dec)`
- YTD total: `=SUM($Jan:CurrentMonth)`
- Variance: `=Actual - Budget`
- Variance %: `=(Actual - Budget) / Budget`

**Macros:**
- Create monthly budget
- Copy budget to new month
- Generate variance report
- Apply conditional formatting

---

### 5. Financial Ratios
**Categories:**

**Liquidity Ratios:**
- Current Ratio
- Quick Ratio
- Cash Ratio

**Profitability Ratios:**
- Gross Profit Margin
- Operating Margin
- Net Profit Margin
- Return on Assets (ROA)
- Return on Equity (ROE)

**Efficiency Ratios:**
- Asset Turnover
- Inventory Turnover
- Receivables Turnover

**Solvency Ratios:**
- Debt-to-Equity
- Debt Ratio
- Times Interest Earned

**Macros:**
- Calculate all ratios
- Create ratio analysis charts
- Generate ratio reports

---

## 🚀 Quick Start

1. **Open a Template** - Double-click any `.xlsx` file
2. **Enter Your Data** - Replace sample data with your numbers
3. **Formulas Auto-Calculate** - All formulas update automatically
4. **Add Macros (Optional)** - Copy from `VBA_MACROS.txt` for automation

---

## 📖 Documentation Guide

### For Beginners
→ Start with **QUICK_START.md**
- Get started guide
- Common tasks
- Basic customization

### For Detailed Reference
→ Read **TEMPLATES_GUIDE.md**
- Complete formula reference
- VBA code examples
- Template structures

### For Automation
→ Use **VBA_MACROS.txt**
- 50+ ready-to-use macros
- Copy-paste code
- Step-by-step instructions

---

## 🔧 Customization

### Easy Customizations
- Change currency symbols
- Add company logo
- Modify color schemes
- Add new categories
- Create custom charts

### Advanced Customizations
- Write custom VBA macros
- Create linked templates
- Build dashboards
- Integrate with other systems

---

## 💡 Use Cases

### Small Business
- Track daily transactions
- Monthly financial statements
- Budget planning
- Cash flow management

### Personal Finance
- Expense tracking
- Budget management
- Investment tracking
- Loan management

### Accounting Professionals
- Client financial statements
- Ratio analysis
- Variance reporting
- Tax preparation support

---

## ⚠️ Security Notes

- ⚠️ Always review VBA code before running
- ⚠️ Only enable macros from trusted sources
- ⚠️ Keep regular backups of your data
- ⚠️ Test with sample data first

---

## 📊 Technical Details

### File Formats
- Excel Templates: `.xlsx` (standard)
- Macro-Enabled: `.xlsm` (for VBA)
- Python Script: `.py` (for generation)

### Dependencies
- Excel 2007 or later
- Python 3.x (for script)
- openpyxl library (for Python script)

### Compatibility
- Windows: Excel 2007+
- Mac: Excel 2011+
- Online: Excel Online (limited macro support)

---

## 🎓 Learning Resources

### Excel Functions Used
- SUM, SUMIF, SUMIFS
- AVERAGE, AVERAGEIF
- IF, IFERROR, nested IFs
- PMT, IPMT, PPMT
- SLN, DB, DDB, SYD
- VLOOKUP, INDEX, MATCH

### VBA Concepts
- Sub procedures and functions
- Worksheet events (Change, SelectionChange)
- Range objects and cells
- Loops and conditionals
- User input (InputBox, MsgBox)

---

## 📈 Future Enhancements

Potential additions:
- Dashboard templates
- Cash flow forecasting
- Multi-currency support
- Pivot table templates
- Chart templates
- Industry-specific templates

---

## 📞 Support

For questions or issues:
1. Check **QUICK_START.md** for common tasks
2. Review **TEMPLATES_GUIDE.md** for detailed formulas
3. Use **VBA_MACROS.txt** for automation code
4. Test with sample data before using with real data

---

## 📝 Version History

**v1.0** (2026-02-21)
- Initial release
- 5 core templates
- Complete documentation
- 50+ VBA macros
- Python generation script

---

**Created by:** OpenClaw
**Date:** 2026-02-21
**Location:** `/home/gem/.openclaw/workspace/financial-templates/`

---

*Happy Accounting! 📊💼*
