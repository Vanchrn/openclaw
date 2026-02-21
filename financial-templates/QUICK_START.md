# Financial Excel Templates - Quick Start Guide

## 📁 What You Have

Your financial templates collection includes:

### Excel Templates (Ready to Use)
1. **General_Ledger.xlsx** - Track all financial transactions with running balance
2. **Balance_Sheet.xlsx** - Assets, liabilities, and equity snapshot
3. **Income_Statement.xlsx** - Revenue, expenses, and profit/loss
4. **Annual_Budget.xlsx** - 12-month budget planning
5. **Financial_Ratios.xlsx** - Calculate key financial metrics

### Documentation
6. **TEMPLATES_GUIDE.md** - Complete formulas and VBA code reference
7. **VBA_MACROS.txt** - Copy-paste macros for automation
8. **QUICK_START.md** - This file

---

## 🚀 Getting Started

### Step 1: Open a Template
Simply double-click any `.xlsx` file to open it in Excel.

### Step 2: Enter Your Data
- Replace sample data with your actual numbers
- Formulas will automatically recalculate
- All templates use professional formatting

### Step 3: Add Automation (Optional)
To add VBA macros:

1. Open your Excel file
2. Press `Alt + F11` to open VBA Editor
3. Insert → Module
4. Copy code from `VBA_MACROS.txt`
5. Close VBA Editor
6. Press `Alt + F8` to run macros

---

## 📊 Template Features

### General Ledger
- ✅ Transaction tracking with auto-generated IDs
- ✅ Running balance calculation
- ✅ Debit/Credit validation
- ✅ Account code dropdowns
- ✅ Export to CSV functionality

### Balance Sheet
- ✅ Current & Non-Current assets
- ✅ Current & Non-Current liabilities
- ✅ Equity section
- ✅ Auto-calculation of totals
- ✅ Balance verification
- ✅ Key financial ratios

### Income Statement
- ✅ Revenue breakdown
- ✅ Cost of Goods Sold
- ✅ Operating expenses
- ✅ Net income calculation
- ✅ Profit margin percentages
- ✅ Variance analysis

### Annual Budget
- ✅ 12-month budget grid
- ✅ Revenue & expense categories
- ✅ Automatic totals
- ✅ Average calculations
- ✅ Budget vs actual comparison

### Financial Ratios
- ✅ Liquidity ratios (Current, Quick, Cash)
- ✅ Profitability ratios (Margins, ROA, ROE)
- ✅ Efficiency ratios (Turnover ratios)
- ✅ Solvency ratios (Debt ratios)
- ✅ All formulas pre-built

---

## 💡 Common Tasks

### Add a New Transaction to General Ledger
1. Open `General_Ledger.xlsx`
2. Enter date, account code, description
3. Enter debit OR credit amount
4. Balance updates automatically

### Create Monthly Financial Reports
1. Update `Balance_Sheet.xlsx` with month-end balances
2. Update `Income_Statement.xlsx` with monthly revenue/expenses
3. Open `Financial_Ratios.xlsx` to analyze performance

### Compare Budget vs Actual
1. Open `Annual_Budget.xlsx`
2. Enter actual amounts in a new column
3. Use variance formulas to see differences
4. Apply conditional formatting for visual indicators

### Generate an Invoice
1. Use the Invoice Generator macros from `VBA_MACROS.txt`
2. Run `GenerateNewInvoice` macro
3. Fill in customer details
4. Add line items
5. Run `CalculateInvoiceTotals` macro

---

## 🔧 Customization Tips

### Change Currency Format
1. Select the cells
2. Right-click → Format Cells
3. Number → Currency
4. Choose your preferred format

### Add New Categories
1. Insert a new row
2. Type category name
3. The total formulas will automatically update

### Create Charts
1. Select your data
2. Insert → Charts
3. Choose chart type
4. Customize as needed

### Protect Formulas
1. Select all cells → Format Cells → Protection → Uncheck Locked
2. Select formula cells → Format Cells → Protection → Check Locked
3. Review → Protect Sheet
4. Set a password

---

## 📝 Formulas Reference

### Common Financial Formulas

**Running Balance:**
```
=Previous_Balance + Debit - Credit
```

**Total Assets:**
```
=Current_Assets + Non_Current_Assets
```

**Working Capital:**
```
=Current_Assets - Current_Liabilities
```

**Current Ratio:**
```
=Current_Assets / Current_Liabilities
```

**Gross Profit Margin:**
```
=Gross_Profit / Net_Sales
```

**Net Profit Margin:**
```
=Net_Income / Net_Sales
```

**Return on Assets (ROA):**
```
=Net_Income / Total_Assets
```

**Return on Equity (ROE):**
```
=Net_Income / Total_Equity
```

---

## ⚠️ Important Notes

### Security
- ⚠️ Always review VBA code before running macros
- ⚠️ Only enable macros from trusted sources
- ⚠️ Keep backups of your data

### Best Practices
- ✅ Save your work frequently
- ✅ Use version control for important files
- ✅ Test formulas with sample data first
- ✅ Document any custom changes you make
- ✅ Regular backup your financial data

### Troubleshooting

**Formula shows #REF! error**
- Check that referenced cells exist
- Verify named ranges are defined

**Macro won't run**
- Check that macros are enabled in Excel
- Verify VBA code is correctly copied
- Check for syntax errors in VBA Editor

**Balance sheet doesn't balance**
- Verify all amounts are entered correctly
- Check for missing transactions
- Review debit/credit entries

---

## 📚 Additional Resources

### For More Help
- **TEMPLATES_GUIDE.md** - Detailed formulas and VBA code
- **VBA_MACROS.txt** - Complete macro library
- Excel Help → Formulas and Functions
- Microsoft Excel documentation

### Learning Resources
- Excel Financial Functions (PMT, IPMT, PPMT, etc.)
- VBA Programming for Excel
- Financial Accounting Principles

---

## 🎯 Next Steps

1. **Explore** - Open each template and review the structure
2. **Customize** - Add your company name and logo
3. **Test** - Enter sample data and verify calculations
4. **Automate** - Add VBA macros for repetitive tasks
5. **Integrate** - Link templates together for comprehensive reporting

---

## 📞 Need Help?

If you need assistance with:
- Customizing templates
- Creating new formulas
- Writing VBA macros
- Troubleshooting issues

Check the detailed documentation in `TEMPLATES_GUIDE.md` and `VBA_MACROS.txt`.

---

**Happy Accounting! 📊💼**

*Last Updated: 2026-02-21*
