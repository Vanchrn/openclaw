# Financial & Accounting Excel Templates - Complete Guide

This guide provides formulas, VBA macros, and instructions for creating professional financial templates in Excel.

---

## 1. GENERAL LEDGER TEMPLATE

### Column Structure
| Column | Header | Description |
|--------|--------|-------------|
| A | Transaction ID | Unique identifier |
| B | Date | Transaction date |
| C | Account Code | Chart of accounts reference |
| D | Description | Transaction details |
| E | Debit | Debit amount |
| F | Credit | Credit amount |
| G | Balance | Running balance |
| H | Category | Asset/Liability/Equity/Revenue/Expense |
| I | Reference | Invoice/check number |
| J | Status | Posted/Pending |

### Key Formulas

**Running Balance (Column G):**
```excel
=IF(ROW()=2, E2-F2, G1+E2-F2)
```

**Total Debits:**
```excel
=SUM(E:E)
```

**Total Credits:**
```excel
=SUM(F:F)
```

**Balance Check:**
```excel
=IF(SUM(E:E)=SUM(F:F), "Balanced", "Out of Balance")
```

### VBA Macros

**Auto-Generate Transaction ID:**
```vba
Sub GenerateTransactionID()
    Dim lastRow As Long
    Dim newID As String
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    newID = "T" & Format(lastRow, "0000")
    
    Cells(lastRow + 1, 1).Value = newID
    Cells(lastRow + 1, 2).Value = Date
End Sub
```

**Validate Debit/Credit Entry:**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 5 Or Target.Column = 6 Then
        If Target.Value < 0 Then
            MsgBox "Negative values not allowed. Please enter positive amounts.", vbExclamation
            Target.Value = ""
        End If
    End If
End Sub
```

**Export to CSV:**
```vba
Sub ExportLedgerToCSV()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Copy
    ActiveWorkbook.SaveAs Filename:="General_Ledger_" & Format(Date, "yyyymmdd") & ".csv", _
        FileFormat:=xlCSV
    ActiveWorkbook.Close
End Sub
```

---

## 2. BALANCE SHEET TEMPLATE

### Structure

**Assets Section:**
```excel
Current Assets:
- Cash                    =SUM(Cash_Range)
- Accounts Receivable     =SUM(AR_Range)
- Inventory               =SUM(Inventory_Range)
- Prepaid Expenses        =SUM(Prepaid_Range)
TOTAL CURRENT ASSETS      =SUM(Current_Assets_Range)

Non-Current Assets:
- Property & Equipment     =SUM(PPE_Range)
- Accumulated Depreciation =-SUM(Accum_Dep_Range)
- Investments             =SUM(Investments_Range)
TOTAL NON-CURRENT ASSETS  =SUM(Non_Current_Assets_Range)

TOTAL ASSETS              =Current_Assets + Non_Current_Assets
```

**Liabilities Section:**
```excel
Current Liabilities:
- Accounts Payable        =SUM(AP_Range)
- Short-term Debt         =SUM(ST_Debt_Range)
- Accrued Expenses        =SUM(Accrued_Range)
TOTAL CURRENT LIABILITIES =SUM(Current_Liab_Range)

Non-Current Liabilities:
- Long-term Debt          =SUM(LT_Debt_Range)
TOTAL LIABILITIES         =Current_Liab + Non_Current_Liab
```

**Equity Section:**
```excel
- Owner's Capital         =Capital_Cell
- Retained Earnings       =Retained_Earnings_Cell
- Current Year Earnings   =Net_Income_Cell
TOTAL EQUITY              =SUM(Equity_Range)

TOTAL LIABILITIES & EQUITY =Total_Liabilities + Total_Equity
```

### Key Formulas

**Working Capital:**
```excel
=Current_Assets - Current_Liabilities
```

**Current Ratio:**
```excel
=Current_Assets / Current_Liabilities
```

**Debt-to-Equity Ratio:**
```excel
=Total_Liabilities / Total_Equity
```

**Asset Turnover:**
```excel
=Revenue / Total_Assets
```

### VBA Macros

**Auto-Calculate Balance Sheet:**
```vba
Sub CalculateBalanceSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Calculate totals
    ws.Range("Current_Assets_Total").Value = _
        Application.WorksheetFunction.Sum(ws.Range("Current_Asits_Range"))
    
    ws.Range("Total_Assets").Value = _
        ws.Range("Current_Assets_Total").Value + ws.Range("Non_Current_Assets_Total").Value
    
    ' Validate balance
    If Abs(ws.Range("Total_Assets").Value - ws.Range("Total_Liab_Equity").Value) > 0.01 Then
        MsgBox "Balance Sheet does not balance! Difference: " & _
            Format(ws.Range("Total_Assets").Value - ws.Range("Total_Liab_Equity").Value, "#,##0.00"), _
            vbExclamation, "Balance Check"
    End If
End Sub
```

---

## 3. INCOME STATEMENT TEMPLATE

### Structure

**Revenue Section:**
```excel
Sales Revenue                =SUM(Sales_Range)
Less: Sales Returns          =SUM(Returns_Range)
Less: Discounts              =SUM(Discounts_Range)
NET SALES                    =Sales - Returns - Discounts

Other Revenue                =SUM(Other_Revenue_Range)
TOTAL REVENUE                =Net_Sales + Other_Revenue
```

**Cost of Goods Sold:**
```excel
Beginning Inventory          =Beg_Inv_Cell
+ Purchases                  =SUM(Purchases_Range)
- Ending Inventory           =End_Inv_Cell
COST OF GOODS SOLD           =Beg_Inv + Purchases - End_Inv

GROSS PROFIT                 =Total_Revenue - COGS
```

**Operating Expenses:**
```excel
Salaries & Wages             =SUM(Salaries_Range)
Rent Expense                 =Rent_Cell
Utilities                    =SUM(Utilities_Range)
Depreciation                 =SUM(Depreciation_Range)
Marketing                    =SUM(Marketing_Range)
Other Operating Expenses      =SUM(Other_Op_Exp_Range)
TOTAL OPERATING EXPENSES     =SUM(Operating_Expenses_Range)

OPERATING INCOME             =Gross_Profit - Operating_Expenses
```

**Other Income/Expenses:**
```excel
Interest Income              =Interest_Income_Cell
Interest Expense             =Interest_Expense_Cell
NET OTHER INCOME/EXPENSE     =Interest_Income - Interest_Expense

INCOME BEFORE TAXES          =Operating_Income + Net_Other
Income Tax Expense           =Income_Tax_Rate * Income_Before_Tax

NET INCOME                   =Income_Before_Tax - Income_Tax
```

### Key Formulas

**Gross Margin %:**
```excel
=Gross_Profit / Net_Sales
```

**Operating Margin %:**
```excel
=Operating_Income / Net_Sales
```

**Net Profit Margin %:**
```excel
=Net_Income / Net_Sales
```

**Return on Sales:**
```excel
=Net_Income / Total_Revenue
```

### VBA Macros

**Generate Income Statement:**
```vba
Sub GenerateIncomeStatement()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Calculate revenue section
    ws.Range("Net_Sales").Value = ws.Range("Sales_Revenue").Value - _
        ws.Range("Sales_Returns").Value - ws.Range("Discounts").Value
    
    ws.Range("Total_Revenue").Value = ws.Range("Net_Sales").Value + _
        ws.Range("Other_Revenue").Value
    
    ' Calculate COGS
    ws.Range("COGS").Value = ws.Range("Beg_Inventory").Value + _
        ws.Range("Purchases").Value - ws.Range("End_Inventory").Value
    
    ' Calculate gross profit
    ws.Range("Gross_Profit").Value = ws.Range("Total_Revenue").Value - _
        ws.Range("COGS").Value
    
    ' Calculate operating income
    ws.Range("Operating_Expenses").Value = _
        Application.WorksheetFunction.Sum(ws.Range("Op_Expenses_Range"))
    
    ws.Range("Operating_Income").Value = ws.Range("Gross_Profit").Value - _
        ws.Range("Operating_Expenses").Value
    
    ' Calculate net income
    ws.Range("Income_Before_Tax").Value = ws.Range("Operating_Income").Value + _
        ws.Range("Net_Other_Income").Value
    
    ws.Range("Income_Tax").Value = ws.Range("Income_Before_Tax").Value * _
        ws.Range("Tax_Rate").Value
    
    ws.Range("Net_Income").Value = ws.Range("Income_Before_Tax").Value - _
        ws.Range("Income_Tax").Value
    
    MsgBox "Income Statement calculated successfully!", vbInformation
End Sub
```

---

## 4. CASH FOW STATEMENT TEMPLATE

### Structure

**Operating Activities:**
```excel
Net Income                    =Net_Income_From_IS
+ Depreciation                =Depreciation_Cell
+ Changes in Working Capital:
  - Increase in A/R           =-Change_AR
  - Decrease in A/R           =+Change_AR
  - Increase in Inventory     =-Change_Inv
  - Decrease in Inventory     =+Change_Inv
  - Increase in A/P           =+Change_AP
  - Decrease in A/P           =-Change_AP
NET CASH FROM OPERATING       =SUM(Operating_Activities)
```

**Investing Activities:**
```excel
Purchase of PPE              =-PPE_Purchases
Sale of Equipment            =+Equipment_Sales
Purchase of Investments      =-Investment_Purchases
Sale of Investments          =+Investment_Sales
NET CASH FROM INVESTING      =SUM(Investing_Activities)
```

**Financing Activities:**
```excel
Proceeds from Loans          =+Loan_Proceeds
Repayment of Loans           =-Loan_Repayments
Issuance of Stock            =+Stock_Issuance
Dividends Paid               =-Dividends
NET CASH FROM FINANCING      =SUM(Financing_Activities)

NET CHANGE IN CASH           =Operating + Investing + Financing
+ Beginning Cash Balance     =Beg_Cash_Balance
= ENDING CASH BALANCE        =Net_Change + Beg_Balance
```

### VBA Macros

**Calculate Cash Flow:**
```vba
Sub CalculateCashFlow()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Operating activities
    ws.Range("Net_Operating_Cash").Formula = _
        "=SUM(Operating_Cash_Flow_Range)"
    
    ' Investing activities
    ws.Range("Net_Investing_Cash").Formula = _
        "=SUM(Investing_Cash_Flow_Range)"
    
    ' Financing activities
    ws.Range("Net_Financing_Cash").Formula = _
        "=SUM(Financing_Cash_Flow_Range)"
    
    ' Net change
    ws.Range("Net_Change_Cash").Formula = _
        "=Net_Operating_Cash + Net_Investing_Cash + Net_Financing_Cash"
    
    ' Ending cash
    ws.Range("Ending_Cash").Formula = _
        "=Net_Change_Cash + Beginning_Cash"
    
    MsgBox "Cash Flow Statement calculated!", vbInformation
End Sub
```

---

## 5. BUDGET TEMPLATE

### Structure

**Monthly Budget Table:**
| Category | Jan | Feb | Mar | ... | Dec | Total |
|----------|-----|-----|-----|-----|-----|-------|
| Revenue  |     |     |     |     |     |       |
| - Sales  |     |     |     |     |     |       |
| - Other  |     |     |     |     |     |       |
| Expenses |     |     |     |     |     |       |
| - COGS   |     |     |     |     |     |       |
| - Salaries|    |     |     |     |     |       |
| - Rent   |     |     |     |     |     |       |
| Net      |     |     |     |     |     |       |

### Key Formulas

**Monthly Total:**
```excel
=SUM(B2:M2)
```

**YTD Total (for any month):**
```excel
=SUM($B2:B2)  ' For January
=SUM($B2:C2)  ' For February
```

**Variance (Actual vs Budget):**
```excel
=Actual - Budget
```

**Variance %:**
```excel
=(Actual - Budget) / Budget
```

### VBA Macros

**Create Monthly Budget:**
```vba
Sub CreateMonthlyBudget()
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    ' Clear previous data
    ws.Range("B2:M100").ClearContents
    
    ' Set up formulas
    For i = 2 To 50
        ' Total for each row
        ws.Cells(i, 14).Formula = "=SUM(B" & i & ":M" & i & ")"
        
        ' YTD formulas
        ws.Cells(i, 15).Formula = "=SUM($B" & i & ":B" & i & ")"
    Next i
    
    ' Column totals
    ws.Range("B51").Formula = "=SUM(B2:B50)"
    ws.Range("B51").AutoFill Destination:=ws.Range("B51:N51"), Type:=xlFillDefault
    
    MsgBox "Budget template created!", vbInformation
End Sub
```

**Budget vs Actual Analysis:**
```vba
Sub BudgetVarianceAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Calculate variance
    ws.Range("P2:P" & lastRow).Formula = "=N2 - O2"  ' Actual - Budget
    ws.Range("Q2:Q" & lastRow).Formula = "=IF(O2<>0, (N2-O2)/O2, 0)"  ' Variance %
    
    ' Conditional formatting
    With ws.Range("P2:P" & lastRow).FormatConditions
        .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .Item(1).Interior.Color = RGB(146, 208, 80)  ' Green for favorable
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(2).Interior.Color = RGB(255, 0, 0)  ' Red for unfavorable
    End With
    
    MsgBox "Variance analysis complete!", vbInformation
End Sub
```

---

## 6. FINANCIAL RATIOS TEMPLATE

### Liquidity Ratios

**Current Ratio:**
```excel
=Current_Assets / Current_Liabilities
```

**Quick Ratio:**
```excel
=(Current_Assets - Inventory) / Current_Liabilities
```

**Cash Ratio:**
```excel
=Cash / Current_Liabilities
```

### Profitability Ratios

**Gross Profit Margin:**
```excel
=Gross_Profit / Net_Sales
```

**Operating Profit Margin:**
```excel
=Operating_Income / Net_Sales
```

**Net Profit Margin:**
```excel
=Net_Income / Net_Sales
```

**Return on Assets (ROA):**
```excel
=Net_Income / Total_Assets
```

**Return on Equity (ROE):**
```excel
=Net_Income / Total_Equity
```

### Efficiency Ratios

**Asset Turnover:**
```excel
=Net_Sales / Total_Assets
```

**Inventory Turnover:**
```excel
=COGS / Average_Inventory
```

**Receivables Turnover:**
```excel
=Net_Sales / Average_Accounts_Receivable
```

### Solvency Ratios

**Debt-to-Equity:**
```excel
=Total_Liabilities / Total_Equity
```

**Debt Ratio:**
```excel
=Total_Liabilities / Total_Assets
```

**Times Interest Earned:**
```excel
=EBIT / Interest_Expense
```

### VBA Macros

**Calculate All Ratios:**
```vba
Sub CalculateFinancialRatios()
    Dim ws As Worksheet
    Dim ratioName As Variant
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    ' Array of ratio names and their formulas
    Dim ratios(1 To 15, 1 To 2) As Variant
    
    ratios(1, 1) = "Current Ratio"
    ratios(1, 2) = "=Current_Assets/Current_Liabilities"
    
    ratios(2, 1) = "Quick Ratio"
    ratios(2, 2) = "=(Current_Assets-Inventory)/Current_Liabilities"
    
    ratios(3, 1) = "Cash Ratio"
    ratios(3, 2) = "=Cash/Current_Liabilities"
    
    ratios(4, 1) = "Gross Profit Margin"
    ratios(4, 2) = "=Gross_Profit/Net_Sales"
    
    ratios(5, 1) = "Operating Margin"
    ratios(5, 2) = "=Operating_Income/Net_Sales"
    
    ratios(6, 1) = "Net Profit Margin"
    ratios(6, 2) = "=Net_Income/Net_Sales"
    
    ratios(7, 1) = "Return on Assets"
    ratios(7, 2) = "=Net_Income/Total_Assets"
    
    ratios(8, 1) = "Return on Equity"
    ratios(8, 2) = "=Net_Income/Total_Equity"
    
    ratios(9, 1) = "Asset Turnover"
    ratios(9, 2) = "=Net_Sales/Total_Assets"
    
    ratios(10, 1) = "Inventory Turnover"
    ratios(10, 2) = "=COGS/Average_Inventory"
    
    ratios(11, 1) = "Receivables Turnover"
    ratios(11, 2) = "=Net_Sales/Average_AR"
    
    ratios(12, 1) = "Debt-to-Equity"
    ratios(12, 2) = "=Total_Liabilities/Total_Equity"
    
    ratios(13, 1) = "Debt Ratio"
    ratios(13, 2) = "=Total_Liabilities/Total_Assets"
    
    ratios(14, 1) = "Times Interest Earned"
    ratios(14, 2) = "=EBIT/Interest_Expense"
    
    ratios(15, 1) = "Working Capital"
    ratios(15, 2) = "=Current_Assets-Current_Liabilities"
    
    ' Output ratios
    For i = 1 To 15
        ws.Cells(i + 1, 1).Value = ratios(i, 1)
        ws.Cells(i + 1, 2).Formula = ratios(i, 2)
    Next i
    
    ' Format as percentage where appropriate
    ws.Range("C3:C8").NumberFormat = "0.00%"
    ws.Range("C10:C11").NumberFormat = "0.00"
    
    MsgBox "Financial ratios calculated!", vbInformation
End Sub
```

---

## 7. INVOICE GENERATOR TEMPLATE

### Structure

**Header Section:**
- Company Name
- Invoice Number
- Date
- Due Date
- Customer Information

**Line Items:**
| Item | Description | Quantity | Unit Price | Amount |
|------|-------------|----------|------------|--------|
| 1 | Product/Service | Qty | Price | Qty*Price |

**Totals:**
- Subtotal
- Tax
- Discount
- Total Due

### Key Formulas

**Line Item Amount:**
```excel
=Quantity * Unit_Price
```

**Subtotal:**
```excel
=SUM(Amount_Column)
```

**Tax Amount:**
```excel
=Subtotal * Tax_Rate
```

**Discount Amount:**
```excel
=Subtotal * Discount_Rate
```

**Total Due:**
```excel
=Subtotal + Tax - Discount
```

### VBA Macros

**Generate New Invoice:**
```vba
Sub GenerateNewInvoice()
    Dim ws As Worksheet
    Dim invNum As String
    Dim lastInv As String
    
    Set ws = ActiveSheet
    
    ' Get last invoice number
    lastInv = ws.Range("Invoice_Number").Value
    
    ' Generate new invoice number
    If IsNumeric(lastInv) Then
        invNum = CStr(CLng(lastInv) + 1)
    Else
        invNum = "INV-" & Format(Date, "yyyymmdd") & "-001"
    End If
    
    ' Set new invoice details
    ws.Range("Invoice_Number").Value = invNum
    ws.Range("Invoice_Date").Value = Date
    ws.Range("Due_Date").Value = Date + 30
    
    ' Clear line items
    ws.Range("Line_Items").ClearContents
    
    ' Set formulas
    ws.Range("Subtotal").Formula = "=SUM(Line_Item_Amounts)"
    ws.Range("Tax").Formula = "=Subtotal * Tax_Rate"
    ws.Range("Total_Due").Formula = "=Subtotal + Tax - Discount"
    
    MsgBox "New invoice " & invNum & " created!", vbInformation
End Sub
```

**Email Invoice:**
```vba
Sub EmailInvoice()
    Dim ws As Worksheet
    Dim OutApp As Object
    Dim OutMail As Object
    Dim invNum As String
    Dim customerEmail As String
    
    Set ws = ActiveSheet
    invNum = ws.Range("Invoice_Number").Value
    customerEmail = ws.Range("Customer_Email").Value
    
    ' Create email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = customerEmail
        .Subject = "Invoice " & invNum
        .Body = "Dear Customer," & vbCrLf & vbCrLf & _
                "Please find attached invoice " & invNum & "." & vbCrLf & vbCrLf & _
                "Payment is due by " & ws.Range("Due_Date").Value & "." & vbCrLf & vbCrLf & _
                "Thank you for your business."
        
        ' Attach invoice (save as PDF first)
        ThisWorkbook.SaveAs Filename:="Invoice_" & invNum & ".pdf", FileFormat:=xlPDF
        .Attachments.Add ThisWorkbook.Path & "\Invoice_" & invNum & ".pdf"
        
        .Display  ' Change to .Send to send automatically
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    MsgBox "Invoice email created!", vbInformation
End Sub
```

---

## 8. LOAN AMORTIZATION SCHEDULE

### Structure

| Period | Payment | Principal | Interest | Balance |
|--------|---------|-----------|----------|---------|
| 0 | - | - | - | Loan Amount |
| 1 | PMT | PPMT | IPMT | Balance - Principal |
| 2 | PMT | PPMT | IPMT | Balance - Principal |
| ... | ... | ... | ... | ... |

### Key Formulas

**Monthly Payment (PMT):**
```excel
=PMT(Interest_Rate/12, Number_of_Periods, -Loan_Amount)
```

**Principal Payment (PPMT):**
```excel
=PPMT(Interest_Rate/12, Period, Number_of_Periods, -Loan_Amount)
```

**Interest Payment (IPMT):**
```excel
=IPMT(Interest_Rate/12, Period, Number_of_Periods, -Loan_Amount)
```

**Remaining Balance:**
```excel
=Previous_Balance - Principal_Payment
```

**Total Interest Paid:**
```excel
=SUM(Interest_Column)
```

### VBA Macros

**Generate Amortization Schedule:**
```vba
Sub GenerateAmortizationSchedule()
    Dim ws As Worksheet
    Dim loanAmount As Double
    Dim annualRate As Double
    Dim loanTerm As Integer
    Dim monthlyPayment As Double
    Dim balance As Double
    Dim interest As Double
    Dim principal As Double
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    ' Get loan parameters
    loanAmount = ws.Range("Loan_Amount").Value
    annualRate = ws.Range("Interest_Rate").Value
    loanTerm = ws.Range("Loan_Term_Years").Value * 12  ' Convert to months
    
    ' Calculate monthly payment
    monthlyPayment = Application.WorksheetFunction.Pmt(annualRate / 12, loanTerm, -loanAmount)
    
    ' Generate schedule
    balance = loanAmount
    
    ws.Range("A2").Value = "Period"
    ws.Range("B2").Value = "Payment"
    ws.Range("C2").Value = "Principal"
    ws.Range("D2").Value = "Interest"
    ws.Range("E2").Value = "Balance"
    
    For i = 1 To loanTerm
        interest = balance * (annualRate / 12)
        principal = monthlyPayment - interest
        balance = balance - principal
        
        ' Prevent negative balance in last period
        If balance < 0 Then
            principal = principal + balance
            balance = 0
        End If
        
        ws.Cells(i + 2, 1).Value = i
        ws.Cells(i + 2, 2).Value = monthlyPayment
        ws.Cells(i + 2, 3).Value = principal
        ws.Cells(i + 2, 4).Value = interest
        ws.Cells(i + 2, 5).Value = balance
    Next i
    
    ' Format as currency
    ws.Range("B3:E" & loanTerm + 2).NumberFormat = "$#,##0.00"
    
    ' Calculate totals
    ws.Cells(loanTerm + 3, 1).Value = "TOTALS"
    ws.Cells(loanTerm + 3, 2).Formula = "=SUM(B3:B" & loanTerm + 2 & ")"
    ws.Cells(loanTerm + 3, 3).Formula = "=SUM(C3:C" & loanTerm + 2 & ")"
    ws.Cells(loanTerm + 3, 4).Formula = "=SUM(D3:D" & loanTerm + 2 & ")"
    
    MsgBox "Amortization schedule generated for " & loanTerm & " months!", vbInformation
End Sub
```

---

## 9. DEPRECIATION CALCULATOR

### Depreciation Methods

**Straight-Line:**
```excel
=(Cost - Salvage_Value) / Useful_Life
```

**Declining Balance:**
```excel
=DB(Cost, Salvage_Value, Life, Period, Month)
```

**Double Declining Balance:**
```excel
=DDB(Cost, Salvage_Value, Life, Period, Factor)
```

**Sum of Years' Digits:**
```excel
=SYD(Cost, Salvage_Value, Life, Period)
```

**MACRS (for tax):**
```excel
=VDB(Cost, 0, Life, Period_Start, Period_End, Factor, No_Switch)
```

### VBA Macros

**Calculate All Depreciation Methods:**
```vba
Sub CalculateDepreciation()
    Dim ws As Worksheet
    Dim cost As Double
    Dim salvage As Double
    Dim life As Integer
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    ' Get asset parameters
    cost = ws.Range("Cost").Value
    salvage = ws.Range("Salvage_Value").Value
    life = ws.Range("Useful_Life").Value
    
    ' Headers
    ws.Range("A1").Value = "Year"
    ws.Range("B1").Value = "Straight-Line"
    ws.Range("C1").Value = "Declining Balance"
    ws.Range("D1").Value = "DDB"
    ws.Range("E1").Value = "Sum of Years"
    
    ' Calculate for each year
    For i = 1 To life
        ws.Cells(i + 1, 1).Value = i
        
        ' Straight-line
        ws.Cells(i + 1, 2).Formula = "=SLN(" & cost & "," & salvage & "," & life & ")"
        
        ' Declining balance
        ws.Cells(i + 1, 3).Formula = "=DB(" & cost & "," & salvage & "," & life & "," & i & ",12)"
        
        ' Double declining balance
        ws.Cells(i + 1, 4).Formula = "=DDB(" & cost & "," & salvage & "," & life & "," & i & ",2)"
        
        ' Sum of years' digits
        ws.Cells(i + 1, 5).Formula = "=SYD(" & cost & "," & salvage & "," & life & "," & i & ")"
    Next i
    
    ' Format as currency
    ws.Range("B2:E" & life + 1).NumberFormat = "$#,##0.00"
    
    MsgBox "Depreciation schedules calculated!", vbInformation
End Sub
```

---

## 10. DATA VALIDATION & DROPDOWN LISTS

### Create Account Code Dropdown
```vba
Sub CreateAccountCodeDropdown()
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    ' Define account codes range
    Set rng = ws.Range("AccountCodes")
    
    ' Add data validation
    With ws.Range("C2:C1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=AccountCodes"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Invalid Account Code"
        .InputMessage = ""
        .ErrorMessage = "Please select a valid account code from the list."
        .ShowInput = True
        .ShowError = True
    End With
End Sub
```

### Create Status Dropdown
```vba
Sub CreateStatusDropdown()
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    With ws.Range("J2:J1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Posted,Pending,Cancelled"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid Status"
        .ErrorMessage = "Please select: Posted, Pending, or Cancelled"
        .ShowError = True
    End With
End Sub
```

---

## HOW TO USE THESE TEMPLATES

### Step 1: Create the Workbook
1. Open Excel
2. Create a new workbook
3. Save as **Excel Macro-Enabled Workbook (*.xlsm)**

### Step 2: Add VBA Code
1. Press `Alt + F11` to open VBA Editor
2. Insert → Module
3. Copy and paste the VBA code
4. Close VBA Editor

### Step 3: Set Up Named Ranges
1. Select the cells for each named range
2. Formulas → Define Name
3. Enter the name (e.g., "Current_Assets")
4. Click OK

### Step 4: Add Formulas
1. Click the cell where you want the formula
2. Type or paste the formula
3. Press Enter
4. Copy down/across as needed

### Step 5: Test the Macros
1. Press `Alt + F8`
2. Select the macro
3. Click Run

### Step 6: Save and Use
1. Save your work
2. Enter your data
3. Let Excel calculate the results
4. Customize as needed

---

## SECURITY NOTES

⚠️ **Always review VBA code before running macros from unknown sources**

These templates are provided as-is. Before using:
1. Review all VBA code
2. Test with sample data
3. Ensure formulas work correctly
4. Backup your data regularly

---

## CUSTOMIZATION TIPS

1. **Add Conditional Formatting** - Highlight important values
2. **Create Charts** - Visualize your data
3. **Add Pivot Tables** - Analyze data dynamically
4. **Protect Formulas** - Lock cells with formulas
5. **Add Password Protection** - Secure sensitive data

---

*Need help with a specific template? Let me know and I can create a customized version!*
