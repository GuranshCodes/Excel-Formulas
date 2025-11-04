📦 Excel Invoice System

By Guransh Dhaliwal

🧾 Overview

This project shows you how to make a simple invoice system in Microsoft Excel using formulas, dropdowns, and VBA code.
It’s great for beginners who want to automate math, organize products, and even convert kilograms to pounds automatically.
You can also check out the latest version of the project in the Releases section!

🚀 Features

Automatically adds up totals, taxes, and grand total

Converts quantities from kilograms to pounds

Lets you pick products from a dropdown menu

Fills in prices and totals automatically

Uses a VBA macro to do tasks automatically

Easy to edit and expand for more products

Developer tools for quick access

📁 Setup Instructions
1. Create Your Workbook

Open Excel → File → New → Blank Workbook

Save it right away:

Go to File → Save As

Choose a folder

Name it: Invoice.xlsm

Set Save as type to: Excel Macro-Enabled Workbook (.xlsm)*

📑 Setting Up Sheets
Sheet 1 → Rename to Invoice

This is the main sheet where you’ll make invoices.

Sheet 2 → Rename to Products

This sheet will hold your product names for the dropdown list.

🛒 Adding Product Names (Sheet2: Products)

Type your product names in column A, like this:

Product A  
Product B  
Product C  
Product D


Optional: Create a named list to make it easier later.

Highlight A1:A10

In the small box above column A, type: ProductsList

Press Enter

📊 Invoice Table Setup (Sheet1: Invoice)

Start your table at row 19.

Column	What It’s For
B19:B30	Product (dropdown list)
M19:M30	Quantity (in kg, auto converts)
O19:O30	Unit Price (default 4.49)
P19:P30	Total (=M*O/100)
Totals Section
Cell	Label / Formula
E31	Subtotal → =SUM(P19:P30)
E32	Tax (13%) → =E31*0.13
E33	Grand Total → =E31+E32
🔽 Adding Product Dropdowns

Highlight B19:B30

Go to Data → Data Validation

Under Allow, choose List

In Source, type:

=ProductsList


If you didn’t name your list, use:

=Products!$A$1:$A$10


Click OK and you’re done!

🧠 Adding the VBA Macro
Open the VBA Editor

Press Alt + F11, or

Go to Developer → Visual Basic

If you don’t see the Developer tab:
File → Options → Customize Ribbon → Check “Developer”

Paste This Code Into Sheet1 (Invoice)
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngM As Range, rngO As Range
    Dim row As Long
    Dim mValue As Double
    Dim pFormula As String

    If Target.CountLarge > 1 Then Exit Sub

    Set rngM = Me.Range("M19:M30")
    Set rngO = Me.Range("O19:O30")

    Application.EnableEvents = False

    If Not Intersect(Target, rngM) Is Nothing Then
        row = Target.Row
        If IsNumeric(Target.Value) And Target.Value <> "" Then
            mValue = Target.Value * 2.20462262
            Target.Value = mValue
        Else
            Target.Value = ""
        End If
        Me.Cells(row, "O").Value = 4.49
        pFormula = "=M" & row & "*O" & row & "/100"
        Me.Cells(row, "P").Formula = pFormula
    End If

    If Not Intersect(Target, rngO) Is Nothing Then
        row = Target.Row
        pFormula = "=M" & row & "*O" & row & "/100"
        Me.Cells(row, "P").Formula = pFormula
    End If

    Application.EnableEvents = True
End Sub


💾 Save the file again as .xlsm to keep your macros working.

🧪 How to Use

Pick a product from the dropdown in B19:B30

Type a quantity in M19:M30 (kg → it’ll change to pounds)

Unit price appears in O19:O30

The total in P19:P30 calculates automatically

Subtotal, Tax, and Grand Total update live

You can also edit prices manually if you want — everything updates automatically.

💡 Tips for Beginners

Always click Enable Content when you open the .xlsm file

To add more products, update the Products sheet

To add more rows, edit the VBA and formula ranges

Formulas to remember:

Subtotal → =SUM(P19:P30)

Tax → =Subtotal * 0.13

Grand Total → =Subtotal + Tax

Open Developer tools quickly:

Alt + F11 → VBA Editor

⚡ Developer Shortcuts
Action	Shortcut / Where to Find It
Open VBA Editor	Alt + F11 or Developer → Visual Basic
Run or Edit Macros	Developer → Macros
Data Validation	Data → Data Validation
🧱 Visual Layout
Sheet1 (Invoice)
+---------+-------------------+------------+----------+
| B19:B30 | Product Name      | M19:M30    | Quantity |
| O19:O30 | Unit Price        | P19:P30    | Total    |
+---------+-------------------+------------+----------+

Sheet2 (Products)
+---------+
| A1:A10  |
| Product A|
| Product B|
| Product C|
| Product D|
+---------+
