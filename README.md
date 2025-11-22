<h1>📦 Excel Invoice System</h1>

<h4>By Guransh Dhaliwal</h4>

<h5>🧾 Overview</h5>

<p>
This project shows you how to make a simple invoice system in Microsoft Excel using formulas, dropdowns, and VBA code.
It’s great for beginners who want to automate math, organize products, and even convert kilograms to pounds automatically.
You can also check out the latest version of the project in the Releases section!
</p>

<h5>🚀 Features</h5>
<p>
- Automatically adds up totals, taxes, and grand total<br>
- Converts quantities from kilograms to pounds<br>
- Lets you pick products from a dropdown menu<br>
- Fills in prices and totals automatically<br>
- Uses a VBA macro to do tasks automatically<br>
- Easy to edit and expand for more products<br>
- Developer tools for quick access
</p>

<h5>📁 Setup Instructions</h5>

<p><strong>1. Create Your Workbook</strong></p>

<p>
Open Excel → File → New → Blank Workbook<br><br>

Save it right away:<br>
File → Save As → choose a folder<br>
Name it: <strong>Invoice.xlsm</strong><br>
Set file type: <strong>Excel Macro-Enabled Workbook (.xlsm)</strong>
</p>

<h5>📑 Setting Up Sheets</h5>

<p>
<strong>Sheet 1 → Rename to: Invoice</strong><br>
This is your main invoice page.<br><br>

<strong>Sheet 2 → Rename to: Products</strong><br>
This sheet stores your product list for dropdowns.
</p>

<h5>🛒 Adding Product Names (Sheet2: Products)</h5>

<p>Type your product names in column A, like:</p>

<p>
Product A<br>
Product B<br>
Product C<br>
Product D
</p>

<p><strong>Optional: Create a Named Range</strong></p>

<p>
Highlight A1:A10 → In the Name Box type <strong>ProductsList</strong> → Press Enter
</p>

<h5>📊 Invoice Table Setup (Sheet1: Invoice)</h5>

<p><strong>Start your table at row 19.</strong></p>

<table border="1" cellpadding="6">
<tr><th>Column</th><th>Description</th></tr>
<tr><td>B19:B30</td><td>Product (dropdown list)</td></tr>
<tr><td>M19:M30</td><td>Quantity (kg → auto converts to lbs)</td></tr>
<tr><td>O19:O30</td><td>Unit Price (default: 4.49)</td></tr>
<tr><td>P19:P30</td><td>Total (=M*O/100)</td></tr>
</table>

<br>

<h5>Totals Section</h5>

<table border="1" cellpadding="6">
<tr><th>Cell</th><th>Label / Formula</th></tr>
<tr><td>E31</td><td>Subtotal → =SUM(P19:P30)</td></tr>
<tr><td>E32</td><td>Tax (13%) → =E31*0.13</td></tr>
<tr><td>E33</td><td>Grand Total → =E31+E32</td></tr>
</table>

<h5>🔽 Adding Product Dropdowns</h5>

<p>
Highlight <strong>B19:B30</strong><br>
Data → Data Validation<br>
Allow → List<br>
Source:
</p>

<p><code>=ProductsList</code></p>

<p>If you didn’t make a named list:</p>

<p><code>=Products!$A$1:$A$10</code></p>

<p>Click OK.</p>

<h5>🧠 Adding the VBA Macro</h5>

<p><strong>Open the VBA Editor</strong></p>
<p>
Alt + F11<br>
or Developer → Visual Basic
</p>

<p><strong>If Developer tab is missing:</strong></p>

<p>File → Options → Customize Ribbon → Check “Developer”</p>

<p><strong>Paste this code into Sheet1 (Invoice)</strong></p>

<pre>
<code>
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
</code>
</pre>

<h5>💾 Save again as .xlsm</h5>

<h5>🧪 How to Use</h5>

<p>
- Pick a product in B19:B30<br>
- Type quantity in M19:M30 (kg → converts to lbs)<br>
- Unit price fills into O19:O30<br>
- Total updates in P19:P30<br>
- Subtotal, Tax, Grand Total update automatically<br>
- You can edit prices — everything still updates
</p>

<h5>💡 Tips for Beginners</h5>

<p>
- Click <strong>Enable Macro Content</strong> when opening<br>
- Add products in the Products sheet<br>
- Extend ranges if you add rows<br><br>

<b>Important formulas:</b><br>
Subtotal → =SUM(P19:P30)<br>
Tax → =Subtotal * 0.13<br>
Grand Total → =Subtotal + Tax<br><br>

<strong>Developer shortcut:</strong><br>
Alt + F11 → VBA Editor
</p>

<h5>⚡ Developer Shortcuts</h5>

<table border="1" cellpadding="6">
<tr><th>Action</th><th>Shortcut</th></tr>
<tr><td>Open VBA Editor</td><td>Alt + F11</td></tr>
<tr><td>Run or Edit Macros</td><td>Developer → Macros</td></tr>
<tr><td>Data Validation</td><td>Data → Data Validation</td></tr>
</table>

<h5>🧱 Visual Layout</h5>
<h3>Sheet1 (Invoice)</h3>
<table border="1">
  <tr>
    <td>B19:B30</td>
    <td>Product Name</td>
    <td>M19:M30</td>
    <td>Quantity</td>
  </tr>
  <tr>
    <td>O19:O30</td>
    <td>Unit Price</td>
    <td>P19:P30</td>
    <td>Total</td>
  </tr>
</table>

<br>

<h3>Sheet2 (Products)</h3>
<table border="1">
  <tr><td>A1:A10</td></tr>
  <tr><td>Product A</td></tr>
  <tr><td>Product B</td></tr>
  <tr><td>Product C</td></tr>
  <tr><td>Product D</td></tr>
</table>
