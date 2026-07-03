<h1>Excel Invoice System</h1>

<h4>By Guransh Dhaliwal</h4>

<hr>

<h2>Overview</h2>

<p>
The Excel Invoice System is a beginner-friendly project that demonstrates how to build a fully functional invoice generator in Microsoft Excel.
It uses formulas, dropdown menus, and VBA automation to handle calculations, product selection, and unit conversion automatically.
</p>

<p>
This project is ideal for learning how to combine Excel features with basic automation to build real-world tools like invoicing systems, product trackers, and billing sheets.
You can also find the latest version of this project in the <strong>Releases</strong> section.
</p>

<h2>Features</h2>

<ul>
  <li>Automatic calculation of subtotal, tax, and grand total</li>
  <li>Product selection using dropdown menus</li>
  <li>Auto-filled pricing system</li>
  <li>Quantity conversion (kilograms to pounds)</li>
  <li>VBA automation for faster updates</li>
  <li>Expandable product system</li>
  <li>Clean and beginner-friendly structure</li>
</ul>

<h2>Setup Instructions</h2>

<h3>1. Create Your Workbook</h3>

<p>
Open Microsoft Excel → Create a new blank workbook<br>
Save immediately as: <strong>Invoice.xlsm</strong><br>
File type: <strong>Excel Macro-Enabled Workbook (.xlsm)</strong>
</p>

<h3>2. Create Sheets</h3>

<p>
Rename Sheet1 → <strong>Invoice</strong> (main invoice page)<br>
Rename Sheet2 → <strong>Products</strong> (product database)
</p>

<h3>3. Add Products</h3>

<p>
In the <strong>Products</strong> sheet, enter product names in column A:
</p>

<pre>
Product A
Product B
Product C
Product D
</pre>

<p>
(Optional) Create a named range:
Highlight A1:A10 → Name Box → type <strong>ProductsList</strong>
</p>

<h3>4. Invoice Table Setup</h3>

<p>Start your invoice table at row 19:</p>

<table border="1" cellpadding="6">
<tr><th>Range</th><th>Purpose</th></tr>
<tr><td>B19:B30</td><td>Product selection (dropdown)</td></tr>
<tr><td>M19:M30</td><td>Quantity (kg → converted to lbs)</td></tr>
<tr><td>O19:O30</td><td>Unit price</td></tr>
<tr><td>P19:P30</td><td>Total calculation</td></tr>
</table>

<h3>5. Totals Section</h3>

<table border="1" cellpadding="6">
<tr><th>Cell</th><th>Formula</th></tr>
<tr><td>E31</td><td>=SUM(P19:P30) (Subtotal)</td></tr>
<tr><td>E32</td><td>=E31*0.13 (Tax)</td></tr>
<tr><td>E33</td><td>=E31+E32 (Grand Total)</td></tr>
</table>

<h3>6. Add Product Dropdown</h3>

<p>
Select <strong>B19:B30</strong><br>
Go to Data → Data Validation → List<br>
Use source:
</p>

<pre>=ProductsList</pre>

<p>
Or use direct range:
</p>

<pre>=Products!$A$1:$A$10</pre>

<h2>VBA Automation</h2>

<h3>Enable Developer Tools</h3>

<p>
File → Options → Customize Ribbon → Enable <strong>Developer</strong> tab
</p>

<h3>Open VBA Editor</h3>

<p>
Press <strong>Alt + F11</strong> or go to Developer → Visual Basic
</p>

<h3>Paste The Code (Sheet1 - Invoice) </h3>
<p>The code is placed in a file called excel_code.vba in this repo!</p>
<h2>How to Use</h2>

<ul>
  <li>Select a product from the dropdown</li>
  <li>Enter quantity in kilograms</li>
  <li>Unit price auto-fills</li>
  <li>Total updates automatically</li>
  <li>Subtotal, tax, and grand total update instantly</li>
</ul>

<h2>Tips</h2>

<ul>
  <li>Enable macros when opening the file</li>
  <li>Extend row ranges if needed</li>
  <li>Modify product list anytime in the Products sheet</li>
</ul>

<h2>Keyboard Shortcuts</h2>

<table border="1" cellpadding="6">
<tr><th>Action</th><th>Shortcut</th></tr>
<tr><td>Open VBA Editor</td><td>Alt + F11</td></tr>
<tr><td>Data Validation</td><td>Data → Validation</td></tr>
<tr><td>Run Macros</td><td>Developer → Macros</td></tr>
</table>
