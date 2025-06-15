# DATA-CLEANING
 Clean and prepare a raw dataset (with nulls, duplicates, inconsistent formats)
1. ✅ Autofit Rows and Columns
📌 Not a formula – this is done manually or via VBA.

How to Use:
Select the entire sheet ➜ Double-click between row/column headers, or

VBA:

vba
Copy
Edit
Sub AutoFitSheet()
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Sub
2. 🔍 Find & Replace
📌 Not a formula, use Ctrl+H or macro.

VBA Example:
vba
Copy
Edit
Sub FindReplace()
    Cells.Replace What:="old", Replacement:="new", LookAt:=xlPart
End Sub
3. 🔠 Lowercase & Uppercase
Function	Formula Example
Lowercase	=LOWER(A1)
Uppercase	=UPPER(A1)

4. ✂️ Trim & Proper Case
Function	Formula Example
Trim extra spaces	=TRIM(A1)
Proper case	=PROPER(A1)

5. 🔡 Text to Columns
UI: Data ➜ Text to Columns
Or use:

excel
Copy
Edit
=SPLIT(A1, ",")   // Google Sheets only
6. 🧹 Removing Duplicates
Method	Formula
Show unique values	=UNIQUE(A1:A100)

🔁 For Excel: Use Data ➜ Remove Duplicates from the ribbon.

7. 🕳️ Filling Empty Cells
Description	Formula
Fill empty with previous value	=IF(A2="",A1,A2)
Google Sheets Array version	=ARRAYFORMULA(IF(A2:A="",A1:A,A2:A))

8. ❓ IFERROR
Scenario	Formula
Handle division error	=IFERROR(A1/B1, "Error")

9. 🎨 Formatting
No formula – use Conditional Formatting rules.

Example Rule:
Highlight blanks:
=ISBLANK(A1) ➜ Apply format (fill color, etc.)

10. 📏 Gridlines
Not a formula – toggle in:

View ➜ Uncheck Gridlines, or use VBA:

vba
Copy
Edit
ActiveWindow.DisplayGridlines = False ' or True
📂 Folder Structure for GitHub
plaintext
Copy
Edit
excel-utils/
│
├── README.md
├── formulas.xlsx         # Sample workbook
├── scripts/
│   ├── autofit.vba
│   ├── find_replace.vba
│   └── formatting.vba
