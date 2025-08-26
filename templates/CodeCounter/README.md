VBA Code Counter (vba-codecount)

A lightweight VBA utility to scan through all VBA-enabled Excel workbooks (.xlsm) in a folder and generate a code metrics report.
- This module counts:
  - 📄 Total lines of code
  - 💬 Comment lines
  - 🧮 Effective code lines (total – comments)
- Results are written to a Data sheet for further analysis.

---

📂 How It Works
1. User selects a folder (via a folder picker).
2. The macro loops through each .xlsm file inside.
3. For each workbook, it:
  - Opens the workbook (read-only, no link updates).
  - Inspects every VBA component (Modules, ClassModules, UserForms).
  - Counts:
    - All lines of code
    - Comment-only lines (lines starting with ')
  - Records results into the Data worksheet.
4. Workbook closes automatically after scanning.
5. Dashboard cell C21 stores the last used folder path.

---

📊 Output
- Results are written into Sheets("Data") starting from row 2:

| Column | Content                         |
| ------ | ------------------------------- |
| A      | Filename                        |
| C      | Total code lines                |
| D      | Code lines (excluding comments) |

---

🔧 Setup
1. Ensure you have the following sheets:
  - Dashboard (with cell C21 for the folder path)
  - Data (for results table)
2. Place the VBA code inside a standard module (e.g. Module1).
3. Enable Trust Access to the VBA Project Object Model:
  - File > Options > Trust Center > Trust Center Settings > Macro Settings > Developer Macro Settings > Check "Trust access to the VBA project object model".

---

🚀 Usage
1. In Dashboard!C21, optionally type the starting folder path.
2. Run the macro.
3. A folder picker will prompt (defaults to Dashboard!C21).
4. Process runs automatically, results appear in Data sheet.

---

📝 Notes
1. Current implementation detects comment lines only by checking if the line begins with '.
2. For more advanced parsing, you can modify the code block in the loop:
```
If Left(Trim(s), 1) = "'" Then
    CodeLineComments = CodeLineComments + 1
End If
```

---

Example ideas:
1. Count Public Sub, Function, or Property signatures.
2. Track block structures (If…End If, For…Next).
3. Count procedure lengths.

---

📌 Example Output
| File            | Total Lines | Effective Lines |
|-----------------|------------:|----------------:|
| ReportTool.xlsm |         580 |             472 |
| Parser.xlsm     |         230 |             180 |
| MacroKit.xlsm   |        1020 |             890 |

---

⚡ Roadmap
 - Add support for .xls (Excel 97–2003)
 - Include procedure count (number of Sub/Function)
 - Breakdown by module type (Module, Class, Form)
 - Summary chart on Dashboard

---
