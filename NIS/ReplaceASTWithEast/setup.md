# 📘 VBA Macro Setup Guide: ReplaceASTWithEast

## ✨ Overview

This macro scans column **F** of the active sheet and replaces every standalone text that **ends with "AST"** (like `"SOUTHWEST AST"`) by inserting a space before `"EAST"`, making it `"SOUTHWEST EAST"`.

It uses **regular expressions** for accurate matching and substitution.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- Open Excel.
- Press `ALT + F11` to launch the **Visual Basic for Applications (VBA)** editor.

### Step 2: Insert the Code

1. In the **Project Explorer** (left panel), right-click your workbook.
2. Select: **Insert > Module**
3. Paste the following macro code exactly as is:

```vba
' === Macro: ReplaceASTWithEast ===
' Description: Replaces any text ending in 'AST' in column F with ' EAST' while preserving the prefix.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 29, 2025

Sub ReplaceASTWithEast()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim regex As Object

    ' Set the worksheet to the currently active one
    Set ws = ActiveSheet

    ' Determine the last used row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True                 ' Replace all matches in a string
        .IgnoreCase = True            ' Case-insensitive match (e.g., 'ast' or 'AST')
        .Pattern = "(\S*)AST\b"       ' Match a word ending in AST (e.g., 'WESTAST')
    End With

    ' Loop through all rows in column F
    For i = 1 To lastRow
        If Not IsError(ws.Cells(i, "F").Value) Then
            cellValue = CStr(ws.Cells(i, "F").Value)
            
            ' If the pattern matches, replace 'AST' with ' EAST'
            If regex.test(cellValue) Then
                ws.Cells(i, "F").Value = regex.Replace(cellValue, "$1 EAST")
            End If
        End If
    Next i

    ' Inform the user that replacement is complete
    MsgBox "AST replacements completed in column F.", vbInformation
End Sub
```

---

## ✅ How to Run the Macro

1. Save the file as a **Macro-Enabled Workbook (`.xlsm`)**.
2. Press `ALT + F8`, select `ReplaceASTWithEast`, and click **Run**.
3. Make sure you're on the worksheet where replacements should happen (it works on the active sheet).
4. It processes column **F** only.

---

## 🧪 Example

| Original Column F         | After Running Macro       |
|---------------------------|---------------------------|
| "SOUTHEAST AST"           | "SOUTHEAST EAST"          |
| "NORTHWEST AST"           | "NORTHWEST EAST"          |
| "ASTONISH"                | _No Change_ (not ending in AST) |
| "EAST AST AST"            | "EAST EAST EAST"          |

- Only whole words or parts ending with `"AST"` are replaced.
- `InStr` is not used, so partial text like `"ASTONISH"` is preserved.

---

## 📝 Notes

- The macro:
  - Ignores cells with errors
  - Works top-down through column **F**
  - Uses **VBScript.RegExp**, so no external libraries are needed
- The regular expression pattern used:
  - `(\S*)AST\b`
    - `\S*` matches any non-space characters (prefix)
    - `AST\b` ensures "AST" appears at the end of a word

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)  
- **Created On:** July 29, 2025  
- **License:** Free for personal, business, or educational use  
- **Tip:** Always back up your worksheet before running bulk-replacement macros.

---
