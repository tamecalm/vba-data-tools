# 📘 VBA Function Setup Guide: SumNumbersInText

## ✨ Overview

This guide walks you through setting up a custom VBA function called `SumNumbersInText` in Excel. This function scans a cell's content, extracts all numeric sequences (even if mixed with text), and returns their sum. It is especially useful for summing embedded numbers from alphanumeric strings like `"abc123xyz45"`.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- Press `ALT + F11` in Excel to launch the **Visual Basic for Applications (VBA)** editor.

### Step 2: Insert a New Module

- In the **Project Explorer** (left panel), right-click on your workbook name (e.g., *VBAProject (YourWorkbook.xlsm)*).
- Click **Insert → Module**.
- A new code window opens. This is where you’ll paste the code.

### Step 3: Paste the Code

Copy and paste the following VBA code into the module **as-is** (do not modify anything):

```vba
' === Function: SumNumbersInText ===
' Extracts and sums numeric values from a cell’s text
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 29, 2025

Function SumNumbersInText(cell As Range) As Double
    Dim total As Double
    Dim temp As String
    Dim i As Integer
    Dim char As String

    total = 0
    temp = ""
    
    ' Ensure the cell is not empty
    If cell.Value = "" Then
        SumNumbersInText = 0
        Exit Function
    End If

    ' Loop through each character in the cell's value
    For i = 1 To Len(cell.Value)
        char = Mid(cell.Value, i, 1)
        
        ' If the character is a digit, add it to the temp string
        If IsNumeric(char) Then
            temp = temp & char
        Else
            ' If a complete number is found, add it to the total
            If temp <> "" Then
                total = total + Val(temp)
                temp = "" ' Reset temp for the next number
            End If
        End If
    Next i
    
    ' If there's a number left in temp, add it to the total
    If temp <> "" Then
        total = total + Val(temp)
    End If
    
    SumNumbersInText = total
End Function

' === Subroutine: num ===
' Demonstrates how to call the function and display result
Sub num()
    Dim result As Double
    result = SumNumbersInText(Range("A1"))
    MsgBox "The total is " & result
End Sub
```

---

## ✅ How to Use the Function in Excel

Once the VBA code is added and saved:

1. Go back to your worksheet.
2. Use the function just like a built-in Excel formula:

```excel
=IF(TRIM(D35)="", "NaN", SumNumbersInText(D35))
```

### Explanation:
- `TRIM(D35)=""` checks if the cell is empty.
- If empty, it returns `"NaN"` (not a number).
- If not, it uses `SumNumbersInText(D35)` to sum all numbers in the cell.

---

## 🧪 Example

If cell `D35` contains:

```
"abc123text45more6"
```

Then:

```
SumNumbersInText(D35) → 123 + 45 + 6 = 174
```

---

## 📝 Notes

- Only whole numbers made up of digits `0-9` are recognized.
- It does **not** support decimals, negative numbers, or embedded commas.
- Works on single cells only.

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)  
- **Created On:** November 25, 2024  
- **Use Case:** Summing numbers embedded in alphanumeric strings in Excel  
- **License:** Free for personal and commercial use

---
