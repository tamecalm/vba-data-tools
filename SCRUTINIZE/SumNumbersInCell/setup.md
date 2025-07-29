# 📊 SumNumbersInCell Excel UDF

## Overview

`SumNumbersInCell` is a custom Excel **User Defined Function (UDF)** written in **VBA**. It extracts and sums **all numeric values (including decimals)** embedded within the text of a single Excel cell.

For example:
- Input: `abc12xyz3.5`
- Output: `15.5`

This is useful when you're working with cells containing text mixed with numbers and want to compute the total of all numbers in the string.

---

## 🛠️ Setup Instructions

### 1. Open the VBA Editor

- Open your Excel workbook.
- Press `ALT + F11` to open the **VBA Editor**.

### 2. Insert the Function Code

- In the VBA editor, go to: `Insert` → `Module`
- Paste the following code:

```vba
Function SumNumbersInCell(cell As Range) As Double
    Dim regex As Object
    Dim matches As Object
    Dim match As Variant
    Dim total As Double

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "\d+(\.\d+)?"
    End With

    If regex.Test(cell.Value) Then
        Set matches = regex.Execute(cell.Value)
        For Each match In matches
            total = total + CDbl(match.Value)
        Next match
    End If

    SumNumbersInCell = total
End Function
```

### 3. Save Your Workbook as Macro-Enabled

- Go to: `File` → `Save As`
- Select file type: `Excel Macro-Enabled Workbook (*.xlsm)`
- Save your file.

---

## ✅ Usage Instructions

1. In any cell (e.g., `A1`), enter a string that contains numbers and text, such as:

   ```
   Order12Box5.5Total
   ```

2. In another cell, enter the formula:

   ```excel
   =SumNumbersInCell(A1)
   ```

3. The result will be:

   ```
   17.5
   ```

---

## 🔍 How It Works

- Uses **regular expressions** to identify all numeric patterns in the text.
- Matches both integers (e.g., `42`) and decimal values (e.g., `3.14`).
- Sums all matched numbers and returns the total.

---

## ⚠️ Limitations

- Does **not support negative numbers** or numbers with commas (e.g., `-12`, `1,000.50`)
- Works only on **single cells** (not ranges)
- Input must be **text or string-compatible content**

---

## 💡 Potential Enhancements

Consider improving the function to:

- Support **negative numbers**
- Handle **thousands separators** (e.g., `1,000.00`)
- Accept and process a **range of cells**

Let me know if you'd like help adding these!

---

## 📎 Author

Created by: *[KEHINDE ILESANMI]*  
Last Updated: July 29, 2025
