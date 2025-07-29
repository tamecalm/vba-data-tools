# 🌍 Region Categorization Script (Excel VBA)

## 📌 Overview

This Excel VBA script assigns a **geopolitical region** to each row based on Nigerian state names found in **Column A** of a worksheet. The corresponding region is inserted into **Column J**.

Supported Regions:
- **NORTH-CENTRAL**
- **NORTH-EAST**
- **NORTH-WEST**

---

## 🛠️ Setup Instructions

### Step 1: Open VBA Editor
- In Excel, press `ALT + F11`.

### Step 2: Insert the Script
- Go to `Insert` → `Module`.
- Paste the full VBA code provided below.

### Step 3: Save Your Workbook
- Use `File` → `Save As`.
- Select `Excel Macro-Enabled Workbook (*.xlsm)`.

---

## ✅ How It Works

1. Script scans values in **Column A** (starting from row 2).
2. It checks if the text contains any known state names using **partial, case-insensitive match**.
3. If a match is found:
   - It assigns the corresponding region name.
   - Inserts the region into **Column J** of the same row.
4. If no match is found, the cell in Column J is **cleared** (not deleted).
5. A message box is displayed after processing.

---

## 🧪 Example

| A (State Input)    | J (Region Output) |
|--------------------|-------------------|
| Benue              | NORTH-CENTRAL     |
| Borno              | NORTH-EAST        |
| Kano               | NORTH-WEST        |
| Unknown State      | *(blank)*         |

---

## 🧠 Customization Tips

- Modify the `Select Case` block to include:
  - More regions (e.g., South-West, South-South)
  - Exact matches or more sophisticated logic
- Adjust `stateCell.Offset(0, 9)` to change output column
- Modify `vbTextCompare` to `vbBinaryCompare` for **case-sensitive** matching

---

## 📎 Author

Created by: *[KEHINDE ILESANMI]*  
Last Updated: July 29, 2025
