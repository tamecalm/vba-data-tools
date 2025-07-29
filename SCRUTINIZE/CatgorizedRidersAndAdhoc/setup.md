# 🧾 Rider, Adhoc & Hub Officer Categorization Script (Excel VBA)

## 📌 Overview

This Excel VBA script categorizes entries in a worksheet into one of the following groups:

- **HUB OFFICER**
- **RIDER**
- **ADHOC**
- **Not In List**

The script reads values from **Column A** of the `Pivot (Old)` worksheet, checks the names against three predefined lists, and writes the corresponding category into **Column K** of the same row.

---

## 📂 File Location & Setup

### Step 1: Open the VBA Editor
- Press `ALT + F11` in Excel.

### Step 2: Insert the Script
- Go to `Insert` > `Module`.
- Paste the full VBA code (see below).

### Step 3: Populate Your Lists
Inside the macro, you’ll see three empty lists:
```vba
adhocList = Array()
hubOfficerList = Array()
ridersList = Array()
```
Populate these with your names, for example:
```vba
adhocList = Array("John Doe", "Jane Smith")
```

---

## 🛠️ How It Works

1. It checks each row in Column A starting from row 2.
2. If the value **contains any known state name**, it is skipped.
3. Otherwise, the name is compared against:
   - `hubOfficerList` (highest priority)
   - `ridersList`
   - `adhocList`
4. If a match is found, the corresponding category is written into **Column K**.
5. If no match is found, it labels the row as `"Not In List"`.

---

## ✅ Example

**Input (Column A):**
```
DANLADI EUGENE
John Doe
Michael Brown
```

**Lists:**
```vba
hubOfficerList = Array("DANLADI EUGENE")
adhocList = Array("John Doe")
ridersList = Array("Michael Brown")
```

**Output (Column K):**
```
HUB OFFICER
ADHOC
RIDER
```

---

## 🧠 Customization Tips

- You can add/remove states in the `statesList` array to skip specific rows.
- Matching is **case-insensitive** and uses **partial match** (`InStr`).
- You can modify the `Offset(0, 10)` to change which column the result goes into.

---

## 🔒 Limitations

- Only works on a **single worksheet** named `Pivot (Old)`
- Matching is based on **partial text matches**
- Will not categorize names if they are also present in the `statesList`

---

## 📎 Author

Created by: *[KEHINDE ILESANMI]*  
Last updated: November 27, 2024
