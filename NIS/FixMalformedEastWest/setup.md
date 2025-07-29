# FixMalformedEastWest Macro

## 🧾 Description

This macro scans **column F** of the active worksheet for malformed directional text patterns such as:
- `E EAST` → corrects to `EAST`
- `W EAST` → corrects to `WEST`

## 👨‍💻 Author

- **Name**: Ilesanmi Kehinde John  
- **Alias**: Calm  
- **Date Created**: July 29, 2025  

## ⚙️ Setup Instructions

1. **Open Excel.**
2. Press `ALT + F11` to open the **VBA Editor**.
3. Go to `Insert > Module` to create a new module.
4. Paste the full VBA code into the module.
5. Close the VBA editor.

## 🚀 How to Run the Macro

1. Go back to Excel.
2. Press `ALT + F8`, select `FixMalformedEastWest`, and click **Run**.

## 📌 Notes

- The macro processes all non-empty cells in column F.
- It counts and displays the number of corrections made via a message box.
- Ensure column F contains directional strings for accurate results.

## 🛠 Example Before and After

| Before     | After     |
|------------|-----------|
| E EAST     | EAST      |
| W EAST     | WEST      |
| SOUTH-W EAST | SOUTH-WEST |

---

Happy cleaning! ✅
