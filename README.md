# 📊 Excel Data Extractor System – EslamHub

Welcome to the **Excel Data Extractor** system by EslamHub! 🚀  
This project provides a smart, reusable Excel tool for extracting filtered data from any table inside Excel files (local or external), using built-in VBA macros and a friendly interface.

---

## 📁 Contents

- 📊 **DataExtractor.xlsm**: The main Excel file for data extraction.
- 💻 **VBA Module**: Includes code for importing filtered data and clearing previous results.
- ⚙️ **Helper Functions**: `GetWorkbookPath()` and `GetSheetName()` to simplify setup.

---

## 🧠 How It Works

The system uses **AdvancedFilter** to extract rows from any source table based on your criteria and outputs only selected columns.

### 🧾 Sheet: `Config`

| Cell | Purpose |
|------|---------|
| `B1` | File path of source workbook (`=GetWorkbookPath()` for current file) |
| `B2` | Sheet name in source workbook (`=GetSheetName()` for the first sheet) |
| `B3` | Starting cell of the source table (e.g., `A2`) |

Make sure these three inputs are correct and point to a valid range. If any of them are incorrect or missing, the extraction will fail.

---

### 📄 Sheet: `Result`

| Row | Description |
|-----|-------------|
| Row 1 | Column headers to match (e.g., `Invoice Date`, `Customer`) |
| Row 2 | Filtering criteria (e.g., `>=45818`, `=Youssef`) |
| Row 4 | Output headers (the columns you want to extract from the source) |

You can pull any combination of columns by simply typing their header names in Row 4. These must exactly match the headers in the source table.

---

## ▶️ Buttons

- ✅ **Get**: Imports the filtered data based on your conditions.
- ❌ **Clear**: Deletes the previously imported data from the result sheet.

---

## 📌 Example

### Filter Area:

| Invoice Date | Customer |
|--------------|----------|
| `>=45818`    | `Youssef` |

### Output Columns (Row 4):

`Invoice Date`, `Invoice No`, `Customer`, `Product`, `Quantity`, `Total`

---

## 🧩 Use Cases

- Sales reports  
- Filtering employee data  
- Extracting product lists  
- Custom reporting systems

---

## 🧪 Macros & Functions

- `ImportFilteredData()` → Main extraction macro
- `ClearImportedData()` → Clears results
- `GetWorkbookPath()` → Returns current workbook path
- `GetSheetName()` → Returns the name of the first sheet

---

## ℹ️ Notes

- Ensure the source file **exists and is accessible**
- Sheet name and start cell **must be correct**
- Row 4 headers must **exactly match** those in the source table
- Criteria can include:
  - `=value`
  - `>=value`, `<=value`
  - Wildcards like `*text*`

---

## 📥 How to Use

1. Go to the `Config` sheet
   - Enter the source file path (or use `=GetWorkbookPath()`)
   - Enter the sheet name (or use `=GetSheetName()`)
   - Set the top-left cell of your data table (e.g., `A2`)
2. Go to the `Result` sheet
   - Fill Row 1 with headers to filter
   - Fill Row 2 with criteria
   - Fill Row 4 with headers to extract
3. Click the **Get** button
4. To remove results, click **Clear**

---

## 🌐 Connect with Me

📺 [YouTube](https://www.youtube.com/@eslamhub)
📱 [TikTok](https://www.tiktok.com/@eslamhub)
💼 [LinkedIn](https://www.linkedin.com/in/eslamhub)
🐦 [X](https://x.com/eslamhub)
📘 [Facebook](https://www.facebook.com/eslamhub1)
📸 [Instagram](https://www.instagram.com/eslam.hub)

---

#Excel #VBA #DataExtraction #EslamHub
