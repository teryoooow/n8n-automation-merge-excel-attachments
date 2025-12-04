![n8n](https://img.shields.io/badge/automation-n8n-orange?logo=n8n)
![Excel](https://img.shields.io/badge/Excel-Merger-green?logo=microsoft-excel)
![Status](https://img.shields.io/badge/Status-Production%20Ready-blue)
![License](https://img.shields.io/badge/License-MIT-lightgrey)


# 📊 n8n Automation – Merge Multiple Excel Email Attachments Into One File

This n8n workflow automatically retrieves incoming emails with Excel attachments and merges **any number** of `.xlsx` files (3, 4, 5, or more) into **one consolidated Excel file** in a single worksheet.  
Perfect for automating reporting, analytics, and dataset aggregation.

---

## 🚀 Features

- ✔️ Automatically detects and processes all Excel attachments  
- ✔️ Works with ANY number of files  
- ✔️ Merges all rows into a single unified dataset  
- ✔️ Supports complex, multi-column spreadsheets  
- ✔️ Fully automated end-to-end  
- ✔️ Output: one cleaned, consolidated Excel file  
- ✔️ Works with Gmail or any email provider supported by n8n  

---

## 🧠 Use Cases

- Aggregating daily/weekly reports from multiple departments  
- Merging multiple exported spreadsheets into a single dataset  
- Processing automated system reports  
- Centralizing Excel data for analytics or BI tools  
- Automated back-office data preparation  

---

## 🏗️ Workflow Architecture

+-----------------------------+
|     Email Trigger Node      |  <- Gmail / IMAP
+-------------+---------------+
              |
              v
+-----------------------------+
|     Filter Excel Files      |  (.xlsx only)
+-------------+---------------+
              |
              v
+-----------------------------+
|      Extract From File      |  Convert Excel -> JSON
+-------------+---------------+
              |  (multiple JSON arrays)
              v
+-----------------------------+
|     Code Node (Merge)       |  Combine all rows
+-------------+---------------+
              |
              v
+-----------------------------+
|       Convert To File       |  Generate Excel
+-------------+---------------+
              |
              v
+-----------------------------+
|   Deliver File (Email,      |
|     Upload, Save, etc.)     |
+--------------------------



---

## 🧩 Key Nodes Used

| Node | Purpose |
|------|---------|
| **Gmail / Email Trigger** | Fetch incoming emails with attachments |
| **IF Node** | Filter `.xlsx` attachments |
| **Extract From File** | Convert Excel → JSON |
| **Code Node** | Merge JSON rows |
| **Convert to File** | Convert JSON → Excel (.xlsx) |
| **Email / Drive / FTP** | Deliver final combined file |

---

## 🧑‍💻 Code Node – Merge Logic

```js
// Collect all rows from every Excel attachment
let merged = [];

for (const item of $input.all()) {
  // Each item.json contains a single row
  merged.push(item.json);
}

// Output as proper n8n items
return merged.map(row => ({ json: row }));

```

## 📁 Output

The workflow generates:

merged-output.xlsx


Containing all rows from all attachments in one sheet.

🛠️ Installation

1. Import this workflow using its exported JSON file
2. Add your Gmail/Email credentials in n8n
3. Set the email search criteria (e.g., label, subject, sender)
4. Activate workflow

🏷️ Repository Topics (GitHub Tags)
n8n
n8n-workflow
automation
excel
merge-excel
email-automation
gmail
data-processing
etl
workflow-automation
no-code
low-code

<!-- End of README -->
