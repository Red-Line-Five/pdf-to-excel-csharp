# PDF to Excel – Breaker Summary Extractor (C#)

## 📌 Project Description

This project is a **C# Windows Forms desktop application** that reads **engineering PDF files** (such as Single Line Diagrams – SLD) and automatically extracts **electrical breaker information**, then exports a **clean summary to Excel**.

The application:

* Reads abbreviations and descriptions directly from the PDF
* Extracts breaker types (ACB, MCCB, etc.)
* Detects current ratings (e.g. 1200A, 250A, 100A)
* Groups and counts breakers by type, current, and poles
* Exports the final result to an **Excel (.xlsx)** file

This tool is ideal for:

* Electrical engineers
* Consultants and contractors
* Panel schedules and BOQ preparation
* Reducing manual data entry from drawings

---

## ⚙️ How It Works

### 1️⃣ Load PDF File

The user selects a PDF file using an **OpenFileDialog**.

### 2️⃣ Read Abbreviations

The application scans the PDF pages to extract:

* Abbreviations (e.g. MCCB, ACB)
* Full descriptions (e.g. Molded Case Circuit Breaker)

These are stored in a `DataTable` for later matching.

### 3️⃣ Extract Breaker Data

* Reads raw text from a specific PDF page
* Normalizes text (fixes OCR errors, units, spacing)
* Uses **Regular Expressions** to detect current values (e.g. `\d+A`)
* Identifies breaker type and poles (default: 3P)

### 4️⃣ Group & Analyze

* Breakers are grouped by **Type, Current, and Poles**
* Quantities are automatically calculated
* Abbreviations are replaced with full descriptions

### 5️⃣ Display & Export

* Results are shown in a **DataGridView**
* User can export the data to **Excel (.xlsx)** using Microsoft Office Interop

---

## 🧰 Technologies Used

* C# (.NET Framework)
* Windows Forms
* iText 7 (PDF parsing)
* Regular Expressions (Regex)
* Microsoft Office Interop Excel
* DataTable & LINQ

---


## 🚀 How to Run the Project

1. Clone or download the repository
2. Open the solution in **Visual Studio**
3. Restore NuGet packages (iText7)
4. Make sure **Microsoft Excel** is installed
5. Build and run the project

---

## 📤 Export Output

The exported Excel file contains:

* Breaker Type (full description)
* Current rating
* Poles
* Quantity (count)

Sorted by current rating in descending order.

---

## 📬 Contact

For customization, enhancements, or freelance work:

* GitHub: [[https://github.com/YOUR_USERNAME](https://github.com/YOUR_USERNAME)](https://github.com/Red-Line-Five)
* Email: charbel.feghaly.rl@gmail.com
* LinkedIn (optional): [your_linkedin_profile](https://www.linkedin.com/in/charbel-feghaly-916473103/)

---

## 💼 Freelance & Customization

This project can be customized to:

* Support multiple PDF formats
* Detect different breaker rules
* Export to formatted BOQ templates
* Add batch processing

Feel free to contact me for professional use or commer
