# Nagarkot Invoice Processor - User Guide

## Introduction
The **Nagarkot Invoice Processor** is a specialized desktop automation tool designed to extract line-item data (codes, quantities, prices) from B. Braun invoice PDFs. It normalizes data from multiple regions (Germany/Spain and Malaysia) into a standardized Excel report for easy import and processing.

## How to Use

### 1. Launching the App
1. Locate the application executable (`.exe`).
2. Double-click to launch. The application will open in full-screen mode.

### 2. The Workflow (Step-by-Step)

1. **Select Files**:
   - Click the **📂 Select PDF Files** button.
   - Navigate to your folder and select one or multiple PDF invoice files.
   - *Note: You can select multiple files at once. Files must be text-based PDFs (not scanned images).*

2. **Choose Format**:
   - Select the appropriate extraction logic based on the invoice origin:
     - **Format 1 (Germany / Spain)**: For invoices from European origins (typically containing "N DE" or similar origin markers).
     - **Format 2 (Malaysia)**: For Asian/Malaysian invoices (typically having a different layout with headers like "HS Code" or "DO No").

3. **Generate Data**:
   - Click the **Generate Invoice Data** button in the bottom right.
   - The application will scan all selected files.
   - **Success**: A message box will appear, and the Excel file will be saved.
   - **Data Preview**: Extracted items will populate the central table for verification.

4. **Output**:
   - The final Excel file is saved **in the same folder as the first selected PDF**.
   - Filename structure: `Combined_[Format]_[Count]files_[Timestamp].xlsx`.

## Interface Reference

| Control / Input | Description | Expected Input |
| :--- | :--- | :--- |
| **Select PDF Files** | Opens file dialog to choose invoices. | `.pdf` files only. |
| **Clear List** | Removes all files and clears the preview. | N/A |
| **Extraction Format** | Toggles regex logic for different styles. | **Format 1**: EU (DE/ES)<br>**Format 2**: Asia (MY) |
| **Generate Invoice Data** | Runs the extraction engine. | Click to execute. |
| **Preview Table** | Shows extracted line items. | Columns: Inv No, Date, Code, Qty, Price. |

## Troubleshooting & Validations

If you see an error or unexpected results, check this table:

| Message / Issue | Possible Cause | Solution |
| :--- | :--- | :--- |
| **"No data extracted from selected files."** | 1. PDF is a scanned image (no selectable text).<br>2. Wrong Format selected.<br>3. Invoice layout has changed. | 1. Ensure PDF text is selectable (copy-paste works).<br>2. Try switching between Format 1 and Format 2.<br>3. Check if file contains standard headers (e.g., "Invoice", "Document"). |
| **"Please select at least one PDF file."** | Tried to run without choosing files. | Click "Select PDF Files" first. |
| **Empty Excel Rows** | The script filters out items without a valid "Code". | Normal behavior. Code structure must be 1-6 digits matching the pattern. |
| **Warning: Nagarkot Logo.png not found** | The logo file is missing from the app directory. | Ensure `logo.png` is in the same folder as the `.exe` (or embedded during build). |

---

**Technical Note**: Results are processed using `pdfplumber`. If a PDF is password-protected or encrypted, extraction may fail.
