# BBraun invoice extractor

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Status](https://img.shields.io/badge/Status-Active-success)

## 📖 Introduction
**BBraun invoice extractor** is a desktop GUI tool designed to eliminate manual data entry for international logistics and supply chain invoices. 

Processing complex PDF invoices with varying layouts can be error-prone and time-consuming. This tool automates the extraction of critical data points—such as Invoice Numbers, Dates, Item Codes, Quantities, and Prices—and consolidates them into a clean, structured Excel spreadsheet for immediate analysis or ERP upload.

## ✨ Key Features
*   **Multi-Region Support:** distinctive parsing logic for different international formats:
    *   **Format 1:** Optimized for European layouts (Germany / Spain).
    *   **Format 2:** Optimized for Asian layouts (Malaysia).
*   **Batch Processing:** Upload and process multiple PDF files simultaneously.
*   **Visual Data Preview:** View extracted data in a grid before saving to ensure accuracy.
*   **Smart Consolidation:** Automatically merges data from all selected files into a single Excel report.
*   **User-Friendly GUI:** Built with a modern, clean interface compliant with corporate design standards.

## 🛠️ Prerequisites
Ensure you have Python installed. The project relies on the following libraries:

*   **GUI:** `tkinter` (Standard library)
*   **Data Handling:** `pandas`, `openpyxl`
*   **PDF Extraction:** `pdfplumber`
*   **Image Processing:** `Pillow` (PIL)

## 📦 Installation

1.  Clone this repository:
    ```bash
    git clone https://github.com/nfpltech3/BBraun-invoice-extractor
    cd BBraun-invoice-extractor
    ```

2.  Install the required dependencies:
    ```bash
    pip install pandas pdfplumber openpyxl Pillow
    ```

## 🚀 Usage Guide

1.  **Launch the Application**
    Run the main script to start the GUI:
    ```bash
    python BBraun_Invoice_forall.py
    ```

2.  **Select Region Format (Crucial Step)**
    *   Look for the **"Invoice Format"** section in the GUI.
    *   Select **Format 1** if your invoices are from **Germany or Spain**.
    *   Select **Format 2** if your invoices are from **Malaysia**.
    *   *Note: Failure to select the correct region may result in incomplete data extraction.*

3.  **Load Files**
    *   Click **"📂 Select Files"** and choose one or more PDF invoices.
    *   The tool will verify the files and display them in the list.

4.  **Extract & Export**
    *   Click **"Generate Invoice Data"**.
    *   Review the **Data Preview** table.
    *   The tool will automatically save a timestamped Excel file (e.g., `Combined_Germany_Spain_5files_20231025.xlsx`) in the same folder as your source PDFs.

---
© Nagarkot Forwarders Pvt Ltd
