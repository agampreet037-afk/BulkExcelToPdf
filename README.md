# Bulk Excel to PDF Converter

A simple Python desktop application that converts one or more Microsoft Excel workbooks into PDF files while preserving formatting. The application uses Microsoft Excel through the Windows COM interface to ensure high-quality PDF output.

## Features

* 📂 Bulk convert multiple Excel files to PDF
* 📄 Export an entire workbook as a single PDF
* 📑 Convert all worksheets while preserving formatting
* ⚡ Fast and easy-to-use graphical interface
* 🖥️ Standalone Windows executable (.exe) available using PyInstaller
* 📁 Automatically saves PDFs to the chosen output location

## Python Libraries

Install the required dependencies:

```bash
pip install pywin32
```

## Installation

Clone the repository:

```bash
git clone https://github.com/<your-username>/BulkExcelToPdf.git
```

Move into the project folder:

```bash
cd BulkExcelToPdf
```

Run the application:

```bash
python main.py
```

## How It Works

1. Select one or more Excel workbooks.
2. Choose the destination folder for the PDFs.
3. The application opens each workbook using Microsoft Excel.
4. Each workbook is exported as a high-quality PDF.
5. Excel closes automatically after the conversion is complete.

## Limitations

* Works only on Windows.
* Requires Microsoft Excel to be installed.
* Password-protected workbooks are not supported.

## Future Improvements

* Drag-and-drop support
* Progress bar
* Dark mode
* Remember last used folders
* Multi-threaded conversion
* Custom PDF naming options
* Installer for easier distribution
* Logging and error reporting

## Author

**Agampreet Singh**

If you find this project useful, feel free to ⭐ the repository and contribute improvements or suggestions.
