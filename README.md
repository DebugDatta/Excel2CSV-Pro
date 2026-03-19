

---

# Excel to CSV Converter

A high-performance multi-sheet Excel-to-CSV converter built with Streamlit. Upload Excel files, process all sheets in parallel, and download clean CSV outputs as individual files or a single combined dataset.

---

## Live App

[https://excel2csvpro.streamlit.app](https://excel2csvpro.streamlit.app)

---

## GitHub Repository

[https://github.com/DebugDatta/Excel2CSV-Pro](https://github.com/DebugDatta/Excel2CSV-Pro)

---

## Overview

This tool enables fast and scalable conversion of Excel files (.xlsx, .xls) into CSV format.

### Core Capabilities

* Upload one or multiple Excel files
* Process all sheets concurrently
* Export each sheet as an individual CSV
* Optionally combine all sheets into one stacked CSV
* Download results as a ZIP archive

---

## Features

* Multi-file upload support
* Parallel sheet processing (ThreadPoolExecutor)
* Per-sheet CSV export
* Optional stacked (combined) CSV mode
* Data preview (first 5 rows)
* ZIP download per input file
* Progress tracking during execution
* Structured logging (rows, columns, processing time)
* Automatic cleanup of temporary files
* File size limit: 50 MB per file

---

## Usage

1. Open the app
   [https://excel2csvpro.streamlit.app](https://excel2csvpro.streamlit.app)

2. Upload one or more Excel files

3. (Optional) Enable Stack Mode

4. Preview the data

5. Click Convert All

6. Download the generated ZIP file

7. Click Clear Results to remove temporary files

---

## Local Setup

### Clone Repository

```bash
git clone https://github.com/DebugDatta/Excel2CSV-Pro.git
cd Excel2CSV-Pro
```

### Install Dependencies

```bash
pip install streamlit pandas openpyxl xlrd
```

### Run Application

```bash
streamlit run app.py
```

### Open in Browser

[http://localhost:8501](http://localhost:8501)

---

## Example

### Input

sales_data.xlsx with sheets:

* Jan
* Feb
* Mar



### Sample Data

Jan

| Product | Revenue |
| ------- | ------- |
| A       | 100     |
| B       | 200     |

Feb

| Product | Revenue |
| ------- | ------- |
| A       | 150     |
| B       | 250     |

Mar

| Product | Revenue |
| ------- | ------- |
| A       | 180     |
| B       | 300     |



## Output

### ZIP Structure

```
sales_data.zip
├── sales_data_Jan.csv
├── sales_data_Feb.csv
├── sales_data_Mar.csv
└── sales_data_stacked.csv (optional)
```



### Individual CSV

sales_data_Jan.csv

```
Product,Revenue
A,100
B,200
```



### Stacked CSV

sales_data_stacked.csv

```
_sheet,Product,Revenue
Jan,A,100
Jan,B,200

Feb,A,150
Feb,B,250

Mar,A,180
Mar,B,300
```



## Notes

* ``` _sheet ``` column identifies the source sheet
* Blank rows separate sheets in stacked mode
* Original schema is preserved
* Sheet order remains unchanged

---

## Technical Details

* Uses ThreadPoolExecutor with as_completed for parallel processing
* Each sheet is read independently using pandas (thread-safe)
* Temporary files cleaned using shutil.rmtree
* CSV encoding: utf-8-sig (Excel compatible)
* Logs written to conversion.log

---

## Dependencies

* streamlit (UI)
* pandas (data processing)
* openpyxl (xlsx support)
* xlrd (xls support)

Python 3.10+ recommended

---

## Use Cases

* Preparing Excel data for databases
* Batch converting multi-sheet workbooks
* Automating Excel export workflows
* Cleaning data for analytics pipelines

---
