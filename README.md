# CSV Data Processor for Production Orders

An automated desktop application designed to clean, validate, and standardize CSV vehicle production datasets. Built with a responsive Tkinter GUI, the application processes raw data sheets, runs data verification slices, and outputs structured Excel workbooks containing both comprehensive validation sheets and clean, production-ready extracts.

## Features

### 1. Advanced Data Processing & Validation
* **VIN Segmentation & Character Extraction:** * Extracts the specific **11th character** (Plant Code designation) directly from the raw `VEHICLE-NUMBER`.
    * Trims and standardizes `VEHICLE-NUMBER` strings to a strict 17-character limit (`VEHICLE-NUMBER-processed`).
    * Calculates string lengths for both original and processed values to flags sequence mismatches.
* **Engine Code Standardisation:** * Slices raw `ENGINE-NUMBER` strings down to a precise 14-character limit.
    * Tracks original vs. processed string lengths across sequential validation columns.
* **Lot Mapping & Translation:** * Translates complex `DELIVERY-NUMBER` values into simplified `LOT-NUMBER` identifiers using customized mid-string slicing rules.
* **Dynamic Code Consolidation:** * Scans, matches, and automatically drops arbitrary trailing un-named split option arrays (`Unnamed: X`), merging them seamlessly with the primary `CODES` column into a unified, hyphen-separated `COMBINED-CODE` column.

### 2. Dual-Sheet Architecture
Every processed file dynamically compiles an Excel workbook (`*_processed.xlsx`) split into two distinct operational views:
* **Processed Data Sheet:** The definitive engineering and troubleshooting workbench containing all original headers, calculated string metrics, isolated character segments, and transformed fields.
* **Clean Data Sheet:** A minimal, highly readable summary layer mapping business-critical keys directly to their cleaned data sources:
    * *Lot No* $\rightarrow$ `LOT-NUMBER`
    * *Commission No* $\rightarrow$ `ORDER-NUMBER`
    * *Body No* $\rightarrow$ `PRODUCTION-NUMBER`
    * *Chassis No* $\rightarrow$ `VEHICLE-NUMBER-processed`
    * *Paint Color* $\rightarrow$ `PAINT`
    * *Upholstery No* $\rightarrow$ `INTERIOR`
    * *Engine No* $\rightarrow$ `ENGINE-NUMBER-processed`
    * *Option* $\rightarrow$ `COMBINED-CODE`

### 3. Professional Excel Theme Matrix
To guarantee rapid row scanning on the shopfloor, headers and columns are styled using strict visual anchors:
* **Global Table Styling:** Dark corporate blue headers (`#366092`) with white bold text, bounded by clean, light grey data cell borders (`#D3D3D3`).
* **Olive Green Matrix (Hex: `#C4D79B`):** Applied exclusively to original source data structures and their relative metrics (`VEHICLE-NUMBER`, `VIN-11th-character`, `VIN-STRING-LENGTH`, `ENGINE-NUMBER`, `ENGINENO-STRING-LENGTH`).
* **Soft Yellow Fill (Hex: `#FFF2CC`):** Applied to secondary operational parameters, calculated target segments, and processed transaction values (`ORDER-NUMBER`, `VEHICLE-NUMBER-processed`, `VIN-processed-STRING-LENGTH`, etc.).
* **Auto-Fit Width Engine:** Column dimensions dynamically calculate structural padding (+3 spacing offset) based on maximum cell string lengths to prevent string truncation.

### 4. High-Fidelity UI/UX App Shell
* **Dark-Mode Interface:** A beautiful, resource-efficient dark industrial UI theme (`#121212`) engineered for desktop operations.
* **Dynamic Multi-Thread Animation:** Features a smooth alpha-composited brand loop embedded inside a static structural background canvas.
* **On-the-Fly Localization:** Supports continuous runtime context switching between **English (EN)** and **Vietnamese (VI)** across all titles, paths, CTA buttons, and interactive error/success dialog popups.

---

## Technical Stack & Dependencies

* **Runtime Environment:** Python 3.x
* **Core Libraries:**
    * `pandas` (Data manipulation, ingestion, and sequence mapping)
    * `openpyxl` (Engine for advanced sheet layout, border structures, and hex coloring tables)
    * `tkinter` (Native GUI layout framework and variable tracking)
    * `Pillow (PIL)` (Image transformation, resizing engines, and GIF frame rendering)

---

## Installation & Setup

1. Clone the repository or extract the project directory locally:
   ```bash
   git clone <repository-url>
   cd csv-data-processor
   
2. Install the necessary system packages via pip:
   ```bash
   pip install pandas openpyxl pillow
   
3. Ensure the asset files match the directory configuration root:
   ```bash
    ├── main.py          # Application execution entry point
    ├── icon.png         # OS Window Title Bar branding asset 
    ├── bg1.png          # App frame background mask
    ├── logo_t.gif       # Multi-frame animated brand loop asset
    ├── EN.png           # Flag iconography asset (English)
    └── VI.png           # Flag iconography asset (Vietnamese)
   
---

## Usage
1. Fire up the processing script via your terminal environment:
   ```bash
    python main.py

2. Press the BROWSE FILE / CHỌN TỆP DỮ LIỆU button to select your target raw semicolon-delimited CSV data spreadsheet.
3. Click EXECUTE PROCESS / BẮT ĐẦU to parse the configuration patterns.
4. Your fully formatted, highlighted .xlsx file will output instantly right next to the original file directory.