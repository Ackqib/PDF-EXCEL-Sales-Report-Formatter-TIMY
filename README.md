# PDF/EXCEL Sales Report Formatter - TIMY

A small Windows desktop app that turns sales reports and inventory/stock reports
(PDF, Excel or CSV, in any layout) into one clean, two-column Excel file:

| Product | Quantity |
|---|---|
| 1012345M | 12 |
| 1023456-34 | 7 |
| 1034567XL3 | 1 |

- **Product** = the 7-digit SKU starting with `10` + the size.
- **Quantity** = the quantity for that product. If a product appears more than once, the quantities are added together.

Everything else (descriptions, prices, store names, totals, page headers) is removed.

## Size format

| In the report | Output |
|---|---|
| Small, SM, S | `S` |
| Medium, MED, MD, M | `M` |
| Large, LG, L | `L` |
| X-Large, Extra Large, XL | `XL` |
| XXL, 2XL, 2X | `XL2` |
| XXXL, 3XL, 3X | `XL3` |
| Waist sizes: 34, W34, 34x32 | `-34` (a hyphen before every number size) |

## What it can read

- **Files:** PDF (with a text layer), `.xlsx`, `.xlsm`, `.xls` and `.csv`. Excel files can have several sheets.
- **One line per item:** the size and quantity are in their own columns.
- **Size grids:** sizes across the top (S M L XL… or 28 30 32 34…) with quantities underneath.
- **Size attached to the SKU:** `1012345-M`, `1012345XL`, `1003169#NAVY\34`.
- **Headings over several lines:** column headings that wrap onto 2 or 3 lines, and descriptions that wrap onto the next line.
- **Choosing the quantity column:** it picks the right one automatically (Qty, On Hand, Available, Units, Sold…). It skips price, cost and value columns, and quantities like "On Order" or "In Transit".
- **Totals:** subtotal and grand-total lines are ignored, so nothing is counted twice.

If the app finds a SKU but can't work out its size or quantity, it doesn't drop the line silently.
The line goes on a separate **Needs Review** sheet in the output file.

## Setup (once)

1. Install [Python 3.10+](https://www.python.org/downloads/). Tick "Add Python to PATH" during setup.
2. Double-click **`Report Converter.bat`**. The first time, it installs the libraries it needs (`requirements.txt`).

## Using it

1. Double-click **`Report Converter.bat`**.
2. Click **Add files…** or drag PDFs and Excel files into the window. You can add as many as you like at once.
3. Click **Convert**.

Each file gets a `<name>_clean.xlsx` saved next to it, or in the folder you choose.

Options:
- **Also make one combined file:** one workbook that adds up every selected report.
- **Keep 0 quantities:** zero quantities are left out unless this is ticked.
- **Also save raw PDF→Excel:** also saves the unfiltered PDF conversion.
- **Quantity column name:** type a column name here (e.g. `Sold`) if a report has several quantity columns and the app picks the wrong one.

## Command line

```bash
python converter.py report1.pdf report2.xlsx -o output_folder --combined
```

Run `python converter.py --help` for all options.

## Files

| File | Purpose |
|---|---|
| `converter.py` | Reads the files and extracts the data. Can also be run from the command line |
| `app.py` | The desktop window (tkinter, with optional drag & drop) |
| `Report Converter.bat` | Double-click launcher. Installs requirements the first time |
| `requirements.txt` | Python libraries needed |

## Limitations

- Scanned PDFs (photos of paper) contain no text and would need OCR first. The app warns you when it finds one.
- Variant suffixes on SKUs (e.g. `1011078-1`) are listed under the 7-digit SKU.
