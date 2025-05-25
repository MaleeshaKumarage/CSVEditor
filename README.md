# CSV/Excel Record Limiter GUI

A simple Python GUI tool to filter and limit rows from CSV or Excel files, with flexible column-based filtering and export.

---

## Features

- Supports both CSV (`.csv`) and Excel (`.xlsx`) files as input.
- Lets you select which columns to filter on, and apply various conditions.
- Real-time display of how many rows match your filters.
- Output file type matches input file type (CSV in, CSV out; Excel in, Excel out).
- Easy-to-use graphical interface.

---

## Requirements

- Python 3.7 or newer
- The following Python packages:
  - `tkinter` (usually included with Python)
  - `openpyxl` (for Excel file support)
  - `pyinstaller` (for building a standalone executable)

---

## Installation

1. **Clone or download this repository.**

2. **Install required packages:**

   Open a terminal or command prompt and run:

   ```sh
   pip install openpyxl pyinstaller
   ```

   (If you only need CSV support, you can skip `openpyxl`.)

---

## Usage

### Run from source

```sh
python csv_limiter_gui.py
```

### Build a Standalone Executable

To create a single-file Windows executable (no console window), run:

```sh
python -m PyInstaller --onefile --windowed --distpath "C:\Users\user\MyAppFolder" csv_limiter_gui.py
```

- The executable will be created in `C:\Users\user\MyAppFolder`.
- You can change the `--distpath` to any folder you like.

---

## How to Use

1. **Open the app.**
2. **Click "Browse..."** and select a CSV or Excel file.
3. **Select columns to filter:**  
   Use the dropdown to pick a column, click "Add", then set the condition and value. Repeat for more columns.
4. **Set a row limit** (optional).
5. **Check the real-time matching row count.**
6. **Click "Process"** to save the filtered/limited file.  
   The output file type will match your input file type.

---

## Notes

- If you select a file type other than `.csv` or `.xlsx`, the app will show an error.
- For Excel support, you must have `openpyxl` installed.
- The output file will always include all columns, but only rows matching your selected filters.

---

## License

MIT License
