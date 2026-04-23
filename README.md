# Excel Tools

A collection of Python utilities for working with Excel files.

## Tools Included

### 1. Excel File Verifier (`/verifier`)
Validates Excel files for data integrity and format compliance.

### 2. Excel Data Generator (`/data-generator`)
Generates random client/test data and saves to Excel using Faker and pandas.

### 3. Excel Comparison App (`/comparison`)
Compares two Excel files and generates a report highlighting differences. Useful for database updates and identifying new entries.

## Quick Start

```bash
git clone https://github.com/Xza85hrf/Excel-Tools.git
cd Excel-Tools
pip install pandas openpyxl faker
```

`tkinter` is part of the Python standard library on Windows and macOS
installers. On Debian/Ubuntu install it once with
`sudo apt install python3-tk`.

## Usage

Each tool is a standalone Tkinter GUI. Launch from the repo root:

```bash
# Run verifier
python verifier/excel_verifier_gui.py

# Run data generator
python "data-generator/Excel Random Data Generator.py"

# Run comparison tool
python comparison/Gui.py
```

## License

MIT License
