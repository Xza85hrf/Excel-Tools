# Excel Tools

A collection of Python utilities for working with Excel files.

## Tools Included

### 1. Excel File Verifier (`/verifier`)
Validates Excel files for data integrity and format compliance.

### 2. Excel Data Generator (`/data-generator`)
Generates random client/test data and saves to Excel using Faker and pandas.

### 3. Excel Comparison App (`/comparison`)
Compares two Excel files and generates a report highlighting differences. Useful for database updates and identifying new entries.

## Requirements

```bash
pip install pandas openpyxl faker tkinter
```

## Usage

Each tool has its own directory with a standalone GUI application:

```bash
# Run verifier
python verifier/main.py

# Run data generator
python data-generator/main.py

# Run comparison tool
python comparison/main.py
```

## License

MIT License
