# Excel Change Identifier

## Project Overview

This project is a Python-based tool that identifies differences between two Excel files and highlights the differences in a new Excel file. The tool can detect new, moved, deleted, or modified test cases and mark them accordingly.

## Project Structure

```plaintext
├── README.md
└── src
    ├── __init__.py
    ├── main.py
    ├── change_identifier.py
    ├── excel_file.py
    ├── excel_reader.py
    └── excel_writer.py
```
- src/main.py: The entry point of the program.
- src/change_identifier.py: Contains the main logic for identifying changes in Excel files.
- src/excel_file.py: Models an Excel file as an object.
- src/excel_reader.py: Reads Excel files.
- src/excel_writer.py: Writes the change results to a new Excel file.

## Usage
Run the program with the following command:

```bash
python src/main.py
```

A dialog will appear allowing you to select the original Excel file and the Excel file to check. The program compares these files, highlights the differences, and saves the result in a new Excel file.

## Dependencies
- openpyxl: For reading and writing Excel files.
- tkinter: For the graphical user interface to display file dialogs.
