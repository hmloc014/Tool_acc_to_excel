## ACC folder import core

`acc2excel.process_acc_folder()` implements the workbook workflow without any
UI dependency:

```python
from acc2excel import process_acc_folder

result = process_acc_folder(
    r"E:\path\to\0_sample_test",
    # The source template is copied to Contingency.xlsm before Excel opens it.
)
```

The function recursively finds `.acc` files, rejects duplicate file names,
copies `CONTINGENCY - Sample 1.xlsm` to `Contingency.xlsm` if the working copy
does not exist, generates the PSS/E available-capacity report, duplicates the
`Sample` sheet to its left, restores the monitored-element order stored in the
ACC summary, replaces short contingency labels with their full PSS/E event
descriptions, writes the five headers (`Monitored Element`, `Contingency`,
`Others`, `Base Flow`, and `Maximum Flow`) in `D4:H4`, and writes the first data
row in `D5:H5`. PSS/E 33 normally requires Python 2.7. Close the target workbook
in Excel before running the import.

## Desktop app

Double-click `Acc to Excel.exe`, choose **Select working folder**, and select a
folder in the Windows File Explorer picker. The folder must contain one
contingency sample `.xlsm` workbook and `.acc` files in that folder or its
subfolders. Enter the desired output workbook name when prompted; `.xlsm` is
added automatically if omitted, and an existing file is never overwritten.
The packaged app starts without opening a CMD window.
The app calls `process_acc_folder()` and creates the requested workbook in the
selected folder. Progress is shown in the white log
area; previous run logs remain visible for manual checking, with each new run
separated clearly. A completion dialog appears when the workbook is ready. Close the
target workbook in Excel before starting.

`acc_to_excel_app.py` is the Python 2.7-compatible UI source and
`Acc_to_Excel.spec` is its PyInstaller build configuration.

## Python file usage in CMD
Place `Contingency-sample.xlsm` and `.acc` files in the same folder;
Open CMD in folder location and type:

```

"C:\Python27\python.exe" acc2excel.py "." --workbook "ten-file.xlsm"

```
