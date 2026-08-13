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

## Usage
Place `Contingency-sample.xlsm` and `.acc` files in the same folder;
Open CMD in folder location and type:

```

"C:\Python27\python.exe" acc2excel.py "." --workbook "ten-file.xlsm"

```