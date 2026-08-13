# -*- coding: utf-8 -*-
"""Import PSS/E ACC available-capacity reports into an Excel template.

The public entry point is :func:`process_acc_folder`.  It intentionally has no
UI code so a Tkinter, wxPython, or other front end can call it later.

PSS/E 33 uses Python 2.7, so this module stays compatible with both Python 2.7
and modern Python 3 versions.
"""
from __future__ import print_function

import io
import os
import re
import shutil
import sys
import tempfile


DEFAULT_PSSE_PATH = r"C:\Program Files (x86)\PTI\PSSE33\PSSBIN"
DEFAULT_SOURCE_WORKBOOK_NAME = "CONTINGENCY - Sample.xlsm"
DEFAULT_WORKBOOK_NAME = "Contingency.xlsm"
DEFAULT_TEMPLATE_SHEET = "Sample"
DEFAULT_HEADER_CELL = "D4"
DEFAULT_START_CELL = "D5"

REPORT_COLUMNS = (
    "Monitored Element",
    "Contingency",
    "Others",
    "Base Flow",
    "Maximum Flow",
)

_INVALID_SHEET_CHARACTERS = re.compile(r"[\[\]:*?/\\]")
_REPORT_NUMBER = re.compile(
    r"^(?:[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|\*+|N/?A)$",
    re.IGNORECASE,
)

try:
    text_type = unicode
except NameError:  # pragma: no cover - Python 3 branch
    text_type = str


class AccImportError(Exception):
    """Base exception for the ACC import workflow."""


class DuplicateAccNameError(AccImportError):
    """Raised when recursive ACC discovery finds duplicate file names."""


class PsseReportError(AccImportError):
    """Raised when PSS/E cannot generate or parse an ACC report."""


class WorkbookImportError(AccImportError):
    """Raised when the Excel template cannot be copied or updated safely."""


def _as_text(value):
    if isinstance(value, text_type):
        return value
    try:
        return value.decode("utf-8")
    except AttributeError:
        return text_type(value)


def find_acc_files(root_folder):
    """Return every ``.acc`` file below *root_folder*, sorted by full path."""
    root_folder = os.path.abspath(root_folder)
    if not os.path.isdir(root_folder):
        raise AccImportError("Folder does not exist: {0}".format(root_folder))

    acc_files = []
    for current_root, _directories, files in os.walk(root_folder):
        for file_name in files:
            if os.path.splitext(file_name)[1].lower() == ".acc":
                acc_files.append(os.path.join(current_root, file_name))

    acc_files.sort(key=lambda path: _as_text(path).lower())
    return acc_files


def validate_unique_acc_names(acc_files):
    """Reject duplicate ACC base names, case-insensitively.

    The worksheet name is based on the file name without ``.acc``.  Detecting
    duplicates before PSS/E or Excel is started prevents partial workbooks.
    """
    grouped = {}
    for acc_path in acc_files:
        stem = os.path.splitext(os.path.basename(acc_path))[0]
        grouped.setdefault(_as_text(stem).lower(), []).append(acc_path)

    duplicates = [paths for paths in grouped.values() if len(paths) > 1]
    if duplicates:
        details = []
        for paths in duplicates:
            details.append("\n  - " + "\n  - ".join(paths))
        raise DuplicateAccNameError(
            "Duplicate .acc file names were found. Rename them before importing:"
            + "".join(details)
        )


def worksheet_title_for_acc(acc_path):
    """Return an Excel-safe worksheet title derived from an ACC file name."""
    stem = _as_text(os.path.splitext(os.path.basename(acc_path))[0])
    title = _INVALID_SHEET_CHARACTERS.sub("_", stem).strip().strip("'")
    title = title[:31]
    if not title:
        raise WorkbookImportError(
            "The ACC file name cannot be converted to a worksheet name: {0}".format(
                acc_path
            )
        )
    return title


def validate_unique_worksheet_titles(acc_files):
    """Reject names that collide after Excel's 31-character/title rules."""
    grouped = {}
    for acc_path in acc_files:
        title = worksheet_title_for_acc(acc_path)
        grouped.setdefault(title.lower(), []).append((title, acc_path))

    collisions = [items for items in grouped.values() if len(items) > 1]
    if collisions:
        details = []
        for items in collisions:
            title = items[0][0]
            paths = [item[1] for item in items]
            details.append(
                "\n  Worksheet '{0}' would be produced by:\n  - {1}".format(
                    title, "\n  - ".join(paths)
                )
            )
        raise DuplicateAccNameError(
            "ACC names collide after applying Excel worksheet-name rules:"
            + "".join(details)
        )


def _coerce_report_number(token):
    """Convert report numeric text to int/float while preserving markers."""
    if not _REPORT_NUMBER.match(token):
        raise ValueError("Not a report number: {0}".format(token))
    if "*" in token or token.upper().replace("/", "") == "NA":
        return token
    value = float(token)
    if value.is_integer() and "." not in token and "e" not in token.lower():
        return int(value)
    return value


def parse_available_capacity_text(report_text):
    """Parse the five requested columns from PSS/E's capacity report.

    Each returned row is ready to be written below the template's existing
    headers.  The first field combines the FROM bus, TO bus, and circuit fields
    exactly as the template formulas expect.
    """
    rows = []
    header_found = False
    contingency_start = None
    others_start = None

    for line in report_text.splitlines():
        if (
            "CONTINGENCY LABEL" in line
            and "OTHERS" in line
            and "AVAILABLE" in line
        ):
            header_found = True
            contingency_start = line.index("<----- CONTINGENCY")
            others_start = line.index("OTHERS")
            continue

        if contingency_start is None:
            continue

        monitored_element = " ".join(line[:contingency_start].split())
        contingency = " ".join(line[contingency_start:others_start].split())
        numeric_tokens = line[others_start:].split()

        if not monitored_element or not contingency or len(numeric_tokens) != 7:
            continue

        try:
            numeric_values = [_coerce_report_number(token) for token in numeric_tokens]
        except ValueError:
            continue

        # The report tail is: Others, Base, Maximum, Impact, Rating, Percent,
        # Available.  Only the first three values are required by the template.
        rows.append(tuple([monitored_element, contingency] + numeric_values[:3]))

    if not header_found:
        raise PsseReportError(
            "PSS/E report does not contain the available-capacity table header."
        )
    return rows


def parse_available_capacity_report(report_path):
    """Read and parse a PSS/E available-capacity report file."""
    with io.open(report_path, "r", encoding="utf-8", errors="ignore") as report_file:
        return parse_available_capacity_text(report_file.read())


def _normalized_acc_key(value):
    """Normalize PSS/E fixed-width text for reliable order/mapping lookups."""
    return " ".join(_as_text(value).split()).upper()


def load_psspy(psse_path=DEFAULT_PSSE_PATH):
    """Load the PSS/E Python API lazily and provide a useful error message."""
    psse_path = os.path.abspath(psse_path)
    if not os.path.isdir(psse_path):
        raise PsseReportError("PSS/E PSSBIN folder does not exist: {0}".format(psse_path))

    if psse_path not in sys.path:
        sys.path.insert(0, psse_path)
    path_parts = os.environ.get("PATH", "").split(os.pathsep)
    if psse_path.lower() not in [part.lower() for part in path_parts]:
        os.environ["PATH"] = psse_path + os.pathsep + os.environ.get("PATH", "")

    try:
        import psspy
    except ImportError as exc:
        raise PsseReportError(
            "Could not import PSS/E from {0}. Run this module with the Python "
            "version supported by your PSS/E installation (PSS/E 33 normally "
            "uses Python 2.7). Original error: {1}".format(psse_path, exc)
        )
    return psspy


def load_pssarrays(psse_path=DEFAULT_PSSE_PATH):
    """Load PSS/E's ACC array API, used for order and descriptions."""
    psse_path = os.path.abspath(psse_path)
    if not os.path.isdir(psse_path):
        raise PsseReportError("PSS/E PSSBIN folder does not exist: {0}".format(psse_path))

    if psse_path not in sys.path:
        sys.path.insert(0, psse_path)
    path_parts = os.environ.get("PATH", "").split(os.pathsep)
    if psse_path.lower() not in [part.lower() for part in path_parts]:
        os.environ["PATH"] = psse_path + os.pathsep + os.environ.get("PATH", "")

    try:
        import pssarrays
    except ImportError as exc:
        raise PsseReportError(
            "Could not import PSS/E pssarrays from {0}. Original error: {1}".format(
                psse_path, exc
            )
        )
    return pssarrays


def order_and_describe_capacity_rows(acc_path, rows, pssarrays_module):
    """Use ACC arrays to restore monitored-element order and event descriptions."""
    summary = pssarrays_module.accc_summary(accfile=acc_path)
    if summary is None or getattr(summary, "ierr", 0):
        raise PsseReportError(
            "PSS/E could not read the ACC summary for {0}.".format(acc_path)
        )

    element_order = {}
    for index, element in enumerate(summary.melement):
        element_order.setdefault(_normalized_acc_key(element), index)

    contingency_order = {}
    for index, label in enumerate(summary.colabel):
        contingency_order.setdefault(
            _normalized_acc_key(label), (index, _as_text(label).strip())
        )

    descriptions = {}
    ordered_rows = []
    for original_index, row in enumerate(rows):
        element_key = _normalized_acc_key(row[0])
        label_key = _normalized_acc_key(row[1])
        if element_key not in element_order:
            raise PsseReportError(
                "Monitored element from the capacity report was not found in "
                "the ACC summary: {0}".format(row[0])
            )
        if label_key not in contingency_order:
            raise PsseReportError(
                "Contingency label from the capacity report was not found in "
                "the ACC summary: {0}".format(row[1])
            )

        if label_key not in descriptions:
            _label_index, acc_label = contingency_order[label_key]
            solution = pssarrays_module.accc_solution(
                accfile=acc_path,
                colabel=acc_label,
                stype="contingency",
                busmsm=0.5,
                sysmsm=5.0,
            )
            if solution is None or getattr(solution, "ierr", 0):
                raise PsseReportError(
                    "PSS/E could not read contingency details for {0}.".format(
                        acc_label
                    )
                )
            codesc = getattr(solution, "codesc", None)
            if isinstance(codesc, (text_type, bytes)):
                codesc = [codesc]
            description_parts = [
                _as_text(part).strip() for part in (codesc or []) if _as_text(part).strip()
            ]
            if not description_parts:
                raise PsseReportError(
                    "PSS/E returned no description for contingency {0}.".format(
                        acc_label
                    )
                )
            descriptions[label_key] = "; ".join(description_parts)

        label_index = contingency_order[label_key][0]
        described_row = tuple([row[0], descriptions[label_key]] + list(row[2:]))
        ordered_rows.append(
            (element_order[element_key], label_index, original_index, described_row)
        )

    ordered_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    return [item[3] for item in ordered_rows]


def generate_available_capacity_report(psspy_module, acc_path, report_path):
    """Ask PSS/E to write the available-capacity table for one ACC file."""
    default_integer = getattr(psspy_module, "_i", -999)
    default_float = getattr(psspy_module, "_f", -999.0)

    # STATUS(1)=2 selects the available-capacity table. STATUS(3)=1 uses Rate A.
    status = [2, 1, 1, 0, 0, 0, 0, 0, 0, 0]
    intval = [0, 0, 0, 0, default_integer]
    realval = [0.5, 5.0, 100.0, 0.0, 0.0, 0.0, default_float]

    ierr = psspy_module.report_output(2, report_path, [1, 0])
    if ierr:
        raise PsseReportError(
            "PSS/E could not open its report output file (error {0}): {1}".format(
                ierr, report_path
            )
        )

    report_error = None
    try:
        report_function = getattr(
            psspy_module,
            "accc_single_run_report_3",
            getattr(psspy_module, "accc_single_run_report_2", None),
        )
        if report_function is None:
            raise PsseReportError(
                "This PSS/E version has no supported ACCC single-run report API."
            )
        report_error = report_function(status, intval, realval, acc_path)
    finally:
        psspy_module.report_output(1, "", [0, 0])

    if report_error:
        raise PsseReportError(
            "PSS/E could not read {0} (ACCC report error {1}).".format(
                acc_path, report_error
            )
        )


def extract_acc_reports(
    acc_files, psspy_module, pssarrays_module, temporary_folder
):
    """Generate and parse all reports before Excel is allowed to modify a file."""
    reports = []
    for index, acc_path in enumerate(acc_files):
        report_path = os.path.join(temporary_folder, "report_{0:04d}.txt".format(index))
        generate_available_capacity_report(psspy_module, acc_path, report_path)
        rows = parse_available_capacity_report(report_path)
        rows = order_and_describe_capacity_rows(
            acc_path, rows, pssarrays_module
        )
        reports.append(
            {
                "acc_path": acc_path,
                "sheet_name": worksheet_title_for_acc(acc_path),
                "rows": rows,
            }
        )
    return reports


def _copy_workbook_for_output(source_path, output_path, overwrite_output):
    source_path = os.path.abspath(source_path)
    output_path = os.path.abspath(output_path)
    if os.path.normcase(source_path) == os.path.normcase(output_path):
        return source_path

    output_folder = os.path.dirname(output_path)
    if output_folder and not os.path.isdir(output_folder):
        os.makedirs(output_folder)
    if os.path.exists(output_path) and not overwrite_output:
        raise WorkbookImportError(
            "Output workbook already exists: {0}".format(output_path)
        )
    shutil.copy2(source_path, output_path)
    return output_path


def _prepare_workbook_copy(root_folder, workbook_name, source_workbook_name):
    """Create the working workbook copy without opening the source template."""
    workbook_path = os.path.abspath(os.path.join(root_folder, workbook_name))
    if os.path.isfile(workbook_path):
        return workbook_path

    source_path = os.path.abspath(os.path.join(root_folder, source_workbook_name))
    if not os.path.isfile(source_path):
        candidates = []
        for file_name in os.listdir(root_folder):
            candidate_path = os.path.join(root_folder, file_name)
            if not os.path.isfile(candidate_path):
                continue
            extension = os.path.splitext(file_name)[1].lower()
            normalized = re.sub(r"[^a-z0-9]", "", file_name.lower())
            if extension in (".xlsx", ".xlsm") and "sample" in normalized and (
                "contingency" in normalized or "contigency" in normalized
            ):
                candidates.append(candidate_path)

        if len(candidates) == 1:
            source_path = candidates[0]
        elif len(candidates) > 1:
            raise WorkbookImportError(
                "Multiple contingency sample workbooks were found. Pass an "
                "explicit source_workbook_name:\n  - {0}".format(
                    "\n  - ".join(sorted(candidates))
                )
            )
        else:
            raise WorkbookImportError(
                "Source workbook does not exist: {0}".format(source_path)
            )

    source_extension = os.path.splitext(source_path)[1].lower()
    destination_extension = os.path.splitext(workbook_path)[1].lower()
    if source_extension != destination_extension:
        raise WorkbookImportError(
            "The copied workbook must keep the template extension. Source is "
            "{0}, but destination is {1}.".format(
                source_extension, destination_extension
            )
        )

    return _copy_workbook_for_output(source_path, workbook_path, False)


def _create_excel_application():
    try:
        import win32com.client
    except ImportError as exc:
        raise WorkbookImportError(
            "Excel automation requires pywin32/win32com. Original error: {0}".format(
                exc
            )
        )
    return win32com.client.DispatchEx("Excel.Application")


def _copy_template_sheet_to_left(template, worksheets):
    """Copy *template* immediately before itself and return the new sheet."""
    sheet_count_before_copy = worksheets.Count
    template_index_before_copy = template.Index
    template.Copy(worksheets.Item(template_index_before_copy), None)
    if worksheets.Count != sheet_count_before_copy + 1:
        raise WorkbookImportError(
            "Excel did not duplicate the template worksheet inside the "
            "target workbook."
        )
    return worksheets.Item(template_index_before_copy)


def write_reports_to_workbook(
    workbook_path,
    reports,
    template_sheet=DEFAULT_TEMPLATE_SHEET,
    header_cell=DEFAULT_HEADER_CELL,
    start_cell=DEFAULT_START_CELL,
    excel_factory=None,
    visible=False,
):
    """Duplicate the template once per report and write each table at *start_cell*.

    Excel COM is used deliberately: it copies the complete worksheet and keeps
    VBA, drawings, formulas, validation, and other features in the ``.xlsm``
    package that third-party workbook libraries can drop.
    """
    excel_factory = excel_factory or _create_excel_application
    excel = None
    workbook = None
    saved = False
    old_automation_security = None
    stage = "starting Excel"

    try:
        stage = "creating Excel application"
        excel = excel_factory()
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            old_automation_security = excel.AutomationSecurity
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            old_automation_security = None

        stage = "opening workbook {0}".format(os.path.abspath(workbook_path))
        workbook = excel.Workbooks.Open(
            os.path.abspath(workbook_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )
        if workbook.ReadOnly:
            raise WorkbookImportError(
                "Workbook opened read-only. Close it in Excel and try again: {0}".format(
                    workbook_path
                )
            )

        stage = "reading worksheets"
        worksheets = workbook.Worksheets
        existing = {}
        template = None
        for index in range(1, worksheets.Count + 1):
            worksheet = worksheets.Item(index)
            existing[_as_text(worksheet.Name).lower()] = worksheet.Name
            if _as_text(worksheet.Name).lower() == _as_text(template_sheet).lower():
                template = worksheet

        if template is None:
            raise WorkbookImportError(
                "Template worksheet '{0}' was not found in {1}.".format(
                    template_sheet, workbook_path
                )
            )

        for report in reports:
            title_key = _as_text(report["sheet_name"]).lower()
            if title_key in existing:
                raise WorkbookImportError(
                    "Worksheet '{0}' already exists in the workbook. Rename the "
                    "ACC file or remove the old worksheet before importing.".format(
                        report["sheet_name"]
                    )
                )
            existing[title_key] = report["sheet_name"]

        for report in reports:
            stage = "duplicating Sample for worksheet {0}".format(
                report["sheet_name"]
            )
            # Positional arguments map to Worksheet.Copy(Before, After). Copy
            # immediately before Sample so the generated tabs appear left of it.
            new_sheet = _copy_template_sheet_to_left(template, worksheets)
            stage = "renaming worksheet to {0}".format(report["sheet_name"])
            new_sheet.Name = report["sheet_name"]

            stage = "writing worksheet {0}".format(report["sheet_name"])
            anchor = new_sheet.Range(start_cell)
            start_row = anchor.Row
            start_column = anchor.Column
            header_anchor = new_sheet.Range(header_cell)
            header_range = new_sheet.Range(
                header_anchor,
                new_sheet.Cells(
                    header_anchor.Row,
                    header_anchor.Column + len(REPORT_COLUMNS) - 1,
                ),
            )
            header_range.Value = (REPORT_COLUMNS,)

            # Headers occupy D4:H4; the first ACC record starts at D5.
            values = list(report["rows"])

            used_range = new_sheet.UsedRange
            used_last_row = used_range.Row + used_range.Rows.Count - 1
            data_last_row = start_row + max(len(values), 1) - 1
            clear_last_row = max(used_last_row, data_last_row)
            new_sheet.Range(
                new_sheet.Cells(start_row, start_column),
                new_sheet.Cells(clear_last_row, start_column + len(REPORT_COLUMNS) - 1),
            ).ClearContents()

            if values:
                destination = new_sheet.Range(
                    new_sheet.Cells(start_row, start_column),
                    new_sheet.Cells(
                        data_last_row, start_column + len(REPORT_COLUMNS) - 1
                    ),
                )
                destination.Value = tuple(tuple(row) for row in values)

        stage = "saving workbook {0}".format(os.path.abspath(workbook_path))
        workbook.Save()
        saved = True
    except AccImportError:
        raise
    except Exception as exc:
        raise WorkbookImportError(
            "Excel could not import the ACC reports while {0}: {1}".format(
                stage, exc
            )
        )
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=saved)
            except Exception:
                pass
        if excel is not None:
            if old_automation_security is not None:
                try:
                    excel.AutomationSecurity = old_automation_security
                except Exception:
                    pass
            try:
                excel.Quit()
            except Exception:
                pass


def process_acc_folder(
    root_folder,
    workbook_name=DEFAULT_WORKBOOK_NAME,
    source_workbook_name=DEFAULT_SOURCE_WORKBOOK_NAME,
    output_path=None,
    template_sheet=DEFAULT_TEMPLATE_SHEET,
    header_cell=DEFAULT_HEADER_CELL,
    start_cell=DEFAULT_START_CELL,
    psse_path=DEFAULT_PSSE_PATH,
    overwrite_output=False,
    visible=False,
    psspy_module=None,
    pssarrays_module=None,
    excel_factory=None,
):
    """Run the complete folder-to-workbook workflow.

    Steps:
      1. Recursively find ACC files and reject duplicate base names.
      2. Copy the macro-enabled sample to the working workbook when needed.
      3. Generate PSS/E available-capacity tables for every ACC file.
      4. Duplicate ``Sample`` once per ACC file and name it after that file.
      5. Write the five-column table at ``D5`` and save the workbook.

    ``output_path=None`` updates the template workbook in place.  Pass a new
    ``output_path`` to preserve the original template.
    """
    root_folder = os.path.abspath(root_folder)
    acc_files = find_acc_files(root_folder)
    if not acc_files:
        raise AccImportError("No .acc files were found below {0}.".format(root_folder))
    validate_unique_acc_names(acc_files)
    validate_unique_worksheet_titles(acc_files)

    workbook_path = _prepare_workbook_copy(
        root_folder, workbook_name, source_workbook_name
    )

    psspy_module = psspy_module or load_psspy(psse_path)
    pssarrays_module = pssarrays_module or load_pssarrays(psse_path)
    if hasattr(psspy_module, "psseinit"):
        psspy_module.psseinit(200000)

    temporary_folder = tempfile.mkdtemp(prefix="acc-capacity-")
    try:
        reports = extract_acc_reports(
            acc_files, psspy_module, pssarrays_module, temporary_folder
        )
    finally:
        shutil.rmtree(temporary_folder, ignore_errors=True)

    destination_path = output_path or workbook_path
    destination_path = _copy_workbook_for_output(
        workbook_path, destination_path, overwrite_output
    )
    write_reports_to_workbook(
        destination_path,
        reports,
        template_sheet=template_sheet,
        header_cell=header_cell,
        start_cell=start_cell,
        excel_factory=excel_factory,
        visible=visible,
    )

    return {
        "workbook_path": destination_path,
        "acc_count": len(reports),
        "sheets": [
            {
                "acc_path": report["acc_path"],
                "sheet_name": report["sheet_name"],
                "row_count": len(report["rows"]),
            }
            for report in reports
        ],
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Import PSS/E ACC available-capacity reports into an XLSM template."
    )
    parser.add_argument("folder", help="Folder containing the workbook and ACC files")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_NAME)
    parser.add_argument(
        "--source-workbook", default=DEFAULT_SOURCE_WORKBOOK_NAME
    )
    parser.add_argument("--output", default=None)
    parser.add_argument("--visible", action="store_true")
    args = parser.parse_args()

    result = process_acc_folder(
        args.folder,
        workbook_name=args.workbook,
        source_workbook_name=args.source_workbook,
        output_path=args.output,
        visible=args.visible,
    )
    print("Imported {0} ACC file(s) into {1}".format(result["acc_count"], result["workbook_path"]))
    for sheet in result["sheets"]:
        print("  {0}: {1} rows".format(sheet["sheet_name"], sheet["row_count"]))
