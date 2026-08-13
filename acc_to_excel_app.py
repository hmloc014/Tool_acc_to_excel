# -*- coding: utf-8 -*-
"""Desktop UI for the ACC-to-Excel workbook workflow."""
from __future__ import print_function

import ctypes, os, re, sys, threading

try:
    import Queue as queue
    import Tkinter as tk
    import tkFileDialog, tkMessageBox, tkSimpleDialog
except ImportError:  # pragma: no cover - Python 3 compatibility
    import queue
    import tkinter as tk
    from tkinter import filedialog as tkFileDialog, messagebox as tkMessageBox
    from tkinter import simpledialog as tkSimpleDialog

from acc2excel import (
    DEFAULT_WORKBOOK_NAME,
    AccImportError,
    find_acc_files,
    process_acc_folder,
)


APP_NAME = "Acc to Excel"
APP_CREDIT = "Power system Department - PECC2"
WINDOW_SIZE = "500x520"
BUTTON_BLUE = "#0067C5"
INSTRUCTION_TEXT = (
    "Select a folder containing a contingency sample workbook and ACC files\n"
    "Close Excel before starting"
)
_WINDOWED_STREAMS = []
_WINDOWED_HANDLES = []
_INVALID_FILE_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_FILE_NAMES = set(
    ["CON", "PRN", "AUX", "NUL"]
    + ["COM{0}".format(number) for number in range(1, 10)]
    + ["LPT{0}".format(number) for number in range(1, 10)]
)

try:
    text_type = unicode
except NameError:  # pragma: no cover - Python 3 branch
    text_type = str


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _hresult_code(result):
    return int(result) & 0xFFFFFFFF


def _check_hresult(result, action):
    code = _hresult_code(result)
    if code & 0x80000000:
        raise OSError("{0} failed (HRESULT 0x{1:08X}).".format(action, code))


def _guid(value):
    result = _GUID()
    ole32 = ctypes.windll.ole32
    ole32.CLSIDFromString.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(_GUID)]
    ole32.CLSIDFromString.restype = ctypes.c_long
    _check_hresult(
        ole32.CLSIDFromString(text_type(value), ctypes.byref(result)),
        "CLSIDFromString",
    )
    return result


def _com_method(pointer, index, return_type, *argument_types):
    vtable = ctypes.cast(
        pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))
    ).contents
    return ctypes.WINFUNCTYPE(
        return_type, ctypes.c_void_p, *argument_types
    )(vtable[index])


def _release_com_object(pointer):
    if pointer:
        _com_method(pointer, 2, ctypes.c_ulong)(pointer)


def _show_windows_folder_dialog(parent):
    """Open Windows' Explorer-style IFileOpenDialog in folder-picking mode."""
    clsid_file_open_dialog = _guid(
        "{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}"
    )
    iid_file_open_dialog = _guid(
        "{D57C7288-D4AD-4768-BE02-9D969532D960}"
    )
    dialog = ctypes.c_void_p()
    shell_item = ctypes.c_void_p()
    display_name = ctypes.c_wchar_p()
    ole32 = ctypes.windll.ole32
    initialized_here = False

    ole32.CoInitialize.argtypes = [ctypes.c_void_p]
    ole32.CoInitialize.restype = ctypes.c_long
    initialize_result = ole32.CoInitialize(None)
    if _hresult_code(initialize_result) in (0, 1):
        initialized_here = True
    elif _hresult_code(initialize_result) != 0x80010106:
        _check_hresult(initialize_result, "CoInitialize")

    try:
        ole32.CoCreateInstance.argtypes = [
            ctypes.POINTER(_GUID),
            ctypes.c_void_p,
            ctypes.c_ulong,
            ctypes.POINTER(_GUID),
            ctypes.POINTER(ctypes.c_void_p),
        ]
        ole32.CoCreateInstance.restype = ctypes.c_long
        _check_hresult(
            ole32.CoCreateInstance(
                ctypes.byref(clsid_file_open_dialog),
                None,
                1,
                ctypes.byref(iid_file_open_dialog),
                ctypes.byref(dialog),
            ),
            "CoCreateInstance",
        )

        get_options = _com_method(
            dialog, 10, ctypes.c_long, ctypes.POINTER(ctypes.c_ulong)
        )
        set_options = _com_method(dialog, 9, ctypes.c_long, ctypes.c_ulong)
        set_title = _com_method(dialog, 17, ctypes.c_long, ctypes.c_wchar_p)
        set_ok_label = _com_method(dialog, 18, ctypes.c_long, ctypes.c_wchar_p)
        set_file_label = _com_method(dialog, 19, ctypes.c_long, ctypes.c_wchar_p)
        show = _com_method(dialog, 3, ctypes.c_long, ctypes.c_void_p)
        get_result = _com_method(
            dialog, 20, ctypes.c_long, ctypes.POINTER(ctypes.c_void_p)
        )

        options = ctypes.c_ulong()
        _check_hresult(get_options(dialog, ctypes.byref(options)), "GetOptions")
        options.value |= 0x00000020  # FOS_PICKFOLDERS
        options.value |= 0x00000040  # FOS_FORCEFILESYSTEM
        options.value |= 0x00000800  # FOS_PATHMUSTEXIST
        options.value |= 0x02000000  # FOS_DONTADDTORECENT
        _check_hresult(set_options(dialog, options.value), "SetOptions")
        _check_hresult(
            set_title(
                dialog,
                u"Choose the folder containing the contingency workbook and ACC files",
            ),
            "SetTitle",
        )
        _check_hresult(set_ok_label(dialog, u"Select Folder"), "SetOkButtonLabel")
        _check_hresult(set_file_label(dialog, u"Folder:"), "SetFileNameLabel")

        show_result = show(dialog, ctypes.c_void_p(parent.winfo_id()))
        if _hresult_code(show_result) == 0x800704C7:
            return None
        _check_hresult(show_result, "Show")
        _check_hresult(
            get_result(dialog, ctypes.byref(shell_item)), "GetResult"
        )

        get_display_name = _com_method(
            shell_item,
            5,
            ctypes.c_long,
            ctypes.c_ulong,
            ctypes.POINTER(ctypes.c_wchar_p),
        )
        _check_hresult(
            get_display_name(
                shell_item, 0x80058000, ctypes.byref(display_name)
            ),
            "GetDisplayName",
        )
        return display_name.value
    finally:
        if display_name:
            ole32.CoTaskMemFree(ctypes.cast(display_name, ctypes.c_void_p))
        _release_com_object(shell_item)
        _release_com_object(dialog)
        if initialized_here:
            ole32.CoUninitialize()


def choose_working_folder(parent):
    """Use the native Explorer folder picker, with a Tk fallback."""
    if sys.platform == "win32":
        try:
            return _show_windows_folder_dialog(parent)
        except Exception:
            pass
    return tkFileDialog.askdirectory(
        parent=parent,
        title="Select working folder",
        mustexist=True,
    )


def normalize_output_workbook_name(value):
    """Return a safe macro-enabled filename, appending ``.xlsm`` if omitted."""
    if value is None:
        raise ValueError("Enter an output workbook name.")
    name = value.strip()
    if not name:
        raise ValueError("Enter an output workbook name.")
    if _INVALID_FILE_NAME.search(name):
        raise ValueError(
            'The output name cannot contain any of these characters: < > : " / \\ | ? *'
        )

    stem, extension = os.path.splitext(name)
    if extension and extension.lower() != ".xlsm":
        raise ValueError("The output workbook must use the .xlsm extension.")
    if not extension:
        stem = name
    if not stem or stem.endswith((" ", ".")):
        raise ValueError("Enter a valid output workbook name.")
    if stem.split(".")[0].upper() in _RESERVED_FILE_NAMES:
        raise ValueError("That file name is reserved by Windows. Choose another name.")
    if len(stem) + len(".xlsm") > 255:
        raise ValueError("The output workbook name is too long.")
    return stem + ".xlsm"


def ask_output_workbook_name(parent, folder):
    """Prompt until the user enters a valid, unused output workbook name."""
    initial_name = DEFAULT_WORKBOOK_NAME
    while True:
        value = tkSimpleDialog.askstring(
            APP_NAME,
            "Enter the desired output workbook name:",
            initialvalue=initial_name,
            parent=parent,
        )
        if value is None:
            return None
        try:
            output_name = normalize_output_workbook_name(value)
        except ValueError as error:
            tkMessageBox.showerror(APP_NAME, str(error), parent=parent)
            initial_name = value
            continue

        output_path = os.path.join(folder, output_name)
        if os.path.exists(output_path):
            tkMessageBox.showerror(
                APP_NAME,
                "A file with that name already exists:\n\n{0}\n\nChoose another name.".format(
                    output_path
                ),
                parent=parent,
            )
            initial_name = output_name
            continue
        return output_name


def configure_windowed_output():
    """Provide Python output streams for a windowed PyInstaller process."""
    if not getattr(sys, "frozen", False):
        return

    for stream_name in ("stdout", "stderr"):
        if getattr(sys, stream_name, None) is None:
            stream = open(os.devnull, "w")
            _WINDOWED_STREAMS.append(stream)
            setattr(sys, stream_name, stream)



def ensure_hidden_psse_console():
    """Attach the console PSS/E expects, keeping it completely hidden."""
    if not getattr(sys, "frozen", False) or sys.platform != "win32":
        return
    try:
        kernel32 = ctypes.windll.kernel32
        console_window = kernel32.GetConsoleWindow()
        if not console_window:
            if not kernel32.AllocConsole():
                return
            console_window = kernel32.GetConsoleWindow()
        if console_window:
            ctypes.windll.user32.ShowWindow(console_window, 0)

        kernel32.CreateFileW.argtypes = [
            ctypes.c_wchar_p,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_void_p,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_void_p,
        ]
        kernel32.CreateFileW.restype = ctypes.c_void_p
        null_handle = kernel32.CreateFileW(
            u"NUL", 0x40000000, 3, None, 3, 0x80, None
        )
        if null_handle not in (None, ctypes.c_void_p(-1).value):
            _WINDOWED_HANDLES.append(null_handle)
            kernel32.SetStdHandle(-11, null_handle)
            kernel32.SetStdHandle(-12, null_handle)
    except Exception:
        pass


def resource_path(file_name):
    """Resolve a bundled PyInstaller resource or a local development asset."""
    base_folder = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    bundled_path = os.path.join(base_folder, file_name)
    if os.path.exists(bundled_path):
        return bundled_path
    return os.path.join(base_folder, "images", file_name)


def run_conversion(
    folder,
    progress_callback=None,
    processor=None,
    file_finder=None,
    workbook_name=DEFAULT_WORKBOOK_NAME,
):
    """Validate *folder*, report queued ACC files, and run the core workflow."""
    processor = processor or process_acc_folder
    file_finder = file_finder or find_acc_files
    progress_callback = progress_callback or (lambda _message: None)

    acc_files = file_finder(folder)
    if not acc_files:
        raise AccImportError("No .acc files were found below {0}.".format(folder))

    progress_callback("Found {0} ACC file(s).".format(len(acc_files)))
    for acc_path in acc_files:
        progress_callback("Queued: {0}".format(os.path.relpath(acc_path, folder)))
    progress_callback("Processing PSS/E reports and writing the workbook...")
    return processor(folder, workbook_name=workbook_name)


class AccToExcelApp(object):
    """Responsive Tk interface around :func:`process_acc_folder`."""

    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.running = False
        self.worker = None
        self._build_window()
        self.root.after(100, self._poll_events)

    def _build_window(self):
        self.root.title(APP_NAME)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(460, 470)
        self.root.configure(background="#F2F2F2")
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        try:
            self.root.iconbitmap(default=resource_path("icon6.ico"))
        except Exception:
            pass

        container = tk.Frame(self.root, background="#F2F2F2", padx=14, pady=14)
        container.pack(fill=tk.BOTH, expand=True)

        self.select_button = tk.Button(
            container,
            text="Select working folder",
            command=self.select_working_folder,
            background=BUTTON_BLUE,
            activebackground="#00549F",
            foreground="white",
            activeforeground="white",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )
        self.select_button.pack(anchor=tk.W)

        tk.Label(
            container,
            text=INSTRUCTION_TEXT,
            background="#F2F2F2",
            foreground="#222222",
            font=("Segoe UI", 10, "bold"),
            justify=tk.LEFT,
            pady=14,
        ).pack(anchor=tk.W)

        tk.Label(
            container,
            text=APP_CREDIT,
            background="#F2F2F2",
            foreground="#555555",
            font=("Segoe UI", 9, "italic"),
        ).pack(anchor=tk.W, pady=(0, 8))

        progress_frame = tk.Frame(container, background="white", bd=1, relief=tk.SOLID)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(progress_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.progress_text = tk.Text(
            progress_frame,
            background="white",
            foreground="#222222",
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            padx=8,
            pady=8,
            yscrollcommand=scrollbar.set,
        )
        self.progress_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.progress_text.yview)
        self._append_progress("Ready. Select a folder containing a contingency sample workbook and ACC files.")

    def select_working_folder(self):
        if self.running:
            return
        folder = choose_working_folder(self.root)
        if not folder:
            return

        workbook_name = ask_output_workbook_name(self.root, folder)
        if not workbook_name:
            return
        self.start_conversion(folder, workbook_name)

    def start_conversion(self, folder, workbook_name=DEFAULT_WORKBOOK_NAME):
        """Start one conversion; used by the folder picker and packaged QA."""
        if self.running:
            return

        self.running = True
        self.select_button.config(state=tk.DISABLED)
        self._append_progress("")
        self._append_progress("-" * 60)
        self._append_progress("New conversion")
        self._append_progress("Selected folder: {0}".format(folder))
        self._append_progress("Output workbook: {0}".format(workbook_name))
        self._append_progress("Beginning conversion. Please wait...")
        self.worker = threading.Thread(
            target=self._worker_run, args=(folder, workbook_name)
        )
        self.worker.daemon = True
        self.worker.start()

    def _worker_run(self, folder, workbook_name):
        pythoncom = None
        try:
            ensure_hidden_psse_console()
            import pythoncom
            pythoncom.CoInitialize()
            result = run_conversion(
                folder,
                self._queue_progress,
                workbook_name=workbook_name,
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", exc))
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _queue_progress(self, message):
        self.events.put(("progress", message))

    def _poll_events(self):
        try:
            while True:
                event_type, payload = self.events.get_nowait()
                if event_type == "progress":
                    self._append_progress(payload)
                elif event_type == "done":
                    self._finish_success(payload)
                elif event_type == "error":
                    self._finish_error(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _finish_success(self, result):
        self.running = False
        self.select_button.config(state=tk.NORMAL)
        for sheet in result["sheets"]:
            self._append_progress(
                "Completed: {0} ({1} rows)".format(
                    sheet["sheet_name"], sheet["row_count"]
                )
            )
        self._append_progress("Finished: {0}".format(result["workbook_path"]))
        tkMessageBox.showinfo(
            APP_NAME,
            "Conversion completed successfully.\n\nOutput:\n{0}".format(
                result["workbook_path"]
            ),
            parent=self.root,
        )

    def _finish_error(self, error):
        self.running = False
        self.select_button.config(state=tk.NORMAL)
        message = str(error)
        self._append_progress("ERROR: {0}".format(message))
        tkMessageBox.showerror(APP_NAME, message, parent=self.root)

    def _append_progress(self, message):
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.progress_text.config(state=tk.DISABLED)

    def _on_window_close(self):
        if self.running:
            tkMessageBox.showwarning(
                APP_NAME,
                "Conversion is still running. Please wait until it finishes.",
                parent=self.root,
            )
            return
        self.root.destroy()


def main():
    configure_windowed_output()
    root = tk.Tk()
    app = AccToExcelApp(root)
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        workbook_name = (
            sys.argv[2] if len(sys.argv) > 2 else DEFAULT_WORKBOOK_NAME
        )
        root.after(
            300,
            lambda: app.start_conversion(
                os.path.abspath(sys.argv[1]), workbook_name
            ),
        )
    root.mainloop()


if __name__ == "__main__":
    main()
