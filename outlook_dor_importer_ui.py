from __future__ import annotations

import base64
import ctypes
import queue
import subprocess
import threading
import traceback
from ctypes import wintypes
from datetime import date
from pathlib import Path

import pythoncom
import tkinter as tk
from tkinter import messagebox, ttk

from outlook_dor_importer import (
    AttachmentInfo,
    DESTINATION_WORKBOOK_NAMES,
    SessionTimer,
    extract_date_from_filename,
    get_outlook_namespace,
    import_attachment,
    list_excel_attachments,
    list_outlook_folder_paths,
    logger,
    parse_date,
    save_attachment,
)

FLASHW_TRAY = 0x00000002
FLASHW_TIMERNOFG = 0x0000000C


class FlashWindowInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwnd", wintypes.HWND),
        ("dwFlags", wintypes.DWORD),
        ("uCount", wintypes.UINT),
        ("dwTimeout", wintypes.DWORD),
    ]

class OutlookDORImporterApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title(
            "Outlook DOR Importer"
        )

        self.geometry(
            "1050x680"
        )

        self.minsize(
            900,
            580,
        )

        self.attachments: list[
            AttachmentInfo
        ] = []

        self.queue_var = tk.StringVar(
            value="Ecrash"
        )

        self.file_filter_var = tk.StringVar()

        self.outlook_folder_var = tk.StringVar(
            value="Inbox"
        )
        self.outlook_folder_paths = ["Inbox"]
        self.outlook_folder_choices = [("Inbox", "Inbox")]
        self.outlook_folder_is_typed = False
        self.folder_matches_var = tk.StringVar(
            value="Type to show matching Outlook folders."
        )
        self.visible_outlook_folder_choices = []
        self.folder_loading = False

        self.destination_var = tk.StringVar(
            value=DESTINATION_WORKBOOK_NAMES[
                "Ecrash"
            ]
        )

        self.production_date_var = (
            tk.StringVar()
        )

        self.week_override_var = (
            tk.StringVar()
        )

        self.row_override_var = (
            tk.StringVar()
        )

        self.status_var = tk.StringVar(
            value=(
                "Ready. Select Refresh Attachments."
            )
        )
        self.activity_var = tk.StringVar(
            value="Ready"
        )
        self.activity_done_var = tk.BooleanVar(
            value=False
        )

        self.progress_target = 0.0
        self.progress_animation_id = None
        self.session_timer = SessionTimer()
        self.ui_events = queue.Queue()
        self.closing = False
        self.action_active = False

        self.protocol(
            "WM_DELETE_WINDOW",
            self.close_application,
        )

        self.create_widgets()
        self.after(
            50,
            self.process_ui_events,
        )
        self.after(
            100,
            self.refresh_outlook_folders,
        )

    def post_ui_event(
        self,
        callback,
        *args,
    ):
        self.ui_events.put(
            (callback, args)
        )

    def process_ui_events(self):
        if self.closing:
            return

        while True:
            try:
                callback, args = self.ui_events.get_nowait()
            except queue.Empty:
                break

            callback(*args)

        self.after(
            50,
            self.process_ui_events,
        )

    def animate_progress(
        self,
        target_value: float,
    ):
        self.progress_target = max(
            0.0,
            min(100.0, float(target_value)),
        )

        if self.progress_animation_id is None:
            self.progress_animation_id = self.after(
                16,
                self._animate_progress_step,
            )

    def _animate_progress_step(self):
        current = float(
            self.progress["value"]
        )

        distance = self.progress_target - current

        if abs(distance) <= 0.25:
            current = self.progress_target
            self.progress["value"] = current
            self.progress_percent_var.set(
                f"{round(current)}%"
            )
            self.progress_animation_id = None
            return

        step = max(
            0.5,
            abs(distance) * 0.12,
        )

        current += step if distance > 0 else -step
        self.progress["value"] = current
        self.progress_percent_var.set(
            f"{round(current)}%"
        )

        self.progress_animation_id = self.after(
            16,
            self._animate_progress_step,
        )


    def notify_action_complete(
        self,
        title: str,
        message: str,
    ):
        """
        Shows a Windows toast and taskbar attention when the app is inactive.
        """
        try:
            window_handle = self.winfo_id()
            user32 = ctypes.windll.user32
            is_inactive = (
                self.state() == "iconic"
                or user32.GetForegroundWindow()
                != window_handle
            )

            if not is_inactive:
                return

            self.bell()

            flash_info = FlashWindowInfo(
                cbSize=ctypes.sizeof(FlashWindowInfo),
                hwnd=window_handle,
                dwFlags=FLASHW_TRAY | FLASHW_TIMERNOFG,
                uCount=3,
                dwTimeout=0,
            )

            user32.FlashWindowEx(
                ctypes.byref(flash_info)
            )

            escaped_title = title.replace(
                "'",
                "''",
            )
            escaped_message = message.replace(
                "'",
                "''",
            )

            notification_icon = Path(
                __file__
            ).with_name(
                "outlook_dor_importer.png"
            )

            if notification_icon.exists():
                toast_template = "ToastImageAndText02"
                image_uri = notification_icon.as_uri().replace(
                    "'",
                    "''",
                )
                image_script = f"""
$image = $xml.GetElementsByTagName('image').Item(0)
$image.SetAttribute('src', '{image_uri}')
"""
            else:
                toast_template = "ToastText02"
                image_script = ""

            powershell_script = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::{toast_template})
$text = $xml.GetElementsByTagName('text')
$text.Item(0).AppendChild($xml.CreateTextNode('{escaped_title}')) > $null
$text.Item(1).AppendChild($xml.CreateTextNode('{escaped_message}')) > $null
{image_script}
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('AutoDOR').Show($toast)
"""

            encoded_script = base64.b64encode(
                powershell_script.encode("utf-16le")
            ).decode("ascii")

            subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded_script,
                ],
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        except Exception:
            logger.debug(
                "Could not notify about completed action",
                exc_info=True,
            )

    
    def create_widgets(self):
        main_frame = ttk.Frame(
            self,
            padding=16,
        )

        main_frame.pack(
            fill="both",
            expand=True,
        )

        title_label = ttk.Label(
            main_frame,
            text="Outlook to DOR Excel Importer",
            font=(
                "Segoe UI",
                17,
                "bold",
            ),
        )

        title_label.pack(
            anchor="w"
        )

        subtitle_label = ttk.Label(
            main_frame,
            text=(
                "Choose an Outlook Excel attachment, "
                "filter its Work Type, and append it "
                "to the open DOR workbook."
            ),
        )

        subtitle_label.pack(
            anchor="w",
            pady=(3, 6),
        )

        activity_frame = ttk.Frame(main_frame)
        activity_frame.pack(
            anchor="w",
            pady=(0, 12),
        )

        ttk.Label(
            activity_frame,
            text="Current task:",
        ).pack(
            side="left",
        )

        ttk.Label(
            activity_frame,
            textvariable=self.activity_var,
        ).pack(
            side="left",
            padx=(6, 12),
        )

        ttk.Checkbutton(
            activity_frame,
            text="Done",
            variable=self.activity_done_var,
            state="disabled",
        ).pack(
            side="left",
        )

        settings_frame = ttk.LabelFrame(
            main_frame,
            text="Import Settings",
            padding=12,
        )

        settings_frame.pack(
            fill="x"
        )

        ttk.Label(
            settings_frame,
            text="Queue:",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )

        queue_dropdown = ttk.Combobox(
            settings_frame,
            textvariable=self.queue_var,
            values=(
                "Ecrash",
                "CGF",
            ),
            state="readonly",
            width=18,
        )

        queue_dropdown.grid(
            row=0,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        queue_dropdown.bind(
            "<<ComboboxSelected>>",
            self.queue_changed,
        )

        ttk.Label(
            settings_frame,
            text="Search/select folder:",
        ).grid(
            row=0,
            column=3,
            sticky="e",
            padx=(16, 0),
            pady=5,
        )

        self.outlook_folder_dropdown = ttk.Combobox(
            settings_frame,
            textvariable=self.outlook_folder_var,
            values=("Inbox",),
            state="normal",
            width=32,
        )

        self.outlook_folder_dropdown.grid(
            row=0,
            column=4,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        self.outlook_folder_dropdown.bind(
            "<<ComboboxSelected>>",
            self.outlook_folder_selection_changed,
        )
        self.outlook_folder_dropdown.bind(
            "<KeyRelease>",
            self.outlook_folder_text_changed,
        )
        self.outlook_folder_dropdown.bind(
            "<FocusOut>",
            self.outlook_folder_focus_lost,
        )

        self.folder_suggestion_window = tk.Toplevel(
            self,
            takefocus=0,
        )
        self.folder_suggestion_window.withdraw()
        self.folder_suggestion_window.overrideredirect(True)

        suggestion_frame = tk.Frame(
            self.folder_suggestion_window,
            borderwidth=1,
            relief="solid",
        )
        suggestion_frame.pack(
            fill="both",
            expand=True,
        )

        self.folder_suggestion_listbox = tk.Listbox(
            suggestion_frame,
            height=7,
            activestyle="dotbox",
            exportselection=False,
            takefocus=0,
        )
        self.folder_suggestion_listbox.pack(
            side="left",
            fill="both",
            expand=True,
        )

        suggestion_scrollbar = ttk.Scrollbar(
            suggestion_frame,
            orient="vertical",
            command=self.folder_suggestion_listbox.yview,
        )
        suggestion_scrollbar.pack(
            side="right",
            fill="y",
        )
        self.folder_suggestion_listbox.configure(
            yscrollcommand=suggestion_scrollbar.set,
        )
        self.folder_suggestion_listbox.bind(
            "<Button-1>",
            self.choose_outlook_folder_suggestion,
        )

        ttk.Label(
            settings_frame,
            textvariable=self.folder_matches_var,
            wraplength=760,
            justify="left",
        ).grid(
            row=5,
            column=0,
            columnspan=6,
            sticky="w",
            padx=(0, 8),
            pady=(0, 5),
        )

        self.refresh_folders_button = ttk.Button(
            settings_frame,
            text="Refresh Folders",
            command=self.refresh_outlook_folders,
        )

        self.refresh_folders_button.grid(
            row=0,
            column=5,
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Additional filename filter:",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )

        file_filter_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.file_filter_var
            ),
        )

        file_filter_entry.grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(8, 0),
            pady=5,
        )

        refresh_button = ttk.Button(
            settings_frame,
            text="Refresh Attachments",
            command=self.refresh_attachments,
        )

        refresh_button.grid(
            row=1,
            column=3,
            padx=(10, 0),
            pady=5,
        )

        latest_button = ttk.Button(
            settings_frame,
            text="Select Latest",
            command=self.select_latest,
        )

        latest_button.grid(
            row=1,
            column=4,
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Open destination workbook:",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=5,
        )

        destination_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.destination_var
            ),
        )

        destination_entry.grid(
            row=2,
            column=1,
            columnspan=4,
            sticky="ew",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Production Date:",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=5,
        )

        production_date_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.production_date_var
            ),
            width=18,
        )

        production_date_entry.grid(
            row=3,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text=(
                "YYYY-MM-DD or MM/DD/YYYY"
            ),
        ).grid(
            row=3,
            column=2,
            sticky="w",
            padx=8,
        )

        ttk.Label(
            settings_frame,
            text="Week override:",
        ).grid(
            row=3,
            column=3,
            sticky="e",
            padx=(8, 0),
        )

        week_override_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.week_override_var
            ),
            width=16,
        )

        week_override_entry.grid(
            row=3,
            column=4,
            sticky="w",
            padx=(8, 0),
        )

        ttk.Label(
            settings_frame,
            text="Insert at row (optional):",
        ).grid(
            row=4,
            column=0,
            sticky="w",
            pady=5,
        )

        row_override_entry = ttk.Entry(
            settings_frame,
            textvariable=self.row_override_var,
            width=18,
        )

        row_override_entry.grid(
            row=4,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Rows at this position will move down.",
        ).grid(
            row=4,
            column=2,
            columnspan=3,
            sticky="w",
            padx=8,
        )

        settings_frame.columnconfigure(
            2,
            weight=1,
        )

        attachment_frame = ttk.LabelFrame(
            main_frame,
            text="Available Excel Attachments",
            padding=10,
        )

        attachment_frame.pack(
            fill="both",
            expand=True,
            pady=(14, 10),
        )

        self.attachment_listbox = tk.Listbox(
            attachment_frame,
            font=(
                "Consolas",
                9,
            ),
            activestyle="dotbox",
        )

        attachment_scrollbar = ttk.Scrollbar(
            attachment_frame,
            orient="vertical",
            command=(
                self
                .attachment_listbox
                .yview
            ),
        )

        self.attachment_listbox.configure(
            yscrollcommand=(
                attachment_scrollbar.set
            )
        )

        self.attachment_listbox.pack(
            side="left",
            fill="both",
            expand=True,
        )

        attachment_scrollbar.pack(
            side="right",
            fill="y",
        )

        self.attachment_listbox.bind(
            "<<ListboxSelect>>",
            self.attachment_selection_changed,
        )

        self.import_button = ttk.Button(
            main_frame,
            text="Import Selected Attachment",
            command=self.run_import,
        )
        
        progress_frame = ttk.Frame(
            main_frame,
        )

        progress_frame.pack(
            fill="x",
            pady=(0, 10),
        )

        self.progress = ttk.Progressbar(
            progress_frame,
            mode="determinate",
        )

        self.progress.pack(
            side="left",
            fill="x",
            expand=True,
        )

        self.progress_percent_var = tk.StringVar(
            value="0%"
        )

        progress_percent_label = ttk.Label(
            progress_frame,
            textvariable=self.progress_percent_var,
            width=5,
            anchor="e",
        )

        progress_percent_label.pack(
            side="right",
            padx=(10, 0),
        )

        self.import_button.pack(
            anchor="e",
            pady=(0, 10),
        )

        status_frame = ttk.LabelFrame(
            main_frame,
            text="Status",
            padding=10,
        )

        status_frame.pack(
            fill="x"
        )

        status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            justify="left",
            wraplength=980,
        )

        status_label.pack(
            anchor="w"
        )


    def set_busy(
        self,
        busy: bool,
    ):
        self.action_active = busy

        if busy:
            self.progress.configure(
                mode="determinate"
            )
            self.progress["value"] = 0
            self.progress_target = 0.0
            self.progress_percent_var.set("0%")

            self.configure(
                cursor="watch"
            )

            self.import_button.configure(
                state="disabled"
            )

        else:
            self.configure(
                cursor=""
            )

            self.import_button.configure(
                state="normal"
            )

        self.update_idletasks()

    def update_progress(
        self,
        value: int,
        message: str = "",
    ):
        value = max(
            0,
            min(100, value),
        )

        self.animate_progress(value)

        if message:
            self.status_var.set(message)
            self.activity_var.set(message)

        self.update_idletasks()

    def set_activity(
        self,
        message: str,
        done: bool = False,
    ):
        self.activity_var.set(message)
        self.activity_done_var.set(done)
        self.update_idletasks()

    def queue_changed(
        self,
        _event=None,
    ):
        selected_queue = (
            self.queue_var.get()
        )

        destination_name = (
            DESTINATION_WORKBOOK_NAMES[
                selected_queue
            ]
        )

        self.destination_var.set(
            destination_name
        )


    def selected_outlook_folder_path(self) -> str:
        if self.outlook_folder_is_typed:
            return self.outlook_folder_var.get().strip() or "Inbox"

        selected_index = self.outlook_folder_dropdown.current()

        if 0 <= selected_index < len(self.outlook_folder_paths):
            return self.outlook_folder_paths[selected_index]

        return self.outlook_folder_var.get().strip() or "Inbox"


    def outlook_folder_text_changed(self, _event=None):
        self.outlook_folder_is_typed = True
        query = self.outlook_folder_var.get()
        matches = filter_outlook_folder_choices(
            self.outlook_folder_choices,
            query,
        )
        self.show_outlook_folder_choices(matches)
        self.show_outlook_folder_suggestions(matches)

        if not query.strip():
            self.folder_matches_var.set(
                f"All {len(matches)} Outlook folders available."
            )
        elif matches:
            match_details = [
                f"{name}: {path}"
                for name, path in matches[:5]
            ]
            if len(matches) > 5:
                match_details.append(
                    f"and {len(matches) - 5} more"
                )
            self.folder_matches_var.set(
                f"{len(matches)} match(es): "
                + " | ".join(match_details)
            )
        else:
            self.folder_matches_var.set(
                "No matching Outlook folders."
            )


    def show_outlook_folder_suggestions(
        self,
        folders: list[tuple[str, str]],
    ):
        self.visible_outlook_folder_choices = folders
        self.folder_suggestion_listbox.delete(
            0,
            tk.END,
        )

        for name, path in folders:
            self.folder_suggestion_listbox.insert(
                tk.END,
                f"{name}    {path}",
            )

        if not folders:
            self.folder_suggestion_window.withdraw()
            return

        self.update_idletasks()
        x_position = self.outlook_folder_dropdown.winfo_rootx()
        y_position = (
            self.outlook_folder_dropdown.winfo_rooty()
            + self.outlook_folder_dropdown.winfo_height()
        )
        popup_width = max(
            self.outlook_folder_dropdown.winfo_width(),
            520,
        )
        popup_height = min(len(folders), 7) * 22 + 4
        self.folder_suggestion_window.geometry(
            f"{popup_width}x{popup_height}+{x_position}+{y_position}"
        )
        self.folder_suggestion_window.deiconify()
        self.folder_suggestion_window.lift()


    def choose_outlook_folder_suggestion(self, event):
        selected_index = self.folder_suggestion_listbox.nearest(
            event.y
        )

        if not 0 <= selected_index < len(
            self.visible_outlook_folder_choices
        ):
            return "break"

        folder_name, folder_path = (
            self.visible_outlook_folder_choices[selected_index]
        )
        self.outlook_folder_var.set(folder_name)
        self.outlook_folder_is_typed = False
        self.show_outlook_folder_choices(
            self.outlook_folder_choices
        )
        full_index = next(
            index
            for index, (_name, path) in enumerate(
                self.outlook_folder_choices
            )
            if path == folder_path
        )
        self.outlook_folder_dropdown.current(full_index)
        self.folder_suggestion_window.withdraw()
        self.folder_matches_var.set(
            f"Selected folder: {folder_path}"
        )
        self.status_var.set(
            f"Selected Outlook folder: {folder_path}"
        )
        self.outlook_folder_dropdown.focus_set()
        return "break"


    def outlook_folder_focus_lost(self, _event=None):
        self.after_idle(
            self.hide_outlook_folder_suggestions_unless_typing
        )


    def hide_outlook_folder_suggestions_unless_typing(self):
        if self.focus_get() is not self.outlook_folder_dropdown:
            self.folder_suggestion_window.withdraw()


    def show_outlook_folder_choices(
        self,
        folders: list[tuple[str, str]],
    ):
        self.outlook_folder_paths = [
            path
            for _name, path in folders
        ]
        self.outlook_folder_dropdown.configure(
            values=[name for name, _path in folders]
        )


    def outlook_folder_selection_changed(self, _event=None):
        selected_index = self.outlook_folder_dropdown.current()

        if not 0 <= selected_index < len(self.outlook_folder_paths):
            return

        selected_path = self.outlook_folder_paths[selected_index]
        self.outlook_folder_is_typed = False
        self.show_outlook_folder_choices(
            self.outlook_folder_choices
        )
        full_index = self.outlook_folder_paths.index(selected_path)
        self.outlook_folder_dropdown.current(full_index)
        self.status_var.set(
            "Selected Outlook folder: "
            f"{self.selected_outlook_folder_path()}"
        )
        self.folder_matches_var.set(
            "Selected folder: "
            f"{self.selected_outlook_folder_path()}"
        )
        self.folder_suggestion_window.withdraw()


    def refresh_outlook_folders(self):
        if self.folder_loading:
            return

        self.folder_loading = True
        self.set_activity("Loading Outlook folders...")
        self.refresh_folders_button.configure(
            state="disabled"
        )

        threading.Thread(
            target=self._outlook_folders_worker,
            daemon=True,
        ).start()

    def _outlook_folders_worker(self):
        folders = None
        error = None
        com_initialized = False

        try:
            pythoncom.CoInitialize()
            com_initialized = True
            namespace = get_outlook_namespace()
            folders = list_outlook_folder_paths(namespace)
        except Exception as caught_error:
            error = caught_error
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        self.post_ui_event(
            self._finish_outlook_folders,
            folders,
            error,
        )

    def _finish_outlook_folders(
        self,
        folders: list[tuple[str, str]] | None,
        error: Exception | None,
    ):
        self.folder_loading = False
        self.refresh_folders_button.configure(
            state="normal"
        )

        if error is not None:
            self.set_activity("Could not load Outlook folders")
            logger.error(
                "Could not load Outlook folders: %s",
                error,
            )
            self.status_var.set(
                "Could not load Outlook folders. "
                "Check Outlook and select Refresh Folders to retry."
            )
            return

        current_value = self.outlook_folder_var.get()
        current_path = self.selected_outlook_folder_path()
        folder_choices = folders or [("Inbox", "Inbox")]
        self.outlook_folder_choices = folder_choices

        try:
            selected_index = next(
                index
                for index, (_name, path) in enumerate(folder_choices)
                if path == current_path
            )
        except StopIteration:
            if self.outlook_folder_is_typed:
                self.outlook_folder_dropdown.set(current_value)
                self.show_outlook_folder_choices(
                    filter_outlook_folder_choices(
                        folder_choices,
                        current_value,
                    )
                )
            else:
                selected_index = next(
                    (
                        index
                        for index, (_name, path) in enumerate(folder_choices)
                        if path == "Inbox"
                    ),
                    0,
                )
                self.show_outlook_folder_choices(folder_choices)
                self.outlook_folder_dropdown.current(selected_index)
        else:
            self.show_outlook_folder_choices(folder_choices)
            self.outlook_folder_dropdown.current(selected_index)

        if not self.outlook_folder_is_typed:
            self.outlook_folder_dropdown.current(selected_index)

        self.status_var.set(
            f"Loaded {len(folder_choices)} Outlook folders."
        )
        self.set_activity("Outlook folders loaded", done=True)


    def refresh_attachments(self):
        folder_path = self.selected_outlook_folder_path()
        self.set_activity("Searching Outlook...")
        self.set_busy(True)
        self.update_progress(
            10,
            f"Searching Outlook {folder_path}..."
        )

        threading.Thread(
            target=self._refresh_attachments_worker,
            args=(
                self.file_filter_var.get(),
                folder_path,
            ),
            daemon=True,
        ).start()

    def _refresh_attachments_worker(
        self,
        file_name_filter: str,
        folder_path: str,
    ):
        attachments = None
        error = None
        process_started_at = self.session_timer.start()

        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True

            attachments = list_excel_attachments(
                file_name_filter,
                folder_path,
            )

        except Exception as caught_error:
            error = caught_error

        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        duration_seconds = self.session_timer.finish(
            "Refresh Attachments",
            process_started_at,
            error is None,
            error,
        )

        self.post_ui_event(
            self._finish_refresh_attachments,
            attachments,
            error,
            duration_seconds,
        )

    def _finish_refresh_attachments(
        self,
        attachments,
        error,
        duration_seconds: float,
    ):
        if error is not None:
            self.set_activity("Attachment refresh failed")
            self.notify_action_complete(
                "Attachment refresh failed",
                "Outlook attachments could not be loaded.",
            )
            self.status_var.set(
                f"Refresh Attachments failed after "
                f"{duration_seconds:.2f} seconds."
            )
            self.show_error(
                "Could not search Outlook",
                error,
            )
            self.set_busy(False)
            return

        self.notify_action_complete(
            "Attachment refresh finished",
            "Outlook attachments are ready to review.",
        )

        self.attachments = attachments or []

        self.update_progress(
            65,
            "Loading matching attachments..."
        )

        self.attachment_listbox.delete(
            0,
            tk.END,
        )

        for attachment in self.attachments:
            self.attachment_listbox.insert(
                tk.END,
                attachment.display_name,
            )

        self.update_progress(
            85,
            "Preparing attachment selection..."
        )

        if not self.attachments:
            self.update_progress(
                100,
                "No matching Excel attachments were found. "
                f"Finished in {duration_seconds:.2f} seconds."
            )
            self.set_activity("Attachment search complete", done=True)
            self.set_busy(False)
            return

        self.select_latest()

        summary = (
            f"Found {len(self.attachments)} "
            "matching attachment(s) in "
            f"{duration_seconds:.2f} seconds."
        )

        self.update_progress(
            100,
            summary,
        )
        self.set_activity("Attachments loaded", done=True)
        self.set_busy(False)



    def select_latest(self):
        if not self.attachments:
            self.refresh_attachments()
            return

        self.attachment_listbox.selection_clear(
            0,
            tk.END,
        )

        self.attachment_listbox.selection_set(
            0
        )

        self.attachment_listbox.activate(
            0
        )

        self.attachment_listbox.see(
            0
        )

        self.attachment_selection_changed()


    def get_selected_attachment(
        self,
    ) -> AttachmentInfo | None:
        selected_indexes = (
            self
            .attachment_listbox
            .curselection()
        )

        if not selected_indexes:
            return None

        selected_index = int(
            selected_indexes[0]
        )

        return self.attachments[
            selected_index
        ]


    def attachment_selection_changed(
        self,
        _event=None,
    ):
        attachment = (
            self.get_selected_attachment()
        )

        if attachment is None:
            return

        detected_date = (
            extract_date_from_filename(
                attachment.file_name
            )
        )

        if detected_date is None:
            self.production_date_var.set(
                ""
            )

            self.status_var.set(
                "No date was detected in the filename. "
                "Enter Production Date manually."
            )

            return

        self.production_date_var.set(
            detected_date.strftime(
                "%Y-%m-%d"
            )
        )

        detected_day = (
            detected_date.strftime(
                "%A"
            )
        )

        self.status_var.set(
            f"Detected "
            f"{detected_date.strftime('%m/%d/%Y')} "
            f"({detected_day}) from the filename."
        )


    def run_import(self):
        selected_attachment = (
            self.get_selected_attachment()
        )

        if selected_attachment is None:
            messagebox.showwarning(
                "No attachment",
                "Select an attachment first.",
            )

            return

        destination_name = (
            self.destination_var
            .get()
            .strip()
        )

        if not destination_name:
            messagebox.showwarning(
                "Missing destination",
                "Enter the name of the open "
                "destination workbook.",
            )

            return

        try:
            production_date = parse_date(
                self.production_date_var.get()
            )

        except ValueError as error:
            messagebox.showerror(
                "Invalid date",
                str(error),
            )

            return

        row_override_text = self.row_override_var.get().strip()
        row_override = None

        if row_override_text:
            try:
                row_override = int(row_override_text)
            except ValueError:
                messagebox.showerror(
                    "Invalid row",
                    "Insert at row must be a whole number.",
                )

                return

            if row_override < 2:
                messagebox.showerror(
                    "Invalid row",
                    "Insert at row must be 2 or greater.",
                )

                return

        self.set_activity("Starting import...")
        self.set_busy(True)
        self.update_progress(
            10,
            f'Downloading "'
            f'{selected_attachment.file_name}'
            f'" from Outlook...'
        )

        threading.Thread(
            target=self._import_worker,
            args=(
                selected_attachment,
                destination_name,
                production_date,
                self.queue_var.get(),
                self.week_override_var.get(),
                row_override,
            ),
            daemon=True,
        ).start()

    def _import_worker(
        self,
        selected_attachment: AttachmentInfo,
        destination_name: str,
        production_date: date,
        selected_queue: str,
        week_override: str,
        row_override: int | None,
    ):
        result = None
        error = None
        process_started_at = self.session_timer.start()
        stage_durations = {}

        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True

            download_started_at = self.session_timer.start()

            try:
                downloaded_path = save_attachment(
                    selected_attachment
                )
            except Exception as caught_error:
                stage_durations["Download attachment"] = (
                    self.session_timer.finish(
                        "Download attachment",
                        download_started_at,
                        False,
                        caught_error,
                    )
                )
                raise
            else:
                stage_durations["Download attachment"] = (
                    self.session_timer.finish(
                        "Download attachment",
                        download_started_at,
                        True,
                    )
                )

            self.post_ui_event(
                self.update_progress,
                45,
                "Copying attachment contents...",
            )

            import_started_at = self.session_timer.start()

            try:
                result = import_attachment(
                    source_path=downloaded_path,
                    selected_queue=selected_queue,
                    destination_workbook_name=destination_name,
                    production_date=production_date,
                    week_override=week_override,
                    row_override=row_override,
                )
            except Exception as caught_error:
                stage_durations["Copy into Excel"] = (
                    self.session_timer.finish(
                        "Copy into Excel",
                        import_started_at,
                        False,
                        caught_error,
                    )
                )
                raise
            else:
                stage_durations["Copy into Excel"] = (
                    self.session_timer.finish(
                        "Copy into Excel",
                        import_started_at,
                        True,
                    )
                )

        except Exception as caught_error:
            error = caught_error

        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        duration_seconds = self.session_timer.finish(
            "Import Selected Attachment",
            process_started_at,
            error is None,
            error,
        )

        self.post_ui_event(
            self._finish_import,
            result,
            error,
            duration_seconds,
            stage_durations,
        )

    def _finish_import(
        self,
        result,
        error,
        duration_seconds: float,
        stage_durations: dict[str, float],
    ):
        if error is not None:
            self.set_activity("Import failed")
            self.notify_action_complete(
                "Excel import failed",
                "The selected report could not be imported.",
            )
            self.status_var.set(
                f"Import failed after {duration_seconds:.2f} seconds."
            )
            self.show_error(
                "Import failed",
                error,
            )
            self.set_busy(False)
            return

        self.notify_action_complete(
            "Excel import finished",
            "The selected report has finished processing.",
        )

        self.update_progress(
            100,
            "Import completed successfully."
        )

        summary = (
            "Import completed successfully.\n\n"
            f"Workbook: {result['workbook']}\n"
            f"Sheet: {result['sheet']}\n"
            f"Rows imported: {result['rows']}\n"
            f"Columns imported: {result['columns']}\n"
            f"Import started at row: {result['first_row']}\n"
            f"Destination rows: "
            f"{result['first_row']} to "
            f"{result['last_row']}\n"
            f"DOR summary row: {result['dor_row']}\n"
            f"Week: {result['week']} "
            f"({result['week_source']})\n"
            f"Day: {result['day']}\n"
            f"Production Date: "
            f"{result['production_date']}\n"
            f"Download time: "
            f"{stage_durations['Download attachment']:.2f} seconds\n"
            f"Excel copy time: "
            f"{stage_durations['Copy into Excel']:.2f} seconds\n"
            f"Total import time: {duration_seconds:.2f} seconds"
        )

        self.status_var.set(summary)
        self.set_activity("Import completed", done=True)

        self.set_busy(False)

        messagebox.showinfo(
            "Import Complete",
            summary,
        )

    def close_application(self):
        if self.action_active:
            messagebox.showwarning(
                "Action in progress",
                "Please wait for the current action to finish "
                "before closing the importer.",
            )
            return

        self.closing = True
        summary = self.session_timer.summary()
        logger.info(summary)
        messagebox.showinfo(
            "Session Summary",
            summary,
        )
        self.destroy()


    def show_error(
        self,
        title: str,
        error: Exception,
    ):
        traceback.print_exc()

        error_message = (
            f"{type(error).__name__}: "
            f"{error}"
        )

        self.status_var.set(
            error_message
        )

        messagebox.showerror(
            title,
            error_message,
        )
