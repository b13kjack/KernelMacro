import ctypes
import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pynput import keyboard

from backend import (
    DEFAULT_CONTROL_KEY_NAMES,
    MacroPlayer,
    MacroRecorder,
    normalize_events,
    _resolve_key_string,
)


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


APP_BASE_DIR = os.path.dirname(
    os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
)
SETTINGS_PATH = os.path.join(APP_BASE_DIR, "settings.json")
FUNCTION_KEY_OPTIONS = tuple(f"F{index}" for index in range(1, 13))
CONTROL_KEY_FIELDS = ("record", "play", "stop_primary", "stop_secondary")
DEFAULT_CONTROL_BINDINGS = {
    "record": DEFAULT_CONTROL_KEY_NAMES[0],
    "play": DEFAULT_CONTROL_KEY_NAMES[1],
    "stop_primary": DEFAULT_CONTROL_KEY_NAMES[2],
    "stop_secondary": DEFAULT_CONTROL_KEY_NAMES[3],
}
APP_BG = "#0c111d"
CARD_BG = "#141c2b"
SURFACE = "#1c2536"
SURFACE_ALT = "#263245"
BORDER_COLOR = "#2a3650"
TEXT_PRIMARY = "#e2e8f0"
TEXT_MUTED = "#64748b"
TEXT_DIM = "#475569"
ACCENT = "#3b82f6"
ACCENT_HOVER = "#2563eb"
DANGER = "#ef4444"
DANGER_HOVER = "#dc2626"
SUCCESS = "#22c55e"
SUCCESS_HOVER = "#16a34a"
WARNING_BG = "#2d1b06"
WARNING_TEXT = "#f59e0b"
WARNING_BORDER = "#854d0e"
FONT_HEADING = ("Segoe UI Semibold", 13)
FONT_BODY = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 11)
FONT_MONO = ("Consolas", 10)
FONT_BTN = ("Segoe UI Semibold", 11)
FONT_BTN_SM = ("Segoe UI", 10)
FONT_STATUS = ("Segoe UI", 11)


class MacroApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.settings, settings_need_save = self._load_settings()

        self.title("Python Macro Recorder")
        self.geometry("1060x680")
        self.minsize(900, 580)
        self.configure(fg_color=APP_BG)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.recorder = MacroRecorder(self._control_key_names())
        self.player = MacroPlayer()
        self.recorded_events = []
        self._global_kb_listener = None
        self.hotkeys_available = False
        self.selected_event_index = None
        self.settings_window = None
        self.settings_vars = {}
        self.settings_key_menus = []
        self.settings_save_button = None
        self._is_admin = self._check_admin()

        self._create_editor_variables()
        self._create_menu()
        self._create_widgets()
        self._refresh_hotkey_ui()
        self.hotkeys_available = self._start_global_keyboard_listener()

        if settings_need_save:
            self._persist_settings(show_error=False)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_editor_variables(self):
        self.action_delay_var = tk.StringVar(value="0")
        self.event_time_var = tk.StringVar(value="0")
        self.event_type_var = tk.StringVar(value="keyboard")
        self.event_action_var = tk.StringVar(value="press")
        self.event_key_var = tk.StringVar(value="a")
        self.event_x_var = tk.StringVar(value="0")
        self.event_y_var = tk.StringVar(value="0")
        self.event_button_var = tk.StringVar(value="left")
        self.event_click_state_var = tk.StringVar(value="press")
        self.event_dx_var = tk.StringVar(value="0")
        self.event_dy_var = tk.StringVar(value="0")

    def _create_menu(self):
        self.menu_bar = tk.Menu(self)
        self.settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.settings_menu.add_command(label="Control Keys", command=self.open_settings_window)
        self.menu_bar.add_cascade(label="Settings", menu=self.settings_menu)
        self.configure(menu=self.menu_bar)

    def _create_widgets(self):
        self.main_frame = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        current_row = 0

        # --- Admin warning bar (conditional) ---
        if not self._is_admin:
            admin_bar = ctk.CTkFrame(
                self.main_frame, fg_color=WARNING_BG, corner_radius=0, height=30,
            )
            admin_bar.grid(row=current_row, column=0, sticky="ew")
            admin_bar.grid_propagate(False)
            ctk.CTkLabel(
                admin_bar, text="  \u26a0  Run as Administrator for game capture",
                font=FONT_SMALL, text_color=WARNING_TEXT,
            ).pack(side="left", padx=8)
            current_row += 1

        # --- Toolbar ---
        toolbar = ctk.CTkFrame(
            self.main_frame, fg_color=SURFACE, corner_radius=0, height=48,
        )
        toolbar.grid(row=current_row, column=0, sticky="ew")
        toolbar.grid_propagate(False)
        current_row += 1

        self.btn_record = ctk.CTkButton(
            toolbar, text="Record", command=self.start_recording,
            fg_color=DANGER, hover_color=DANGER_HOVER,
            font=FONT_BTN, height=32, corner_radius=6,
        )
        self.btn_record.pack(side="left", padx=(10, 4), pady=8)

        self.btn_stop = ctk.CTkButton(
            toolbar, text="Stop", command=self.stop_action, state="disabled",
            fg_color=SURFACE_ALT, hover_color=BORDER_COLOR,
            font=FONT_BTN, height=32, corner_radius=6,
        )
        self.btn_stop.pack(side="left", padx=4, pady=8)

        self.btn_play = ctk.CTkButton(
            toolbar, text="Play", command=self.play_macro, state="disabled",
            fg_color=SUCCESS, hover_color=SUCCESS_HOVER,
            font=FONT_BTN, height=32, corner_radius=6,
        )
        self.btn_play.pack(side="left", padx=(4, 12), pady=8)

        sep = ctk.CTkFrame(toolbar, fg_color=BORDER_COLOR, width=2, height=26)
        sep.pack(side="left", padx=8, pady=11)

        ctk.CTkLabel(
            toolbar, text="Loops", font=FONT_SMALL, text_color=TEXT_MUTED,
        ).pack(side="left", padx=(8, 4))
        self.loop_entry = ctk.CTkEntry(
            toolbar, width=50, height=28, font=FONT_SMALL, corner_radius=4,
        )
        self.loop_entry.pack(side="left", padx=(0, 12))
        self.loop_entry.insert(0, "1")

        ctk.CTkLabel(
            toolbar, text="Delay (ms)", font=FONT_SMALL, text_color=TEXT_MUTED,
        ).pack(side="left", padx=(0, 4))
        self.delay_entry = ctk.CTkEntry(
            toolbar, width=60, height=28, font=FONT_SMALL, corner_radius=4,
            textvariable=self.action_delay_var,
        )
        self.delay_entry.pack(side="left", padx=(0, 8))

        self.btn_save = ctk.CTkButton(
            toolbar, text="Save", command=self.save_macro, state="disabled",
            fg_color="transparent", hover_color=SURFACE_ALT,
            border_width=1, border_color=BORDER_COLOR,
            font=FONT_BTN_SM, height=28, width=60, corner_radius=6,
        )
        self.btn_save.pack(side="right", padx=(4, 10), pady=10)

        self.btn_load = ctk.CTkButton(
            toolbar, text="Load", command=self.load_macro,
            fg_color="transparent", hover_color=SURFACE_ALT,
            border_width=1, border_color=BORDER_COLOR,
            font=FONT_BTN_SM, height=28, width=60, corner_radius=6,
        )
        self.btn_load.pack(side="right", padx=4, pady=10)

        # --- Status bar ---
        status_bar = ctk.CTkFrame(
            self.main_frame, fg_color=CARD_BG, corner_radius=0, height=30,
        )
        status_bar.grid(row=current_row, column=0, sticky="ew")
        status_bar.grid_propagate(False)
        current_row += 1

        self.status_label = ctk.CTkLabel(
            status_bar, text=self._ready_status_text(),
            font=FONT_STATUS, text_color=TEXT_PRIMARY,
        )
        self.status_label.pack(side="left", padx=12, pady=4)

        self.info_label = ctk.CTkLabel(
            status_bar, text="Events: 0",
            font=FONT_SMALL, text_color=TEXT_MUTED,
        )
        self.info_label.pack(side="right", padx=12, pady=4)

        # --- Content area ---
        self.main_frame.grid_rowconfigure(current_row, weight=1)

        self.actions_frame = ctk.CTkFrame(
            self.main_frame, fg_color=CARD_BG, corner_radius=0,
        )
        self.actions_frame.grid(row=current_row, column=0, sticky="nsew")
        self.actions_frame.grid_columnconfigure(0, weight=1)
        self.actions_frame.grid_columnconfigure(1, weight=1)
        self.actions_frame.grid_rowconfigure(1, weight=1)

        self.actions_label = ctk.CTkLabel(
            self.actions_frame, text="Recorded Actions",
            font=FONT_HEADING, text_color=TEXT_PRIMARY,
        )
        self.actions_label.grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")

        self.editor_label = ctk.CTkLabel(
            self.actions_frame, text="Edit Selected Action",
            font=FONT_HEADING, text_color=TEXT_PRIMARY,
        )
        self.editor_label.grid(row=0, column=1, padx=12, pady=(8, 4), sticky="w")

        self._create_event_list()
        self._create_event_editor()
        self._update_info()
        self._refresh_event_list()
        self._set_idle_controls()

    def _create_event_list(self):
        self.list_frame = ctk.CTkFrame(
            self.actions_frame, fg_color=SURFACE, corner_radius=8,
        )
        self.list_frame.grid(row=1, column=0, padx=(12, 6), pady=(0, 12), sticky="nsew")
        self.list_frame.grid_columnconfigure(0, weight=1)
        self.list_frame.grid_rowconfigure(0, weight=1)

        list_container = tk.Frame(self.list_frame, bg=SURFACE)
        list_container.grid(row=0, column=0, padx=8, pady=(8, 4), sticky="nsew")
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(0, weight=1)

        self.event_listbox = tk.Listbox(
            list_container,
            activestyle="none",
            bg="#0f1729",
            fg=TEXT_PRIMARY,
            selectbackground=ACCENT,
            selectforeground=TEXT_PRIMARY,
            relief="flat",
            highlightthickness=0,
            exportselection=False,
            font=FONT_MONO,
        )
        self.event_listbox.grid(row=0, column=0, sticky="nsew")
        self.event_listbox.bind("<<ListboxSelect>>", self.on_event_select)

        self.event_scrollbar = tk.Scrollbar(list_container, command=self.event_listbox.yview)
        self.event_scrollbar.grid(row=0, column=1, sticky="ns")
        self.event_listbox.configure(yscrollcommand=self.event_scrollbar.set)

        buttons_frame = ctk.CTkFrame(self.list_frame, fg_color="transparent")
        buttons_frame.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="ew")
        buttons_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        _btn_kw = dict(
            corner_radius=6, height=28, font=FONT_BTN_SM,
            fg_color=SURFACE_ALT, hover_color=BORDER_COLOR,
        )

        self.btn_add_action = ctk.CTkButton(
            buttons_frame, text="Add", command=self.add_action, **_btn_kw,
        )
        self.btn_add_action.grid(row=0, column=0, padx=(0, 3), sticky="ew")

        self.btn_delete_action = ctk.CTkButton(
            buttons_frame, text="Delete", command=self.delete_selected_action,
            state="disabled", **_btn_kw,
        )
        self.btn_delete_action.grid(row=0, column=1, padx=3, sticky="ew")

        self.btn_move_up = ctk.CTkButton(
            buttons_frame, text="\u2191 Up",
            command=lambda: self.move_selected_action(-1),
            state="disabled", **_btn_kw,
        )
        self.btn_move_up.grid(row=0, column=2, padx=3, sticky="ew")

        self.btn_move_down = ctk.CTkButton(
            buttons_frame, text="\u2193 Down",
            command=lambda: self.move_selected_action(1),
            state="disabled", **_btn_kw,
        )
        self.btn_move_down.grid(row=0, column=3, padx=(3, 0), sticky="ew")

    def _create_event_editor(self):
        self.editor_frame = ctk.CTkFrame(
            self.actions_frame, fg_color=SURFACE, corner_radius=8,
        )
        self.editor_frame.grid(row=1, column=1, padx=(6, 12), pady=(0, 12), sticky="nsew")
        self.editor_frame.grid_columnconfigure(1, weight=1)

        row = 0
        self.time_entry = self._add_editor_entry("Time (s)", self.event_time_var, row)
        row += 1

        self.type_menu = self._add_editor_menu(
            "Type", self.event_type_var,
            ["keyboard", "mouse", "timeout"],
            self._on_type_changed, row,
        )
        row += 1

        self.action_menu = self._add_editor_menu(
            "Action", self.event_action_var,
            ["press", "release"],
            self._on_action_changed, row,
        )
        row += 1

        self.key_entry = self._add_key_capture_field("Key", self.event_key_var, row)
        row += 1

        self.x_entry = self._add_editor_entry("X", self.event_x_var, row)
        row += 1
        self.y_entry = self._add_editor_entry("Y", self.event_y_var, row)
        row += 1

        self.button_menu = self._add_editor_menu(
            "Button", self.event_button_var,
            ["left", "right", "middle"], None, row,
        )
        row += 1

        self.click_state_menu = self._add_editor_menu(
            "Click State", self.event_click_state_var,
            ["press", "release"], None, row,
        )
        row += 1

        self.dx_entry = self._add_editor_entry("Scroll X", self.event_dx_var, row)
        row += 1
        self.dy_entry = self._add_editor_entry("Scroll Y", self.event_dy_var, row)
        row += 1

        self.editor_buttons_frame = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        self.editor_buttons_frame.grid(row=row, column=0, columnspan=2, padx=8, pady=(8, 8), sticky="ew")
        self.editor_buttons_frame.grid_columnconfigure((0, 1), weight=1)

        self.btn_revert_action = ctk.CTkButton(
            self.editor_buttons_frame, text="Revert",
            command=self.revert_selected_action, state="disabled",
            fg_color=SURFACE_ALT, hover_color=BORDER_COLOR,
            font=FONT_BTN_SM, height=30, corner_radius=6,
        )
        self.btn_revert_action.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self.btn_apply_action = ctk.CTkButton(
            self.editor_buttons_frame, text="Apply",
            command=self.apply_selected_action, state="disabled",
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=FONT_BTN_SM, height=30, corner_radius=6,
        )
        self.btn_apply_action.pack(side="left", expand=True, fill="x", padx=(4, 0))

        self._sync_editor_state()

    def _add_editor_entry(self, label_text, variable, row):
        label = ctk.CTkLabel(
            self.editor_frame, text=label_text,
            font=FONT_SMALL, text_color=TEXT_MUTED,
        )
        label.grid(row=row, column=0, padx=(10, 6), pady=4, sticky="w")

        entry = ctk.CTkEntry(
            self.editor_frame, textvariable=variable,
            height=28, font=FONT_SMALL, corner_radius=4,
        )
        entry.grid(row=row, column=1, padx=(0, 10), pady=4, sticky="ew")
        return entry

    def _add_editor_menu(self, label_text, variable, values, command, row):
        label = ctk.CTkLabel(
            self.editor_frame, text=label_text,
            font=FONT_SMALL, text_color=TEXT_MUTED,
        )
        label.grid(row=row, column=0, padx=(10, 6), pady=4, sticky="w")

        menu = ctk.CTkOptionMenu(
            self.editor_frame, values=values, variable=variable, command=command,
            height=28, font=FONT_SMALL, corner_radius=4,
            fg_color=SURFACE_ALT, button_color=BORDER_COLOR,
            button_hover_color=ACCENT,
        )
        menu.grid(row=row, column=1, padx=(0, 10), pady=4, sticky="ew")
        return menu

    def _add_key_capture_field(self, label_text, variable, row):
        label = ctk.CTkLabel(
            self.editor_frame, text=label_text,
            font=FONT_SMALL, text_color=TEXT_MUTED,
        )
        label.grid(row=row, column=0, padx=(10, 6), pady=4, sticky="w")

        container = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        container.grid(row=row, column=1, padx=(0, 10), pady=4, sticky="ew")
        container.grid_columnconfigure(0, weight=1)

        entry = ctk.CTkEntry(
            container, textvariable=variable, state="disabled",
            text_color=TEXT_PRIMARY, height=28, font=FONT_SMALL, corner_radius=4,
        )
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        btn = ctk.CTkButton(
            container, text="Capture", width=65, height=28, corner_radius=6,
            command=self._capture_key,
            fg_color=SURFACE_ALT, hover_color=BORDER_COLOR, font=FONT_BTN_SM,
        )
        btn.grid(row=0, column=1, sticky="e")

        self.btn_capture_key = btn
        return entry

    def _capture_key(self):
        self.btn_capture_key.configure(text="Listening...", state="disabled")

        def on_press(key):
            try:
                key_str = key.char
            except AttributeError:
                key_str = str(key)

            self.event_key_var.set(key_str)
            self._sync_editor_state()
            self.after(0, lambda: self.btn_capture_key.configure(text="Capture", state="normal"))
            return False

        # Start a temporary listener
        listener = keyboard.Listener(on_press=on_press)
        listener.start()

    @staticmethod
    def _check_admin():
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    def _ready_status_text(self):
        return "Ready"

    def _set_ready_status(self):
        self.status_label.configure(text=self._ready_status_text(), text_color="white")

    def _default_settings(self):
        return {"control_keys": dict(DEFAULT_CONTROL_BINDINGS)}

    def _load_settings(self):
        default_settings = self._default_settings()

        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as file:
                loaded_settings = json.load(file)
        except FileNotFoundError:
            return default_settings, True
        except (OSError, json.JSONDecodeError):
            return default_settings, True

        if not isinstance(loaded_settings, dict):
            return default_settings, True

        try:
            normalized_keys = self._normalize_control_keys(
                loaded_settings.get("control_keys", {})
            )
        except ValueError:
            return default_settings, True

        return {"control_keys": normalized_keys}, False

    def _persist_settings(self, show_error):
        payload = {"control_keys": dict(self.settings["control_keys"])}

        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as file:
                json.dump(payload, file, indent=2)
        except OSError as exc:
            if show_error:
                messagebox.showwarning(
                    "Settings not saved",
                    f"Settings were applied for this session only.\n\n{exc}",
                    parent=self.settings_window if self._settings_window_exists() else self,
                )
            return False

        return True

    def _normalize_control_keys(self, control_keys):
        if not isinstance(control_keys, dict):
            raise ValueError("Control key settings must be an object.")

        normalized = {}
        allowed = {label.lower() for label in FUNCTION_KEY_OPTIONS}

        for field in CONTROL_KEY_FIELDS:
            raw_value = control_keys.get(field, DEFAULT_CONTROL_BINDINGS[field])
            key_name = str(raw_value).strip().lower()

            if key_name not in allowed:
                raise ValueError(f"Unsupported key '{raw_value}'. Use F1-F12.")

            normalized[field] = key_name

        if len(set(normalized.values())) != len(normalized):
            raise ValueError("Control keys must all be unique.")

        return normalized

    def _control_key_names(self):
        control_keys = self.settings["control_keys"]
        return (
            control_keys["record"],
            control_keys["play"],
            control_keys["stop_primary"],
            control_keys["stop_secondary"],
        )

    def _display_control_key(self, field_name):
        return self.settings["control_keys"][field_name].upper()

    def _resolve_control_key_set(self):
        """Build a mapping from pynput Key objects to hotkey handler names."""
        control_keys = self.settings["control_keys"]
        key_map = {}
        for field in ("record", "play", "stop_primary", "stop_secondary"):
            key_obj = getattr(keyboard.Key, control_keys[field], None)
            if key_obj is not None:
                action = field if field in ("record", "play") else "stop"
                key_map[key_obj] = action
        return key_map

    def _on_global_key_press(self, key):
        """Single global keyboard hook: handles hotkeys AND recording."""
        hotkey_map = self._resolve_control_key_set()
        action = hotkey_map.get(key)
        if action == "record":
            self.after(0, self.on_record_hotkey)
            return
        if action == "play":
            self.after(0, self.on_play_hotkey)
            return
        if action == "stop":
            self.after(0, self.on_stop_hotkey)
            return

        if self.recorder.running:
            self.recorder.on_press(key)

    def _on_global_key_release(self, key):
        """Forward release events to the recorder when recording."""
        hotkey_map = self._resolve_control_key_set()
        if key in hotkey_map:
            return

        if self.recorder.running:
            self.recorder.on_release(key)

    def _start_global_keyboard_listener(self):
        self._stop_global_keyboard_listener()

        try:
            self._global_kb_listener = keyboard.Listener(
                on_press=self._on_global_key_press,
                on_release=self._on_global_key_release,
            )
            self._global_kb_listener.start()
        except Exception:
            self._global_kb_listener = None
            return False

        return True

    def _stop_global_keyboard_listener(self):
        if not self._global_kb_listener:
            return

        try:
            self._global_kb_listener.stop()
        except RuntimeError:
            pass
        finally:
            self._global_kb_listener = None

    def _refresh_hotkey_ui(self):
        self.btn_record.configure(text=f"Record ({self._display_control_key('record')})")
        self.btn_play.configure(text=f"Play ({self._display_control_key('play')})")
        self.btn_stop.configure(
            text=(
                f"Stop ({self._display_control_key('stop_primary')}/"
                f"{self._display_control_key('stop_secondary')})"
            )
        )

        if not self.recorder.running and not self.player.running:
            self._set_ready_status()

    def open_settings_window(self):
        if self._settings_window_exists():
            self.settings_window.focus()
            self.settings_window.lift()
            return

        window = ctk.CTkToplevel(self)
        window.title("Control Key Settings")
        window.geometry("420x320")
        window.resizable(False, False)
        window.transient(self)
        window.protocol("WM_DELETE_WINDOW", self._close_settings_window)

        self.settings_window = window
        self.settings_vars = {}
        self.settings_key_menus = []

        frame = ctk.CTkFrame(window)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(frame, text="Control Keys", font=("Arial", 18, "bold"))
        title.grid(row=0, column=0, columnspan=2, pady=(8, 8), sticky="w")

        helper = ctk.CTkLabel(
            frame,
            text="Choose unique keys for Record, Play, and Stop.",
            justify="left",
        )
        helper.grid(row=1, column=0, columnspan=2, pady=(0, 12), sticky="w")

        row = 2
        labels = {
            "record": "Record",
            "play": "Play",
            "stop_primary": "Stop Key 1",
            "stop_secondary": "Stop Key 2",
        }

        for field in CONTROL_KEY_FIELDS:
            label = ctk.CTkLabel(frame, text=labels[field])
            label.grid(row=row, column=0, padx=(0, 10), pady=6, sticky="w")

            variable = tk.StringVar(value=self._display_control_key(field))
            menu = ctk.CTkOptionMenu(
                frame,
                values=list(FUNCTION_KEY_OPTIONS),
                variable=variable,
            )
            menu.grid(row=row, column=1, pady=6, sticky="ew")

            self.settings_vars[field] = variable
            self.settings_key_menus.append(menu)
            row += 1

        buttons_frame = ctk.CTkFrame(frame, fg_color="transparent")
        buttons_frame.grid(row=row, column=0, columnspan=2, pady=(16, 0), sticky="ew")
        buttons_frame.grid_columnconfigure((0, 1), weight=1)

        self.settings_save_button = ctk.CTkButton(
            buttons_frame,
            text="Save Settings",
            command=self.save_control_settings,
        )
        self.settings_save_button.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        cancel_button = ctk.CTkButton(
            buttons_frame,
            text="Close",
            command=self._close_settings_window,
        )
        cancel_button.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        self._sync_settings_window_state()

    def _settings_window_exists(self):
        return self.settings_window is not None and self.settings_window.winfo_exists()

    def _close_settings_window(self):
        if self._settings_window_exists():
            self.settings_window.destroy()

        self.settings_window = None
        self.settings_vars = {}
        self.settings_key_menus = []
        self.settings_save_button = None

    def _sync_settings_window_state(self):
        if not self._settings_window_exists():
            return

        busy = self.recorder.running or self.player.running
        control_state = "disabled" if busy else "normal"

        for menu in self.settings_key_menus:
            menu.configure(state=control_state)

        if self.settings_save_button:
            self.settings_save_button.configure(state=control_state)

    def save_control_settings(self):
        if self.recorder.running or self.player.running:
            messagebox.showwarning(
                "Busy",
                "Stop recording or playback before changing control keys.",
                parent=self.settings_window if self._settings_window_exists() else self,
            )
            return

        raw_control_keys = {
            field: self.settings_vars[field].get().strip().lower()
            for field in CONTROL_KEY_FIELDS
        }

        try:
            normalized_keys = self._normalize_control_keys(raw_control_keys)
        except ValueError as exc:
            messagebox.showerror(
                "Invalid settings",
                str(exc),
                parent=self.settings_window if self._settings_window_exists() else self,
            )
            return

        self.settings["control_keys"] = normalized_keys
        self.recorder.set_control_keys(self._control_key_names())
        self.hotkeys_available = self._start_global_keyboard_listener()
        self._refresh_hotkey_ui()
        self._set_idle_controls()
        self._persist_settings(show_error=True)
        self.status_label.configure(text="Settings saved", text_color="white")

        if not self.hotkeys_available:
            messagebox.showwarning(
                "Hotkeys unavailable",
                "Control keys were updated, but the global hotkey listener could not be started.\n"
                "The on-screen buttons still work.",
                parent=self.settings_window if self._settings_window_exists() else self,
            )

        self._close_settings_window()

    def on_record_hotkey(self):
        if self.btn_record.cget("state") == "normal":
            self.after(0, self.start_recording)

    def on_play_hotkey(self):
        if self.btn_play.cget("state") == "normal":
            self.after(0, self.play_macro)

    def on_stop_hotkey(self):
        if self.btn_stop.cget("state") == "normal":
            self.after(0, self.stop_action)

    def start_recording(self):
        if self.recorder.running or self.player.running:
            return

        try:
            self.recorder.start()
        except Exception as exc:
            self.status_label.configure(text="Recording failed", text_color="red")
            messagebox.showerror("Recording failed", str(exc), parent=self)
            self._set_idle_controls()
            return

        self.recorded_events = []
        self.selected_event_index = None
        self._refresh_event_list()
        self.status_label.configure(text="Recording...", text_color="red")
        self._set_idle_controls()

    def stop_action(self):
        was_recording = self.recorder.running
        was_playing = self.player.running

        if was_recording:
            self.recorder.stop()
            try:
                self.recorded_events = normalize_events(self.recorder.events)
            except ValueError as exc:
                self.recorded_events = []
                messagebox.showerror("Invalid recording", str(exc), parent=self)
            self.status_label.configure(text="Recorded", text_color="white")

        if was_playing:
            self.player.stop()
            self._set_ready_status()

        select_index = 0 if self.recorded_events else None
        self._refresh_event_list(select_index=select_index)
        self._set_idle_controls()

    def play_macro(self):
        if not self.recorded_events or self.player.running or self.recorder.running:
            return

        if not self._commit_editor_changes():
            return

        try:
            loops = int(self.loop_entry.get().strip())
            if loops < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid loops",
                "Loop count must be a positive integer.",
                parent=self,
            )
            self.loop_entry.delete(0, tk.END)
            self.loop_entry.insert(0, "1")
            return

        try:
            action_delay_ms = float(self.action_delay_var.get().strip() or "0")
            if action_delay_ms < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid delay",
                "Randomize delay must be 0 or greater.",
                parent=self,
            )
            self.action_delay_var.set("0")
            return

        try:
            self.recorded_events = normalize_events(self.recorded_events)
        except ValueError as exc:
            messagebox.showerror("Invalid macro", str(exc), parent=self)
            return

        self._refresh_event_list(select_index=self.selected_event_index)
        self.status_label.configure(
            text=f"Playing ({loops} loop{'s' if loops != 1 else ''})...",
            text_color="green",
        )

        try:
            started = self.player.play(
                self.recorded_events,
                loops=loops,
                action_delay=action_delay_ms / 1000.0,
                on_finish=self.on_playback_finish,
            )
        except ValueError as exc:
            messagebox.showerror("Playback failed", str(exc), parent=self)
            self._set_idle_controls()
            return

        if started:
            self._set_idle_controls()
        else:
            self._reset_ui_after_playback()

    def on_playback_finish(self):
        try:
            if self.winfo_exists():
                self.after(0, self._reset_ui_after_playback)
        except tk.TclError:
            return

    def _reset_ui_after_playback(self):
        if self.recorder.running:
            return

        self._set_ready_status()
        self._set_idle_controls()

    def _set_idle_controls(self):
        busy = self.recorder.running or self.player.running

        self.btn_record.configure(state="disabled" if busy else "normal")
        self.btn_stop.configure(state="normal" if busy else "disabled")
        self.btn_load.configure(state="disabled" if busy else "normal")
        self._set_widget_enabled(self.loop_entry, not busy)
        self._set_widget_enabled(self.delay_entry, not busy)

        has_events = bool(self.recorded_events)
        ready_state = "normal" if has_events and not busy else "disabled"
        self.btn_play.configure(state=ready_state)
        self.btn_save.configure(state=ready_state)

        self.settings_menu.entryconfigure(0, state="disabled" if busy else "normal")

        self._update_action_controls()
        self._sync_editor_state()
        self._sync_settings_window_state()

    def _update_info(self):
        self.info_label.configure(text=f"Events: {len(self.recorded_events)}")

    def _refresh_event_list(self, select_index=None):
        # Listbox must be NORMAL to accept insert/delete operations.
        self.event_listbox.configure(state=tk.NORMAL)
        self.event_listbox.delete(0, tk.END)

        for index, event in enumerate(self.recorded_events):
            self.event_listbox.insert(tk.END, self._format_event_label(index, event))

        self._update_info()

        if not self.recorded_events:
            self.selected_event_index = None
            self._clear_editor()
            self._update_action_controls()
            return

        if select_index is None and self.selected_event_index is not None:
            select_index = self.selected_event_index

        if select_index is None:
            self.event_listbox.selection_clear(0, tk.END)
            self.selected_event_index = None
            self._clear_editor()
            self._update_action_controls()
            return

        bounded_index = max(0, min(select_index, len(self.recorded_events) - 1))
        self.event_listbox.selection_clear(0, tk.END)
        self.event_listbox.selection_set(bounded_index)
        self.event_listbox.see(bounded_index)
        self._load_event_into_editor(bounded_index)
        self._update_action_controls()

    def _format_event_label(self, index, event):
        base = f"{index + 1:>3}. {event['time']:>7.3f}s | "

        if event["type"] == "timeout":
            return f"{base}Timeout pause"

        if event["type"] == "keyboard":
            return f"{base}Key {event['action']} | {event['key']}"

        if event["action"] == "move":
            return f"{base}Mouse move | ({event['x']}, {event['y']})"

        if event["action"] == "click":
            return (
                f"{base}Mouse click {event['action_type']} | "
                f"{event['button']} @ ({event['x']}, {event['y']})"
            )

        return (
            f"{base}Mouse scroll | "
            f"({event['x']}, {event['y']}) dx={event['dx']} dy={event['dy']}"
        )

    def on_event_select(self, _event):
        if self.recorder.running or self.player.running:
            return

        selection = self.event_listbox.curselection()
        if not selection:
            self.selected_event_index = None
            self._clear_editor()
            self._update_action_controls()
            return

        self._load_event_into_editor(selection[0])
        self._update_action_controls()

    def _load_event_into_editor(self, index):
        event = self.recorded_events[index]
        self.selected_event_index = index

        self.event_time_var.set(self._format_number(event["time"]))
        self.event_type_var.set(event["type"])
        self._refresh_action_menu()
        self.event_action_var.set(event["action"])
        self.event_key_var.set(event.get("key", "a"))
        self.event_x_var.set(str(event.get("x", 0)))
        self.event_y_var.set(str(event.get("y", 0)))
        self.event_button_var.set(self._button_label(event.get("button", "left")))
        self.event_click_state_var.set(event.get("action_type", "press"))
        self.event_dx_var.set(str(event.get("dx", 0)))
        self.event_dy_var.set(str(event.get("dy", 0)))
        self._sync_editor_state()

    def _clear_editor(self):
        self.event_time_var.set("0")
        self.event_type_var.set("keyboard")
        self._refresh_action_menu()
        self.event_action_var.set("press")
        self.event_key_var.set("a")
        self.event_x_var.set("0")
        self.event_y_var.set("0")
        self.event_button_var.set("left")
        self.event_click_state_var.set("press")
        self.event_dx_var.set("0")
        self.event_dy_var.set("0")
        self._sync_editor_state()

    def _refresh_action_menu(self):
        event_type = self.event_type_var.get()
        if event_type == "keyboard":
            values = ["press", "release"]
        elif event_type == "mouse":
            values = ["move", "click", "scroll"]
        else:
            values = ["pause"]

        self.action_menu.configure(values=values)
        if self.event_action_var.get() not in values:
            self.event_action_var.set(values[0])

    def _on_type_changed(self, _value):
        self._refresh_action_menu()
        self._sync_editor_state()

    def _on_action_changed(self, _value):
        self._sync_editor_state()

    def _sync_editor_state(self):
        has_selection = self.selected_event_index is not None and bool(self.recorded_events)
        busy = self.recorder.running or self.player.running

        editor_enabled = has_selection and not busy

        self._set_widget_enabled(self.time_entry, editor_enabled)
        self._set_widget_enabled(self.type_menu, editor_enabled)
        
        event_type = self.event_type_var.get()
        action = self.event_action_var.get()

        self._set_widget_enabled(self.action_menu, editor_enabled and event_type != "timeout")
        self._set_widget_enabled(self.key_entry, editor_enabled and event_type == "keyboard")
        if self.btn_capture_key:
             self.btn_capture_key.configure(state="normal" if editor_enabled and event_type == "keyboard" else "disabled")

        mouse_enabled = editor_enabled and event_type == "mouse"
        self._set_widget_enabled(self.x_entry, mouse_enabled)
        self._set_widget_enabled(self.y_entry, mouse_enabled)
        self._set_widget_enabled(self.button_menu, mouse_enabled and action == "click")
        self._set_widget_enabled(self.click_state_menu, mouse_enabled and action == "click")
        self._set_widget_enabled(self.dx_entry, mouse_enabled and action == "scroll")
        self._set_widget_enabled(self.dy_entry, mouse_enabled and action == "scroll")

        self.btn_apply_action.configure(state="normal" if editor_enabled else "disabled")
        self.btn_revert_action.configure(state="normal" if editor_enabled else "disabled")

    def _update_action_controls(self):
        busy = self.recorder.running or self.player.running
        has_events = bool(self.recorded_events)
        has_selection = self.selected_event_index is not None and has_events

        self.btn_add_action.configure(state="disabled" if busy else "normal")
        self.btn_delete_action.configure(state="normal" if has_selection and not busy else "disabled")
        self.btn_move_up.configure(
            state="normal" if has_selection and not busy and self.selected_event_index > 0 else "disabled"
        )
        self.btn_move_down.configure(
            state=(
                "normal"
                if has_selection and not busy and self.selected_event_index < len(self.recorded_events) - 1
                else "disabled"
            )
        )
        self.event_listbox.configure(state=tk.DISABLED if busy else tk.NORMAL)

    def apply_selected_action(self):
        if self._commit_editor_changes():
            self.status_label.configure(text="Action updated", text_color="white")

    def _commit_editor_changes(self):
        if self.selected_event_index is None:
            return True

        try:
            updated_event = self._build_event_from_editor()
        except ValueError as exc:
            messagebox.showerror("Invalid action", str(exc), parent=self)
            return False

        updated_events = [dict(event) for event in self.recorded_events]
        updated_events[self.selected_event_index] = updated_event

        signature = self._event_signature(updated_event)
        occurrence = self._event_occurrence(updated_events, self.selected_event_index, signature)

        try:
            normalized_events = normalize_events(updated_events)
        except ValueError as exc:
            messagebox.showerror("Invalid action", str(exc), parent=self)
            return False

        self.recorded_events = normalized_events
        new_index = self._find_event_index(normalized_events, signature, occurrence)
        self._refresh_event_list(select_index=new_index)
        self._set_idle_controls()
        return True

    def _build_event_from_editor(self):
        raw_event = {
            "time": self.event_time_var.get().strip(),
            "type": self.event_type_var.get(),
            "action": self.event_action_var.get(),
        }

        if raw_event["type"] == "timeout":
            pass
        elif raw_event["type"] == "keyboard":
            raw_event["key"] = self.event_key_var.get().strip()
        else:
            raw_event["x"] = self.event_x_var.get().strip()
            raw_event["y"] = self.event_y_var.get().strip()

            if raw_event["action"] == "click":
                raw_event["button"] = self.event_button_var.get().strip()
                raw_event["action_type"] = self.event_click_state_var.get()

            if raw_event["action"] == "scroll":
                raw_event["dx"] = self.event_dx_var.get().strip()
                raw_event["dy"] = self.event_dy_var.get().strip()

        return normalize_events([raw_event])[0]

    def revert_selected_action(self):
        if self.selected_event_index is None or not self.recorded_events:
            return

        self._load_event_into_editor(self.selected_event_index)
        self.status_label.configure(text="Reverted action edits", text_color="white")

    def add_action(self):
        if self.recorder.running or self.player.running:
            return

        if not self._commit_editor_changes():
            return

        next_time = self.recorded_events[-1]["time"] + 0.1 if self.recorded_events else 0.0
        new_event = {
            "time": round(next_time, 3),
            "type": "keyboard",
            "action": "press",
            "key": "a",
        }

        updated_events = [dict(event) for event in self.recorded_events]
        updated_events.append(new_event)

        signature = self._event_signature(new_event)
        occurrence = self._event_occurrence(updated_events, len(updated_events) - 1, signature)

        try:
            normalized_events = normalize_events(updated_events)
        except ValueError as exc:
            messagebox.showerror("Invalid action", str(exc), parent=self)
            return

        self.recorded_events = normalized_events
        new_index = self._find_event_index(normalized_events, signature, occurrence)
        self._refresh_event_list(select_index=new_index)
        self._set_idle_controls()
        self.status_label.configure(text="Action added", text_color="white")

    def delete_selected_action(self):
        if self.selected_event_index is None or self.recorder.running or self.player.running:
            return

        del self.recorded_events[self.selected_event_index]
        next_index = None
        if self.recorded_events:
            next_index = min(self.selected_event_index, len(self.recorded_events) - 1)

        self._refresh_event_list(select_index=next_index)
        self._set_idle_controls()
        self.status_label.configure(text="Action deleted", text_color="white")

    def move_selected_action(self, direction):
        if self.selected_event_index is None or self.recorder.running or self.player.running:
            return

        if not self._commit_editor_changes():
            return

        current_index = self.selected_event_index
        target_index = current_index + direction
        if target_index < 0 or target_index >= len(self.recorded_events):
            return

        current_events = [dict(event) for event in self.recorded_events]
        current_events[current_index], current_events[target_index] = (
            current_events[target_index],
            current_events[current_index],
        )

        time_slots = sorted(event["time"] for event in self.recorded_events)
        for index, event in enumerate(current_events):
            event["time"] = time_slots[index]

        try:
            self.recorded_events = normalize_events(current_events)
        except ValueError as exc:
            messagebox.showerror("Invalid action", str(exc), parent=self)
            return

        self._refresh_event_list(select_index=target_index)
        self._set_idle_controls()
        self.status_label.configure(text="Action moved", text_color="white")

    def save_macro(self):
        if not self.recorded_events:
            return

        if not self._commit_editor_changes():
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as file:
                json.dump(self.recorded_events, file, indent=2)
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc), parent=self)
            return

        self.status_label.configure(text="Saved", text_color="white")

    def load_macro(self):
        if self.recorder.running or self.player.running:
            return

        file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as file:
                loaded_events = json.load(file)
            self.recorded_events = normalize_events(loaded_events)
        except (OSError, json.JSONDecodeError, ValueError) as exc:
            messagebox.showerror("Load failed", str(exc), parent=self)
            return

        select_index = 0 if self.recorded_events else None
        self._refresh_event_list(select_index=select_index)
        self._set_idle_controls()
        self.status_label.configure(text="Loaded", text_color="white")

    @staticmethod
    def _set_widget_enabled(widget, enabled):
        widget.configure(state="normal" if enabled else "disabled")

    @staticmethod
    def _button_label(raw_value):
        lowered = str(raw_value).lower()
        if "right" in lowered:
            return "right"
        if "middle" in lowered:
            return "middle"
        return "left"

    @staticmethod
    def _format_number(value):
        formatted = f"{value:.6f}".rstrip("0").rstrip(".")
        return formatted or "0"

    @staticmethod
    def _event_signature(event):
        return tuple(sorted(event.items()))

    def _event_occurrence(self, events, index, signature):
        matches = 0
        for event in events[: index + 1]:
            if self._event_signature(event) == signature:
                matches += 1
        return max(0, matches - 1)

    def _find_event_index(self, events, signature, occurrence):
        matches = 0
        for index, event in enumerate(events):
            if self._event_signature(event) != signature:
                continue

            if matches == occurrence:
                return index

            matches += 1

        return 0 if events else None

    def _on_close(self):
        if self.recorder.running:
            self.recorder.stop()
        if self.player.running:
            self.player.stop()
        self.destroy()

    def destroy(self):
        self._close_settings_window()
        self._stop_global_keyboard_listener()
        super().destroy()


if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()
