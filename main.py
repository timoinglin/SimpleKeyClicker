"""
SimpleKeyClicker - Modern Edition
A gorgeous, modern key/mouse automation tool with a premium dark UI.

Input is driven by pydirectinput (SendInput / scancodes) so clicks, key presses,
moves and drags register inside games. pyautogui is used only for free-text typing
(pydirectinput can't emit shifted characters) and for reading the screen.
"""

import threading
import time
import os
import sys
import json
import random
import atexit
import pydirectinput
import keyboard
import customtkinter as ctk
from tkinter import filedialog
import pyautogui
from PIL import ImageGrab, Image
from pynput import mouse
import ctypes

# Performance settings
try:
    pydirectinput.PAUSE = 0
    pyautogui.PAUSE = 0
    pydirectinput.FAILSAFE = False
    pyautogui.FAILSAFE = False
except Exception:
    pass


def _begin_high_res_timer():
    """Request 1ms timer resolution; paired with _end_high_res_timer on exit."""
    if os.name == 'nt':
        try:
            ctypes.windll.winmm.timeBeginPeriod(1)
        except Exception:
            pass


def _end_high_res_timer():
    if os.name == 'nt':
        try:
            ctypes.windll.winmm.timeEndPeriod(1)
        except Exception:
            pass


_begin_high_res_timer()
atexit.register(_end_high_res_timer)


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def settings_path():
    """Per-user settings file in %APPDATA%/SimpleKeyClicker (falls back to home)."""
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    folder = os.path.join(base, "SimpleKeyClicker")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        folder = os.path.abspath(".")
    return os.path.join(folder, "settings.json")


# App constants
TOOL_NAME = "SimpleKeyClicker"
VERSION = "2.2.0"
ICON_PATH = resource_path("logo.ico")
LOGO_PATH = resource_path("logo.png")

SINGLE_ACTION_KEYS = {
    'tab', 'space', 'enter', 'esc', 'backspace', 'delete', 'insert',
    'up', 'down', 'left', 'right', 'home', 'end', 'pageup', 'pagedown',
    'capslock', 'numlock', 'scrolllock', 'printscreen', 'prntscrn', 'prtsc', 'pause',
    'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12',
    'shift', 'ctrl', 'alt', 'win', 'cmd'
}

MODIFIER_KEYS = {'ctrl', 'alt', 'shift', 'win', 'cmd'}
DANGEROUS_KEYS = {'alt', 'ctrl', 'shift', 'win', 'cmd', 'f4', 'delete', 'tab'}
# Control-flow pseudo-commands handled by the interpreter (never sent as input).
CONTROL_COMMANDS = {'repeat', 'endrepeat', 'ifcolor', 'ifnotcolor'}
COLOR_MATCH_TOLERANCE = 10
WAITCOLOR_TIMEOUT = 30

# Accent presets (accent, accent_hover)
ACCENTS = {
    "Purple": ("#6c5ce7", "#8b7cf7"),
    "Blue":   ("#3b82f6", "#60a5fa"),
    "Cyan":   ("#06b6d4", "#22d3ee"),
    "Green":  ("#10b981", "#34d399"),
    "Pink":   ("#ec4899", "#f472b6"),
    "Orange": ("#f59e0b", "#fbbf24"),
    "Red":    ("#ef4444", "#f87171"),
}

# Modern color palette (accent is swapped in by apply_accent)
COLORS = {
    "bg_dark": "#0f0f1a",
    "bg_card": "#1a1a2e",
    "bg_hover": "#252542",
    "accent": "#6c5ce7",
    "accent_hover": "#8b7cf7",
    "success": "#00d26a",
    "danger": "#ff4757",
    "warning": "#ffa502",
    "text": "#ffffff",
    "text_dim": "#8892b0",
    "border": "#2d2d44",
}


def apply_accent(name):
    """Swap the accent colours used across the UI."""
    accent, hover = ACCENTS.get(name, ACCENTS["Purple"])
    COLORS["accent"] = accent
    COLORS["accent_hover"] = hover


# Configure CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ModernActionRow(ctk.CTkFrame):
    """A single action row with modern styling."""

    def __init__(self, parent, app, key="", hold="0.0", delay="0.5",
                 enabled=True, is_first=False, **kwargs):
        super().__init__(parent, fg_color=COLORS["bg_card"], corner_radius=12, **kwargs)
        self.app = app
        self.is_first = is_first
        self.is_active = False

        # Variables
        self.key_var = ctk.StringVar(value=key)
        self.hold_var = ctk.StringVar(value=hold)
        self.delay_var = ctk.StringVar(value=delay)
        self.enabled_var = ctk.BooleanVar(value=bool(enabled))

        self._create_widgets()
        self._update_enabled_style()

    def _create_widgets(self):
        # Main container with padding
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=15, pady=12)

        # Enable / disable toggle (skip this row without deleting it)
        self.enable_chk = ctk.CTkCheckBox(
            container, text="", width=24, checkbox_width=20, checkbox_height=20,
            corner_radius=6, fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
            variable=self.enabled_var, command=self._on_enabled_toggle
        )
        self.enable_chk.pack(side="left", padx=(0, 6))

        # Status indicator
        self.status_label = ctk.CTkLabel(container, text="", width=20, font=("Segoe UI", 14))
        self.status_label.pack(side="left", padx=(0, 8))

        # Key/Button input section
        key_frame = ctk.CTkFrame(container, fg_color="transparent")
        key_frame.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(key_frame, text="Action", font=("Segoe UI", 11),
                     text_color=COLORS["text_dim"]).pack(anchor="w")

        key_input_frame = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_input_frame.pack(fill="x")

        self.key_entry = ctk.CTkEntry(
            key_input_frame, textvariable=self.key_var, width=200,
            placeholder_text="click, key, moveto(x,y)...",
            font=("Segoe UI", 13), height=36, corner_radius=8,
            border_color=COLORS["border"]
        )
        self.key_entry.pack(side="left", padx=(0, 8))

        self.capture_btn = ctk.CTkButton(
            key_input_frame, text="🎯", width=40, height=36,
            font=("Segoe UI", 14), corner_radius=8,
            fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
            command=self._capture
        )
        self.capture_btn.pack(side="left")

        # Timing inputs
        timing_frame = ctk.CTkFrame(container, fg_color="transparent")
        timing_frame.pack(side="left", padx=18)

        hold_frame = ctk.CTkFrame(timing_frame, fg_color="transparent")
        hold_frame.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(hold_frame, text="Hold (s)", font=("Segoe UI", 11),
                     text_color=COLORS["text_dim"]).pack(anchor="w")
        self.hold_entry = ctk.CTkEntry(hold_frame, textvariable=self.hold_var, width=70,
                                       height=36, corner_radius=8, font=("Segoe UI", 13),
                                       border_color=COLORS["border"])
        self.hold_entry.pack()

        delay_frame = ctk.CTkFrame(timing_frame, fg_color="transparent")
        delay_frame.pack(side="left")
        ctk.CTkLabel(delay_frame, text="Delay (s)", font=("Segoe UI", 11),
                     text_color=COLORS["text_dim"]).pack(anchor="w")
        self.delay_entry = ctk.CTkEntry(delay_frame, textvariable=self.delay_var, width=70,
                                        height=36, corner_radius=8, font=("Segoe UI", 13),
                                        border_color=COLORS["border"])
        self.delay_entry.pack()

        # Action buttons (right side)
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(side="right")

        btn_style = {"width": 36, "height": 36, "corner_radius": 8, "font": ("Segoe UI", 14)}

        self.up_btn = ctk.CTkButton(btn_frame, text="▲", fg_color=COLORS["bg_hover"],
                                    hover_color=COLORS["border"], command=self._move_up, **btn_style)
        self.up_btn.pack(side="left", padx=2)

        self.down_btn = ctk.CTkButton(btn_frame, text="▼", fg_color=COLORS["bg_hover"],
                                      hover_color=COLORS["border"], command=self._move_down, **btn_style)
        self.down_btn.pack(side="left", padx=2)

        self.dup_btn = ctk.CTkButton(btn_frame, text="❏", fg_color=COLORS["accent"],
                                     hover_color=COLORS["accent_hover"], command=self._duplicate, **btn_style)
        self.dup_btn.pack(side="left", padx=2)

        self.del_btn = ctk.CTkButton(btn_frame, text="✕", fg_color=COLORS["danger"],
                                     hover_color="#ff6b7a", command=self._delete, **btn_style)
        self.del_btn.pack(side="left", padx=2)

        if self.is_first:
            self.del_btn.configure(state="disabled", fg_color=COLORS["bg_hover"])

        # Live validation
        for entry, var in ((self.key_entry, self.key_var),
                           (self.hold_entry, self.hold_var),
                           (self.delay_entry, self.delay_var)):
            entry.bind("<KeyRelease>", self._validate)
            entry.bind("<FocusOut>", self._validate)

    # --- callbacks ---
    def _capture(self):
        self.app.start_capture(self.key_var)

    def _move_up(self):
        self.app.move_row(self, -1)

    def _move_down(self):
        self.app.move_row(self, 1)

    def _duplicate(self):
        self.app.duplicate_row(self)

    def _delete(self):
        if not self.is_first:
            self.app.delete_row(self)

    def _on_enabled_toggle(self):
        self._update_enabled_style()
        self.app.mark_dirty()

    def _update_enabled_style(self):
        enabled = self.enabled_var.get()
        color = COLORS["text"] if enabled else COLORS["text_dim"]
        try:
            self.key_entry.configure(text_color=color)
        except Exception:
            pass

    def _validate(self, *_):
        """Highlight invalid fields with a red border; returns True if all valid."""
        ok = True
        # Hold time
        try:
            if float(self.hold_var.get()) < 0:
                raise ValueError
            self.hold_entry.configure(border_color=COLORS["border"])
        except Exception:
            self.hold_entry.configure(border_color=COLORS["danger"])
            ok = False
        # Delay
        try:
            KeyClickerApp._parse_delay(self.delay_var.get())
            self.delay_entry.configure(border_color=COLORS["border"])
        except Exception:
            self.delay_entry.configure(border_color=COLORS["danger"])
            ok = False
        # Action text
        if self.key_var.get().strip():
            self.key_entry.configure(border_color=COLORS["border"])
        else:
            self.key_entry.configure(border_color=COLORS["danger"])
            ok = False
        return ok

    def set_active(self, active):
        self.is_active = active
        if active:
            self.configure(fg_color=COLORS["accent"], border_width=0)
            self.status_label.configure(text="▶", text_color=COLORS["text"])
        else:
            self.configure(fg_color=COLORS["bg_card"], border_width=0)
            self.status_label.configure(text="")

    def get_data(self):
        return {
            "key": self.key_var.get(),
            "hold": self.hold_var.get(),
            "delay": self.delay_var.get(),
            "enabled": bool(self.enabled_var.get()),
        }


class KeyClickerApp(ctk.CTk):
    """Modern SimpleKeyClicker application."""

    def __init__(self):
        super().__init__()

        # State
        self.running = False
        self.paused = False
        self.safe_mode = True
        self.humanize = False
        self.rows = []
        self.thread = None
        self._start_lock = threading.Lock()
        self._toast_frame = None

        # Live run stats (plain attributes, read by the main-thread UI poller)
        self._active_row = None
        self._highlighted_row = None
        self._stat_loops = 0
        self._stat_actions = 0
        self._run_start = 0.0
        self._run_mode = "infinite"
        self._run_reps = 0

        # Variables
        self.run_mode = ctk.StringVar(value="infinite")
        self.repetitions = ctk.IntVar(value=10)

        # Custom keybinds
        self.hotkey_start = ctk.StringVar(value="ctrl+f2")
        self.hotkey_stop = ctk.StringVar(value="ctrl+f3")
        self.hotkey_pause = ctk.StringVar(value="ctrl+f4")
        self.hotkey_emergency = ctk.StringVar(value="esc")

        # Persisted preferences (loaded before UI so accent applies)
        self.accent_name = "Purple"
        self.appearance = "dark"
        self.always_on_top = False
        self._loaded_settings = self._read_settings()
        self._apply_preferences(self._loaded_settings)

        self.title(TOOL_NAME)
        self.geometry(self._loaded_settings.get("geometry", "1000x680"))
        self.minsize(920, 560)
        self.configure(fg_color=COLORS["bg_dark"])
        self._set_icon(self)

        self._create_ui()
        self._restore_session(self._loaded_settings)
        self._setup_hotkeys()
        self.attributes('-topmost', self.always_on_top)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ helpers
    def _set_icon(self, window):
        try:
            window.after(200, lambda: self._safe_iconbitmap(window))
        except Exception:
            pass

    @staticmethod
    def _safe_iconbitmap(window):
        try:
            window.iconbitmap(ICON_PATH)
        except Exception:
            pass

    def mark_dirty(self):
        # Session is auto-persisted on close; this hook exists for future use.
        pass

    # ------------------------------------------------------------------- UI build
    def _create_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=25, pady=20)
        self._main_frame = main

        self._create_header(main)
        self._create_controls(main)
        self._create_rows_section(main)
        self._create_status_bar(main)

    def _create_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", pady=(0, 18))

        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left")

        try:
            logo_img = Image.open(LOGO_PATH).resize((50, 50), Image.LANCZOS)
            self.logo = ctk.CTkImage(logo_img, size=(50, 50))
            ctk.CTkLabel(title_frame, image=self.logo, text="").pack(side="left", padx=(0, 15))
        except Exception:
            self.logo = None

        text_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        text_frame.pack(side="left")
        ctk.CTkLabel(text_frame, text=TOOL_NAME, font=("Segoe UI", 28, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text="Modern Automation Tool", font=("Segoe UI", 13),
                     text_color=COLORS["text_dim"]).pack(anchor="w")

        # Menu buttons (right side)
        menu_frame = ctk.CTkFrame(header, fg_color="transparent")
        menu_frame.pack(side="right")

        btn_style = {"width": 96, "height": 36, "corner_radius": 8, "font": ("Segoe UI", 12)}

        self.pin_btn = ctk.CTkButton(menu_frame, text="📌", width=40, height=36, corner_radius=8,
                                     font=("Segoe UI", 14),
                                     fg_color=COLORS["accent"] if self.always_on_top else COLORS["bg_card"],
                                     hover_color=COLORS["accent_hover"], command=self._toggle_topmost)
        self.pin_btn.pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="📂 Load", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"],
                      command=self.load_configuration, **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="💾 Save", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"],
                      command=self.save_configuration, **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="⚙ Settings", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"],
                      command=self.show_settings, **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="❓ Help", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"],
                      command=self.show_help, **btn_style).pack(side="left", padx=5)

    def _create_controls(self, parent):
        controls = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=16)
        controls.pack(fill="x", pady=(0, 15))

        inner = ctk.CTkFrame(controls, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)

        # Start / Stop / Pause buttons
        btn_frame = ctk.CTkFrame(inner, fg_color="transparent")
        btn_frame.pack(side="left")

        self.start_btn = ctk.CTkButton(
            btn_frame, text="▶  START", width=130, height=50,
            font=("Segoe UI", 15, "bold"), corner_radius=12,
            fg_color=COLORS["success"], hover_color="#00b85c", command=self.start_action)
        self.start_btn.pack(side="left", padx=(0, 8))

        self.pause_btn = ctk.CTkButton(
            btn_frame, text="⏸  PAUSE", width=120, height=50,
            font=("Segoe UI", 14, "bold"), corner_radius=12,
            fg_color=COLORS["warning"], hover_color="#e6940a",
            command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", padx=(0, 8))

        self.stop_btn = ctk.CTkButton(
            btn_frame, text="⏹  STOP", width=120, height=50,
            font=("Segoe UI", 15, "bold"), corner_radius=12,
            fg_color=COLORS["danger"], hover_color="#ff6b7a",
            command=self.stop_action, state="disabled")
        self.stop_btn.pack(side="left")

        # Run mode options
        mode_frame = ctk.CTkFrame(inner, fg_color="transparent")
        mode_frame.pack(side="left", padx=36)

        ctk.CTkRadioButton(mode_frame, text="Run Indefinitely", variable=self.run_mode,
                           value="infinite", font=("Segoe UI", 13),
                           command=self._update_rep_state).pack(anchor="w")

        rep_frame = ctk.CTkFrame(mode_frame, fg_color="transparent")
        rep_frame.pack(anchor="w", pady=(5, 0))
        ctk.CTkRadioButton(rep_frame, text="Run", variable=self.run_mode, value="limited",
                           font=("Segoe UI", 13), command=self._update_rep_state).pack(side="left")
        self.rep_entry = ctk.CTkEntry(rep_frame, textvariable=self.repetitions, width=60, height=30,
                                      corner_radius=6, font=("Segoe UI", 13))
        self.rep_entry.pack(side="left", padx=8)
        ctk.CTkLabel(rep_frame, text="times", font=("Segoe UI", 13)).pack(side="left")

        # Toggles (right side)
        toggle_frame = ctk.CTkFrame(inner, fg_color="transparent")
        toggle_frame.pack(side="right")

        self.humanize_switch = ctk.CTkSwitch(
            toggle_frame, text="Humanize", font=("Segoe UI", 13),
            command=self._toggle_humanize, progress_color=COLORS["accent"])
        if self.humanize:
            self.humanize_switch.select()
        self.humanize_switch.pack(anchor="e", pady=(0, 6))

        self.safe_switch = ctk.CTkSwitch(
            toggle_frame, text="Safe Mode", font=("Segoe UI", 13),
            command=self._toggle_safe_mode, progress_color=COLORS["success"])
        if self.safe_mode:
            self.safe_switch.select()
        self.safe_switch.pack(anchor="e")

        self._update_rep_state()

    def _create_rows_section(self, parent):
        section = ctk.CTkFrame(parent, fg_color="transparent")
        section.pack(fill="both", expand=True)

        header = ctk.CTkFrame(section, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(header, text="Automation Actions", font=("Segoe UI", 18, "bold")).pack(side="left")

        ctk.CTkButton(header, text="+ Add Action", width=120, height=36, font=("Segoe UI", 13),
                      corner_radius=8, fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
                      command=lambda: self.add_row()).pack(side="right")

        self.rows_frame = ctk.CTkScrollableFrame(
            section, fg_color="transparent",
            scrollbar_button_color=COLORS["border"],
            scrollbar_button_hover_color=COLORS["accent"])
        self.rows_frame.pack(fill="both", expand=True)

    def _create_status_bar(self, parent):
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="x", pady=(15, 0))

        # Progress bar (only shown for limited runs)
        self.progress = ctk.CTkProgressBar(wrapper, height=6, corner_radius=3,
                                           progress_color=COLORS["accent"])
        self.progress.set(0)

        status = ctk.CTkFrame(wrapper, fg_color=COLORS["bg_card"], corner_radius=10, height=42)
        status.pack(fill="x")
        status.pack_propagate(False)
        self._status_frame = status

        inner = ctk.CTkFrame(status, fg_color="transparent")
        inner.pack(fill="x", padx=15, pady=8)

        self.status_label = ctk.CTkLabel(inner, text="⏸ Ready", font=("Segoe UI", 13),
                                         text_color=COLORS["text_dim"])
        self.status_label.pack(side="left")

        self.stats_label = ctk.CTkLabel(inner, text="", font=("Segoe UI", 12),
                                        text_color=COLORS["text_dim"])
        self.stats_label.pack(side="left", padx=20)

        self.hotkey_hint_label = ctk.CTkLabel(inner, text="", font=("Segoe UI", 11),
                                              text_color=COLORS["text_dim"])
        self.hotkey_hint_label.pack(side="right")
        self._update_hotkey_hints()

    # ------------------------------------------------------------------- toggles
    def _update_rep_state(self):
        self.rep_entry.configure(state="normal" if self.run_mode.get() == "limited" else "disabled")

    def _toggle_safe_mode(self):
        self.safe_mode = bool(self.safe_switch.get())

    def _toggle_humanize(self):
        self.humanize = bool(self.humanize_switch.get())

    def _toggle_topmost(self):
        self.always_on_top = not self.always_on_top
        self.attributes('-topmost', self.always_on_top)
        self.pin_btn.configure(fg_color=COLORS["accent"] if self.always_on_top else COLORS["bg_card"])

    # ------------------------------------------------------------------- rows
    def add_row(self, is_first=False, key="", hold="0.0", delay="0.5", enabled=True):
        row = ModernActionRow(self.rows_frame, self, key=key, hold=hold, delay=delay,
                              enabled=enabled, is_first=is_first)
        row.pack(fill="x", pady=5)
        self.rows.append(row)
        self._update_row_buttons()

    def delete_row(self, row):
        if row in self.rows and not row.is_first:
            self.rows.remove(row)
            row.destroy()
            self._update_row_buttons()

    def duplicate_row(self, row):
        if row in self.rows:
            data = row.get_data()
            idx = self.rows.index(row)
            new_row = ModernActionRow(self.rows_frame, self, **data)
            self.rows.insert(idx + 1, new_row)
            self._repack_rows()

    def move_row(self, row, direction):
        if row in self.rows:
            idx = self.rows.index(row)
            new_idx = idx + direction
            if 0 <= new_idx < len(self.rows):
                self.rows[idx], self.rows[new_idx] = self.rows[new_idx], self.rows[idx]
                self._repack_rows()

    def _repack_rows(self):
        for r in self.rows:
            r.pack_forget()
        for r in self.rows:
            r.pack(fill="x", pady=5)
        self._update_row_buttons()

    def _update_row_buttons(self):
        for i, row in enumerate(self.rows):
            row.up_btn.configure(state="normal" if i > 0 else "disabled")
            row.down_btn.configure(state="normal" if i < len(self.rows) - 1 else "disabled")

    # ------------------------------------------------------------------- hotkeys
    def _setup_hotkeys(self):
        """Register hotkeys independently. Each callback is marshalled onto the Tk
        main thread (the `keyboard` library fires them on its own thread)."""
        bindings = [
            ("Start", self.hotkey_start.get(), self.start_action),
            ("Stop", self.hotkey_stop.get(), self.stop_action),
            ("Pause", self.hotkey_pause.get(), self.toggle_pause),
            ("Emergency Stop", self.hotkey_emergency.get(), self.emergency_stop),
        ]
        failed = []
        for label, combo, callback in bindings:
            if not combo:
                continue
            try:
                keyboard.add_hotkey(combo, lambda cb=callback: self.after(0, cb))
            except Exception:
                failed.append(f"{label} ('{combo}')")
        return failed

    def _rebind_hotkeys(self):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        failed = self._setup_hotkeys()
        self._update_hotkey_hints()
        return failed

    def _update_hotkey_hints(self):
        try:
            self.hotkey_hint_label.configure(
                text=(f"{self.hotkey_start.get().upper()}: Start  •  "
                      f"{self.hotkey_pause.get().upper()}: Pause  •  "
                      f"{self.hotkey_stop.get().upper()}: Stop  •  "
                      f"{self.hotkey_emergency.get().upper()}: Emergency"))
        except Exception:
            pass

    # ------------------------------------------------------------------- run control
    def start_action(self):
        # Guard against double-start (button + hotkey racing).
        with self._start_lock:
            if self.running or (self.thread and self.thread.is_alive()):
                return
            self.running = True

        # Validate every row on the main thread.
        invalid = []
        for i, row in enumerate(self.rows):
            if not row._validate():
                invalid.append(i + 1)
        if invalid:
            self.running = False
            self._toast(f"Fix invalid values in row(s): {', '.join(map(str, invalid))}", "danger")
            return

        mode = self.run_mode.get()
        reps = 0
        if mode == "limited":
            try:
                reps = self.repetitions.get()
                if reps <= 0:
                    raise ValueError
            except Exception:
                self.running = False
                self._toast("Invalid repetition count.", "danger")
                return

        # Snapshot all Tk-backed data on the main thread; the worker never reads Tk.
        actions = [{
            "row": row,
            "key": row.key_var.get().strip(),
            "hold": row.hold_var.get(),
            "delay": row.delay_var.get(),
            "enabled": bool(row.enabled_var.get()),
        } for row in self.rows]

        self.paused = False
        self._run_mode = mode
        self._run_reps = reps
        self._stat_loops = 0
        self._stat_actions = 0
        self._run_start = time.time()
        self._active_row = None

        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="normal", text="⏸  PAUSE")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="▶ Running...", text_color=COLORS["success"])
        if mode == "limited":
            self.progress.set(0)
            self.progress.pack(fill="x", pady=(0, 8), before=self._status_frame)

        self.thread = threading.Thread(target=self._run_loop, args=(mode, reps, actions), daemon=True)
        self.thread.start()
        self._poll_ui()

    def toggle_pause(self):
        if not self.running:
            return
        self.paused = not self.paused
        if self.paused:
            self.pause_btn.configure(text="▶  RESUME")
            self.status_label.configure(text="⏸ Paused", text_color=COLORS["warning"])
        else:
            self.pause_btn.configure(text="⏸  PAUSE")
            self.status_label.configure(text="▶ Running...", text_color=COLORS["success"])

    def stop_action(self):
        if self.running:
            self.running = False
            self.paused = False

    def emergency_stop(self):
        if self.running:
            self.running = False
            self.paused = False
            self.status_label.configure(text="⚠ Emergency Stop", text_color=COLORS["danger"])

    def _update_stopped(self):
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="⏸  PAUSE")
        self.stop_btn.configure(state="disabled")
        try:
            self.progress.pack_forget()
        except Exception:
            pass
        if self._highlighted_row is not None:
            try:
                if self._highlighted_row.winfo_exists():
                    self._highlighted_row.set_active(False)
            except Exception:
                pass
            self._highlighted_row = None

    def _on_run_finished(self, completed):
        if completed:
            self.status_label.configure(text="✓ Completed", text_color=COLORS["success"])
        elif "Emergency" not in self.status_label.cget("text"):
            self.status_label.configure(text="⏸ Stopped", text_color=COLORS["text_dim"])
        self._update_stopped()
        self._update_stats_labels(final=True)

    # ------------------------------------------------------------------- UI poller
    def _poll_ui(self):
        """Main-thread poller: moves the active-row highlight and refreshes stats
        at ~8Hz, so the worker never has to schedule per-action UI callbacks."""
        if not self.running:
            return
        row = self._active_row
        if row is not self._highlighted_row:
            if self._highlighted_row is not None:
                try:
                    if self._highlighted_row.winfo_exists():
                        self._highlighted_row.set_active(False)
                except Exception:
                    pass
            self._highlighted_row = row
            if row is not None:
                try:
                    if row.winfo_exists():
                        row.set_active(True)
                except Exception:
                    pass
        self._update_stats_labels()
        self.after(120, self._poll_ui)

    def _update_stats_labels(self, final=False):
        elapsed = max(0.0, time.time() - self._run_start) if self._run_start else 0.0
        loops = self._stat_loops
        actions = self._stat_actions
        cps = actions / elapsed if elapsed > 0.2 else 0.0
        mins, secs = divmod(int(elapsed), 60)
        parts = [f"Loops: {loops}", f"Actions: {actions}",
                 f"Time: {mins:02d}:{secs:02d}", f"CPS: {cps:.1f}"]
        if self._run_mode == "limited" and self._run_reps:
            frac = min(1.0, loops / self._run_reps)
            try:
                self.progress.set(frac)
            except Exception:
                pass
            if cps and loops < self._run_reps and not final:
                remaining = (self._run_reps - loops) * (elapsed / loops) if loops else 0
                rm, rs = divmod(int(remaining), 60)
                parts.append(f"ETA: {rm:02d}:{rs:02d}")
        self.stats_label.configure(text="   •   ".join(parts))

    # ------------------------------------------------------------------- worker
    def _run_loop(self, mode, reps, actions):
        loop = 0
        completed = False
        try:
            while self.running:
                if mode == "limited" and loop >= reps:
                    completed = True
                    break
                loop += 1
                self._stat_loops = loop
                self._exec_sequence(actions)
            if mode == "limited" and loop >= reps:
                completed = True
        finally:
            self.running = False
            self._active_row = None
            self.after(0, lambda: self._on_run_finished(completed))

    def _exec_sequence(self, actions):
        """Execute one pass through the action list, honouring control flow
        (repeat/endrepeat, ifcolor/ifnotcolor) and enable toggles."""
        i = 0
        n = len(actions)
        loop_stack = []   # [body_start_index, remaining_iterations]
        skip_next = False
        while i < n and self.running:
            self._wait_if_paused()
            if not self.running:
                break
            act = actions[i]
            if not act["enabled"]:
                i += 1
                continue

            key = act["key"]
            cmd = (key.split('(')[0] if '(' in key else key).strip().lower()

            # --- control flow -------------------------------------------------
            if cmd == "repeat":
                count = 1
                if '(' in key:
                    try:
                        a = self._args(key)
                        count = int(a[0]) if a else 1
                    except Exception:
                        count = 1
                if count < 1:
                    count = 1
                loop_stack.append([i + 1, count - 1])
                i += 1
                continue
            if cmd == "endrepeat":
                if loop_stack and loop_stack[-1][1] > 0:
                    loop_stack[-1][1] -= 1
                    i = loop_stack[-1][0]
                else:
                    if loop_stack:
                        loop_stack.pop()
                    i += 1
                continue
            if cmd in ("ifcolor", "ifnotcolor"):
                matched = False
                try:
                    r, g, b, x, y = map(int, self._args(key))
                    matched = self._check_color(r, g, b, x, y)
                except Exception:
                    matched = False
                condition = matched if cmd == "ifcolor" else (not matched)
                if not condition:
                    skip_next = True
                self._apply_delay(act)
                i += 1
                continue

            if skip_next:
                skip_next = False
                i += 1
                continue

            # --- real action --------------------------------------------------
            self._active_row = act["row"]
            try:
                hold = float(act["hold"])
            except Exception:
                hold = 0.0
            if not self._perform_action(key, hold):
                self.running = False
                break
            self._stat_actions += 1
            self._apply_delay(act)
            i += 1

    def _wait_if_paused(self):
        while self.paused and self.running:
            time.sleep(0.05)

    def _apply_delay(self, act):
        try:
            delay = self._parse_delay(act["delay"])
        except Exception:
            delay = 0.0
        self._sleep_responsive(delay)

    def _sleep_responsive(self, secs):
        """Sleep in small chunks so Stop/Pause stay responsive even on long delays."""
        if secs <= 0:
            return
        end = time.time() + secs
        while self.running:
            remaining = end - time.time()
            if remaining <= 0:
                break
            self._wait_if_paused()
            if not self.running:
                break
            time.sleep(min(0.02, remaining))

    @staticmethod
    def _args(key):
        """Return the comma-separated args inside parentheses, stripped."""
        inner = key[key.index('(') + 1:key.rindex(')')]
        return [a.strip() for a in inner.split(',')] if inner.strip() else []

    # ------------------------------------------------------------------- input
    def _move_to(self, x, y):
        """Move the cursor with pydirectinput; optional humanized curved path."""
        if not self.humanize:
            pydirectinput.moveTo(x, y)
            return
        try:
            cx, cy = pydirectinput.position()
        except Exception:
            cx, cy = x, y
        dist = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5
        steps = max(1, min(60, int(dist / 30)))
        for s in range(1, steps + 1):
            if not self.running:
                break
            t = s / steps
            ease = t * t * (3 - 2 * t)  # smoothstep
            nx = int(cx + (x - cx) * ease + random.uniform(-1.5, 1.5))
            ny = int(cy + (y - cy) * ease + random.uniform(-1.5, 1.5))
            pydirectinput.moveTo(nx, ny)
            time.sleep(random.uniform(0.001, 0.004))
        pydirectinput.moveTo(x, y)

    def _is_combo(self, key):
        """A '+'-joined string is a combo only if it includes a modifier; that way
        literal text like 'a+b' is typed instead of fired as keystrokes."""
        if '+' not in key:
            return False
        parts = [p.strip().lower() for p in key.split('+')]
        if any(not p for p in parts):
            return False
        has_modifier = any(p in MODIFIER_KEYS for p in parts)
        all_valid = all(p in SINGLE_ACTION_KEYS or len(p) == 1 for p in parts)
        return has_modifier and all_valid

    def _is_blocked(self, key):
        """Safe Mode blocks only genuinely dangerous keystrokes (modifiers, F4,
        delete, tab) — never harmless commands like click/moveto/waitcolor."""
        if self._is_combo(key):
            parts = [p.strip().lower() for p in key.split('+')]
            return any(p in DANGEROUS_KEYS for p in parts)
        name = (key.split('(')[0] if '(' in key else key).strip().lower()
        return name in DANGEROUS_KEYS

    def _perform_action(self, key, hold_time):
        if not self.running:
            return False

        if self.safe_mode and self._is_blocked(key):
            self.after(0, lambda: self._toast(f"'{key}' blocked by Safe Mode.", "danger"))
            return False

        try:
            # Commands with coordinates: click(x,y), moveto(x,y), waitcolor(...), drag(...)
            if '(' in key and ')' in key:
                cmd = key.split('(')[0].strip().lower()
                args = self._args(key)

                if cmd in {'click', 'rclick', 'mclick'}:
                    x, y = map(int, args)
                    self._move_to(x, y)
                    button = {'click': 'left', 'rclick': 'right', 'mclick': 'middle'}[cmd]
                    if hold_time > 0:
                        pydirectinput.mouseDown(button=button)
                        self._sleep_responsive(hold_time)
                        pydirectinput.mouseUp(button=button)
                    else:
                        pydirectinput.click(button=button)
                    return True

                if cmd == 'moveto':
                    x, y = map(int, args)
                    self._move_to(x, y)
                    return True

                if cmd == 'waitcolor':
                    r, g, b, x, y = map(int, args)
                    return self._wait_for_color(r, g, b, x, y)

                if cmd == 'drag':
                    x1, y1, x2, y2 = map(int, args)
                    self._move_to(x1, y1)
                    time.sleep(0.05)
                    pydirectinput.mouseDown()
                    time.sleep(0.05)
                    if self.humanize:
                        self._move_to(x2, y2)
                    else:
                        pydirectinput.moveTo(x2, y2, duration=max(hold_time, 0.2))
                    time.sleep(0.04)
                    pydirectinput.mouseUp()
                    return True

            # Keyboard combos (ctrl+c, alt+f4, shift+a ...)
            if self._is_combo(key):
                parts = [p.strip() for p in key.split('+')]
                for p in parts:
                    pydirectinput.keyDown(p)
                self._sleep_responsive(max(hold_time, 0.05))
                for p in reversed(parts):
                    pydirectinput.keyUp(p)
                return True

            # Simple mouse / key actions
            k_lower = key.lower()
            if k_lower == "click":
                self._click_button('left', hold_time)
            elif k_lower == "rclick":
                self._click_button('right', hold_time)
            elif k_lower == "mclick":
                self._click_button('middle', hold_time)
            elif k_lower in SINGLE_ACTION_KEYS or len(key) == 1:
                if hold_time > 0:
                    pydirectinput.keyDown(key)
                    self._sleep_responsive(hold_time)
                    pydirectinput.keyUp(key)
                else:
                    pydirectinput.press(key)
            else:
                # Free text. pydirectinput can't emit shifted chars, so use pyautogui.
                pyautogui.write(key, interval=0.0)

            return True
        except Exception as e:
            self.after(0, lambda msg=str(e): self._toast(f"Error: {msg}", "danger"))
            return False

    def _click_button(self, button, hold_time):
        if hold_time > 0:
            pydirectinput.mouseDown(button=button)
            self._sleep_responsive(hold_time)
            pydirectinput.mouseUp(button=button)
        else:
            pydirectinput.click(button=button)

    # ------------------------------------------------------------------- color
    @staticmethod
    def _check_color(r, g, b, x, y):
        try:
            pixel = ImageGrab.grab(bbox=(x, y, x + 1, y + 1)).getpixel((0, 0))
            return all(abs(pixel[i] - (r, g, b)[i]) <= COLOR_MATCH_TOLERANCE for i in range(3))
        except Exception:
            return False

    def _wait_for_color(self, r, g, b, x, y, timeout=WAITCOLOR_TIMEOUT):
        start = time.time()
        while time.time() - start < timeout:
            if not self.running:
                return False
            self._wait_if_paused()
            if self._check_color(r, g, b, x, y):
                return True
            time.sleep(0.05)
        self.after(0, lambda: self._toast(f"Color not found at ({x},{y})", "danger"))
        return False

    @staticmethod
    def _parse_delay(value):
        """Parse a delay: fixed ('0.5') or random range ('0.3-0.8')."""
        value = value.strip()
        if '-' in value and not value.startswith('-'):
            lo, hi = value.split('-', 1)
            lo, hi = float(lo), float(hi)
            if lo > hi:
                lo, hi = hi, lo
            return random.uniform(lo, hi)
        return float(value)

    # ------------------------------------------------------------------- capture
    def start_capture(self, key_var):
        if self.running:
            self._toast("Stop automation before capturing.", "danger")
            return

        self.attributes('-alpha', 0.3)
        data = {'x': None, 'y': None, 'color': None}

        def on_click(x, y, button, pressed):
            if button == mouse.Button.left and pressed:
                data['x'], data['y'] = int(x), int(y)
                try:
                    data['color'] = pyautogui.pixel(data['x'], data['y'])
                except Exception:
                    data['color'] = (0, 0, 0)
                return False

        def worker():
            # Runs off the main thread so the listener never freezes the UI.
            try:
                with mouse.Listener(on_click=on_click) as listener:
                    listener.join()
            except Exception:
                pass
            self.after(0, finish)

        def finish():
            self.attributes('-alpha', 1.0)
            self.lift()
            self.focus_force()
            if data['x'] is not None:
                self._show_capture_dialog(data, key_var)

        threading.Thread(target=worker, daemon=True).start()

    def _show_capture_dialog(self, data, key_var):
        dialog = self._make_dialog("Capture Options", "360x340")
        x, y = data['x'], data['y']
        r, g, b = data['color']

        info_frame = ctk.CTkFrame(dialog, fg_color=COLORS["bg_card"], corner_radius=12)
        info_frame.pack(fill="x", padx=20, pady=20)
        inner = ctk.CTkFrame(info_frame, fg_color="transparent")
        inner.pack(padx=15, pady=15)

        color_hex = f"#{r:02x}{g:02x}{b:02x}"
        ctk.CTkFrame(inner, width=40, height=40, corner_radius=8, fg_color=color_hex).pack(side="left", padx=(0, 15))
        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left")
        ctk.CTkLabel(text_frame, text=f"Position: ({x}, {y})", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text=f"Color: RGB({r}, {g}, {b})", font=("Segoe UI", 12),
                     text_color=COLORS["text_dim"]).pack(anchor="w")

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20)

        def insert(cmd):
            key_var.set(cmd)
            dialog.destroy()

        btn_style = {"height": 40, "corner_radius": 10, "font": ("Segoe UI", 13)}
        ctk.CTkButton(btn_frame, text=f"🖱 Click at ({x}, {y})", fg_color=COLORS["accent"],
                      hover_color=COLORS["accent_hover"], command=lambda: insert(f"click({x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text=f"➡ Move to ({x}, {y})", fg_color=COLORS["accent"],
                      hover_color=COLORS["accent_hover"], command=lambda: insert(f"moveto({x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text="🎨 Wait for color", fg_color=COLORS["accent"],
                      hover_color=COLORS["accent_hover"], command=lambda: insert(f"waitcolor({r},{g},{b},{x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text="Cancel", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"],
                      command=dialog.destroy, **btn_style).pack(fill="x", pady=(15, 5))
        self._center_window(dialog)

    # ------------------------------------------------------------------- settings dialog
    def show_settings(self):
        dialog = self._make_dialog("Settings", "500x560", grab=True)

        ctk.CTkLabel(dialog, text="⚙  Settings", font=("Segoe UI", 20, "bold")).pack(pady=(20, 4))
        ctk.CTkLabel(dialog, text="Keybinds, theme and behaviour", font=("Segoe UI", 12),
                     text_color=COLORS["text_dim"]).pack(pady=(0, 12))

        # --- Keybinds ---
        entries_frame = ctk.CTkFrame(dialog, fg_color=COLORS["bg_card"], corner_radius=12)
        entries_frame.pack(fill="x", padx=25, pady=5)

        temp_vars = {
            "start": ctk.StringVar(value=self.hotkey_start.get()),
            "stop": ctk.StringVar(value=self.hotkey_stop.get()),
            "pause": ctk.StringVar(value=self.hotkey_pause.get()),
            "emergency": ctk.StringVar(value=self.hotkey_emergency.get()),
        }
        labels = {"start": "Start", "stop": "Stop", "pause": "Pause / Resume", "emergency": "Emergency Stop"}

        def record_hotkey(var, btn):
            btn.configure(text="Press keys...", fg_color=COLORS["danger"])

            def capture():
                try:
                    combo = keyboard.read_hotkey(suppress=False)
                    var.set(combo)
                except Exception:
                    pass
                finally:
                    self.after(0, lambda: btn.configure(text="⏺ Record", fg_color=COLORS["accent"]))

            threading.Thread(target=capture, daemon=True).start()

        for key_name in ["start", "stop", "pause", "emergency"]:
            row = ctk.CTkFrame(entries_frame, fg_color="transparent")
            row.pack(fill="x", padx=15, pady=8)
            ctk.CTkLabel(row, text=labels[key_name], font=("Segoe UI", 13, "bold"),
                         width=130, anchor="w").pack(side="left")
            ctk.CTkEntry(row, textvariable=temp_vars[key_name], width=150, height=34,
                         corner_radius=8, font=("Segoe UI", 13)).pack(side="left", padx=(0, 10))
            rec_btn = ctk.CTkButton(row, text="⏺ Record", width=90, height=34, corner_radius=8,
                                    font=("Segoe UI", 12), fg_color=COLORS["accent"],
                                    hover_color=COLORS["accent_hover"])
            rec_btn.configure(command=lambda v=temp_vars[key_name], b=rec_btn: record_hotkey(v, b))
            rec_btn.pack(side="left")

        # --- Theme ---
        theme_frame = ctk.CTkFrame(dialog, fg_color=COLORS["bg_card"], corner_radius=12)
        theme_frame.pack(fill="x", padx=25, pady=12)
        trow = ctk.CTkFrame(theme_frame, fg_color="transparent")
        trow.pack(fill="x", padx=15, pady=12)
        ctk.CTkLabel(trow, text="Accent", font=("Segoe UI", 13, "bold"), width=130, anchor="w").pack(side="left")
        accent_menu = ctk.CTkOptionMenu(trow, values=list(ACCENTS.keys()),
                                        fg_color=COLORS["accent"], button_color=COLORS["accent_hover"],
                                        command=self._change_accent)
        accent_menu.set(self.accent_name)
        accent_menu.pack(side="left")

        # --- Buttons ---
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=25, pady=18)

        def save():
            prev = (self.hotkey_start.get(), self.hotkey_stop.get(),
                    self.hotkey_pause.get(), self.hotkey_emergency.get())
            self.hotkey_start.set(temp_vars["start"].get())
            self.hotkey_stop.set(temp_vars["stop"].get())
            self.hotkey_pause.set(temp_vars["pause"].get())
            self.hotkey_emergency.set(temp_vars["emergency"].get())
            failed = self._rebind_hotkeys()
            if failed:
                self.hotkey_start.set(prev[0])
                self.hotkey_stop.set(prev[1])
                self.hotkey_pause.set(prev[2])
                self.hotkey_emergency.set(prev[3])
                self._rebind_hotkeys()
                self._toast("Could not bind: " + ", ".join(failed) + ". Kept previous keybinds.", "danger")
                return
            dialog.destroy()
            self._toast("Settings saved!", "success")

        def reset_defaults():
            temp_vars["start"].set("ctrl+f2")
            temp_vars["stop"].set("ctrl+f3")
            temp_vars["pause"].set("ctrl+f4")
            temp_vars["emergency"].set("esc")

        btn_style = {"height": 40, "corner_radius": 10, "font": ("Segoe UI", 13)}
        ctk.CTkButton(btn_frame, text="Reset Defaults", width=130, fg_color=COLORS["bg_card"],
                      hover_color=COLORS["bg_hover"], command=reset_defaults, **btn_style).pack(side="left")
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color=COLORS["bg_card"],
                      hover_color=COLORS["bg_hover"], command=dialog.destroy, **btn_style).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Save", width=100, fg_color=COLORS["success"],
                      hover_color="#00b85c", command=save, **btn_style).pack(side="right")

        self._center_window(dialog)

    def _change_accent(self, name):
        if name == self.accent_name:
            return
        if self.running:
            self._toast("Stop automation before changing theme.", "danger")
            return
        self.accent_name = name
        apply_accent(name)
        self._rebuild_ui()
        self._toast(f"{name} theme applied", "success")

    def _rebuild_ui(self):
        """Recreate the UI after an accent change, preserving the action rows.
        Only the main content frame is rebuilt — open dialogs are left intact."""
        state = {
            "run_mode": self.run_mode.get(),
            "repetitions": self._safe_reps(),
            "rows": [r.get_data() for r in self.rows],
        }
        self.rows.clear()
        try:
            self._main_frame.destroy()
        except Exception:
            pass
        self.configure(fg_color=COLORS["bg_dark"])
        self._create_ui()
        self._restore_session(state, restore_hotkeys=False)
        self._update_hotkey_hints()

    # ------------------------------------------------------------------- persistence
    def _safe_reps(self):
        try:
            return int(self.repetitions.get())
        except Exception:
            return 10

    def save_configuration(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        config = self._serialize(include_window=False)
        try:
            with open(path, 'w', encoding="utf-8") as f:
                json.dump(config, f, indent=2)
            self._toast("Configuration saved!", "success")
        except Exception as e:
            self._toast(f"Save failed: {e}", "danger")

    def load_configuration(self):
        if self.running:
            self._toast("Stop automation before loading.", "danger")
            return
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                config = json.load(f)
            self._restore_session(config, restore_hotkeys=True)
            self._toast("Configuration loaded!", "success")
        except Exception as e:
            self._toast(f"Load failed: {e}", "danger")

    def _serialize(self, include_window=True):
        config = {
            'version': VERSION,
            'run_mode': self.run_mode.get(),
            'repetitions': self._safe_reps(),
            'rows': [r.get_data() for r in self.rows],
            'hotkey_start': self.hotkey_start.get(),
            'hotkey_stop': self.hotkey_stop.get(),
            'hotkey_pause': self.hotkey_pause.get(),
            'hotkey_emergency': self.hotkey_emergency.get(),
            'safe_mode': self.safe_mode,
            'humanize': self.humanize,
            'accent': self.accent_name,
            'always_on_top': self.always_on_top,
        }
        if include_window:
            try:
                config['geometry'] = self.geometry()
            except Exception:
                pass
        return config

    def _read_settings(self):
        try:
            with open(settings_path(), encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_settings(self):
        try:
            with open(settings_path(), 'w', encoding="utf-8") as f:
                json.dump(self._serialize(include_window=True), f, indent=2)
        except Exception:
            pass

    def _apply_preferences(self, config):
        """Apply non-widget preferences before the UI is built."""
        self.accent_name = config.get('accent', 'Purple')
        if self.accent_name not in ACCENTS:
            self.accent_name = 'Purple'
        apply_accent(self.accent_name)
        self.safe_mode = config.get('safe_mode', True)
        self.humanize = config.get('humanize', False)
        self.always_on_top = config.get('always_on_top', False)

    def _restore_session(self, config, restore_hotkeys=True):
        """Rebuild rows + state from a settings/config dict (handles legacy 'sleep')."""
        self.run_mode.set(config.get('run_mode', 'infinite'))
        try:
            self.repetitions.set(int(config.get('repetitions', 10)))
        except Exception:
            self.repetitions.set(10)
        self._update_rep_state()

        if restore_hotkeys:
            self.hotkey_start.set(config.get('hotkey_start', self.hotkey_start.get()))
            self.hotkey_stop.set(config.get('hotkey_stop', self.hotkey_stop.get()))
            self.hotkey_pause.set(config.get('hotkey_pause', self.hotkey_pause.get()))
            self.hotkey_emergency.set(config.get('hotkey_emergency', self.hotkey_emergency.get()))
            self._rebind_hotkeys()

        # Reflect toggles in widgets
        self.safe_mode = config.get('safe_mode', self.safe_mode)
        self.humanize = config.get('humanize', self.humanize)
        try:
            self.safe_switch.select() if self.safe_mode else self.safe_switch.deselect()
            self.humanize_switch.select() if self.humanize else self.humanize_switch.deselect()
        except Exception:
            pass

        # Clear and rebuild rows
        for r in self.rows:
            try:
                r.destroy()
            except Exception:
                pass
        self.rows.clear()

        rows_data = config.get('rows', [])
        if not rows_data:
            self.add_row(is_first=True)
        else:
            for i, data in enumerate(rows_data):
                self.add_row(
                    is_first=(i == 0),
                    key=data.get('key', ''),
                    hold=data.get('hold', '0.0'),
                    delay=data.get('delay', data.get('sleep', '0.5')),
                    enabled=data.get('enabled', True),
                )

    # ------------------------------------------------------------------- help
    def show_help(self):
        dialog = self._make_dialog("Help — Actions Reference", "640x640", grab=False)

        hdr = ctk.CTkFrame(dialog, fg_color="transparent")
        hdr.pack(fill="x", padx=25, pady=(20, 10))
        if self.logo is not None:
            ctk.CTkLabel(hdr, image=self.logo, text="").pack(side="left", padx=(0, 12))
        hdr_text = ctk.CTkFrame(hdr, fg_color="transparent")
        hdr_text.pack(side="left")
        ctk.CTkLabel(hdr_text, text=TOOL_NAME, font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ctk.CTkLabel(hdr_text, text=f"Version {VERSION}", font=("Segoe UI", 12),
                     text_color=COLORS["text_dim"]).pack(anchor="w")

        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent",
                                        scrollbar_button_color=COLORS["border"],
                                        scrollbar_button_hover_color=COLORS["accent"])
        scroll.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        def section(title):
            ctk.CTkLabel(scroll, text=title, font=("Segoe UI", 15, "bold"),
                         text_color=COLORS["accent"]).pack(anchor="w", padx=10, pady=(14, 6))

        def card():
            c = ctk.CTkFrame(scroll, fg_color=COLORS["bg_card"], corner_radius=10)
            c.pack(fill="x", padx=8, pady=2)
            return c

        def line(c, action, desc):
            r = ctk.CTkFrame(c, fg_color="transparent")
            r.pack(fill="x", padx=14, pady=5)
            ctk.CTkLabel(r, text=action, font=("Consolas", 13, "bold"), text_color=COLORS["text"],
                         width=210, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=desc, font=("Segoe UI", 12), text_color=COLORS["text_dim"]).pack(side="left")

        section("⌨  Keyboard")
        c = card()
        line(c, "a, b, 1, 2 …", "Single key press")
        line(c, "space  enter  tab  esc", "Special keys")
        line(c, "up  down  left  right", "Arrow keys")
        line(c, "f1 – f12", "Function keys")
        line(c, "shift  ctrl  alt  win", "Modifier keys (use Hold Time)")
        line(c, "ctrl+c, alt+f4", "Key combos (need a modifier)")
        line(c, "Hello World!", "Type a text string")

        section("🖱  Mouse")
        c = card()
        line(c, "click / rclick / mclick", "Left / right / middle click")
        line(c, "click(x,y)", "Left click at coordinates")
        line(c, "rclick(x,y) / mclick(x,y)", "Right / middle click at coords")
        line(c, "moveto(x,y)", "Move cursor to coordinates")
        line(c, "drag(x1,y1,x2,y2)", "Drag from A to B")

        section("🎨  Color & Conditions")
        c = card()
        line(c, "waitcolor(r,g,b,x,y)", "Wait until color appears at (x,y)")
        line(c, "ifcolor(r,g,b,x,y)", "Run next row only if color matches")
        line(c, "ifnotcolor(r,g,b,x,y)", "Run next row only if color is absent")

        section("🔁  Control Flow")
        c = card()
        line(c, "repeat(N)", "Repeat the rows below N times…")
        line(c, "endrepeat", "…until this marker")

        section("⏱  Timing")
        c = card()
        line(c, "0.5", "Fixed delay (seconds)")
        line(c, "0.3-0.8", "Random delay between min and max")

        section("💡  Tips")
        tips_card = ctk.CTkFrame(scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        tips_card.pack(fill="x", padx=8, pady=2)
        tips = [
            ("🎯", "Use Capture to grab coordinates & colors"),
            ("☑", "Uncheck a row to skip it without deleting"),
            ("🧍", "Humanize adds curved, jittered mouse moves"),
            ("⏸", "Pause/resume keeps your place in the sequence"),
            ("🛡", "Safe Mode blocks dangerous keys (ctrl, alt …)"),
            ("💾", "Your session auto-restores on next launch"),
        ]
        for icon, tip in tips:
            tr = ctk.CTkFrame(tips_card, fg_color="transparent")
            tr.pack(fill="x", padx=14, pady=5)
            ctk.CTkLabel(tr, text=icon, font=("Segoe UI", 15), width=28).pack(side="left", padx=(0, 8))
            ctk.CTkLabel(tr, text=tip, font=("Segoe UI", 12), text_color=COLORS["text_dim"]).pack(side="left")

        self._center_window(dialog)

    # ------------------------------------------------------------------- toast / dialogs
    def _toast(self, message, kind="success"):
        """Non-blocking transient notification in the top-right of the window."""
        color = {"success": COLORS["success"], "danger": COLORS["danger"],
                 "info": COLORS["accent"]}.get(kind, COLORS["accent"])
        icon = {"success": "✓", "danger": "⚠", "info": "ℹ"}.get(kind, "ℹ")
        try:
            if self._toast_frame is not None and self._toast_frame.winfo_exists():
                self._toast_frame.destroy()
        except Exception:
            pass
        toast = ctk.CTkFrame(self, fg_color=COLORS["bg_card"], corner_radius=10, border_width=2,
                             border_color=color)
        ctk.CTkLabel(toast, text=f"{icon}  {message}", font=("Segoe UI", 13),
                     text_color=COLORS["text"]).pack(padx=16, pady=10)
        toast.place(relx=0.5, rely=0.02, anchor="n")
        self._toast_frame = toast
        self.after(2600, lambda: toast.destroy() if toast.winfo_exists() else None)

    def _make_dialog(self, title, geometry, grab=True):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry(geometry)
        dialog.transient(self)
        if grab:
            dialog.grab_set()
        dialog.configure(fg_color=COLORS["bg_dark"])
        self._set_icon(dialog)
        return dialog

    def _center_window(self, window):
        window.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - window.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - window.winfo_height()) // 2
        window.geometry(f"+{max(0, x)}+{max(0, y)}")

    # ------------------------------------------------------------------- close
    def _on_close(self):
        if self.running:
            self.running = False
            self.paused = False
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        self._save_settings()
        _end_high_res_timer()
        self.destroy()


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = KeyClickerApp()
    app.mainloop()
