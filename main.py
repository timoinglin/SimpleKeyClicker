"""
SimpleKeyClicker - Modern Edition
A gorgeous, modern key/mouse automation tool with a premium dark UI.
"""

import threading
import time
import os
import sys
import json
import pydirectinput
import keyboard
import customtkinter as ctk
from tkinter import filedialog, PhotoImage
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

# Set Windows timer resolution to 1ms for high-precision timing
if os.name == 'nt':
    try:
        ctypes.windll.winmm.timeBeginPeriod(1)
    except Exception:
        pass

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# App constants
TOOL_NAME = "SimpleKeyClicker"
VERSION = "2.0 Modern"
ICON_PATH = resource_path("logo.ico")
LOGO_PATH = resource_path("logo.png")
EMERGENCY_STOP_KEY = 'esc'

SINGLE_ACTION_KEYS = {
    'tab', 'space', 'enter', 'esc', 'backspace', 'delete', 'insert',
    'up', 'down', 'left', 'right', 'home', 'end', 'pageup', 'pagedown',
    'capslock', 'numlock', 'scrolllock', 'printscreen', 'prntscrn', 'prtsc', 'pause',
    'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12',
    'shift', 'ctrl', 'alt', 'win', 'cmd'
}

DANGEROUS_KEYS = {'alt', 'ctrl', 'shift', 'win', 'cmd', 'f4', 'delete', 'tab'}
SYSTEM_COMMANDS = {'waitcolor', 'ifcolor'}
COLOR_MATCH_TOLERANCE = 10
WAITCOLOR_TIMEOUT = 30

# Modern color palette
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

# Configure CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ModernActionRow(ctk.CTkFrame):
    """A single action row with modern styling."""
    
    def __init__(self, parent, app, key="", hold="0.0", delay="0.5", is_first=False, **kwargs):
        super().__init__(parent, fg_color=COLORS["bg_card"], corner_radius=12, **kwargs)
        self.app = app
        self.is_first = is_first
        self.is_active = False
        
        # Variables
        self.key_var = ctk.StringVar(value=key)
        self.hold_var = ctk.StringVar(value=hold)
        self.delay_var = ctk.StringVar(value=delay)
        
        self._create_widgets()
        
    def _create_widgets(self):
        # Main container with padding
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="x", padx=15, pady=12)
        
        # Status indicator (left side)
        self.status_label = ctk.CTkLabel(container, text="", width=24, font=("Segoe UI", 14))
        self.status_label.pack(side="left", padx=(0, 10))
        
        # Key/Button input section
        key_frame = ctk.CTkFrame(container, fg_color="transparent")
        key_frame.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(key_frame, text="Action", font=("Segoe UI", 11), text_color=COLORS["text_dim"]).pack(anchor="w")
        
        key_input_frame = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_input_frame.pack(fill="x")
        
        self.key_entry = ctk.CTkEntry(
            key_input_frame, textvariable=self.key_var, width=200,
            placeholder_text="click, key, moveto(x,y)...",
            font=("Segoe UI", 13), height=36, corner_radius=8
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
        timing_frame.pack(side="left", padx=20)
        
        # Hold time
        hold_frame = ctk.CTkFrame(timing_frame, fg_color="transparent")
        hold_frame.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(hold_frame, text="Hold (s)", font=("Segoe UI", 11), text_color=COLORS["text_dim"]).pack(anchor="w")
        self.hold_entry = ctk.CTkEntry(hold_frame, textvariable=self.hold_var, width=70, height=36, corner_radius=8, font=("Segoe UI", 13))
        self.hold_entry.pack()
        
        # Delay time
        delay_frame = ctk.CTkFrame(timing_frame, fg_color="transparent")
        delay_frame.pack(side="left")
        ctk.CTkLabel(delay_frame, text="Delay (s)", font=("Segoe UI", 11), text_color=COLORS["text_dim"]).pack(anchor="w")
        self.delay_entry = ctk.CTkEntry(delay_frame, textvariable=self.delay_var, width=70, height=36, corner_radius=8, font=("Segoe UI", 13))
        self.delay_entry.pack()
        
        # Action buttons (right side)
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(side="right")
        
        btn_style = {"width": 36, "height": 36, "corner_radius": 8, "font": ("Segoe UI", 14)}
        
        self.up_btn = ctk.CTkButton(btn_frame, text="▲", fg_color=COLORS["bg_hover"], hover_color=COLORS["border"], command=self._move_up, **btn_style)
        self.up_btn.pack(side="left", padx=2)
        
        self.down_btn = ctk.CTkButton(btn_frame, text="▼", fg_color=COLORS["bg_hover"], hover_color=COLORS["border"], command=self._move_down, **btn_style)
        self.down_btn.pack(side="left", padx=2)
        
        self.dup_btn = ctk.CTkButton(btn_frame, text="❏", fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"], command=self._duplicate, **btn_style)
        self.dup_btn.pack(side="left", padx=2)
        
        self.del_btn = ctk.CTkButton(btn_frame, text="✕", fg_color=COLORS["danger"], hover_color="#ff6b7a", command=self._delete, **btn_style)
        self.del_btn.pack(side="left", padx=2)
        
        if self.is_first:
            self.del_btn.configure(state="disabled", fg_color=COLORS["bg_hover"])
    
    def _capture(self):
        self.app._start_capture(self.key_var)
    
    def _move_up(self):
        self.app._move_row(self, -1)
    
    def _move_down(self):
        self.app._move_row(self, 1)
    
    def _duplicate(self):
        self.app._duplicate_row(self)
    
    def _delete(self):
        if not self.is_first:
            self.app._delete_row(self)
    
    def set_active(self, active):
        self.is_active = active
        if active:
            self.configure(fg_color=COLORS["accent"], border_width=0)
            self.status_label.configure(text="▶")
        else:
            self.configure(fg_color=COLORS["bg_card"], border_width=0)
            self.status_label.configure(text="")
    
    def set_completed(self):
        self.status_label.configure(text="✓", text_color=COLORS["success"])
    
    def get_data(self):
        return {
            "key": self.key_var.get(),
            "hold": self.hold_var.get(),
            "delay": self.delay_var.get()
        }


class KeyClickerApp(ctk.CTk):
    """Modern SimpleKeyClicker application."""
    
    def __init__(self):
        super().__init__()
        
        self.title(f"{TOOL_NAME} {VERSION}")
        self.geometry("950x650")
        self.minsize(900, 550)
        self.configure(fg_color=COLORS["bg_dark"])
        
        try:
            self.iconbitmap(ICON_PATH)
        except:
            pass
        
        # State
        self.running = False
        self.safe_mode = True
        self.rows = []
        self.thread = None
        self.error_acknowledged = threading.Event()
        
        # Variables
        self.run_mode = ctk.StringVar(value="infinite")
        self.repetitions = ctk.IntVar(value=10)
        
        self._create_ui()
        self._setup_hotkeys()
    
    def _create_ui(self):
        # Main container
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=25, pady=20)
        
        # Header section
        self._create_header(main)
        
        # Controls section
        self._create_controls(main)
        
        # Action rows section
        self._create_rows_section(main)
        
        # Status bar
        self._create_status_bar(main)
    
    def _create_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", pady=(0, 20))
        
        # Logo and title
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left")
        
        try:
            logo_img = Image.open(LOGO_PATH)
            logo_img = logo_img.resize((50, 50), Image.LANCZOS)
            self.logo = ctk.CTkImage(logo_img, size=(50, 50))
            ctk.CTkLabel(title_frame, image=self.logo, text="").pack(side="left", padx=(0, 15))
        except:
            pass
        
        text_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        text_frame.pack(side="left")
        ctk.CTkLabel(text_frame, text=TOOL_NAME, font=("Segoe UI", 28, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text="Modern Automation Tool", font=("Segoe UI", 13), text_color=COLORS["text_dim"]).pack(anchor="w")
        
        # Menu buttons (right side)
        menu_frame = ctk.CTkFrame(header, fg_color="transparent")
        menu_frame.pack(side="right")
        
        btn_style = {"width": 100, "height": 36, "corner_radius": 8, "font": ("Segoe UI", 12)}
        
        ctk.CTkButton(menu_frame, text="📂 Load", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"], command=self.load_configuration, **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="💾 Save", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"], command=self.save_configuration, **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(menu_frame, text="❓ Help", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"], command=self.show_help, **btn_style).pack(side="left", padx=5)
    
    def _create_controls(self, parent):
        controls = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=16)
        controls.pack(fill="x", pady=(0, 15))
        
        inner = ctk.CTkFrame(controls, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)
        
        # Start/Stop buttons
        btn_frame = ctk.CTkFrame(inner, fg_color="transparent")
        btn_frame.pack(side="left")
        
        self.start_btn = ctk.CTkButton(
            btn_frame, text="▶  START", width=140, height=50,
            font=("Segoe UI", 15, "bold"), corner_radius=12,
            fg_color=COLORS["success"], hover_color="#00b85c",
            command=self.start_action
        )
        self.start_btn.pack(side="left", padx=(0, 10))
        
        self.stop_btn = ctk.CTkButton(
            btn_frame, text="⏹  STOP", width=140, height=50,
            font=("Segoe UI", 15, "bold"), corner_radius=12,
            fg_color=COLORS["danger"], hover_color="#ff6b7a",
            command=self.stop_action
        )
        self.stop_btn.pack(side="left")
        
        # Run mode options
        mode_frame = ctk.CTkFrame(inner, fg_color="transparent")
        mode_frame.pack(side="left", padx=40)
        
        ctk.CTkRadioButton(mode_frame, text="Run Indefinitely", variable=self.run_mode, value="infinite", font=("Segoe UI", 13), command=self._update_rep_state).pack(anchor="w")
        
        rep_frame = ctk.CTkFrame(mode_frame, fg_color="transparent")
        rep_frame.pack(anchor="w", pady=(5, 0))
        ctk.CTkRadioButton(rep_frame, text="Run", variable=self.run_mode, value="limited", font=("Segoe UI", 13), command=self._update_rep_state).pack(side="left")
        self.rep_entry = ctk.CTkEntry(rep_frame, textvariable=self.repetitions, width=60, height=30, corner_radius=6, font=("Segoe UI", 13))
        self.rep_entry.pack(side="left", padx=8)
        ctk.CTkLabel(rep_frame, text="times", font=("Segoe UI", 13)).pack(side="left")
        
        # Safe mode toggle
        safe_frame = ctk.CTkFrame(inner, fg_color="transparent")
        safe_frame.pack(side="right")
        
        self.safe_switch = ctk.CTkSwitch(
            safe_frame, text="Safe Mode", font=("Segoe UI", 13),
            command=self._toggle_safe_mode,
            progress_color=COLORS["success"]
        )
        self.safe_switch.select()
        self.safe_switch.pack()
        
        self._update_rep_state()
    
    def _create_rows_section(self, parent):
        section = ctk.CTkFrame(parent, fg_color="transparent")
        section.pack(fill="both", expand=True)
        
        # Header with Add button
        header = ctk.CTkFrame(section, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(header, text="Automation Actions", font=("Segoe UI", 18, "bold")).pack(side="left")
        
        ctk.CTkButton(
            header, text="+ Add Action", width=120, height=36,
            font=("Segoe UI", 13), corner_radius=8,
            fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
            command=self._add_row
        ).pack(side="right")
        
        # Scrollable container for rows
        self.rows_frame = ctk.CTkScrollableFrame(
            section, fg_color="transparent",
            scrollbar_button_color=COLORS["border"],
            scrollbar_button_hover_color=COLORS["accent"]
        )
        self.rows_frame.pack(fill="both", expand=True)
        
        # Add initial row
        self._add_row(is_first=True)
    
    def _create_status_bar(self, parent):
        status = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10, height=40)
        status.pack(fill="x", pady=(15, 0))
        status.pack_propagate(False)
        
        inner = ctk.CTkFrame(status, fg_color="transparent")
        inner.pack(fill="x", padx=15, pady=8)
        
        self.status_label = ctk.CTkLabel(inner, text="⏸ Ready", font=("Segoe UI", 13), text_color=COLORS["text_dim"])
        self.status_label.pack(side="left")
        
        ctk.CTkLabel(inner, text="Ctrl+F2: Start  •  Ctrl+F3: Stop  •  ESC: Emergency Stop", font=("Segoe UI", 11), text_color=COLORS["text_dim"]).pack(side="right")
    
    def _update_rep_state(self):
        if self.run_mode.get() == "limited":
            self.rep_entry.configure(state="normal")
        else:
            self.rep_entry.configure(state="disabled")
    
    def _toggle_safe_mode(self):
        self.safe_mode = self.safe_switch.get()
    
    def _add_row(self, is_first=False, key="", hold="0.0", delay="0.5"):
        row = ModernActionRow(self.rows_frame, self, key=key, hold=hold, delay=delay, is_first=is_first)
        row.pack(fill="x", pady=5)
        self.rows.append(row)
        self._update_row_buttons()
    
    def _delete_row(self, row):
        if row in self.rows and not row.is_first:
            self.rows.remove(row)
            row.destroy()
            self._update_row_buttons()
    
    def _duplicate_row(self, row):
        if row in self.rows:
            data = row.get_data()
            idx = self.rows.index(row)
            new_row = ModernActionRow(self.rows_frame, self, **data)
            new_row.pack(fill="x", pady=5)
            self.rows.insert(idx + 1, new_row)
            self._repack_rows()
    
    def _move_row(self, row, direction):
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
    
    def _setup_hotkeys(self):
        try:
            keyboard.add_hotkey('ctrl+f2', self.start_action)
            keyboard.add_hotkey('ctrl+f3', self.stop_action)
            keyboard.add_hotkey(EMERGENCY_STOP_KEY, self.emergency_stop)
        except Exception as e:
            print(f"Hotkey warning: {e}")
    
    def start_action(self):
        if self.running:
            return
        
        # Validate rows
        for i, row in enumerate(self.rows):
            if not row.key_var.get().strip():
                self._show_error(f"Row {i+1}: Please specify an action.")
                return
            try:
                float(row.hold_var.get())
                float(row.delay_var.get())
            except ValueError:
                self._show_error(f"Row {i+1}: Invalid timing values.")
                return
        
        if self.run_mode.get() == "limited":
            try:
                if self.repetitions.get() <= 0:
                    raise ValueError()
            except:
                self._show_error("Invalid repetition count.")
                return
        
        self.running = True
        self.status_label.configure(text="▶ Running...", text_color=COLORS["success"])
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()
    
    def stop_action(self):
        if self.running:
            self.running = False
            self.after(0, self._update_stopped)
    
    def _update_stopped(self):
        self.status_label.configure(text="⏸ Stopped", text_color=COLORS["text_dim"])
        for row in self.rows:
            row.set_active(False)
    
    def emergency_stop(self):
        if self.running:
            self.running = False
            self.after(0, lambda: self.status_label.configure(text="⚠ Emergency Stop", text_color=COLORS["danger"]))
            self.after(0, self._update_stopped)
    
    def _run_loop(self):
        mode = self.run_mode.get()
        reps = self.repetitions.get() if mode == "limited" else 0
        loop = 0
        
        try:
            while self.running:
                loop += 1
                if mode == "limited" and loop > reps:
                    break
                
                status = f"▶ Running (Loop {loop})" if mode == "infinite" else f"▶ Running ({loop}/{reps})"
                self.after(0, lambda s=status: self.status_label.configure(text=s, text_color=COLORS["success"]))
                
                for i, row in enumerate(self.rows):
                    if not self.running:
                        break
                    
                    self.after(0, lambda r=row: r.set_active(True))
                    
                    key = row.key_var.get().strip()
                    try:
                        hold = float(row.hold_var.get())
                        delay = float(row.delay_var.get())
                    except:
                        self.running = False
                        break
                    
                    if not self._perform_action(key, hold):
                        self.running = False
                        break
                    
                    if delay > 0.05:
                        self.after(0, lambda r=row: r.set_completed())
                    
                    if delay > 0:
                        time.sleep(delay)
                    
                    if delay > 0.05:
                        self.after(0, lambda r=row: r.set_active(False))
        finally:
            if mode == "limited" and loop > reps:
                self.after(0, lambda: self.status_label.configure(text="✓ Completed", text_color=COLORS["success"]))
            self.after(0, self._update_stopped)
            self.running = False
    
    def _perform_action(self, key, hold_time):
        if not self.running:
            return False
        
        # Safe mode check
        k_lower = key.lower()
        is_dangerous = k_lower in DANGEROUS_KEYS or any(cmd in k_lower for cmd in SYSTEM_COMMANDS)
        if self.safe_mode and is_dangerous:
            self.after(0, lambda: self._show_error(f"'{key}' blocked by Safe Mode."))
            return False
        
        try:
            # Handle commands with coordinates
            if '(' in key and ')' in key:
                cmd = key.split('(')[0].lower()
                args_str = key[key.index('(')+1:key.rindex(')')]
                args = [a.strip() for a in args_str.split(',')]
                
                if cmd in {'click', 'rclick', 'mclick'}:
                    x, y = map(int, args)
                    pydirectinput.moveTo(x, y)
                    button = {'click': 'left', 'rclick': 'right', 'mclick': 'middle'}[cmd]
                    if hold_time > 0:
                        pydirectinput.mouseDown(button=button)
                        time.sleep(hold_time)
                        pydirectinput.mouseUp(button=button)
                    else:
                        pydirectinput.click(button=button)
                    return True
                
                elif cmd == 'moveto':
                    x, y = map(int, args)
                    pyautogui.moveTo(x, y)
                    return True
                
                elif cmd == 'waitcolor':
                    r, g, b, x, y = map(int, args)
                    return self._wait_for_color(r, g, b, x, y)
            
            # Simple actions
            if k_lower == "click":
                if hold_time > 0:
                    pydirectinput.mouseDown()
                    time.sleep(hold_time)
                    pydirectinput.mouseUp()
                else:
                    pydirectinput.click()
            elif k_lower == "rclick":
                if hold_time > 0:
                    pydirectinput.mouseDown(button='right')
                    time.sleep(hold_time)
                    pydirectinput.mouseUp(button='right')
                else:
                    pydirectinput.rightClick()
            elif k_lower == "mclick":
                if hold_time > 0:
                    pydirectinput.mouseDown(button='middle')
                    time.sleep(hold_time)
                    pydirectinput.mouseUp(button='middle')
                else:
                    pydirectinput.middleClick()
            elif k_lower in SINGLE_ACTION_KEYS or len(key) == 1:
                if hold_time > 0:
                    pydirectinput.keyDown(key)
                    time.sleep(hold_time)
                    pydirectinput.keyUp(key)
                else:
                    pydirectinput.press(key)
            else:
                pyautogui.write(key, interval=0.0)
            
            return True
        except Exception as e:
            self.after(0, lambda: self._show_error(f"Error: {e}"))
            return False
    
    def _wait_for_color(self, r, g, b, x, y, timeout=WAITCOLOR_TIMEOUT):
        start = time.time()
        while time.time() - start < timeout:
            if not self.running:
                return False
            try:
                pixel = ImageGrab.grab(bbox=(x, y, x+1, y+1)).getpixel((0, 0))
                if all(abs(pixel[i] - [r, g, b][i]) <= COLOR_MATCH_TOLERANCE for i in range(3)):
                    return True
            except:
                pass
            time.sleep(0.1)
        self.after(0, lambda: self._show_error(f"Color not found at ({x},{y})"))
        return False
    
    def _start_capture(self, key_var):
        if self.running:
            return
        
        self.attributes('-alpha', 0.3)
        
        # Capture data
        data = {'x': None, 'y': None, 'color': None}
        
        def on_click(x, y, button, pressed):
            if button == mouse.Button.left and pressed:
                data['x'], data['y'] = int(x), int(y)
                try:
                    data['color'] = pyautogui.pixel(data['x'], data['y'])
                except:
                    data['color'] = (0, 0, 0)
                return False
        
        listener = mouse.Listener(on_click=on_click)
        listener.start()
        listener.join()
        
        self.attributes('-alpha', 1.0)
        self.lift()
        self.focus_force()
        
        if data['x'] is not None:
            self._show_capture_dialog(data, key_var)
    
    def _show_capture_dialog(self, data, key_var):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Capture Options")
        dialog.geometry("350x320")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(fg_color=COLORS["bg_dark"])
        
        x, y = data['x'], data['y']
        r, g, b = data['color']
        
        # Color preview and info
        info_frame = ctk.CTkFrame(dialog, fg_color=COLORS["bg_card"], corner_radius=12)
        info_frame.pack(fill="x", padx=20, pady=20)
        
        inner = ctk.CTkFrame(info_frame, fg_color="transparent")
        inner.pack(padx=15, pady=15)
        
        color_hex = f"#{r:02x}{g:02x}{b:02x}"
        color_preview = ctk.CTkFrame(inner, width=40, height=40, corner_radius=8, fg_color=color_hex)
        color_preview.pack(side="left", padx=(0, 15))
        
        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left")
        ctk.CTkLabel(text_frame, text=f"Position: ({x}, {y})", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ctk.CTkLabel(text_frame, text=f"Color: RGB({r}, {g}, {b})", font=("Segoe UI", 12), text_color=COLORS["text_dim"]).pack(anchor="w")
        
        # Action buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20)
        
        def insert(cmd):
            key_var.set(cmd)
            dialog.destroy()
        
        btn_style = {"height": 40, "corner_radius": 10, "font": ("Segoe UI", 13)}
        
        ctk.CTkButton(btn_frame, text=f"🖱 Click at ({x}, {y})", fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"], command=lambda: insert(f"click({x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text=f"➡ Move to ({x}, {y})", fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"], command=lambda: insert(f"moveto({x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text=f"🎨 Wait for color", fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"], command=lambda: insert(f"waitcolor({r},{g},{b},{x},{y})"), **btn_style).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text="Cancel", fg_color=COLORS["bg_card"], hover_color=COLORS["bg_hover"], command=dialog.destroy, **btn_style).pack(fill="x", pady=(15, 5))
        
        self._center_window(dialog)
    
    def save_configuration(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            config = {
                'run_mode': self.run_mode.get(),
                'repetitions': self.repetitions.get(),
                'rows': [r.get_data() for r in self.rows]
            }
            with open(path, 'w') as f:
                json.dump(config, f, indent=2)
            self._show_success("Configuration saved!")
    
    def load_configuration(self):
        if self.running:
            self._show_error("Stop automation before loading.")
            return
        
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            try:
                with open(path) as f:
                    config = json.load(f)
                
                self.run_mode.set(config.get('run_mode', 'infinite'))
                self.repetitions.set(config.get('repetitions', 10))
                self._update_rep_state()
                
                # Clear existing rows
                for r in self.rows:
                    r.destroy()
                self.rows.clear()
                
                # Add loaded rows
                rows_data = config.get('rows', [])
                if not rows_data:
                    self._add_row(is_first=True)
                else:
                    for i, data in enumerate(rows_data):
                        self._add_row(is_first=(i == 0), key=data.get('key', ''), hold=data.get('hold', '0.0'), delay=data.get('delay', '0.5'))
                
                self._show_success("Configuration loaded!")
            except Exception as e:
                self._show_error(f"Load failed: {e}")
    
    def show_help(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Help - Actions Reference")
        dialog.geometry("600x500")
        dialog.transient(self)
        dialog.configure(fg_color=COLORS["bg_dark"])
        
        text = ctk.CTkTextbox(dialog, font=("Consolas", 12), fg_color=COLORS["bg_card"], corner_radius=12)
        text.pack(fill="both", expand=True, padx=20, pady=20)
        
        help_text = """
═══════════════════════════════════════════════════
                 KEYBOARD ACTIONS
═══════════════════════════════════════════════════

Letters & Numbers:  a, b, c, 1, 2, 3, etc.
Special Keys:       space, enter, tab, esc, backspace
Arrow Keys:         up, down, left, right
Function Keys:      f1, f2, ... f12
Modifiers:          shift, ctrl, alt, win


═══════════════════════════════════════════════════
                  MOUSE ACTIONS
═══════════════════════════════════════════════════

click           Left click at current position
rclick          Right click at current position
mclick          Middle click at current position

click(x,y)      Click at coordinates (x, y)
rclick(x,y)     Right click at (x, y)
mclick(x,y)     Middle click at (x, y)
moveto(x,y)     Move cursor to (x, y)


═══════════════════════════════════════════════════
               COLOR DETECTION
═══════════════════════════════════════════════════

waitcolor(r,g,b,x,y)   Wait until color RGB appears at (x,y)


═══════════════════════════════════════════════════
                    TIPS
═══════════════════════════════════════════════════

• Use the 🎯 Capture button to get coordinates/colors
• Hold Time > 0 holds the key/button down
• Delay is the pause AFTER the action
• Safe Mode blocks dangerous keys (ctrl, alt, etc.)
• ESC = Emergency Stop (always works)
"""
        text.insert("1.0", help_text)
        text.configure(state="disabled")
        
        self._center_window(dialog)
    
    def _show_error(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Error")
        dialog.geometry("400x180")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(fg_color=COLORS["bg_dark"])
        
        ctk.CTkLabel(dialog, text="⚠", font=("Segoe UI", 48), text_color=COLORS["danger"]).pack(pady=(20, 10))
        ctk.CTkLabel(dialog, text=message, font=("Segoe UI", 13), wraplength=350).pack(pady=5)
        ctk.CTkButton(dialog, text="OK", width=100, fg_color=COLORS["accent"], command=dialog.destroy).pack(pady=15)
        
        self._center_window(dialog)
        self.error_acknowledged.set()
    
    def _show_success(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry("350x150")
        dialog.transient(self)
        dialog.configure(fg_color=COLORS["bg_dark"])
        
        ctk.CTkLabel(dialog, text="✓", font=("Segoe UI", 48), text_color=COLORS["success"]).pack(pady=(20, 10))
        ctk.CTkLabel(dialog, text=message, font=("Segoe UI", 13)).pack()
        ctk.CTkButton(dialog, text="OK", width=100, fg_color=COLORS["success"], command=dialog.destroy).pack(pady=15)
        
        self._center_window(dialog)
    
    def _center_window(self, window):
        window.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - window.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - window.winfo_height()) // 2
        window.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = KeyClickerApp()
    app.mainloop()
