import os
import sys
import json
import csv
import shutil
import tempfile
import subprocess
import re
import importlib
import platform
import threading
import urllib.request
import math
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    Image = None
    ImageDraw = None
    ImageTk = None

PDF_IMPORT_ERROR = ""
PdfReader = None
PdfWriter = None
try:
    from pypdf import PdfMerger, PdfReader, PdfWriter
except ImportError as pypdf_error:
    try:
        from PyPDF2 import PdfMerger, PdfReader, PdfWriter  # fallback for older installs
    except ImportError as pypdf2_error:
        PdfMerger = None
        PdfReader = None
        PdfWriter = None
        PDF_IMPORT_ERROR = (
            f"pypdf error: {pypdf_error}; PyPDF2 error: {pypdf2_error}"
        )

try:
    from send2trash import send2trash
except ImportError:
    send2trash = None

try:
    from win10toast import ToastNotifier as Win10ToastNotifier
except ImportError:
    Win10ToastNotifier = None

try:
    from winotify import Notification as WinotifyNotification
except ImportError:
    WinotifyNotification = None

try:
    pdfplumber = importlib.import_module("pdfplumber")
except Exception:
    pdfplumber = None

try:
    load_workbook = importlib.import_module("openpyxl").load_workbook
except Exception:
    load_workbook = None

# ------------------------
# Main Window
# ------------------------
root = tk.Tk()
root.title("PDF Record Manager")
BASE_WINDOW_WIDTH = 880
BASE_WINDOW_HEIGHT = 820
MIN_WINDOW_WIDTH = 720
MIN_WINDOW_HEIGHT = 620
DEFAULT_MARGIN_X = 60
DEFAULT_MARGIN_Y = 120

ACCENT_COLOR = "#2D7FF9"
BG_COLOR = "#0F172A"
SURFACE_COLOR = "#1E293B"
TEXT_COLOR = "#E2E8F0"
SUBTEXT_COLOR = "#94A3B8"
SHELL_BG_COLOR = BG_COLOR
SURFACE_ELEVATED_COLOR = "#24344f"
FIELD_BG_COLOR = "#131c2f"
BORDER_COLOR = "#334155"
ACTION_BAR_BG = "#18243a"
STATUS_ACCENT_COLOR = "#93c5fd"
MUTED_BUTTON_BG = "#334155"
MUTED_BUTTON_HOVER_BG = "#415775"
STATUS_BADGE_BG = "#2b5f98"
STATUS_BADGE_TEXT = "#f8fbff"
STATUS_SURFACE_BG = "#20314c"
PENDING_ROW_BG = "#121d31"
PENDING_ROW_HOVER_BG = "#1a2a45"
PENDING_ROW_SELECTED_BG = "#27456f"
PENDING_ROW_SELECTED_HOVER_BG = "#2f5487"
PENDING_ROW_TEXT = "#e5edff"
LISTBOX_BG = "#101a2b"
LISTBOX_BORDER = "#2f405d"
LISTBOX_TEXT = "#e5ecf8"
FOCUS_RING_COLOR = "#7fb4ff"

AUTO_REFRESH_INTERVAL_MS = 1000
APP_ICON_PREFERRED_NAMES = ("app.ico", "application.ico", "icon.ico")
APP_VERSION = "1.2.0"
APP_BUILD_COMMIT = os.environ.get("PDF_AUTOTOOL_COMMIT", "unknown")
APP_BUILD_DATE = os.environ.get("PDF_AUTOTOOL_BUILD_DATE", "unknown")
APP_BUILD_INFO_FILENAME = "build_info.json"
DEFAULT_UPDATE_MANIFEST_URL = ""
UPDATE_CHECK_TIMEOUT_SEC = 8
TOOLBAR_ICON_SIZE = 20
TOOLBAR_ICON_COLOR = "#FFFFFF"
TOOLBAR_ICON_STROKE_MULTIPLIER = 1.55
TOOLBAR_ICON_RENDER_SCALE = 6
TOOLBAR_ICON_PREVIEW = "\u25c9"
TOOLBAR_ICON_SELECT_ALL = "\u2611"
TOOLBAR_ICON_SELECT_NONE = "\u2610"
TOOLBAR_ICON_SELECT_PARTIAL = "\u25a3"
TOOLBAR_ICON_SOURCE_ADD = "\u2795"
TOOLBAR_ICON_SOURCE_REMOVE = "\u2796"
TOOLBAR_ICON_SOURCE_CLEAR = "\u2715"
TOOLBAR_ICON_EDIT = "\u270e"
TOOLBAR_ICON_ROTATE = "\u21ba"

ROTATION_PREVIEW_DEFAULT_WIDTH = 1080
ROTATION_PREVIEW_DEFAULT_HEIGHT = 740
ROTATION_PREVIEW_MIN_WIDTH = 760
ROTATION_PREVIEW_MIN_HEIGHT = 520
ROTATION_PREVIEW_THUMB_MIN_WIDTH = 110
ROTATION_PREVIEW_THUMB_MAX_WIDTH = 180
ROTATION_PREVIEW_THUMB_MAX_LIMIT = 320


def _find_app_icon_path():
    search_dirs = []

    if getattr(sys, "frozen", False):
        bundle_dir = getattr(sys, "_MEIPASS", None)
        if bundle_dir:
            search_dirs.append(bundle_dir)
        search_dirs.append(os.path.dirname(sys.executable))

    search_dirs.append(os.path.dirname(os.path.abspath(__file__)))

    checked = set()
    for directory in search_dirs:
        normalized_dir = os.path.normpath(directory)
        if normalized_dir in checked:
            continue
        checked.add(normalized_dir)

        if not os.path.isdir(directory):
            continue

        try:
            entries = os.listdir(directory)
        except OSError:
            continue

        lower_name_map = {name.lower(): name for name in entries}
        for preferred in APP_ICON_PREFERRED_NAMES:
            matched_name = lower_name_map.get(preferred)
            if matched_name:
                return os.path.normpath(os.path.join(directory, matched_name))

        ico_files = sorted(name for name in entries if name.lower().endswith(".ico"))
        if ico_files:
            return os.path.normpath(os.path.join(directory, ico_files[0]))

    return ""


def _apply_app_icon(window):
    icon_path = _find_app_icon_path()
    if not icon_path:
        return False

    try:
        window.iconbitmap(default=icon_path)
        return True
    except tk.TclError:
        try:
            window.iconbitmap(icon_path)
            return True
        except tk.TclError:
            return False


def apply_theme(window):
    window.configure(bg=SHELL_BG_COLOR)
    style = ttk.Style(window)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "TFrame",
        background=SHELL_BG_COLOR,
    )
    style.configure(
        "Card.TFrame",
        background=SURFACE_COLOR,
        relief="flat",
    )
    style.configure(
        "Shell.TFrame",
        background=SHELL_BG_COLOR,
    )
    style.configure(
        "HeaderCard.TFrame",
        background=SURFACE_ELEVATED_COLOR,
        relief="flat",
    )
    style.configure(
        "ActionBar.TFrame",
        background=ACTION_BAR_BG,
        relief="flat",
    )
    style.configure(
        "Title.TLabel",
        background=SHELL_BG_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 20),
    )
    style.configure(
        "HeaderTitle.TLabel",
        background=SURFACE_ELEVATED_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 20),
    )
    style.configure(
        "Subheading.TLabel",
        background=SHELL_BG_COLOR,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 10),
    )
    style.configure(
        "HeaderSubheading.TLabel",
        background=SURFACE_ELEVATED_COLOR,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 10),
    )
    style.configure(
        "HeaderStatus.TLabel",
        background=SURFACE_ELEVATED_COLOR,
        foreground=STATUS_BADGE_TEXT,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Card.TLabel",
        background=SURFACE_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI", 11),
    )
    style.configure(
        "SectionTitle.TLabel",
        background=SURFACE_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 11),
    )
    style.configure(
        "FieldLabel.TLabel",
        background=SURFACE_COLOR,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 9),
    )
    style.configure(
        "CardMuted.TLabel",
        background=SURFACE_COLOR,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 10),
    )
    style.configure(
        "StatusStrong.TLabel",
        background=SURFACE_COLOR,
        foreground="#f3f8ff",
        font=("Segoe UI Semibold", 11),
    )
    style.configure(
        "StatusText.TLabel",
        background=SHELL_BG_COLOR,
        foreground="#deebff",
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "StatusBadge.TFrame",
        background=STATUS_BADGE_BG,
        relief="flat",
    )
    style.configure(
        "StatusBadge.TLabel",
        background=STATUS_BADGE_BG,
        foreground=STATUS_BADGE_TEXT,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "StatusSurface.TFrame",
        background=STATUS_SURFACE_BG,
        relief="flat",
    )
    style.configure(
        "StatusSurface.TLabel",
        background=STATUS_SURFACE_BG,
        foreground="#f2f7ff",
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "ActionTitle.TLabel",
        background=ACTION_BAR_BG,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 9),
    )
    style.configure(
        "TLabel",
        background=SHELL_BG_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI", 11),
    )
    style.configure(
        "PendingRow.TFrame",
        background=PENDING_ROW_BG,
        relief="flat",
    )
    style.configure(
        "PendingRowHover.TFrame",
        background=PENDING_ROW_HOVER_BG,
        relief="flat",
    )
    style.configure(
        "PendingRowSelected.TFrame",
        background=PENDING_ROW_SELECTED_BG,
        relief="flat",
    )
    style.configure(
        "PendingRowSelectedHover.TFrame",
        background=PENDING_ROW_SELECTED_HOVER_BG,
        relief="flat",
    )
    style.configure(
        "PendingFile.TLabel",
        background=PENDING_ROW_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI", 10),
    )
    style.configure(
        "PendingFileHover.TLabel",
        background=PENDING_ROW_HOVER_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI", 10),
    )
    style.configure(
        "PendingFileSelected.TLabel",
        background=PENDING_ROW_SELECTED_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "PendingFileSelectedHover.TLabel",
        background=PENDING_ROW_SELECTED_HOVER_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "PendingFile.TCheckbutton",
        background=PENDING_ROW_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI", 10),
        padding=(2, 0),
    )
    style.map(
        "PendingFile.TCheckbutton",
        background=[("active", PENDING_ROW_HOVER_BG), ("selected", PENDING_ROW_BG)],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "PendingFileHover.TCheckbutton",
        background=PENDING_ROW_HOVER_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI", 10),
        padding=(2, 0),
    )
    style.map(
        "PendingFileHover.TCheckbutton",
        background=[("active", PENDING_ROW_HOVER_BG), ("selected", PENDING_ROW_HOVER_BG)],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "PendingFileSelected.TCheckbutton",
        background=PENDING_ROW_SELECTED_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI Semibold", 10),
        padding=(2, 0),
    )
    style.map(
        "PendingFileSelected.TCheckbutton",
        background=[("active", PENDING_ROW_SELECTED_HOVER_BG), ("selected", PENDING_ROW_SELECTED_BG)],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "PendingFileSelectedHover.TCheckbutton",
        background=PENDING_ROW_SELECTED_HOVER_BG,
        foreground=PENDING_ROW_TEXT,
        font=("Segoe UI Semibold", 10),
        padding=(2, 0),
    )
    style.map(
        "PendingFileSelectedHover.TCheckbutton",
        background=[("active", PENDING_ROW_SELECTED_HOVER_BG), ("selected", PENDING_ROW_SELECTED_HOVER_BG)],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "TButton",
        background=ACCENT_COLOR,
        foreground="white",
        font=("Segoe UI Semibold", 11),
        padding=8,
        borderwidth=0,
    )
    style.map(
        "TButton",
        background=[("active", "#5090ff"), ("disabled", "#4c566a")],
    )
    style.configure(
        "Accent.TButton",
        background=ACCENT_COLOR,
        foreground="white",
    )
    style.configure(
        "PrimaryAction.TButton",
        background=ACCENT_COLOR,
        foreground="white",
        font=("Segoe UI Semibold", 11),
        padding=(14, 8),
        borderwidth=0,
    )
    style.map(
        "PrimaryAction.TButton",
        background=[("active", "#5090ff"), ("disabled", "#4c566a")],
    )
    style.configure(
        "SecondaryAction.TButton",
        background=MUTED_BUTTON_BG,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 11),
        padding=(14, 8),
        borderwidth=0,
    )
    style.map(
        "SecondaryAction.TButton",
        background=[("active", MUTED_BUTTON_HOVER_BG), ("disabled", "#2b3446")],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "Subtle.TButton",
        background=MUTED_BUTTON_BG,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 10),
        padding=(10, 7),
        borderwidth=0,
    )
    style.map(
        "Subtle.TButton",
        background=[("active", MUTED_BUTTON_HOVER_BG), ("disabled", "#2b3446")],
        foreground=[("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "ToolbarIcon.TButton",
        background=ACCENT_COLOR,
        foreground="white",
        font=("Segoe UI Semibold", 10),
        padding=(9, 6),
        borderwidth=0,
    )
    style.map(
        "ToolbarIcon.TButton",
        background=[("active", "#5090ff"), ("disabled", "#4c566a")],
    )
    style.configure(
        "TEntry",
        fieldbackground=FIELD_BG_COLOR,
        foreground=TEXT_COLOR,
        insertcolor=TEXT_COLOR,
        bordercolor=BORDER_COLOR,
        relief="flat",
    )
    style.configure(
        "TCombobox",
        fieldbackground=FIELD_BG_COLOR,
        foreground=TEXT_COLOR,
        arrowcolor=ACCENT_COLOR,
    )
    style.map(
        "TCombobox",
        foreground=[("readonly", TEXT_COLOR), ("disabled", SUBTEXT_COLOR)],
        fieldbackground=[("readonly", FIELD_BG_COLOR), ("disabled", "#1b273d")],
        background=[("readonly", FIELD_BG_COLOR), ("active", FIELD_BG_COLOR)],
        selectforeground=[("readonly", TEXT_COLOR)],
        selectbackground=[("readonly", FIELD_BG_COLOR)],
        arrowcolor=[("readonly", ACCENT_COLOR), ("disabled", SUBTEXT_COLOR)],
    )
    style.configure(
        "Success.Horizontal.TProgressbar",
        troughcolor="#0b1220",
        background="#15803d",
        bordercolor="#0b1220",
        lightcolor="#16a34a",
        darkcolor="#14532d",
    )
    style.configure(
        "Vertical.TScrollbar",
        gripcount=0,
        background=BORDER_COLOR,
        troughcolor="#0b1220",
        bordercolor="#0b1220",
        lightcolor=BORDER_COLOR,
        darkcolor=BORDER_COLOR,
        arrowcolor="#64748b",
        relief="flat",
        arrowsize=10,
        width=12,
    )
    style.map(
        "Vertical.TScrollbar",
        background=[("active", "#475569"), ("pressed", ACCENT_COLOR)],
        arrowcolor=[("active", "#cbd5e1"), ("pressed", "#ffffff")],
    )
    style.configure(
        "Horizontal.TScrollbar",
        gripcount=0,
        background=BORDER_COLOR,
        troughcolor="#0b1220",
        bordercolor="#0b1220",
        lightcolor=BORDER_COLOR,
        darkcolor=BORDER_COLOR,
        arrowcolor="#64748b",
        relief="flat",
        arrowsize=10,
        width=12,
    )
    style.map(
        "Horizontal.TScrollbar",
        background=[("active", "#475569"), ("pressed", ACCENT_COLOR)],
        arrowcolor=[("active", "#cbd5e1"), ("pressed", "#ffffff")],
    )

    # Prefer a clean modern thumb-on-track look when the theme supports custom layouts.
    try:
        style.layout(
            "Vertical.TScrollbar",
            [
                (
                    "Vertical.Scrollbar.trough",
                    {
                        "sticky": "ns",
                        "children": [
                            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})
                        ],
                    },
                )
            ],
        )
        style.layout(
            "Horizontal.TScrollbar",
            [
                (
                    "Horizontal.Scrollbar.trough",
                    {
                        "sticky": "we",
                        "children": [
                            ("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})
                        ],
                    },
                )
            ],
        )
    except tk.TclError:
        pass

    window.option_add("*TCombobox*Listbox.foreground", TEXT_COLOR)
    window.option_add("*TCombobox*Listbox.background", SURFACE_COLOR)
    window.option_add("*TCombobox*Listbox.selectBackground", ACCENT_COLOR)
    window.option_add("*TCombobox*Listbox.selectForeground", "white")
    window.option_add("*TCombobox*Listbox.highlightColor", FOCUS_RING_COLOR)
    window.option_add("*TCombobox*Listbox.highlightBackground", LISTBOX_BORDER)


def _hex_color_to_rgba(hex_color, alpha=255):
    color = (hex_color or "").strip().lstrip("#")
    if len(color) == 3:
        color = "".join(ch * 2 for ch in color)
    if len(color) != 6:
        return 255, 255, 255, alpha
    try:
        return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16), alpha
    except ValueError:
        return 255, 255, 255, alpha


def _get_pillow_lanczos_filter():
    if Image is None:
        return None
    resampling = getattr(Image, "Resampling", None)
    if resampling is not None:
        return resampling.LANCZOS
    return getattr(Image, "LANCZOS", getattr(Image, "BICUBIC"))


def _create_toolbar_icon_image(icon_name, size=TOOLBAR_ICON_SIZE, color=TOOLBAR_ICON_COLOR):
    if Image is None or ImageDraw is None or ImageTk is None:
        return None

    scale = max(2, int(TOOLBAR_ICON_RENDER_SCALE))
    canvas_size = max(16, int(size)) * scale
    stroke = max(2, int(TOOLBAR_ICON_STROKE_MULTIPLIER * scale))
    padding = int(3.0 * scale)
    rgba = _hex_color_to_rgba(color)
    center = canvas_size // 2

    icon_canvas = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(icon_canvas)

    def _rounded_outline(bounds, radius):
        if hasattr(draw, "rounded_rectangle"):
            draw.rounded_rectangle(bounds, radius=radius, outline=rgba, width=stroke)
        else:
            draw.rectangle(bounds, outline=rgba, width=stroke)

    if icon_name == "refresh":
        outer = (padding, padding, canvas_size - padding, canvas_size - padding)
        draw.arc(
            outer,
            start=40,
            end=328,
            fill=rgba,
            width=stroke,
        )

        radius = (canvas_size - (2 * padding)) // 2
        arrow_angle = math.radians(328)
        arrow_tip_x = center + int(radius * math.cos(arrow_angle))
        arrow_tip_y = center + int(radius * math.sin(arrow_angle))
        draw.polygon(
            [
                (arrow_tip_x, arrow_tip_y),
                (arrow_tip_x - int(5.3 * scale), arrow_tip_y - int(0.8 * scale)),
                (arrow_tip_x - int(2.1 * scale), arrow_tip_y + int(4.1 * scale)),
            ],
            fill=rgba,
        )
    elif icon_name == "preview":
        eye_top = int(canvas_size * 0.31)
        eye_bottom = int(canvas_size * 0.69)
        eye_left = int(canvas_size * 0.13)
        eye_right = int(canvas_size * 0.87)
        draw.ellipse(
            (eye_left, eye_top, eye_right, eye_bottom),
            outline=rgba,
            width=stroke,
        )
        iris_radius = int(2.7 * scale)
        pupil_radius = int(1.5 * scale)
        draw.ellipse(
            (
                center - iris_radius,
                center - iris_radius,
                center + iris_radius,
                center + iris_radius,
            ),
            outline=rgba,
            width=max(2, stroke - int(0.7 * scale)),
        )
        draw.ellipse(
            (
                center - pupil_radius,
                center - pupil_radius,
                center + pupil_radius,
                center + pupil_radius,
            ),
            fill=rgba,
        )
    elif icon_name == "select_all":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
        draw.line(
            [
                (box_left + int(2.5 * scale), center + int(0.9 * scale)),
                (box_left + int(6.1 * scale), box_bottom - int(2.8 * scale)),
                (box_right - int(1.9 * scale), box_top + int(3.3 * scale)),
            ],
            fill=rgba,
            width=stroke,
        )
    elif icon_name == "select_none":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
    elif icon_name == "select_partial":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
        mid_y = (box_top + box_bottom) // 2
        draw.line(
            (
                box_left + int(2.8 * scale),
                mid_y,
                box_right - int(2.8 * scale),
                mid_y,
            ),
            fill=rgba,
            width=stroke,
        )
    elif icon_name == "source_add":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
        mid_x = (box_left + box_right) // 2
        mid_y = (box_top + box_bottom) // 2
        draw.line(
            (
                mid_x,
                box_top + int(2.8 * scale),
                mid_x,
                box_bottom - int(2.8 * scale),
            ),
            fill=rgba,
            width=stroke,
        )
        draw.line(
            (
                box_left + int(2.8 * scale),
                mid_y,
                box_right - int(2.8 * scale),
                mid_y,
            ),
            fill=rgba,
            width=stroke,
        )
    elif icon_name == "source_remove":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
        mid_y = (box_top + box_bottom) // 2
        draw.line(
            (
                box_left + int(2.8 * scale),
                mid_y,
                box_right - int(2.8 * scale),
                mid_y,
            ),
            fill=rgba,
            width=stroke,
        )
    elif icon_name == "clear_selection":
        box_left = int(canvas_size * 0.19)
        box_top = int(canvas_size * 0.19)
        box_right = int(canvas_size * 0.81)
        box_bottom = int(canvas_size * 0.81)
        inset = int(3.8 * scale)
        _rounded_outline((box_left, box_top, box_right, box_bottom), radius=int(2.4 * scale))
        draw.line(
            (
                box_left + inset,
                box_top + inset,
                box_right - inset,
                box_bottom - inset,
            ),
            fill=rgba,
            width=stroke,
        )
        draw.line(
            (
                box_right - inset,
                box_top + inset,
                box_left + inset,
                box_bottom - inset,
            ),
            fill=rgba,
            width=stroke,
        )
    elif icon_name == "edit":
        shaft_start = (int(canvas_size * 0.30), int(canvas_size * 0.70))
        shaft_end = (int(canvas_size * 0.70), int(canvas_size * 0.30))
        draw.line(
            (
                shaft_start[0],
                shaft_start[1],
                shaft_end[0],
                shaft_end[1],
            ),
            fill=rgba,
            width=stroke,
        )

        tip = [
            (int(canvas_size * 0.70), int(canvas_size * 0.30)),
            (int(canvas_size * 0.79), int(canvas_size * 0.22)),
            (int(canvas_size * 0.76), int(canvas_size * 0.35)),
        ]
        draw.polygon(tip, fill=rgba)

        eraser_bounds = (
            int(canvas_size * 0.22),
            int(canvas_size * 0.69),
            int(canvas_size * 0.33),
            int(canvas_size * 0.80),
        )
        draw.rectangle(eraser_bounds, outline=rgba, width=max(2, stroke - int(0.5 * scale)))

        baseline_y = int(canvas_size * 0.84)
        draw.line(
            (
                int(canvas_size * 0.18),
                baseline_y,
                int(canvas_size * 0.44),
                baseline_y,
            ),
            fill=rgba,
            width=max(2, stroke - int(0.7 * scale)),
        )
    else:
        return None

    resample_filter = _get_pillow_lanczos_filter()
    if resample_filter is not None:
        icon_canvas = icon_canvas.resize((size, size), resample_filter)
    else:
        icon_canvas = icon_canvas.resize((size, size))

    return ImageTk.PhotoImage(icon_canvas)


def _build_pending_toolbar_icon_images():
    icon_images = {}
    for icon_name in (
        "refresh",
        "preview",
        "edit",
        "select_all",
        "select_none",
        "select_partial",
        "source_add",
        "source_remove",
        "clear_selection",
    ):
        icon_image = _create_toolbar_icon_image(icon_name)
        if icon_image is not None:
            icon_images[icon_name] = icon_image
    return icon_images


def _get_display_work_area(window):
    screen_w = max(window.winfo_screenwidth(), 320)
    screen_h = max(window.winfo_screenheight(), 320)
    work_x = 0
    work_y = 0
    work_w = screen_w
    work_h = screen_h

    if sys.platform.startswith("win"):
        try:
            import ctypes

            class _Rect(ctypes.Structure):
                _fields_ = [
                    ("left", ctypes.c_long),
                    ("top", ctypes.c_long),
                    ("right", ctypes.c_long),
                    ("bottom", ctypes.c_long),
                ]

            rect = _Rect()
            spi_get_work_area = 0x0030
            if ctypes.windll.user32.SystemParametersInfoW(
                spi_get_work_area, 0, ctypes.byref(rect), 0
            ):
                width = rect.right - rect.left
                height = rect.bottom - rect.top
                if width > 0 and height > 0:
                    work_x = rect.left
                    work_y = rect.top
                    work_w = width
                    work_h = height
        except Exception:
            pass

    return work_x, work_y, max(work_w, 320), max(work_h, 320)


def configure_window_geometry(
    window,
    base_width,
    base_height,
    min_width=None,
    min_height=None,
    margin_x=DEFAULT_MARGIN_X,
    margin_y=DEFAULT_MARGIN_Y,
):
    """Size a window relative to the current display, leaving a bit of padding."""

    work_x, work_y, work_w, work_h = _get_display_work_area(window)

    usable_w = max(work_w - margin_x, int(work_w * 0.9))
    usable_h = max(work_h - margin_y, int(work_h * 0.9))

    width = min(base_width, usable_w)
    height = min(base_height, usable_h)

    min_width = base_width if min_width is None else min_width
    min_height = base_height if min_height is None else min_height

    if usable_w >= min_width:
        width = max(width, min_width)
    width = max(320, min(width, work_w - 10))

    if usable_h >= min_height:
        height = max(height, min_height)
    height = max(320, min(height, work_h - 10))

    x_pos = work_x + max(0, (work_w - int(width)) // 2)
    y_pos = work_y + max(0, (work_h - int(height)) // 2)

    window.geometry(f"{int(width)}x{int(height)}+{x_pos}+{y_pos}")
    window.minsize(int(min(width, min_width)), int(min(height, min_height)))


def _center_window_to_current_size(window):
    if not window.winfo_exists():
        return

    try:
        window.update_idletasks()
    except tk.TclError:
        return

    work_x, work_y, work_w, work_h = _get_display_work_area(window)

    width = max(window.winfo_width(), window.winfo_reqwidth(), 320)
    height = max(window.winfo_height(), window.winfo_reqheight(), 320)

    frame_x = max(0, window.winfo_rootx() - window.winfo_x())
    frame_top = max(0, window.winfo_rooty() - window.winfo_y())
    frame_bottom = max(frame_x, 4)

    if frame_x == 0 and frame_top == 0:
        if sys.platform.startswith("win"):
            frame_x = 8
            frame_top = 32
            frame_bottom = 8
        else:
            frame_x = 4
            frame_top = 28
            frame_bottom = 4

    safety_x = 8
    safety_y = 10

    max_client_w = max(320, work_w - (frame_x * 2) - safety_x)
    max_client_h = max(320, work_h - frame_top - frame_bottom - safety_y)

    try:
        min_w, min_h = window.minsize()
    except tk.TclError:
        min_w, min_h = (320, 320)

    target_min_w = min(max(320, int(min_w)), int(max_client_w))
    target_min_h = min(max(320, int(min_h)), int(max_client_h))
    if target_min_w != int(min_w) or target_min_h != int(min_h):
        window.minsize(target_min_w, target_min_h)

    width = max(width, target_min_w)
    height = max(height, target_min_h)

    width = min(int(width), int(max_client_w))
    height = min(int(height), int(max_client_h))

    outer_w = width + (frame_x * 2)
    outer_h = height + frame_top + frame_bottom

    x_pos = work_x + max(0, (work_w - outer_w) // 2)
    y_pos = work_y + max(0, (work_h - outer_h) // 2)

    x_pos = max(work_x, min(x_pos, work_x + max(0, work_w - outer_w)))
    y_pos = max(work_y, min(y_pos, work_y + max(0, work_h - outer_h)))

    window.geometry(f"{width}x{height}+{x_pos}+{y_pos}")


configure_window_geometry(
    root,
    BASE_WINDOW_WIDTH,
    BASE_WINDOW_HEIGHT,
    min_width=MIN_WINDOW_WIDTH,
    min_height=MIN_WINDOW_HEIGHT,
)

_apply_app_icon(root)

apply_theme(root)


class SystemTrayNotifier:
    def __init__(self):
        self._mode = None
        self._toaster = None
        self._failure_note = ""
        self._available = False
        self._init_backend()

    def _init_backend(self):
        if not sys.platform.startswith("win"):
            self._failure_note = "Tray alerts require Windows."
            return

        if WinotifyNotification is not None:
            self._mode = "winotify"
            self._available = True
            return

        supports_win10_toast = (
            Win10ToastNotifier is not None and sys.version_info < (3, 11)
        )
        if supports_win10_toast:
            try:
                self._toaster = Win10ToastNotifier()
                self._mode = "win10toast"
                self._available = True
                return
            except Exception as exc:
                self._failure_note = f"Tray alerts unavailable ({exc})."

        if not self._available:
            if WinotifyNotification is None:
                self._failure_note = "Install winotify (pip install winotify) for tray alerts."
            elif not supports_win10_toast:
                self._failure_note = "win10toast is unsupported on Python 3.11+; install winotify."

    def is_available(self):
        return self._available and self._mode is not None


    def notify(self, title, message):
        if not self.is_available():
            return False
        try:
            if self._mode == "winotify":
                notification = WinotifyNotification(
                    app_id="PDF AutoTool",
                    title=title,
                    msg=message,
                )
                notification.show()
            elif self._mode == "win10toast" and self._toaster is not None:
                self._toaster.show_toast(title, message, duration=5, threaded=True)
            else:
                return False
            return True
        except Exception as exc:
            self._available = False
            self._failure_note = f"Tray alerts failed: {exc}"
            return False

    def status_message(self):
        if self.is_available():
            backend = "Windows notifications" if self._mode == "winotify" else "win10toast"
            return f"Tray notifications active ({backend})."
        return self._failure_note or "Tray notifications unavailable."


def _mark_widget_as_scroll_canvas(widget):
    setattr(widget, "_scroll_canvas_enabled", True)


def _mark_widget_as_scroll_list(widget):
    setattr(widget, "_scroll_list_enabled", True)


def _find_scroll_canvas_for_widget(widget):
    current = widget
    while current is not None:
        if getattr(current, "_scroll_list_enabled", False):
            return current
        if getattr(current, "_scroll_canvas_enabled", False):
            return current
        parent_name = current.winfo_parent()
        if not parent_name:
            break
        try:
            current = current.nametowidget(parent_name)
        except Exception:
            break
    return None


_COMBOBOX_POPDOWN_SENTINEL = "__combobox_popdown__"


def _register_combobox_for_popdown_guard(combobox):
    tracked = getattr(root, "_combobox_popdown_guard_widgets", None)
    if tracked is None:
        tracked = []
        root._combobox_popdown_guard_widgets = tracked

    if combobox in tracked:
        return

    tracked.append(combobox)

    def _remove_combobox(_event=None):
        active = getattr(root, "_combobox_popdown_guard_widgets", None)
        if not active:
            return
        try:
            active.remove(combobox)
        except ValueError:
            pass

    combobox.bind("<Destroy>", _remove_combobox, add="+")


def _is_combobox_popdown_visible(combobox):
    try:
        popdown_name = str(
            combobox.tk.call("ttk::combobox::PopdownWindow", str(combobox))
        )
    except tk.TclError:
        return False

    if not popdown_name:
        return False

    try:
        return bool(int(combobox.tk.call("winfo", "ismapped", popdown_name)))
    except (tk.TclError, ValueError, TypeError):
        return False


def _any_combobox_popdown_visible():
    tracked = getattr(root, "_combobox_popdown_guard_widgets", None)
    if not tracked:
        return False

    live_widgets = []
    for widget in tracked:
        try:
            if not widget.winfo_exists():
                continue
        except Exception:
            continue

        live_widgets.append(widget)
        if _is_combobox_popdown_visible(widget):
            if len(live_widgets) != len(tracked):
                root._combobox_popdown_guard_widgets = live_widgets
            return True

    if len(live_widgets) != len(tracked):
        root._combobox_popdown_guard_widgets = live_widgets
    return False


def _hide_visible_combobox_popdowns():
    tracked = getattr(root, "_combobox_popdown_guard_widgets", None)
    if not tracked:
        return False

    closed_any = False
    live_widgets = []

    for widget in tracked:
        try:
            if not widget.winfo_exists():
                continue
        except Exception:
            continue

        live_widgets.append(widget)
        if not _is_combobox_popdown_visible(widget):
            continue

        closed = False
        try:
            widget.tk.call("ttk::combobox::Unpost", str(widget))
            closed = True
        except tk.TclError:
            pass

        if not closed:
            try:
                widget.event_generate("<Escape>")
                closed = True
            except tk.TclError:
                pass

        if closed:
            closed_any = True

    if len(live_widgets) != len(tracked):
        root._combobox_popdown_guard_widgets = live_widgets

    return closed_any


def _get_pointer_widget_safely(x_root, y_root):
    try:
        widget_name = root.tk.call("winfo", "containing", x_root, y_root)
    except tk.TclError:
        return None

    if not widget_name:
        return None

    widget_name = str(widget_name)

    # ttk Combobox dropdown widgets ("popdown") are not normal Tkinter children.
    # Let native handling consume wheel events there instead of forcing routing.
    if ".popdown" in widget_name or widget_name.endswith("popdown"):
        return _COMBOBOX_POPDOWN_SENTINEL

    try:
        return root.nametowidget(widget_name)
    except (KeyError, tk.TclError):
        return None


def _dispatch_global_mousewheel(event):
    if not root.winfo_exists():
        return None

    x_root = getattr(event, "x_root", None)
    y_root = getattr(event, "y_root", None)
    if x_root is None or y_root is None:
        x_root, y_root = root.winfo_pointerxy()

    target_widget = _get_pointer_widget_safely(x_root, y_root)

    if target_widget == _COMBOBOX_POPDOWN_SENTINEL:
        return None

    _hide_visible_combobox_popdowns()

    if target_widget is None:
        return None

    canvas = _find_scroll_canvas_for_widget(target_widget)
    if canvas is None or not canvas.winfo_exists():
        return None

    if hasattr(event, "delta") and event.delta:
        steps = int(-1 * (event.delta / 120))
        if steps != 0:
            canvas.yview_scroll(steps, "units")
            return "break"
        return None

    event_num = getattr(event, "num", None)
    if event_num == 4:
        canvas.yview_scroll(-1, "units")
        return "break"
    if event_num == 5:
        canvas.yview_scroll(1, "units")
        return "break"
    return None


def _ensure_global_mousewheel_binding():
    if getattr(root, "_global_mousewheel_bound", False):
        return
    root.bind_all("<MouseWheel>", _dispatch_global_mousewheel, add="+")
    root.bind_all("<Button-4>", _dispatch_global_mousewheel, add="+")
    root.bind_all("<Button-5>", _dispatch_global_mousewheel, add="+")
    root._global_mousewheel_bound = True


def _invoke_focused_widget_if_activatable(widget):
    if widget is None:
        return None

    try:
        if not widget.winfo_exists():
            return None
    except Exception:
        return None

    widget_class = str(widget.winfo_class()).lower()
    activatable_classes = {
        "button",
        "ttk::button",
        "checkbutton",
        "ttk::checkbutton",
        "radiobutton",
        "ttk::radiobutton",
    }

    if widget_class not in activatable_classes and not hasattr(widget, "invoke"):
        return None

    try:
        ttk_state = widget.state()
        if isinstance(ttk_state, (tuple, list, set)) and "disabled" in ttk_state:
            return None
    except Exception:
        pass

    try:
        if str(widget.cget("state")).lower() == "disabled":
            return None
    except Exception:
        pass

    try:
        widget.invoke()
        return "break"
    except Exception:
        return None


def _handle_global_enter_activation(event):
    focused_widget = root.focus_get()
    if focused_widget is None:
        focused_widget = getattr(event, "widget", None)
    return _invoke_focused_widget_if_activatable(focused_widget)


def _ensure_global_enter_activation_binding():
    if getattr(root, "_global_enter_activation_bound", False):
        return
    root.bind_all("<Return>", _handle_global_enter_activation, add="+")
    root.bind_all("<KP_Enter>", _handle_global_enter_activation, add="+")
    root._global_enter_activation_bound = True


def _apply_modern_listbox_style(listbox, *, compact=False, export_selection=None):
    font_size = 10 if compact else 10
    config = {
        "bg": LISTBOX_BG,
        "fg": LISTBOX_TEXT,
        "selectbackground": ACCENT_COLOR,
        "selectforeground": "white",
        "relief": "flat",
        "borderwidth": 0,
        "highlightthickness": 1,
        "highlightbackground": LISTBOX_BORDER,
        "highlightcolor": FOCUS_RING_COLOR,
        "activestyle": "none",
        "selectborderwidth": 0,
        "disabledforeground": SUBTEXT_COLOR,
        "font": ("Segoe UI", font_size),
    }
    if export_selection is not None:
        config["exportselection"] = bool(export_selection)
    listbox.configure(**config)


def create_scrollable_panel(parent):
    container = ttk.Frame(parent, style="TFrame")
    canvas = tk.Canvas(
        container,
        highlightthickness=0,
        bg=BG_COLOR,
        bd=0,
    )
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollable_frame = ttk.Frame(canvas, style="TFrame")

    def _update_scrollregion(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", _update_scrollregion)
    window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    def _resize_canvas(event):
        canvas.itemconfig(window_id, width=event.width)

    canvas.bind("<Configure>", _resize_canvas)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    _mark_widget_as_scroll_canvas(canvas)
    _ensure_global_mousewheel_binding()

    return container, scrollable_frame


class HoverTooltip:
    def __init__(self, widget, text, delay_ms=450):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self._after_id = None
        self._tip_window = None

        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")
        self.widget.bind("<Destroy>", self._on_destroy, add="+")

    def _on_enter(self, _event=None):
        self._schedule_show()

    def _on_leave(self, _event=None):
        self._cancel_show()
        self._hide()

    def _on_destroy(self, _event=None):
        self._on_leave()

    def _schedule_show(self):
        self._cancel_show()
        try:
            self._after_id = self.widget.after(self.delay_ms, self._show)
        except tk.TclError:
            self._after_id = None

    def _cancel_show(self):
        if self._after_id is None:
            return
        try:
            self.widget.after_cancel(self._after_id)
        except tk.TclError:
            pass
        self._after_id = None

    def _show(self):
        self._after_id = None
        if self._tip_window is not None:
            return
        if not self.widget.winfo_exists() or not self.text:
            return

        x_pos = self.widget.winfo_rootx() + 10
        y_pos = self.widget.winfo_rooty() + self.widget.winfo_height() + 8

        self._tip_window = tk.Toplevel(self.widget)
        self._tip_window.wm_overrideredirect(True)
        self._tip_window.geometry(f"+{x_pos}+{y_pos}")

        label = tk.Label(
            self._tip_window,
            text=self.text,
            bg="#111827",
            fg="#e2e8f0",
            relief="solid",
            bd=1,
            padx=8,
            pady=4,
            justify="left",
            anchor="w",
            font=("Segoe UI", 9),
        )
        label.pack()

    def _hide(self):
        if self._tip_window is None:
            return
        try:
            self._tip_window.destroy()
        except tk.TclError:
            pass
        self._tip_window = None


def _attach_hover_tooltip(widget, text):
    widget._hover_tooltip = HoverTooltip(widget, text)

# ------------------------
# Variables
# ------------------------
pending_folder = tk.StringVar()
root_folder = tk.StringVar()
auto_refresh_var = tk.BooleanVar(value=True)
tray_notifications_enabled_var = tk.BooleanVar(value=True)
keep_backup_preference_var = tk.BooleanVar(value=False)
show_text_with_icons_var = tk.BooleanVar(value=False)

employee_sources_listbox = None
employee_source_paths = []
employee_list_status_var = tk.StringVar(value="No employee sources selected.")
employee_sources_selection_count_var = tk.StringVar(value="Selected: 0/0")
employee_name_suggestions = []
pending_files_count_var = tk.StringVar(value="(0 selected / 0 total)")
name_filter_mode = tk.StringVar(value="strict")
_suppress_name_filter_refresh = False
rotation_preview_window_width = ROTATION_PREVIEW_DEFAULT_WIDTH
rotation_preview_window_height = ROTATION_PREVIEW_DEFAULT_HEIGHT
rotation_preview_thumb_width = ROTATION_PREVIEW_THUMB_MAX_WIDTH

tray_notifier = SystemTrayNotifier()
tray_status_var = tk.StringVar()
update_status_var = tk.StringVar(value="Updates unavailable")
_latest_prompted_update_version = ""


def _set_tray_status_message(custom=None):
    if custom is not None:
        tray_status_var.set(custom)
        return
    if tray_notifier.is_available():
        tray_status_var.set("Tray notifications active.")
    else:
        tray_status_var.set(tray_notifier.status_message())


_set_tray_status_message()
_ensure_global_enter_activation_binding()


menubar = tk.Menu(root)

file_menu = tk.Menu(menubar, tearoff=0)
preferences_menu = tk.Menu(file_menu, tearoff=0)
preferences_menu.add_checkbutton(
    label="Pending Files Auto-refresh",
    variable=auto_refresh_var,
    command=lambda: _on_auto_refresh_preference_changed(),
)
preferences_menu.add_checkbutton(
    label="Tray Notifications",
    variable=tray_notifications_enabled_var,
    command=lambda: _on_tray_notifications_preference_changed(),
)
preferences_menu.add_checkbutton(
    label="Merge: Keep timestamped backup",
    variable=keep_backup_preference_var,
    command=lambda: _on_keep_backup_preference_changed(),
)
preferences_menu.add_checkbutton(
    label="Show text with icons",
    variable=show_text_with_icons_var,
    command=lambda: _on_show_text_with_icons_preference_changed(),
)

name_filter_menu = tk.Menu(preferences_menu, tearoff=0)
name_filter_menu.add_radiobutton(
    label="Strict",
    variable=name_filter_mode,
    value="strict",
)
name_filter_menu.add_radiobutton(
    label="Lenient",
    variable=name_filter_mode,
    value="lenient",
)

preferences_menu.add_cascade(label="Name Filter Mode", menu=name_filter_menu)
file_menu.add_cascade(label="Preference", menu=preferences_menu)

application_menu = tk.Menu(file_menu, tearoff=0)
application_menu.add_command(label="Restart", command=lambda: restart_application())
application_menu.add_separator()
application_menu.add_command(label="Exit", command=lambda: exit_application())
file_menu.add_cascade(label="Application", menu=application_menu)

menubar.add_cascade(label="File", menu=file_menu)

help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="Check for Updates", command=lambda: check_for_updates(manual=True))
help_menu.add_command(label="About", command=lambda: show_about_dialog())

menubar.add_cascade(label="Help", menu=help_menu)
root.config(menu=menubar)


def _resolve_config_path():
    """Return a writable, user-specific settings path that also works when frozen."""

    if getattr(sys, "frozen", False):
        base_dir = (
            os.environ.get("APPDATA")
            or os.environ.get("LOCALAPPDATA")
            or os.path.expanduser("~")
        )
        config_dir = os.path.join(base_dir, "PDF_AutoTool")
    else:
        config_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "settings.json")


CONFIG_PATH = _resolve_config_path()


def normalize_path(path):
    return os.path.normpath(path) if path else ""


def _update_employee_list_status(message):
    employee_list_status_var.set(message)


def _normalize_candidate_line(raw_line):
    return " ".join(raw_line.strip().split())


def _is_header_row(normalized_line):
    lowered = normalized_line.lower().replace(":", " ").strip()
    tokens = {token for token in lowered.split() if token}
    if lowered in {"name", "employee name", "employee names"}:
        return True
    return {"name", "nickname", "position"}.issubset(tokens)


_TABLE_SPLIT_RE = re.compile(r"\t+|\s{2,}")
_NAME_TOKEN_RE = re.compile(r"^[A-Za-z][A-Za-z'.-]*$")
_PDF_NAME_START_PATTERN = re.compile(
    r"^\s*([A-Z][A-Za-z'\- ]+,\s+[A-Z][A-Za-z.'\- ]+?)(?=\s{2,}|\t|$)"
)
_NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}
_MAX_GIVEN_NAME_TOKENS = 10
_PDF_COLUMN_HEADER_LABELS = {
    "name",
    "employee name",
    "employee names",
    "nickname",
    "position",
    "admin asst local",
    "direct",
    "division",
    "photo",
}


def _split_table_columns(raw_line):
    prepared = (raw_line or "").replace("\u00a0", " ").strip()
    if not prepared:
        return []
    return [segment.strip() for segment in _TABLE_SPLIT_RE.split(prepared) if segment.strip()]


def _extract_pdf_names_from_text(text):
    candidates = []
    for raw_line in (text or "").splitlines():
        raw = (raw_line or "").replace("\u00a0", " ").strip()
        if not raw:
            continue

        normalized = _normalize_candidate_line(raw)
        if _is_header_row(normalized):
            continue
        if _canonical_header_label(normalized) in _PDF_COLUMN_HEADER_LABELS:
            continue

        match = _PDF_NAME_START_PATTERN.match(raw)
        if match:
            base_candidate = _normalize_candidate_line(match.group(1))
        else:
            columns = _split_table_columns(raw)
            if columns:
                base_candidate = _normalize_candidate_line(columns[0])
            else:
                base_candidate = normalized

        if "," in base_candidate:
            candidate = _extract_comma_prefix_candidate(
                base_candidate,
                max_given_tokens=_MAX_GIVEN_NAME_TOKENS,
            )
        else:
            candidate = base_candidate

        if _line_passes_filter(candidate, allow_mononym=True):
            candidates.append(candidate)

    return candidates


def _extract_pdf_names_from_tables(page):
    extracted = []
    try:
        tables = page.extract_tables()
    except Exception:
        return extracted

    for table in tables or []:
        if not table:
            continue

        header_row_index = None
        name_column_index = None

        scan_limit = min(len(table), 5)
        for row_index in range(scan_limit):
            row = table[row_index] or []
            for column_index, cell in enumerate(row):
                text = _normalize_candidate_line("" if cell is None else str(cell))
                if _canonical_header_label(text) == "name":
                    header_row_index = row_index
                    name_column_index = column_index
                    break
            if name_column_index is not None:
                break

        if name_column_index is None:
            continue

        for row in table[header_row_index + 1 :]:
            if name_column_index >= len(row):
                continue
            raw_value = "" if row[name_column_index] is None else str(row[name_column_index])
            normalized = _normalize_candidate_line(raw_value)
            if not normalized or _is_header_row(normalized):
                continue

            if "," in normalized:
                candidate = _extract_comma_prefix_candidate(
                    normalized,
                    max_given_tokens=_MAX_GIVEN_NAME_TOKENS,
                )
            else:
                candidate = normalized

            if _line_passes_filter(candidate, allow_mononym=True):
                extracted.append(candidate)

    return extracted


def _read_pdf_name_lines(path):
    extracted = []

    if pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf_file:
                for page in pdf_file.pages:
                    table_names = _extract_pdf_names_from_tables(page)
                    if table_names:
                        extracted.extend(table_names)
                        continue

                    page_text = page.extract_text(layout=True) or page.extract_text() or ""
                    extracted.extend(_extract_pdf_names_from_text(page_text))
        except Exception as exc:
            raise RuntimeError(str(exc)) from exc

        if extracted:
            return list(dict.fromkeys(extracted))

    if PdfReader is None:
        raise RuntimeError("Install pdfplumber (recommended) or pypdf for PDF sources.")

    try:
        reader = PdfReader(path)
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc

    for page in getattr(reader, "pages", []):
        try:
            page_text = page.extract_text() or ""
        except Exception:
            continue
        extracted.extend(_extract_pdf_names_from_text(page_text))

    if extracted:
        return list(dict.fromkeys(extracted))

    fallback_bucket = set()
    for page in getattr(reader, "pages", []):
        try:
            page_text = page.extract_text() or ""
        except Exception:
            continue
        _collect_lines_from_text(page_text, fallback_bucket)

    return sorted(fallback_bucket, key=lambda value: value.lower())


def _clean_name_token(token):
    return re.sub(r"^[^A-Za-z]+|[^A-Za-z'.-]+$", "", token or "")


def _is_initial_token(token):
    core = (token or "").rstrip(".")
    return len(core) == 1 and core.isalpha()


def _extract_comma_prefix_candidate(raw_value, max_given_tokens=None):
    prepared = _normalize_candidate_line(raw_value)
    if not prepared or "," not in prepared:
        return ""

    surname_part, remainder = prepared.split(",", 1)
    surname = _normalize_candidate_line(surname_part)
    if not surname or not any(ch.isalpha() for ch in surname):
        return ""
    if any(ch.isdigit() for ch in surname):
        return ""

    given_tokens = []
    for raw_token in remainder.strip().split():
        token = _clean_name_token(raw_token)
        if not token:
            if given_tokens:
                break
            continue
        if not _NAME_TOKEN_RE.fullmatch(token):
            if given_tokens:
                break
            continue

        # In many table exports the nickname column repeats a prior name token.
        if given_tokens and token.lower() in {existing.lower() for existing in given_tokens}:
            break

        # Stop before compact uppercase codes from following columns.
        if (
            given_tokens
            and token.isupper()
            and len(token) in (2, 3)
            and token.lower() not in _NAME_SUFFIXES
        ):
            break

        given_tokens.append(token)

        if max_given_tokens and len(given_tokens) >= max_given_tokens:
            break
        if len(given_tokens) >= _MAX_GIVEN_NAME_TOKENS:
            break

    if not given_tokens:
        return ""
    return _normalize_candidate_line(f"{surname}, {' '.join(given_tokens)}")


def _canonical_header_label(value):
    lowered = re.sub(r"[^a-z]+", " ", value.lower()).strip()
    normalized = " ".join(lowered.split())
    if normalized in {"name", "employee name", "employee names"}:
        return "name"
    return normalized


def _looks_like_header_value(value):
    if not value or "," in value:
        return False
    if any(ch.isdigit() for ch in value):
        return False
    words = value.split()
    if not words or len(words) > 5:
        return False
    return any(ch.isalpha() for ch in value)


def _normalize_source_lines(raw_lines):
    prepared = []
    for raw_line in raw_lines:
        normalized = _normalize_candidate_line((raw_line or "").replace("\u00a0", " "))
        if normalized:
            prepared.append(normalized)
    return prepared


def _find_name_header_block(lines):
    max_scan = min(len(lines), 120)
    for start in range(max_scan):
        if _canonical_header_label(lines[start]) != "name":
            continue

        headers = []
        end = start
        while end < len(lines) and len(headers) < 12:
            value = lines[end]
            if not _looks_like_header_value(value):
                break
            headers.append(_canonical_header_label(value))
            end += 1

        if len(headers) >= 2 and "name" in headers:
            return start, end, headers
    return None


def _extract_name_column_candidates(lines):
    header_block = _find_name_header_block(lines)
    if not header_block:
        return []

    _, start_data, headers = header_block
    column_count = len(headers)
    name_index = headers.index("name")

    candidates = []
    row = []
    idx = start_data
    while idx < len(lines):
        window = lines[idx : idx + column_count]
        if len(window) == column_count and [_canonical_header_label(v) for v in window] == headers:
            row = []
            idx += column_count
            continue

        row.append(lines[idx])
        if len(row) == column_count:
            candidate = _normalize_candidate_line(row[name_index])
            if _line_passes_filter(candidate, allow_mononym=True) and not _is_header_row(candidate):
                candidates.append(candidate)
            row = []
        idx += 1
    return candidates


def _extract_first_column_candidates(lines):
    candidates = []
    for line in lines:
        columns = _split_table_columns(line)
        if len(columns) < 2:
            continue
        candidate = _normalize_candidate_line(columns[0])
        if _line_passes_filter(candidate, allow_mononym=True) and not _is_header_row(candidate):
            candidates.append(candidate)
    return candidates


def _infer_given_name_token_limit(candidates):
    counts = []
    for candidate in candidates:
        if "," not in candidate:
            continue
        _, given = candidate.split(",", 1)
        token_count = len(given.strip().split())
        if token_count:
            counts.append(token_count)

    if not counts:
        return min(6, _MAX_GIVEN_NAME_TOKENS)
    counts.sort()
    percentile_index = int((len(counts) - 1) * 0.9)
    inferred = counts[percentile_index]
    return max(2, min(_MAX_GIVEN_NAME_TOKENS, inferred))


def _extract_comma_line_candidates(lines, max_given_tokens):
    candidates = []
    for line in lines:
        if _is_header_row(line):
            continue
        candidate = _extract_comma_prefix_candidate(line, max_given_tokens=max_given_tokens)
        if _line_passes_filter(candidate):
            candidates.append(candidate)
    return candidates


def _extract_single_cell_candidates(lines):
    candidates = []
    for line in lines:
        if _is_header_row(line):
            continue
        candidate = _normalize_candidate_line(line)
        if _line_passes_filter(candidate, allow_mononym=True):
            candidates.append(candidate)
    return candidates


def _looks_like_mononym_name(normalized_line):
    value = _normalize_candidate_line(normalized_line)
    if not value or "," in value or " " in value:
        return False
    if len(value) < 2:
        return False
    if any(ch.isdigit() for ch in value):
        return False
    if not _NAME_TOKEN_RE.fullmatch(value):
        return False
    if value.isupper() and len(value) <= 3:
        return False
    return any(ch.isalpha() for ch in value)


def _line_passes_filter(normalized_line, allow_mononym=False):
    if not normalized_line:
        return False
    mode = name_filter_mode.get()
    if mode == "strict":
        if "," in normalized_line and any(ch.isalpha() for ch in normalized_line):
            return True
        if allow_mononym and _looks_like_mononym_name(normalized_line):
            return True
        return False
    return any(ch.isalpha() for ch in normalized_line) and len(normalized_line) >= 3


def _collect_extracted_candidates(raw_lines, bucket):
    lines = _normalize_source_lines(raw_lines)
    if not lines:
        return

    name_column_candidates = _extract_name_column_candidates(lines)
    first_column_candidates = _extract_first_column_candidates(lines)
    inferred_limit = _infer_given_name_token_limit(name_column_candidates + first_column_candidates)
    comma_line_candidates = _extract_comma_line_candidates(lines, max_given_tokens=inferred_limit)

    combined_candidates = (
        name_column_candidates
        + first_column_candidates
        + comma_line_candidates
    )

    if not combined_candidates:
        combined_candidates = _extract_single_cell_candidates(lines)

    for candidate in combined_candidates:
        if _line_passes_filter(candidate, allow_mononym=True):
            bucket.add(candidate)


def _collect_lines_from_text(text, bucket):
    _collect_extracted_candidates(text.splitlines(), bucket)


def _collect_lines_from_iterable(lines, bucket):
    _collect_extracted_candidates(list(lines), bucket)


def _read_text_file_lines(path):
    encodings = ("utf-8-sig", "utf-16", "latin-1")
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as file:
                return file.readlines()
        except UnicodeDecodeError:
            continue
        except OSError:
            raise
    with open(path, "r", errors="ignore") as file:
        return file.readlines()


def _read_excel_name_lines(path):
    if load_workbook is None:
        raise RuntimeError("Install openpyxl (pip install openpyxl) for Excel sources.")

    extension = os.path.splitext(path)[1].lower()
    if extension == ".xls":
        raise RuntimeError("Legacy .xls is not supported; save as .xlsx/.xlsm and re-add it.")

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc

    extracted = []
    try:
        for sheet in workbook.worksheets:
            name_column_index = None

            for row in sheet.iter_rows(values_only=True):
                values = []
                for cell in row:
                    text = "" if cell is None else str(cell)
                    values.append(_normalize_candidate_line(text))

                if not any(values):
                    continue

                if name_column_index is None:
                    for idx, value in enumerate(values):
                        if _canonical_header_label(value) == "name":
                            name_column_index = idx
                            break
                    if name_column_index is not None:
                        continue

                if name_column_index is not None and name_column_index < len(values):
                    candidate = values[name_column_index]
                else:
                    candidate = next((value for value in values if value), "")

                if candidate and not _is_header_row(candidate):
                    extracted.append(candidate)
    finally:
        workbook.close()

    return extracted


def load_employee_name_suggestions(progress_callback=None):
    global employee_name_suggestions
    employee_name_suggestions = []

    if not employee_source_paths:
        _update_employee_list_status("No employee sources selected.")
        if progress_callback is not None:
            progress_callback("No employee sources selected.", 1, 1)
        return

    suggestions = set()
    errors = []
    total_sources = len(employee_source_paths)

    if progress_callback is not None:
        progress_callback("Reading employee name sources...", 0, total_sources)

    for index, path in enumerate(employee_source_paths, start=1):
        source_name = os.path.basename(path)
        if progress_callback is not None:
            progress_callback(f"Parsing source {index}/{total_sources}: {source_name}", index - 1, total_sources)

        extension = os.path.splitext(path)[1].lower()
        if extension == ".pdf":
            try:
                pdf_name_lines = _read_pdf_name_lines(path)
            except RuntimeError as exc:
                errors.append(f"PDF load failed for {source_name}: {exc}")
                if progress_callback is not None:
                    progress_callback(f"Processed source {index}/{total_sources}: {source_name}", index, total_sources)
                continue
            for candidate in pdf_name_lines:
                if _line_passes_filter(candidate, allow_mononym=True):
                    suggestions.add(candidate)
        elif extension in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"):
            try:
                excel_name_lines = _read_excel_name_lines(path)
            except RuntimeError as exc:
                errors.append(f"Excel load failed for {source_name}: {exc}")
                if progress_callback is not None:
                    progress_callback(f"Processed source {index}/{total_sources}: {source_name}", index, total_sources)
                continue
            _collect_lines_from_iterable(excel_name_lines, suggestions)
        elif extension in (".txt", ".csv"):
            try:
                lines = _read_text_file_lines(path)
            except OSError as exc:
                errors.append(f"Unable to read {source_name}: {exc}")
                if progress_callback is not None:
                    progress_callback(f"Processed source {index}/{total_sources}: {source_name}", index, total_sources)
                continue
            _collect_lines_from_iterable(lines, suggestions)
        else:
            errors.append(f"Unsupported file type: {source_name}")

        if progress_callback is not None:
            progress_callback(f"Processed source {index}/{total_sources}: {source_name}", index, total_sources)

    employee_name_suggestions = sorted(suggestions, key=lambda value: value.lower())

    if employee_name_suggestions:
        msg = f"{len(employee_name_suggestions)} names loaded from {len(employee_source_paths)} source(s)."
    else:
        msg = "No names detected in selected sources."

    if errors:
        unique_errors = list(dict.fromkeys(errors))
        msg += " Issues: " + "; ".join(unique_errors[:3])
        if len(unique_errors) > 3:
            msg += " ..."

    filter_label = "Strict" if name_filter_mode.get() == "strict" else "Lenient"
    msg += f" (Filter: {filter_label})."
    _update_employee_list_status(msg)

    if progress_callback is not None:
        progress_callback("Employee names loaded.", total_sources, total_sources)


def _refresh_employee_sources_listbox():
    if employee_sources_listbox is None:
        employee_sources_selection_count_var.set("Selected: 0/0")
        return
    employee_sources_listbox.delete(0, tk.END)
    for path in employee_source_paths:
        employee_sources_listbox.insert(tk.END, os.path.basename(path))
    _refresh_employee_sources_selection_count()


def _refresh_employee_sources_selection_count(_event=None):
    if employee_sources_listbox is None or not employee_sources_listbox.winfo_exists():
        employee_sources_selection_count_var.set("Selected: 0/0")
        return

    total_count = int(employee_sources_listbox.size())
    selected_count = len(employee_sources_listbox.curselection())
    employee_sources_selection_count_var.set(
        f"Selected: {selected_count}/{total_count}"
    )


def _set_employee_sources(paths, persist=False, progress_callback=None):
    global employee_source_paths
    normalized = []
    for raw_path in paths:
        path = normalize_path(raw_path)
        if path and path not in normalized:
            normalized.append(path)
    employee_source_paths = normalized
    _refresh_employee_sources_listbox()
    load_employee_name_suggestions(progress_callback=progress_callback)
    if persist:
        save_settings()


def _on_filter_mode_change(*_args):
    if _suppress_name_filter_refresh:
        return
    load_employee_name_suggestions()
    save_settings()


name_filter_mode.trace_add("write", _on_filter_mode_change)


def _normalize_name_for_search(value):
    return " ".join((value or "").lower().replace(",", " ").split())


def get_filtered_name_suggestions(prefix):
    if not employee_name_suggestions:
        return []
    normalized_query = _normalize_name_for_search(prefix)
    if not normalized_query:
        return employee_name_suggestions

    prefix_matches = []
    contains_matches = []

    for name in employee_name_suggestions:
        searchable = _normalize_name_for_search(name)
        if searchable.startswith(normalized_query):
            prefix_matches.append(name)
        elif normalized_query in searchable:
            contains_matches.append(name)

    return prefix_matches + contains_matches


def _get_suggestion_popup_state(combobox):
    state = getattr(combobox, "_suggestion_popup_state", None)
    if state is None:
        state = {"popup": None, "listbox": None, "tracker_after_id": None}
        setattr(combobox, "_suggestion_popup_state", state)
    return state


def _hide_suggestion_popup(combobox):
    state = _get_suggestion_popup_state(combobox)
    tracker_after_id = state.get("tracker_after_id")
    if tracker_after_id is not None:
        try:
            combobox.after_cancel(tracker_after_id)
        except tk.TclError:
            pass
    state["tracker_after_id"] = None

    popup = state.get("popup")
    if popup is not None and popup.winfo_exists():
        popup.destroy()
    state["popup"] = None
    state["listbox"] = None


def _scroll_list_widget_from_event(scroll_widget, event=None):
    if scroll_widget is None or not scroll_widget.winfo_exists():
        return "break"

    if event is None:
        return "break"

    if hasattr(event, "delta") and event.delta:
        steps = int(-1 * (event.delta / 120))
        if steps != 0:
            scroll_widget.yview_scroll(steps, "units")
        return "break"

    event_num = getattr(event, "num", None)
    if event_num == 4:
        scroll_widget.yview_scroll(-1, "units")
        return "break"
    if event_num == 5:
        scroll_widget.yview_scroll(1, "units")
        return "break"
    return "break"


def _position_suggestion_popup(combobox):
    state = _get_suggestion_popup_state(combobox)
    popup = state.get("popup")
    listbox = state.get("listbox")
    try:
        if popup is None or listbox is None:
            return False
        if not popup.winfo_exists() or not listbox.winfo_exists() or not combobox.winfo_exists():
            return False

        combobox.update_idletasks()
        popup.update_idletasks()

        visible_rows = max(1, int(listbox.cget("height")))
        popup_width = max(combobox.winfo_width(), 320)
        popup_height = min(240, (visible_rows * 22) + 6)

        x_pos = combobox.winfo_rootx()
        y_pos = combobox.winfo_rooty() + combobox.winfo_height()
        screen_w = combobox.winfo_screenwidth()
        screen_h = combobox.winfo_screenheight()

        if x_pos + popup_width > screen_w - 8:
            x_pos = max(0, screen_w - popup_width - 8)
        if y_pos + popup_height > screen_h - 8:
            y_pos = max(0, combobox.winfo_rooty() - popup_height)

        popup.geometry(f"{popup_width}x{popup_height}+{x_pos}+{y_pos}")
        popup.deiconify()
        popup.lift()
        return True
    except (tk.TclError, ValueError, TypeError):
        return False


def _start_suggestion_popup_tracking(combobox, interval_ms=40):
    state = _get_suggestion_popup_state(combobox)

    existing_after_id = state.get("tracker_after_id")
    if existing_after_id is not None:
        try:
            combobox.after_cancel(existing_after_id)
        except tk.TclError:
            pass
        state["tracker_after_id"] = None

    def _track_position():
        try:
            if not combobox.winfo_exists():
                state["tracker_after_id"] = None
                return
        except tk.TclError:
            state["tracker_after_id"] = None
            return

        if not _position_suggestion_popup(combobox):
            state["tracker_after_id"] = None
            return
        try:
            state["tracker_after_id"] = combobox.after(interval_ms, _track_position)
        except tk.TclError:
            state["tracker_after_id"] = None

    try:
        state["tracker_after_id"] = combobox.after(interval_ms, _track_position)
    except tk.TclError:
        state["tracker_after_id"] = None


def _select_suggestion_from_popup(combobox, _event=None):
    state = _get_suggestion_popup_state(combobox)
    listbox = state.get("listbox")
    if listbox is None or not listbox.winfo_exists():
        return "break"

    selection = ()
    if _event is not None and hasattr(_event, "y") and listbox.size() > 0:
        try:
            clicked_index = listbox.nearest(_event.y)
        except tk.TclError:
            clicked_index = -1
        if 0 <= clicked_index < listbox.size():
            selection = (clicked_index,)

    if not selection:
        selection = listbox.curselection()

    if selection:
        chosen_value = listbox.get(selection[0])
    elif listbox.size() > 0:
        chosen_value = listbox.get(0)
    else:
        _hide_suggestion_popup(combobox)
        return "break"

    combobox.set(chosen_value)
    combobox.icursor(tk.END)

    try:
        combobox.event_generate("<<ComboboxSelected>>")
    except tk.TclError:
        pass

    selection_callback = getattr(combobox, "_on_suggestion_selected", None)
    if callable(selection_callback):
        try:
            selection_callback(chosen_value)
        except Exception:
            pass

    _hide_suggestion_popup(combobox)
    combobox.focus_set()
    return "break"


def _schedule_suggestion_popup_close(combobox, delay_ms=140):
    def _close_if_focus_left():
        state = _get_suggestion_popup_state(combobox)
        popup = state.get("popup")
        listbox = state.get("listbox")
        try:
            if popup is None or not popup.winfo_exists():
                return
        except tk.TclError:
            return

        try:
            focus_name = str(combobox.tk.call("focus") or "")
        except tk.TclError:
            focus_name = ""

        # ttk Combobox popdown is a detached Tcl widget; keep the popup while it owns focus.
        if focus_name and (".popdown" in focus_name or focus_name.endswith("popdown")):
            return

        try:
            focus_widget = combobox.focus_get()
        except (tk.TclError, KeyError):
            focus_widget = None

        if focus_widget is combobox:
            return
        if listbox is not None and focus_widget is listbox:
            return
        if focus_widget is not None and str(focus_widget).startswith(str(popup)):
            return

        _hide_suggestion_popup(combobox)

    try:
        combobox.after(delay_ms, _close_if_focus_left)
    except tk.TclError:
        pass


def _focus_suggestion_popup_list(combobox):
    state = _get_suggestion_popup_state(combobox)
    listbox = state.get("listbox")
    if listbox is None or not listbox.winfo_exists() or listbox.size() == 0:
        return "break"

    selection = listbox.curselection()
    target_index = selection[0] if selection else 0
    listbox.selection_clear(0, tk.END)
    listbox.selection_set(target_index)
    listbox.activate(target_index)
    listbox.focus_set()
    return "break"


def _show_suggestion_popup(combobox, suggestions):
    if not suggestions:
        _hide_suggestion_popup(combobox)
        return

    state = _get_suggestion_popup_state(combobox)
    popup = state.get("popup")
    listbox = state.get("listbox")

    if popup is None or not popup.winfo_exists() or listbox is None or not listbox.winfo_exists():
        popup = tk.Toplevel(combobox)
        popup.wm_overrideredirect(True)
        popup.transient(combobox.winfo_toplevel())
        popup.attributes("-topmost", True)

        list_container = ttk.Frame(popup, style="Card.TFrame")
        list_container.pack(fill="both", expand=True)

        listbox = tk.Listbox(list_container)
        _apply_modern_listbox_style(listbox, compact=True, export_selection=False)
        _mark_widget_as_scroll_list(listbox)
        popup_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=popup_scrollbar.set)

        listbox.pack(side="left", fill="both", expand=True)
        popup_scrollbar.pack(side="right", fill="y")

        def _scroll_suggestion_list(event=None):
            return _scroll_list_widget_from_event(listbox, event)

        popup.bind("<MouseWheel>", _scroll_suggestion_list, add="+")
        popup.bind("<Button-4>", _scroll_suggestion_list, add="+")
        popup.bind("<Button-5>", _scroll_suggestion_list, add="+")
        list_container.bind("<MouseWheel>", _scroll_suggestion_list, add="+")
        list_container.bind("<Button-4>", _scroll_suggestion_list, add="+")
        list_container.bind("<Button-5>", _scroll_suggestion_list, add="+")
        listbox.bind("<MouseWheel>", _scroll_suggestion_list, add="+")
        listbox.bind("<Button-4>", _scroll_suggestion_list, add="+")
        listbox.bind("<Button-5>", _scroll_suggestion_list, add="+")
        popup_scrollbar.bind("<MouseWheel>", _scroll_suggestion_list, add="+")
        popup_scrollbar.bind("<Button-4>", _scroll_suggestion_list, add="+")
        popup_scrollbar.bind("<Button-5>", _scroll_suggestion_list, add="+")

        listbox.bind("<Button-1>", lambda event: _select_suggestion_from_popup(combobox, event))
        listbox.bind("<Double-Button-1>", lambda event: _select_suggestion_from_popup(combobox, event))
        listbox.bind("<Return>", lambda event: _select_suggestion_from_popup(combobox, event))
        listbox.bind("<Escape>", lambda _event: (_hide_suggestion_popup(combobox), combobox.focus_set(), "break")[2])
        listbox.bind("<FocusOut>", lambda _event: _schedule_suggestion_popup_close(combobox), add="+")

        state["popup"] = popup
        state["listbox"] = listbox

        if not getattr(combobox, "_suggestion_popup_bindings_set", False):
            combobox.bind("<Destroy>", lambda _event: _hide_suggestion_popup(combobox), add="+")
            combobox.bind("<FocusOut>", lambda _event: _schedule_suggestion_popup_close(combobox), add="+")
            combobox.bind("<Down>", lambda _event: _focus_suggestion_popup_list(combobox), add="+")
            combobox.bind("<Escape>", lambda _event: (_hide_suggestion_popup(combobox), "break")[1], add="+")
            setattr(combobox, "_suggestion_popup_bindings_set", True)

    limited_suggestions = suggestions[:200]
    listbox.delete(0, tk.END)
    for suggestion in limited_suggestions:
        listbox.insert(tk.END, suggestion)

    listbox.selection_clear(0, tk.END)
    if listbox.size() > 0:
        listbox.selection_set(0)
        listbox.activate(0)

    visible_rows = min(8, max(1, len(limited_suggestions)))
    listbox.configure(height=visible_rows)

    _position_suggestion_popup(combobox)
    _start_suggestion_popup_tracking(combobox)


def _update_combobox_suggestions(combobox, query_text, event=None):
    suggestions = get_filtered_name_suggestions(query_text)
    combobox["values"] = suggestions

    keysym = getattr(event, "keysym", "") if event is not None else ""
    navigation_keys = {"Up", "Down", "Prior", "Next", "Return", "Tab", "Escape"}
    if keysym in navigation_keys:
        if keysym in {"Return", "Tab", "Escape"}:
            _hide_suggestion_popup(combobox)
        return

    if query_text.strip() and suggestions:
        combobox.after_idle(lambda: _show_suggestion_popup(combobox, suggestions))
    else:
        combobox.after_idle(lambda: _hide_suggestion_popup(combobox))


def _prevent_combobox_mousewheel_value_change(combobox):
    _register_combobox_for_popdown_guard(combobox)

    def _on_wheel(event=None):
        state = _get_suggestion_popup_state(combobox)
        listbox = state.get("listbox")
        if listbox is not None and listbox.winfo_exists():
            return _scroll_list_widget_from_event(listbox, event)

        if event is not None:
            _dispatch_global_mousewheel(event)

        return "break"

    combobox.bind("<MouseWheel>", _on_wheel, add="+")
    combobox.bind("<Button-4>", _on_wheel, add="+")
    combobox.bind("<Button-5>", _on_wheel, add="+")


def _set_update_status(message):
    update_status_var.set(message)


def _refresh_update_status():
    if DEFAULT_UPDATE_MANIFEST_URL.strip():
        _set_update_status("Update service available")
    else:
        _set_update_status("Updates unavailable")


def _normalize_version_tuple(version_text):
    numbers = [int(chunk) for chunk in re.findall(r"\d+", str(version_text))]
    if not numbers:
        return (0,)
    while len(numbers) > 1 and numbers[-1] == 0:
        numbers.pop()
    return tuple(numbers)


def _is_newer_version(candidate_version, current_version):
    candidate = list(_normalize_version_tuple(candidate_version))
    current = list(_normalize_version_tuple(current_version))
    width = max(len(candidate), len(current))
    candidate.extend([0] * (width - len(candidate)))
    current.extend([0] * (width - len(current)))
    return tuple(candidate) > tuple(current)


def _get_installation_scope():
    if not getattr(sys, "frozen", False):
        return "source"

    executable_dir = os.path.abspath(os.path.dirname(sys.executable)).lower()
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    if local_appdata and executable_dir.startswith(os.path.abspath(local_appdata).lower()):
        return "user setup"

    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base_path = os.environ.get(env_name, "")
        if base_path and executable_dir.startswith(os.path.abspath(base_path).lower()):
            return "system setup"

    return "portable"


_about_build_metadata_cache = None


def _normalize_build_metadata_value(value):
    text = str(value or "").strip()
    if not text:
        return "unknown"
    if text.lower() in {"unknown", "none", "null", "n/a"}:
        return "unknown"
    return text


def _candidate_build_info_paths():
    paths = []

    if getattr(sys, "frozen", False):
        bundle_dir = getattr(sys, "_MEIPASS", None)
        if bundle_dir:
            paths.append(os.path.join(bundle_dir, APP_BUILD_INFO_FILENAME))
        paths.append(os.path.join(os.path.dirname(sys.executable), APP_BUILD_INFO_FILENAME))

    paths.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), APP_BUILD_INFO_FILENAME))

    deduped_paths = []
    seen_paths = set()
    for candidate_path in paths:
        normalized_path = os.path.normcase(os.path.normpath(candidate_path))
        if normalized_path in seen_paths:
            continue
        seen_paths.add(normalized_path)
        deduped_paths.append(candidate_path)

    return deduped_paths


def _read_build_info_file():
    for info_path in _candidate_build_info_paths():
        if not os.path.isfile(info_path):
            continue

        try:
            with open(info_path, "r", encoding="utf-8") as info_file:
                data = json.load(info_file)
            if isinstance(data, dict):
                return data
        except Exception:
            continue

    return {}


def _run_git_text_command(command):
    kwargs = {
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
        "check": False,
        "cwd": os.path.dirname(os.path.abspath(__file__)),
    }
    create_no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if create_no_window:
        kwargs["creationflags"] = create_no_window

    try:
        result = subprocess.run(command, **kwargs)
    except Exception:
        return ""

    if result.returncode != 0:
        return ""
    return str(result.stdout or "").strip()


def _resolve_about_build_metadata():
    global _about_build_metadata_cache
    if _about_build_metadata_cache is not None:
        return _about_build_metadata_cache

    commit_value = _normalize_build_metadata_value(APP_BUILD_COMMIT)
    date_value = _normalize_build_metadata_value(APP_BUILD_DATE)

    if commit_value == "unknown" or date_value == "unknown":
        build_info = _read_build_info_file()
        if commit_value == "unknown":
            commit_value = _normalize_build_metadata_value(
                build_info.get("commit") or build_info.get("build_commit")
            )
        if date_value == "unknown":
            date_value = _normalize_build_metadata_value(
                build_info.get("build_date") or build_info.get("date")
            )

    if not getattr(sys, "frozen", False):
        if commit_value == "unknown":
            commit_value = _normalize_build_metadata_value(
                _run_git_text_command(["git", "rev-parse", "--short=12", "HEAD"])
            )
        if date_value == "unknown":
            date_value = _normalize_build_metadata_value(
                _run_git_text_command(["git", "show", "-s", "--format=%cI", "HEAD"])
            )

    _about_build_metadata_cache = {
        "commit": commit_value,
        "build_date": date_value,
    }
    return _about_build_metadata_cache


def _get_about_date_text():
    metadata = _resolve_about_build_metadata()
    resolved_date = metadata.get("build_date", "unknown")
    if resolved_date != "unknown":
        return resolved_date

    target_path = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    try:
        modified_dt = datetime.fromtimestamp(os.path.getmtime(target_path)).astimezone()
        return modified_dt.isoformat(timespec="seconds")
    except Exception:
        return "unknown"


def show_about_dialog():
    build_metadata = _resolve_about_build_metadata()
    os_name = os.environ.get("OS") or platform.system() or "unknown"
    architecture = platform.machine() or "unknown"
    os_version = platform.version() or "unknown"
    update_feed_value = DEFAULT_UPDATE_MANIFEST_URL.strip() or "Unavailable in this build"

    details = [
        "App: PDF Record Manager",
        f"Version: {APP_VERSION}",
        f"Install Type: {_get_installation_scope()}",
        f"Build Commit: {build_metadata.get('commit', 'unknown')}",
        f"Build Date: {_get_about_date_text()}",
        f"Update Feed: {update_feed_value}",
        f"Python Runtime: {platform.python_version()}",
        f"OS: {os_name} {architecture} {os_version}",
    ]

    messagebox.showinfo("About PDF Record Manager", "\n".join(details), parent=root)


def _download_update_manifest(manifest_url):
    request = urllib.request.Request(
        manifest_url,
        headers={"User-Agent": f"PDFRecordManager/{APP_VERSION}"},
    )
    with urllib.request.urlopen(request, timeout=UPDATE_CHECK_TIMEOUT_SEC) as response:
        payload = response.read()
        encoding = response.headers.get_content_charset() or "utf-8"

    manifest = json.loads(payload.decode(encoding, errors="replace"))
    if not isinstance(manifest, dict):
        raise RuntimeError("Update feed must return a JSON object.")
    return manifest


def _handle_update_manifest(manual, manifest):
    global _latest_prompted_update_version

    latest_version = str(manifest.get("version", "")).strip()
    installer_url = str(manifest.get("installer_url", "")).strip()
    portable_url = str(manifest.get("portable_url", "")).strip()
    release_page_url = str(manifest.get("release_page_url", "")).strip()
    notes = str(manifest.get("notes", "")).strip()

    if not latest_version:
        _set_update_status("Invalid update feed")
        if manual:
            messagebox.showwarning("Update Check", "Update feed is missing the 'version' field.")
        return

    if _is_newer_version(latest_version, APP_VERSION):
        _set_update_status("Update available")
        should_prompt = manual or latest_version != _latest_prompted_update_version
        if not should_prompt:
            return

        _latest_prompted_update_version = latest_version
        message_lines = [
            f"Current version: {APP_VERSION}",
            f"Latest version: {latest_version}",
        ]
        if notes:
            message_lines.extend(["", notes])

        if installer_url and portable_url:
            message_lines.extend(
                [
                    "",
                    "Choose update package:",
                    "Yes = Installer (recommended)",
                    "No = Portable package",
                    "Cancel = Later",
                ]
            )
            choice = messagebox.askyesnocancel("Update Available", "\n".join(message_lines))
            target_url = installer_url if choice is True else portable_url if choice is False else ""
            if target_url:
                try:
                    _launch_path(target_url)
                except RuntimeError as exc:
                    messagebox.showerror("Update Download", f"Unable to open update link: {exc}")
        elif installer_url:
            message_lines.extend(["", "Open installer download link now?"])
            if messagebox.askyesno("Update Available", "\n".join(message_lines)):
                try:
                    _launch_path(installer_url)
                except RuntimeError as exc:
                    messagebox.showerror("Update Download", f"Unable to open installer link: {exc}")
        elif portable_url:
            message_lines.extend(["", "Open portable download link now?"])
            if messagebox.askyesno("Update Available", "\n".join(message_lines)):
                try:
                    _launch_path(portable_url)
                except RuntimeError as exc:
                    messagebox.showerror("Update Download", f"Unable to open portable link: {exc}")
        elif release_page_url:
            message_lines.extend(["", "Open release page now?"])
            if messagebox.askyesno("Update Available", "\n".join(message_lines)):
                try:
                    _launch_path(release_page_url)
                except RuntimeError as exc:
                    messagebox.showerror("Update Download", f"Unable to open release page: {exc}")
        elif manual:
            messagebox.showinfo(
                "Update Available",
                "\n".join(
                    message_lines
                    + ["", "No installer_url, portable_url, or release_page_url was provided in the update feed."]
                ),
            )
    else:
        _set_update_status("Up to date")
        if manual:
            messagebox.showinfo("No Update", f"You are on the latest version ({APP_VERSION}).")


def _handle_update_check_error(manual, error_message):
    _set_update_status("Update check failed")
    if manual:
        messagebox.showwarning("Update Check", f"Unable to check for updates.\n\n{error_message}")


def check_for_updates(manual=False):
    manifest_url = DEFAULT_UPDATE_MANIFEST_URL.strip()

    if not manifest_url:
        _refresh_update_status()
        if manual:
            messagebox.showinfo(
                "Update Check",
                "Updates are unavailable in this build.",
            )
        return

    _set_update_status("Checking for updates...")

    def _worker():
        try:
            manifest = _download_update_manifest(manifest_url)
        except Exception as exc:
            root.after(0, lambda error_text=str(exc): _handle_update_check_error(manual, error_text))
            return
        root.after(0, lambda: _handle_update_manifest(manual, manifest))

    threading.Thread(target=_worker, daemon=True).start()


def load_settings(progress_callback=None):
    global _suppress_name_filter_refresh
    global rotation_preview_window_width, rotation_preview_window_height, rotation_preview_thumb_width

    if not os.path.exists(CONFIG_PATH):
        if progress_callback is not None:
            progress_callback("No saved settings found.", 1, 1)
        return

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
            data = json.load(config_file)
    except (json.JSONDecodeError, OSError):
        if progress_callback is not None:
            progress_callback("Unable to read saved settings; using defaults.", 1, 1)
        return

    pending_value = data.get("pending_folder", "")
    root_value = data.get("root_folder", "")
    sources_value = data.get("employee_sources", [])
    filter_value = data.get("name_filter_mode", name_filter_mode.get())
    backup_pref_value = data.get("keep_backup_before_replace", False)
    auto_refresh_value = data.get("auto_refresh_pending_files", auto_refresh_var.get())
    tray_notifications_value = data.get(
        "tray_notifications_enabled",
        tray_notifications_enabled_var.get(),
    )
    show_text_with_icons_value = data.get(
        "show_text_with_icons",
        data.get("pending_toolbar_text_labels", show_text_with_icons_var.get()),
    )
    rotation_preview_width_value = data.get(
        "rotation_preview_window_width",
        rotation_preview_window_width,
    )
    rotation_preview_height_value = data.get(
        "rotation_preview_window_height",
        rotation_preview_window_height,
    )
    rotation_preview_thumb_value = data.get(
        "rotation_preview_thumb_width",
        rotation_preview_thumb_width,
    )

    if pending_value:
        pending_folder.set(normalize_path(pending_value))
    if root_value:
        root_folder.set(normalize_path(root_value))

    if filter_value in ("strict", "lenient"):
        _suppress_name_filter_refresh = True
        try:
            name_filter_mode.set(filter_value)
        finally:
            _suppress_name_filter_refresh = False

    keep_backup_preference_var.set(bool(backup_pref_value))
    auto_refresh_var.set(bool(auto_refresh_value))
    tray_notifications_enabled_var.set(bool(tray_notifications_value))
    show_text_with_icons_var.set(bool(show_text_with_icons_value))

    try:
        parsed_rotation_width = int(rotation_preview_width_value)
        if parsed_rotation_width >= ROTATION_PREVIEW_MIN_WIDTH:
            rotation_preview_window_width = parsed_rotation_width
    except (TypeError, ValueError):
        pass

    try:
        parsed_rotation_height = int(rotation_preview_height_value)
        if parsed_rotation_height >= ROTATION_PREVIEW_MIN_HEIGHT:
            rotation_preview_window_height = parsed_rotation_height
    except (TypeError, ValueError):
        pass

    try:
        parsed_thumb_width = int(rotation_preview_thumb_value)
        parsed_thumb_width = max(ROTATION_PREVIEW_THUMB_MIN_WIDTH, parsed_thumb_width)
        parsed_thumb_width = min(ROTATION_PREVIEW_THUMB_MAX_LIMIT, parsed_thumb_width)
        rotation_preview_thumb_width = parsed_thumb_width
    except (TypeError, ValueError):
        pass

    _set_employee_sources(sources_value or [], persist=False, progress_callback=progress_callback)
    _update_icon_button_labels()
    _refresh_update_status()


def save_settings():
    data = {
        "pending_folder": pending_folder.get(),
        "root_folder": root_folder.get(),
        "employee_sources": employee_source_paths,
        "name_filter_mode": name_filter_mode.get(),
        "keep_backup_before_replace": keep_backup_preference_var.get(),
        "auto_refresh_pending_files": auto_refresh_var.get(),
        "tray_notifications_enabled": tray_notifications_enabled_var.get(),
        "show_text_with_icons": show_text_with_icons_var.get(),
        "rotation_preview_window_width": rotation_preview_window_width,
        "rotation_preview_window_height": rotation_preview_window_height,
        "rotation_preview_thumb_width": rotation_preview_thumb_width,
    }
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as config_file:
            json.dump(data, config_file, indent=2)
    except OSError:
        pass


def exit_application():
    try:
        save_settings()
    except Exception:
        pass

    cancel_refresh = globals().get("_cancel_auto_refresh_job")
    if callable(cancel_refresh):
        try:
            cancel_refresh()
        except Exception:
            pass

    try:
        root.destroy()
    except tk.TclError:
        pass


def restart_application():
    try:
        save_settings()
    except Exception:
        pass

    if getattr(sys, "frozen", False):
        launch_cmd = [sys.executable] + sys.argv[1:]
        launch_cwd = os.path.dirname(sys.executable)
    else:
        launch_cmd = [sys.executable, os.path.abspath(__file__)] + sys.argv[1:]
        launch_cwd = os.path.dirname(os.path.abspath(__file__))

    try:
        subprocess.Popen(launch_cmd, cwd=launch_cwd)
    except Exception as exc:
        messagebox.showerror("Restart Failed", f"Unable to restart the app: {exc}")
        return

    exit_application()


def _launch_path(target_path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(target_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target_path])
        else:
            subprocess.Popen(["xdg-open", target_path])
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc


def _ensure_pdf_merger_available():
    if PdfMerger is None:
        message = "pypdf/PyPDF2 is not available. Install with 'pip install pypdf'."
        if PDF_IMPORT_ERROR:
            message += f"\nDetails: {PDF_IMPORT_ERROR}"
        raise RuntimeError(message)


def _ensure_pdf_rotation_available():
    if PdfReader is None or PdfWriter is None:
        message = (
            "PDF rotation requires pypdf/PyPDF2 with PdfWriter support. "
            "Install with 'pip install pypdf'."
        )
        if PDF_IMPORT_ERROR:
            message += f"\nDetails: {PDF_IMPORT_ERROR}"
        raise RuntimeError(message)


def _create_pdf_page_thumbnail(pdf_page, max_width=ROTATION_PREVIEW_THUMB_MAX_WIDTH):
    if Image is None:
        return None

    try:
        page_image = pdf_page.to_image(resolution=90)
        pil_image = page_image.original.convert("RGB")
    except Exception:
        return None

    width, height = pil_image.size
    if width <= 0 or height <= 0:
        return None

    if width > max_width:
        target_height = max(1, int((max_width / float(width)) * height))
        resample_filter = _get_pillow_lanczos_filter()
        if resample_filter is not None:
            pil_image = pil_image.resize((max_width, target_height), resample_filter)
        else:
            pil_image = pil_image.resize((max_width, target_height))

    return pil_image


def _build_pdf_thumbnail_photo(pil_image, max_width=ROTATION_PREVIEW_THUMB_MAX_WIDTH):
    if pil_image is None or ImageTk is None:
        return None

    image_to_render = pil_image
    width, height = image_to_render.size
    if width <= 0 or height <= 0:
        return None

    if width > max_width:
        target_height = max(1, int((max_width / float(width)) * height))
        resample_filter = _get_pillow_lanczos_filter()
        if resample_filter is not None:
            image_to_render = image_to_render.resize((max_width, target_height), resample_filter)
        else:
            image_to_render = image_to_render.resize((max_width, target_height))

    return ImageTk.PhotoImage(image_to_render)


def _rotate_pdf_pages_in_place(
    file_path,
    degrees=None,
    pages_to_rotate=None,
    page_rotation_map=None,
):
    _ensure_pdf_rotation_available()

    normalized_rotation_map = {}
    if page_rotation_map is not None:
        for raw_page_index, raw_rotation in dict(page_rotation_map).items():
            try:
                page_index = int(raw_page_index)
                if page_index < 0:
                    continue
            except (TypeError, ValueError):
                continue

            try:
                rotation_value = int(raw_rotation) % 360
            except (TypeError, ValueError):
                continue

            if rotation_value % 90 != 0:
                raise ValueError("Rotation angle must be a multiple of 90 degrees.")
            if rotation_value == 0:
                continue

            normalized_rotation_map[page_index] = rotation_value
    else:
        if degrees is None:
            raise ValueError("Rotation angle is required when page_rotation_map is not provided.")

        normalized_degrees = int(degrees)
        if normalized_degrees % 90 != 0:
            raise ValueError("Rotation angle must be a multiple of 90 degrees.")

        normalized_degrees = normalized_degrees % 360
        if normalized_degrees == 0:
            return 0

        if pages_to_rotate is None:
            selected_page_indexes = None
        else:
            selected_page_indexes = {
                int(page_index)
                for page_index in pages_to_rotate
                if int(page_index) >= 0
            }
            if not selected_page_indexes:
                return 0

        if selected_page_indexes is None:
            normalized_rotation_map = None
        else:
            normalized_rotation_map = {
                page_index: normalized_degrees for page_index in selected_page_indexes
            }

    try:
        reader = PdfReader(file_path)
    except Exception as exc:
        raise RuntimeError(f"Unable to read PDF: {exc}") from exc

    writer = PdfWriter()
    rotated_pages = 0

    for page_index, page in enumerate(reader.pages):
        if normalized_rotation_map is None:
            page_rotation = normalized_degrees
        else:
            page_rotation = normalized_rotation_map.get(page_index, 0)

        if page_rotation:
            try:
                rotated_page = page.rotate(page_rotation)
            except Exception:
                rotated_page = page
                try:
                    rotated_page = page.rotate_clockwise(page_rotation)
                except Exception as exc:
                    raise RuntimeError(
                        f"Unable to rotate page {page_index + 1}: {exc}"
                    ) from exc

            writer.add_page(rotated_page)
            rotated_pages += 1
        else:
            writer.add_page(page)

    metadata = getattr(reader, "metadata", None)
    if metadata:
        try:
            cleaned_metadata = {
                str(key): str(value)
                for key, value in dict(metadata).items()
                if key and value is not None
            }
            if cleaned_metadata:
                writer.add_metadata(cleaned_metadata)
        except Exception:
            pass

    temp_output_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            delete=False,
            suffix=".pdf",
            dir=os.path.dirname(file_path) or None,
        ) as temp_output:
            temp_output_path = temp_output.name
            writer.write(temp_output)

        shutil.move(temp_output_path, file_path)
        temp_output_path = None
    except Exception as exc:
        raise RuntimeError(f"Unable to write rotated PDF: {exc}") from exc
    finally:
        if temp_output_path and os.path.exists(temp_output_path):
            try:
                os.remove(temp_output_path)
            except OSError:
                pass

    return rotated_pages


def get_pdf_page_count(pdf_path):
    if PdfReader is None:
        return None
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except Exception:
        return None


def create_backup_file(source_path):
    if not os.path.exists(source_path):
        return None
    backup_dir = os.path.join(os.path.dirname(source_path), "_backups")
    os.makedirs(backup_dir, exist_ok=True)
    base_name = os.path.basename(source_path)
    name_root, ext = os.path.splitext(base_name)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_name = f"{name_root}_backup_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)
    shutil.copy2(source_path, backup_path)
    return backup_path


def merge_pdf_files(new_pdf_path, existing_pdf_path, output_path):
    _ensure_pdf_merger_available()
    merger = PdfMerger()
    try:
        merger.append(new_pdf_path)
        merger.append(existing_pdf_path)
        with open(output_path, "wb") as merged_file:
            merger.write(merged_file)
    finally:
        merger.close()


def parse_filename_metadata(filename):
    base_name = os.path.splitext(filename)[0]
    parts = base_name.split("_")
    if len(parts) < 2:
        return None

    latest_year = ""
    earliest_year = ""
    employee_name = ""

    if len(parts) >= 3 and parts[-1].isdigit() and parts[-2].isdigit():
        latest_val = max(int(parts[-2]), int(parts[-1]))
        earliest_val = min(int(parts[-2]), int(parts[-1]))
        latest_year = str(latest_val)
        earliest_year = str(earliest_val)
        employee_name = "_".join(parts[:-2])
    elif parts[-1].isdigit():
        latest_year = str(int(parts[-1]))
        earliest_year = ""
        employee_name = "_".join(parts[:-1])
    else:
        return None

    if not employee_name.strip():
        return None

    return {
        "name": employee_name,
        "latest": latest_year,
        "earliest": earliest_year,
    }


def _normalize_record_year_inputs(latest_text, earliest_text):
    latest = str(latest_text or "").strip()
    earliest = str(earliest_text or "").strip()

    if latest and not latest.isdigit():
        raise ValueError("Latest Year must be numeric.")
    if earliest and not earliest.isdigit():
        raise ValueError("Oldest Year must be numeric.")

    if not latest and not earliest:
        return None

    if latest and earliest:
        latest_year = max(int(latest), int(earliest))
        earliest_year = min(int(latest), int(earliest))
    else:
        single_year = int(latest or earliest)
        latest_year = single_year
        earliest_year = single_year

    return latest_year, earliest_year


def _get_year_input_guidance(latest_text, earliest_text):
    latest = str(latest_text or "").strip()
    earliest = str(earliest_text or "").strip()

    if not latest and not earliest:
        return "Enter Latest Year or Oldest Year (at least one is required)."
    if latest and not latest.isdigit():
        return "Latest Year must contain digits only (example: 2025)."
    if earliest and not earliest.isdigit():
        return "Oldest Year must contain digits only (example: 2018)."

    years = _normalize_record_year_inputs(latest, earliest)
    if years is None:
        return "Enter Latest Year or Oldest Year (at least one is required)."

    latest_year, earliest_year = years
    if latest_year == earliest_year:
        return f"Year will be saved as {latest_year}."
    return f"Year range will be normalized to {latest_year} - {earliest_year}."


def _build_record_filename(name, latest_year, earliest_year):
    latest_int = int(latest_year)
    earliest_int = int(earliest_year)
    if latest_int == earliest_int:
        return f"{name}_{latest_int}.pdf"
    return f"{name}_{latest_int}_{earliest_int}.pdf"


def _validate_filesystem_component_name(raw_value, value_label):
    value = str(raw_value or "").strip()
    if not value:
        raise ValueError(f"{value_label} cannot be empty.")
    if value in {".", ".."}:
        raise ValueError(f"{value_label} cannot be '.' or '..'.")
    if value[-1] in {" ", "."}:
        raise ValueError(f"{value_label} cannot end with a space or period.")

    invalid_chars = '<>:"/\\|?*'
    bad_chars = [ch for ch in invalid_chars if ch in value]
    if bad_chars:
        raise ValueError(
            f"{value_label} contains invalid characters: {' '.join(bad_chars)}"
        )

    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "COM1",
        "COM2",
        "COM3",
        "COM4",
        "COM5",
        "COM6",
        "COM7",
        "COM8",
        "COM9",
        "LPT1",
        "LPT2",
        "LPT3",
        "LPT4",
        "LPT5",
        "LPT6",
        "LPT7",
        "LPT8",
        "LPT9",
    }
    if value.split(".")[0].upper() in reserved_names:
        raise ValueError(f"{value_label} uses a reserved Windows name.")

    return value


def _format_editor_datetime(dt_value):
    return dt_value.strftime("%Y-%m-%d %H:%M:%S")


def _parse_editor_datetime_input(raw_value, label):
    value = str(raw_value or "").strip()
    if not value:
        raise ValueError(f"{label} is required.")

    accepted_formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    )
    for dt_format in accepted_formats:
        try:
            parsed = datetime.strptime(value, dt_format)
            if dt_format == "%Y-%m-%d":
                parsed = parsed.replace(hour=0, minute=0, second=0)
            return parsed
        except ValueError:
            continue

    raise ValueError(f"{label} format is invalid. Use YYYY-MM-DD HH:MM[:SS].")


def _set_file_creation_and_modified_time(file_path, created_dt, modified_dt):
    if modified_dt is not None:
        stats = os.stat(file_path)
        os.utime(file_path, (stats.st_atime, modified_dt.timestamp()))

    if not sys.platform.startswith("win"):
        return

    if created_dt is None and modified_dt is None:
        return

    import ctypes
    from ctypes import wintypes

    class FILETIME(ctypes.Structure):
        _fields_ = [
            ("dwLowDateTime", wintypes.DWORD),
            ("dwHighDateTime", wintypes.DWORD),
        ]

    def _to_filetime(dt_value):
        intervals = int(dt_value.timestamp() * 10_000_000) + 116_444_736_000_000_000
        return FILETIME(intervals & 0xFFFFFFFF, intervals >> 32)

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    set_file_time = kernel32.SetFileTime
    close_handle = kernel32.CloseHandle

    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    set_file_time.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
        ctypes.POINTER(FILETIME),
    ]
    set_file_time.restype = wintypes.BOOL

    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    FILE_WRITE_ATTRIBUTES = 0x0100
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x00000080
    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

    file_handle = create_file(
        file_path,
        FILE_WRITE_ATTRIBUTES,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )

    if file_handle == INVALID_HANDLE_VALUE:
        raise OSError(f"Unable to access file attributes for {os.path.basename(file_path)}.")

    created_filetime = _to_filetime(created_dt) if created_dt is not None else None
    modified_filetime = _to_filetime(modified_dt) if modified_dt is not None else None

    try:
        success = set_file_time(
            file_handle,
            ctypes.byref(created_filetime) if created_filetime is not None else None,
            None,
            ctypes.byref(modified_filetime) if modified_filetime is not None else None,
        )
        if not success:
            raise OSError(f"Unable to update timestamps for {os.path.basename(file_path)}.")
    finally:
        close_handle(file_handle)


def archive_pending_file(pending_path):
    processed_dir = os.path.join(os.path.dirname(pending_path), "processed")
    os.makedirs(processed_dir, exist_ok=True)

    base_name = os.path.basename(pending_path)
    destination = os.path.join(processed_dir, base_name)
    name_root, ext = os.path.splitext(destination)
    counter = 1
    while os.path.exists(destination):
        destination = f"{name_root}_{counter}{ext}"
        counter += 1

    try:
        shutil.move(pending_path, destination)
    except OSError:
        # Fallback path in case move fails on this filesystem/session.
        shutil.copy2(pending_path, destination)
        os.remove(pending_path)

    return destination

# ------------------------
# Functions
# ------------------------
def select_pending_folder():
    folder = filedialog.askdirectory()
    if folder:
        pending_folder.set(normalize_path(folder))
        save_settings()
        load_pending_files() 

def select_root_folder():
    folder = filedialog.askdirectory()
    if folder:
        root_folder.set(normalize_path(folder))
        save_settings()


def _apply_employee_sources_with_progress(
    paths,
    persist,
    title="Loading Employee Sources",
    heading="Parsing employee source files",
    initial_status="Preparing selected sources...",
):
    loading_win, loading_status_var, loading_detail_var, loading_progress_bar = _create_startup_loading_window(
        title=title,
        heading=heading,
        initial_status=initial_status,
    )

    try:
        _update_startup_loading_ui(
            loading_win,
            loading_status_var,
            loading_detail_var,
            loading_progress_bar,
            initial_status,
            0,
            0,
        )

        _set_employee_sources(
            paths,
            persist=persist,
            progress_callback=lambda message, current=None, total=None: _update_startup_loading_ui(
                loading_win,
                loading_status_var,
                loading_detail_var,
                loading_progress_bar,
                message,
                current,
                total,
            ),
        )
    except Exception as exc:
        messagebox.showerror("Error", f"Unable to load employee sources: {exc}")
    finally:
        if getattr(loading_progress_bar, "_is_running", False):
            loading_progress_bar.stop()
            loading_progress_bar._is_running = False
        if loading_win.winfo_exists():
            loading_win.destroy()


def select_employee_sources():
    files = filedialog.askopenfilenames(
        filetypes=[
            ("Supported files", "*.pdf *.xlsx *.xlsm *.xltx *.xltm *.xls *.csv *.txt"),
            ("PDF files", "*.pdf"),
            ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm *.xls"),
            ("CSV files", "*.csv"),
            ("Text files", "*.txt"),
        ]
    )
    if not files:
        return
    new_paths = employee_source_paths + [normalize_path(path) for path in files]
    _apply_employee_sources_with_progress(
        new_paths,
        persist=True,
        title="Loading Employee Sources",
        heading="Parsing employee source files",
        initial_status="Preparing selected sources...",
    )


def remove_selected_employee_source():
    if employee_sources_listbox is None:
        return
    selection = employee_sources_listbox.curselection()
    if not selection:
        messagebox.showwarning("Warning", "Select at least one source to remove.")
        return
    selected_indexes = set(selection)
    remaining = [path for idx, path in enumerate(employee_source_paths) if idx not in selected_indexes]
    _apply_employee_sources_with_progress(
        remaining,
        persist=True,
        title="Updating Employee Sources",
        heading="Refreshing source list",
        initial_status="Applying selected removals...",
    )


def clear_employee_sources():
    if not employee_source_paths:
        return
    confirmed = messagebox.askyesno("Clear Sources", "Remove all employee name sources?")
    if confirmed:
        _apply_employee_sources_with_progress(
            [],
            persist=True,
            title="Clearing Employee Sources",
            heading="Removing all source files",
            initial_status="Clearing source list...",
        )


def show_parsed_names_window():
    win = tk.Toplevel(root)
    _apply_app_icon(win)
    win.title("Parsed Employee Names")
    configure_window_geometry(
        win,
        520,
        620,
        min_width=420,
        min_height=360,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(root)
    win.lift()
    win.focus_force()
    apply_theme(win)

    content = ttk.Frame(win, padding=16, style="TFrame")
    content.pack(fill="both", expand=True)

    parsed_count = len(employee_name_suggestions)
    parsed_text = "name" if parsed_count == 1 else "names"

    ttk.Label(
        content,
        text=f"Parsed employee names: {parsed_count} {parsed_text}",
        style="Card.TLabel",
    ).pack(anchor="w", pady=(0, 10))

    if parsed_count == 0:
        ttk.Label(
            content,
            text="No names parsed yet. Add employee source files to populate this list.",
            style="Subheading.TLabel",
            justify="left",
            wraplength=460,
        ).pack(anchor="w")
        ttk.Button(content, text="Close", command=win.destroy).pack(anchor="e", pady=(12, 0))
        return

    list_container = ttk.Frame(content, style="Card.TFrame", padding=8)
    list_container.pack(fill="both", expand=True)

    names_scrollbar = ttk.Scrollbar(list_container, orient="vertical")
    names_scrollbar.pack(side="right", fill="y")

    names_listbox = tk.Listbox(
        list_container,
        yscrollcommand=names_scrollbar.set,
    )
    _apply_modern_listbox_style(names_listbox, compact=True, export_selection=False)
    names_listbox.pack(side="left", fill="both", expand=True)
    names_scrollbar.configure(command=names_listbox.yview)

    for name in employee_name_suggestions:
        names_listbox.insert(tk.END, name)

    def export_parsed_names():
        names_to_export = [str(name).strip() for name in employee_name_suggestions if str(name).strip()]
        if not names_to_export:
            messagebox.showwarning("Export Parsed Names", "There are no parsed names to export.", parent=win)
            return

        target_path = filedialog.asksaveasfilename(
            parent=win,
            title="Export Parsed Names",
            initialfile="parsed_names.txt",
            defaultextension=".txt",
            filetypes=[
                ("Text Files", "*.txt"),
                ("CSV Files", "*.csv"),
                ("JSON Files", "*.json"),
                ("All Files", "*.*"),
            ],
        )
        if not target_path:
            return

        output_path = normalize_path(target_path)
        _, extension = os.path.splitext(output_path)
        extension = extension.lower()

        if extension not in {".txt", ".csv", ".json"}:
            output_path = output_path + ".txt"
            extension = ".txt"

        try:
            if extension == ".csv":
                with open(output_path, "w", newline="", encoding="utf-8") as output_file:
                    writer = csv.writer(output_file)
                    writer.writerow(["Name"])
                    for parsed_name in names_to_export:
                        writer.writerow([parsed_name])
            elif extension == ".json":
                with open(output_path, "w", encoding="utf-8") as output_file:
                    json.dump(names_to_export, output_file, indent=2, ensure_ascii=False)
            else:
                with open(output_path, "w", encoding="utf-8") as output_file:
                    output_file.write("\n".join(names_to_export))

            messagebox.showinfo(
                "Export Parsed Names",
                f"Exported {len(names_to_export)} parsed name(s) to:\n{output_path}",
                parent=win,
            )
        except OSError as exc:
            messagebox.showerror(
                "Export Parsed Names",
                f"Unable to export parsed names:\n{exc}",
                parent=win,
            )

    actions = ttk.Frame(content, style="TFrame")
    actions.pack(fill="x", pady=(12, 0))

    ttk.Button(actions, text="Export", command=export_parsed_names, style="Accent.TButton").pack(side="left")
    ttk.Button(actions, text="Close", command=win.destroy).pack(side="right")
        
pending_items_frame = None
pending_canvas_widget = None
pending_file_vars = {}
pending_file_order = []
pending_selection_anchor_filename = None
pending_snapshot = set()
auto_refresh_job_id = None
ui_icon_images = {}
pending_row_preview_buttons = []
name_buttons_container = None
add_sources_button = None
remove_sources_button = None
clear_sources_button = None
view_parsed_names_button = None
pending_rotate_button = None
pending_preview_button = None
pending_master_toggle_button = None
pending_master_selection_state = "none"
_pending_selection_update_in_progress = False


def _set_pending_snapshot(file_names=None):
    global pending_snapshot
    pending_snapshot = set(file_names or [])


def _list_pending_files_on_disk():
    folder = normalize_path(pending_folder.get())
    if not folder or not os.path.isdir(folder):
        return set()
    try:
        return {file for file in os.listdir(folder) if file.lower().endswith(".pdf")}
    except OSError:
        return set()


def _notify_new_pending_files(new_files):
    if not new_files:
        return
    if not tray_notifications_enabled_var.get():
        return
    count = len(new_files)
    plural = "s" if count != 1 else ""
    message = f"{count} new pending PDF{plural} detected."
    if tray_notifier.notify("Pending PDFs Updated", message):
        _set_tray_status_message()
    else:
        _set_tray_status_message(tray_notifier.status_message())


def _auto_refresh_handler():
    global auto_refresh_job_id
    auto_refresh_job_id = None
    if not auto_refresh_var.get():
        return
    current_files = _list_pending_files_on_disk()
    added = current_files - pending_snapshot
    removed = pending_snapshot - current_files
    if added or removed:
        load_pending_files()
        if added:
            _notify_new_pending_files(added)
    _set_pending_snapshot(current_files)
    _schedule_auto_refresh()


def _schedule_auto_refresh():
    global auto_refresh_job_id
    if auto_refresh_job_id is not None or not auto_refresh_var.get():
        return
    auto_refresh_job_id = root.after(AUTO_REFRESH_INTERVAL_MS, _auto_refresh_handler)


def _cancel_auto_refresh_job():
    global auto_refresh_job_id
    if auto_refresh_job_id is not None:
        try:
            root.after_cancel(auto_refresh_job_id)
        except Exception:
            pass
        auto_refresh_job_id = None


def _toggle_auto_refresh():
    if auto_refresh_var.get():
        _cancel_auto_refresh_job()
        _auto_refresh_handler()
    else:
        _cancel_auto_refresh_job()


def _on_auto_refresh_preference_changed():
    _toggle_auto_refresh()
    save_settings()


def _on_tray_notifications_preference_changed():
    save_settings()


def _on_keep_backup_preference_changed():
    save_settings()


def _configure_icon_button(button, icon_name, fallback_icon, label):
    icon_image = ui_icon_images.get(icon_name)
    show_text_label = show_text_with_icons_var.get()

    if icon_image is not None:
        if show_text_label:
            button.configure(
                text=label,
                image=icon_image,
                compound="left",
                width=0,
                padding=(10, 6),
            )
        else:
            button.configure(
                text="",
                image=icon_image,
                compound="image",
                width=0,
                padding=(8, 6),
            )
        return

    if show_text_label:
        button.configure(
            text=f"{fallback_icon} {label}",
            image="",
            compound="none",
            width=max(10, len(label) + 4),
            padding=(10, 6),
        )
    else:
        button.configure(
            text=fallback_icon,
            image="",
            compound="none",
            width=3,
            padding=(8, 6),
        )


def _refresh_name_toolbar_layout():
    if name_buttons_container is None or not name_buttons_container.winfo_exists():
        return

    buttons = [
        add_sources_button,
        remove_sources_button,
        clear_sources_button,
        view_parsed_names_button,
    ]
    live_buttons = [button for button in buttons if button is not None and button.winfo_exists()]
    if not live_buttons:
        return

    for button in live_buttons:
        try:
            button.pack_forget()
        except tk.TclError:
            pass
        try:
            button.grid_forget()
        except tk.TclError:
            pass

    show_text_label = show_text_with_icons_var.get()
    if show_text_label:
        name_buttons_container.columnconfigure(0, weight=1, uniform="name_toolbar")
        name_buttons_container.columnconfigure(1, weight=1, uniform="name_toolbar")
        for index, button in enumerate(live_buttons):
            row_index = index // 2
            col_index = index % 2
            button.grid(row=row_index, column=col_index, sticky="ew", padx=4, pady=4)
        return

    name_buttons_container.columnconfigure(0, weight=0)
    name_buttons_container.columnconfigure(1, weight=0)
    for index, button in enumerate(live_buttons):
        button.pack(side="left", padx=(0, 6) if index == 0 else 6)


def _refresh_pending_toolbar_layout():
    if (
        pending_master_actions is None
        or not pending_master_actions.winfo_exists()
        or pending_master_toggle_button is None
        or not pending_master_toggle_button.winfo_exists()
    ):
        return

    try:
        pending_master_toggle_button.pack_forget()
        pending_master_actions.pack_forget()
    except tk.TclError:
        return

    pending_master_toggle_button.pack(side="left", anchor="w")
    pending_master_actions.pack(side="right")


def _update_icon_button_labels():
    button_specs = (
        (add_sources_button, "source_add", TOOLBAR_ICON_SOURCE_ADD, "Add Sources"),
        (remove_sources_button, "source_remove", TOOLBAR_ICON_SOURCE_REMOVE, "Remove Selected"),
        (clear_sources_button, "clear_selection", TOOLBAR_ICON_SOURCE_CLEAR, "Clear All"),
        (view_parsed_names_button, "preview", TOOLBAR_ICON_PREVIEW, "View Parsed"),
        (pending_rotate_button, "refresh", TOOLBAR_ICON_ROTATE, "Rotate"),
        (pending_preview_button, "preview", TOOLBAR_ICON_PREVIEW, "Preview"),
    )
    for button, icon_name, fallback_icon, label in button_specs:
        if button is None or not button.winfo_exists():
            continue
        _configure_icon_button(button, icon_name, fallback_icon, label)

    for button in list(pending_row_preview_buttons):
        if button is None or not button.winfo_exists():
            continue
        _configure_icon_button(button, "preview", TOOLBAR_ICON_PREVIEW, "Preview")

    _refresh_name_toolbar_layout()
    _refresh_pending_toolbar_layout()
    _refresh_pending_master_toggle_state()


def _get_pending_selection_counts():
    total_count = len(pending_file_vars)
    selected_count = sum(1 for var in pending_file_vars.values() if var.get())
    return total_count, selected_count


def _configure_master_toggle_button(selection_state, enabled):
    if pending_master_toggle_button is None or not pending_master_toggle_button.winfo_exists():
        return

    show_text_label = show_text_with_icons_var.get()
    icon_key_map = {
        "none": "select_none",
        "partial": "select_partial",
        "all": "select_all",
    }
    fallback_icon_map = {
        "none": TOOLBAR_ICON_SELECT_NONE,
        "partial": TOOLBAR_ICON_SELECT_PARTIAL,
        "all": TOOLBAR_ICON_SELECT_ALL,
    }
    label_map = {
        "none": "Select",
        "partial": "Partial",
        "all": "All",
    }

    icon_key = icon_key_map.get(selection_state, "select_none")
    fallback_icon = fallback_icon_map.get(selection_state, TOOLBAR_ICON_SELECT_NONE)
    label = label_map.get(selection_state, "Select All")
    icon_image = ui_icon_images.get(icon_key)

    if icon_image is not None:
        if show_text_label:
            pending_master_toggle_button.configure(
                text=label,
                image=icon_image,
                compound="left",
                width=0,
                padding=(10, 6),
            )
        else:
            pending_master_toggle_button.configure(
                text="",
                image=icon_image,
                compound="image",
                width=0,
                padding=(8, 6),
            )
    else:
        if show_text_label:
            pending_master_toggle_button.configure(
                text=f"{fallback_icon} {label}",
                image="",
                compound="none",
                width=max(14, len(label) + 4),
                padding=(10, 6),
            )
        else:
            pending_master_toggle_button.configure(
                text=fallback_icon,
                image="",
                compound="none",
                width=3,
                padding=(8, 6),
            )

    if enabled:
        pending_master_toggle_button.state(["!disabled"])
    else:
        pending_master_toggle_button.state(["disabled"])


def _set_all_pending_file_selections(selected):
    global _pending_selection_update_in_progress

    if not pending_file_vars:
        _refresh_pending_master_toggle_state()
        return

    _pending_selection_update_in_progress = True
    try:
        for var in pending_file_vars.values():
            var.set(bool(selected))
    finally:
        _pending_selection_update_in_progress = False

    _refresh_pending_master_toggle_state()


def _set_pending_selection_state(selected_names):
    global _pending_selection_update_in_progress

    selected_lookup = set(selected_names or [])
    if not pending_file_vars:
        _refresh_pending_master_toggle_state()
        return

    _pending_selection_update_in_progress = True
    try:
        for file_name, var in pending_file_vars.items():
            var.set(file_name in selected_lookup)
    finally:
        _pending_selection_update_in_progress = False

    _refresh_pending_master_toggle_state()


def _get_selected_pending_file_names_set():
    return {name for name, var in pending_file_vars.items() if var.get()}


def _is_pending_ctrl_pressed(event):
    return bool(int(getattr(event, "state", 0)) & 0x0004)


def _is_pending_shift_pressed(event):
    return bool(int(getattr(event, "state", 0)) & 0x0001)


def _select_pending_range_to(target_filename, additive=False):
    global pending_selection_anchor_filename

    if target_filename not in pending_file_vars:
        return

    ordered_names = [name for name in pending_file_order if name in pending_file_vars]
    if not ordered_names:
        ordered_names = list(pending_file_vars.keys())
    if target_filename not in ordered_names:
        return

    anchor_filename = pending_selection_anchor_filename
    if anchor_filename not in ordered_names:
        anchor_filename = target_filename
        pending_selection_anchor_filename = target_filename

    start_index = ordered_names.index(anchor_filename)
    end_index = ordered_names.index(target_filename)
    left = min(start_index, end_index)
    right = max(start_index, end_index)
    range_selection = set(ordered_names[left : right + 1])

    if additive:
        selected_names = _get_selected_pending_file_names_set()
        selected_names.update(range_selection)
    else:
        selected_names = range_selection

    _set_pending_selection_state(selected_names)


def _focus_pending_selection_surface():
    if pending_canvas_widget is None:
        return
    try:
        if pending_canvas_widget.winfo_exists():
            pending_canvas_widget.focus_set()
    except Exception:
        pass


def _handle_pending_item_click(event=None, target_filename=None):
    global pending_selection_anchor_filename

    if not target_filename or target_filename not in pending_file_vars:
        return "break"

    _focus_pending_selection_surface()

    ctrl_pressed = _is_pending_ctrl_pressed(event)
    shift_pressed = _is_pending_shift_pressed(event)

    if shift_pressed:
        _select_pending_range_to(target_filename, additive=ctrl_pressed)
    elif ctrl_pressed:
        selected_names = _get_selected_pending_file_names_set()
        if target_filename in selected_names:
            selected_names.remove(target_filename)
        else:
            selected_names.add(target_filename)
        _set_pending_selection_state(selected_names)
        pending_selection_anchor_filename = target_filename
    else:
        _set_pending_selection_state({target_filename})
        pending_selection_anchor_filename = target_filename

    return "break"


def _is_widget_in_pending_list(widget):
    if widget is None:
        return False

    try:
        widget_path = str(widget)
    except Exception:
        return False

    for candidate in (pending_canvas_widget, pending_items_frame):
        if candidate is None:
            continue
        try:
            if not candidate.winfo_exists():
                continue
            candidate_path = str(candidate)
        except Exception:
            continue

        if widget_path == candidate_path or widget_path.startswith(candidate_path + "."):
            return True

    return False


def _on_pending_ctrl_select_all(event=None):
    global pending_selection_anchor_filename

    if event is not None and not _is_widget_in_pending_list(getattr(event, "widget", None)):
        return None

    if not pending_file_vars:
        _refresh_pending_master_toggle_state()
        return "break"

    _set_all_pending_file_selections(True)
    if pending_file_order:
        pending_selection_anchor_filename = pending_file_order[0]

    _focus_pending_selection_surface()
    return "break"


def _refresh_pending_master_toggle_state():
    global pending_master_selection_state

    total_count, selected_count = _get_pending_selection_counts()
    pending_files_count_var.set(f"({selected_count} selected / {total_count} total)")

    if pending_master_toggle_button is None or not pending_master_toggle_button.winfo_exists():
        pending_master_selection_state = "none"
        return

    if total_count == 0:
        pending_master_selection_state = "none"
        _configure_master_toggle_button("none", enabled=False)
        return

    if selected_count == 0:
        pending_master_selection_state = "none"
        _configure_master_toggle_button("none", enabled=True)
    elif selected_count == total_count:
        pending_master_selection_state = "all"
        _configure_master_toggle_button("all", enabled=True)
    else:
        pending_master_selection_state = "partial"
        _configure_master_toggle_button("partial", enabled=True)


def _on_pending_file_check_state_changed(*_args):
    if _pending_selection_update_in_progress:
        return
    _refresh_pending_master_toggle_state()


def _on_pending_master_toggle_clicked():
    total_count, selected_count = _get_pending_selection_counts()
    if total_count == 0:
        _refresh_pending_master_toggle_state()
        return

    target_state_selected = selected_count != total_count
    _set_all_pending_file_selections(target_state_selected)


def _on_show_text_with_icons_preference_changed():
    _update_icon_button_labels()
    save_settings()


def _format_pending_filename_for_display(filename, max_length=62):
    if not filename or len(filename) <= max_length:
        return filename

    head_len = max(16, int(max_length * 0.65))
    tail_len = max(8, max_length - head_len - 3)
    return f"{filename[:head_len]}...{filename[-tail_len:]}"


def _set_pending_row_hover_state(row_widget, name_label, check_widget, hovered):
    selected = bool(getattr(check_widget, "_pending_selected", False))

    if selected and hovered:
        row_widget.configure(style="PendingRowSelectedHover.TFrame")
        name_label.configure(style="PendingFileSelectedHover.TLabel")
        check_widget.configure(style="PendingFileSelectedHover.TCheckbutton")
    elif selected:
        row_widget.configure(style="PendingRowSelected.TFrame")
        name_label.configure(style="PendingFileSelected.TLabel")
        check_widget.configure(style="PendingFileSelected.TCheckbutton")
    elif hovered:
        row_widget.configure(style="PendingRowHover.TFrame")
        name_label.configure(style="PendingFileHover.TLabel")
        check_widget.configure(style="PendingFileHover.TCheckbutton")
    else:
        row_widget.configure(style="PendingRow.TFrame")
        name_label.configure(style="PendingFile.TLabel")
        check_widget.configure(style="PendingFile.TCheckbutton")


def _ensure_auto_refresh_job():
    if auto_refresh_var.get():
        _schedule_auto_refresh()


def _set_pending_files_count(count):
    total_count = max(0, int(count))
    pending_files_count_var.set(f"(0 selected / {total_count} total)")


def load_pending_files():
    global pending_file_vars, pending_row_preview_buttons
    global pending_file_order, pending_selection_anchor_filename
    if pending_items_frame is None:
        pending_row_preview_buttons = []
        pending_file_order = []
        pending_selection_anchor_filename = None
        _set_pending_files_count(0)
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    pending_row_preview_buttons = []

    for child in pending_items_frame.winfo_children():
        child.destroy()

    folder = normalize_path(pending_folder.get())
    if not folder:
        pending_file_vars = {}
        pending_file_order = []
        pending_selection_anchor_filename = None
        _set_pending_files_count(0)
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    try:
        files = sorted(file for file in os.listdir(folder) if file.lower().endswith(".pdf"))
    except OSError as exc:
        messagebox.showerror("Error", f"Unable to load pending PDFs: {exc}")
        pending_file_vars = {}
        pending_file_order = []
        pending_selection_anchor_filename = None
        _set_pending_files_count(0)
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    _set_pending_files_count(len(files))
    pending_file_order = list(files)

    previous_state = {name: var.get() for name, var in pending_file_vars.items()}
    pending_file_vars = {}

    if pending_selection_anchor_filename not in pending_file_order:
        pending_selection_anchor_filename = None

    if not files:
        pending_file_order = []
        pending_selection_anchor_filename = None
        _set_pending_snapshot()
        _refresh_pending_master_toggle_state()
        ttk.Label(
            pending_items_frame,
            text="No pending PDFs found.",
            style="Subheading.TLabel",
            padding=6,
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=4, pady=4)
        return

    _set_pending_snapshot(files)

    for filename in files:
        var = tk.BooleanVar(value=previous_state.get(filename, False))
        var.trace_add("write", _on_pending_file_check_state_changed)
        pending_file_vars[filename] = var

        row = ttk.Frame(pending_items_frame, style="PendingRow.TFrame", padding=(10, 6))
        row.pack(fill="x", padx=2, pady=3)

        checkbutton = ttk.Checkbutton(row, variable=var, style="PendingFile.TCheckbutton")
        checkbutton.pack(side="left", padx=(2, 8))
        checkbutton._pending_selected = bool(var.get())

        display_name = _format_pending_filename_for_display(filename)
        name_label = ttk.Label(row, text=display_name, style="PendingFile.TLabel", anchor="w")
        name_label.pack(side="left", fill="x", expand=True, padx=(0, 8))

        if display_name != filename:
            _attach_hover_tooltip(name_label, filename)

        def _on_pending_item_click(_event=None, target_filename=filename):
            return _handle_pending_item_click(_event, target_filename)

        hover_state = {"value": False}

        def _apply_row_visual(
            target_row=row,
            target_label=name_label,
            target_check=checkbutton,
            target_var=var,
            target_hover_state=hover_state,
        ):
            target_check._pending_selected = bool(target_var.get())
            _set_pending_row_hover_state(
                target_row,
                target_label,
                target_check,
                target_hover_state["value"],
            )

        def _on_row_selection_changed(
            *_args,
            target_apply_row_visual=_apply_row_visual,
        ):
            target_apply_row_visual()

        def _on_row_enter(_event=None, target_hover_state=hover_state, target_apply_row_visual=_apply_row_visual):
            target_hover_state["value"] = True
            target_apply_row_visual()

        def _on_row_leave(_event=None, target_hover_state=hover_state, target_apply_row_visual=_apply_row_visual):
            target_hover_state["value"] = False
            target_apply_row_visual()

        var.trace_add("write", _on_row_selection_changed)

        for click_widget in (row, name_label, checkbutton):
            click_widget.bind("<Button-1>", _on_pending_item_click, add="+")

        for hover_widget in (row, name_label, checkbutton):
            hover_widget.bind("<Enter>", _on_row_enter, add="+")
            hover_widget.bind("<Leave>", _on_row_leave, add="+")

        _apply_row_visual()

        preview_button = ttk.Button(
            row,
            style="ToolbarIcon.TButton",
            command=lambda f=filename: preview_specific_pending_pdf(f),
        )
        preview_button.pack(side="right", padx=(8, 0))
        preview_button.bind("<Enter>", _on_row_enter, add="+")
        preview_button.bind("<Leave>", _on_row_leave, add="+")
        _attach_hover_tooltip(preview_button, "Preview this pending PDF")
        _configure_icon_button(preview_button, "preview", TOOLBAR_ICON_PREVIEW, "Preview")
        pending_row_preview_buttons.append(preview_button)

    if pending_selection_anchor_filename is None:
        for file_name in pending_file_order:
            var = pending_file_vars.get(file_name)
            if var is not None and var.get():
                pending_selection_anchor_filename = file_name
                break

    _refresh_pending_master_toggle_state()

def preview_selected_pdf():
    selected = get_selected_pending_files()
    if not selected:
        messagebox.showwarning("Warning", "Select at least one pending PDF using the checkboxes.")
        return

    for filename in selected:
        preview_specific_pending_pdf(filename)


def rotate_selected_pending_pdfs(
    selected_file_infos_override=None,
    window_title="Rotate Pending PDFs",
    post_save_callback=None,
    parent_window=None,
):
    global rotation_preview_window_width, rotation_preview_window_height, rotation_preview_thumb_width

    selected_file_infos = []
    missing_files = []

    if selected_file_infos_override is None:
        selected = get_selected_pending_files()
        if not selected:
            messagebox.showwarning("Warning", "Select at least one pending PDF using the checkboxes.")
            return

        pending_dir = normalize_path(pending_folder.get())
        if not pending_dir:
            messagebox.showwarning("Warning", "Select the pending folder first.")
            return
        if not os.path.isdir(pending_dir):
            messagebox.showerror("Error", "The pending folder no longer exists.")
            return

        for filename in selected:
            file_path = normalize_path(os.path.join(pending_dir, filename))
            if not os.path.exists(file_path):
                missing_files.append(filename)
                continue
            selected_file_infos.append(
                {
                    "name": filename,
                    "path": file_path,
                    "page_count": 0,
                }
            )
    else:
        for file_info in list(selected_file_infos_override):
            if isinstance(file_info, dict):
                raw_path = file_info.get("path", "")
                file_path = normalize_path(raw_path)
                filename = str(file_info.get("name") or os.path.basename(file_path or ""))
            else:
                file_path = normalize_path(str(file_info))
                filename = os.path.basename(file_path)

            if not file_path or not os.path.exists(file_path):
                missing_files.append(filename or "(missing file)")
                continue

            selected_file_infos.append(
                {
                    "name": filename,
                    "path": file_path,
                    "page_count": 0,
                }
            )

    try:
        _ensure_pdf_rotation_available()
    except RuntimeError as exc:
        messagebox.showerror("Rotate PDF", str(exc))
        return

    if missing_files:
        preview = ", ".join(missing_files[:3])
        suffix = "" if len(missing_files) <= 3 else ", ..."
        messagebox.showwarning(
            "Missing Files",
            f"Some selected files were skipped because they no longer exist:\n{preview}{suffix}",
        )

    if not selected_file_infos:
        messagebox.showerror("Rotate PDF", "None of the selected PDFs are available.")
        if selected_file_infos_override is None:
            load_pending_files()
        return

    owner_window = root
    if parent_window is not None:
        try:
            if parent_window.winfo_exists():
                owner_window = parent_window
        except Exception:
            pass

    win = tk.Toplevel(owner_window)
    _apply_app_icon(win)
    win.title(window_title)
    configure_window_geometry(
        win,
        rotation_preview_window_width,
        rotation_preview_window_height,
        min_width=ROTATION_PREVIEW_MIN_WIDTH,
        min_height=ROTATION_PREVIEW_MIN_HEIGHT,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(owner_window)
    win.lift()
    win.focus_force()
    apply_theme(win)

    def _remember_rotation_window_size(event=None):
        global rotation_preview_window_width, rotation_preview_window_height
        if event is not None and event.widget is not win:
            return
        if not win.winfo_exists():
            return

        width = max(ROTATION_PREVIEW_MIN_WIDTH, int(win.winfo_width()))
        height = max(ROTATION_PREVIEW_MIN_HEIGHT, int(win.winfo_height()))
        rotation_preview_window_width = width
        rotation_preview_window_height = height

    def _close_rotation_window():
        global rotation_preview_thumb_width
        try:
            _remember_rotation_window_size()
            try:
                rotation_preview_thumb_width = int(float(thumbnail_size_var.get()))
                rotation_preview_thumb_width = max(ROTATION_PREVIEW_THUMB_MIN_WIDTH, rotation_preview_thumb_width)
                rotation_preview_thumb_width = min(ROTATION_PREVIEW_THUMB_MAX_LIMIT, rotation_preview_thumb_width)
            except Exception:
                pass
            save_settings()
        except Exception:
            pass
        if win.winfo_exists():
            win.destroy()

    win.protocol("WM_DELETE_WINDOW", _close_rotation_window)
    win.bind("<Configure>", _remember_rotation_window_size, add="+")

    status_var = tk.StringVar(value="Preparing page previews...")
    backup_before_rotate_var = tk.BooleanVar(value=True)
    initial_thumb_width = rotation_preview_thumb_width
    if initial_thumb_width < ROTATION_PREVIEW_THUMB_MIN_WIDTH or initial_thumb_width > ROTATION_PREVIEW_THUMB_MAX_LIMIT:
        initial_thumb_width = ROTATION_PREVIEW_THUMB_MAX_WIDTH
    thumbnail_size_var = tk.IntVar(value=int(initial_thumb_width))

    selected_row_paths = set()
    selected_pages_by_file = {}
    row_state_by_file = {}
    page_entries_by_file = {}
    page_entries = []
    row_widgets = []
    active_row_path = selected_file_infos[0]["path"] if selected_file_infos else None
    row_selection_anchor_path = active_row_path
    page_selection_anchor_by_file = {}
    selection_context = "rows"
    total_loaded_pages = 0
    preview_error_count = 0

    CTRL_MASK = 0x0004
    SHIFT_MASK = 0x0001

    row_bg_normal = BG_COLOR
    row_bg_active = "#17253c"
    row_border_normal = LISTBOX_BORDER
    row_border_selected = ACCENT_COLOR
    page_bg_normal = "#111a2b"
    page_bg_selected = "#1e3a67"
    page_border_normal = LISTBOX_BORDER
    page_border_selected = ACCENT_COLOR

    content = ttk.Frame(win, padding=(14, 12, 14, 14), style="Shell.TFrame")
    content.pack(fill="both", expand=True)

    header_card = ttk.Frame(content, style="HeaderCard.TFrame", padding=(12, 10))
    header_card.pack(fill="x", pady=(0, 10))
    ttk.Label(
        header_card,
        text="Batch PDF Rotation",
        style="HeaderSubheading.TLabel",
    ).pack(anchor="w")
    ttk.Label(
        header_card,
        text=(
            "Click a PDF row to target all pages, Ctrl+Click to multi-select, Shift+Click to range-select, "
            "Ctrl+A to select all in the active context, Shift+Scroll to move a row horizontally, "
            "and Ctrl+Scroll over previews to resize page thumbnails."
        ),
        style="HeaderSubheading.TLabel",
        wraplength=1020,
        justify="left",
    ).pack(anchor="w", pady=(4, 0))

    controls_card = ttk.Frame(content, style="Card.TFrame", padding=(12, 10))
    controls_card.pack(fill="x", pady=(0, 8))

    controls = ttk.Frame(controls_card, style="Card.TFrame")
    controls.pack(fill="x")

    ttk.Label(controls, text="Preview Page Size", style="FieldLabel.TLabel").pack(side="left")
    page_size_scale = ttk.Scale(
        controls,
        from_=ROTATION_PREVIEW_THUMB_MIN_WIDTH,
        to=ROTATION_PREVIEW_THUMB_MAX_LIMIT,
        orient="horizontal",
        length=220,
        variable=thumbnail_size_var,
    )
    page_size_scale.pack(side="left", padx=(8, 6))

    page_size_value_label = ttk.Label(controls, text=f"{thumbnail_size_var.get()} px", style="CardMuted.TLabel")
    page_size_value_label.pack(side="left", padx=(0, 12))

    ttk.Checkbutton(
        controls,
        text="Create Backup",
        variable=backup_before_rotate_var,
    ).pack(side="right")

    rotate_actions = ttk.Frame(content, style="ActionBar.TFrame", padding=(12, 10))
    rotate_actions.pack(fill="x", pady=(0, 8))
    ttk.Label(rotate_actions, text="Rotation Actions", style="ActionTitle.TLabel").pack(anchor="w")

    rotate_buttons = ttk.Frame(rotate_actions, style="ActionBar.TFrame")
    rotate_buttons.pack(fill="x", pady=(6, 0))

    rotate_left_button = ttk.Button(
        rotate_buttons,
        text="Rotate Left 90\N{DEGREE SIGN}",
        style="SecondaryAction.TButton",
    )
    rotate_left_button.pack(side="left")

    rotate_right_button = ttk.Button(
        rotate_buttons,
        text="Rotate Right 90\N{DEGREE SIGN}",
        style="SecondaryAction.TButton",
    )
    rotate_right_button.pack(side="left", padx=(8, 0))

    save_rotation_button = ttk.Button(
        rotate_buttons,
        text="Save Rotation",
        style="PrimaryAction.TButton",
    )
    save_rotation_button.pack(side="left", padx=(12, 0))

    ttk.Button(
        rotate_buttons,
        text="Close",
        command=_close_rotation_window,
        style="SecondaryAction.TButton",
    ).pack(side="right")

    ttk.Label(content, textvariable=status_var, style="StatusText.TLabel").pack(anchor="w", pady=(0, 8))

    preview_area = ttk.Frame(content, style="Card.TFrame", padding=8)
    preview_area.pack(fill="both", expand=True)

    preview_hscrollbar = ttk.Scrollbar(preview_area, orient="horizontal")
    preview_hscrollbar.pack(side="bottom", fill="x")

    preview_canvas = tk.Canvas(
        preview_area,
        highlightthickness=0,
        bg=BG_COLOR,
        bd=0,
    )
    preview_canvas.pack(side="left", fill="both", expand=True)

    preview_scrollbar = ttk.Scrollbar(preview_area, orient="vertical", command=preview_canvas.yview)
    preview_scrollbar.pack(side="right", fill="y")
    preview_canvas.configure(
        yscrollcommand=preview_scrollbar.set,
        xscrollcommand=preview_hscrollbar.set,
    )
    preview_hscrollbar.configure(command=preview_canvas.xview)

    _mark_widget_as_scroll_canvas(preview_canvas)
    _ensure_global_mousewheel_binding()

    preview_cards_frame = ttk.Frame(preview_canvas, style="Card.TFrame")
    preview_cards_window_id = preview_canvas.create_window((0, 0), window=preview_cards_frame, anchor="nw")
    preview_cards_frame.columnconfigure(0, weight=1)

    def _update_preview_scrollregion(_event=None):
        preview_canvas.configure(scrollregion=preview_canvas.bbox("all"))

    def _sync_preview_cards_width(_event=None):
        try:
            preview_canvas.itemconfigure(
                preview_cards_window_id,
                width=max(1, int(preview_canvas.winfo_width())),
            )
        except tk.TclError:
            return
        _update_preview_scrollregion()

    preview_canvas.bind("<Configure>", _sync_preview_cards_width, add="+")
    preview_cards_frame.bind("<Configure>", _update_preview_scrollregion)

    def _is_ctrl_pressed(event):
        return bool(int(getattr(event, "state", 0)) & CTRL_MASK)

    def _is_shift_pressed(event):
        return bool(int(getattr(event, "state", 0)) & SHIFT_MASK)

    def _ordered_selected_row_paths():
        ordered = []
        for info in selected_file_infos:
            file_path = info["path"]
            if file_path in selected_row_paths:
                ordered.append(file_path)
        return ordered

    def _ordered_display_row_paths():
        ordered = []
        for info in selected_file_infos:
            file_path = info["path"]
            if file_path in row_state_by_file:
                ordered.append(file_path)
        return ordered

    def _clear_page_selections_except(file_path):
        for existing_file_path, page_indexes in selected_pages_by_file.items():
            if existing_file_path != file_path and page_indexes:
                page_indexes.clear()

    def _apply_row_span_selection(anchor_file_path, end_file_path, additive=False):
        ordered_paths = _ordered_display_row_paths()
        if not ordered_paths or end_file_path not in ordered_paths:
            return

        if anchor_file_path not in ordered_paths:
            anchor_file_path = end_file_path

        start_index = ordered_paths.index(anchor_file_path)
        end_index = ordered_paths.index(end_file_path)
        left = min(start_index, end_index)
        right = max(start_index, end_index)
        span_selection = set(ordered_paths[left : right + 1])

        if additive:
            selected_row_paths.update(span_selection)
        else:
            selected_row_paths.clear()
            selected_row_paths.update(span_selection)

    def _apply_page_span_selection(file_path, anchor_page_index, end_page_index, additive=False):
        entries = page_entries_by_file.get(file_path, [])
        if not entries:
            return

        max_page_index = max(entry["page_index"] for entry in entries)
        start_page_index = max(0, min(int(anchor_page_index), max_page_index))
        target_page_index = max(0, min(int(end_page_index), max_page_index))

        left = min(start_page_index, target_page_index)
        right = max(start_page_index, target_page_index)
        span_selection = set(range(left, right + 1))

        selected_pages = selected_pages_by_file.setdefault(file_path, set())
        if additive:
            selected_pages.update(span_selection)
        else:
            selected_pages.clear()
            selected_pages.update(span_selection)

    def _row_has_unsaved_changes(file_path):
        for entry in page_entries_by_file.get(file_path, []):
            if int(entry.get("rotation_degrees", 0)) % 360:
                return True
        return False

    def _count_unsaved_files():
        return sum(1 for info in selected_file_infos if _row_has_unsaved_changes(info["path"]))

    def _count_selected_pages_total():
        return sum(len(indexes) for indexes in selected_pages_by_file.values())

    def _set_status(message=None):
        if message:
            status_var.set(message)
            return
        row_total = len(row_state_by_file)
        row_selected = sum(1 for path in selected_row_paths if path in row_state_by_file)
        unsaved_files = _count_unsaved_files()
        selected_pages_total = _count_selected_pages_total()
        status_var.set(
            f"Loaded {total_loaded_pages} page(s). Selected rows: {row_selected}/{row_total}. "
            f"Selected pages: {selected_pages_total}. Unsaved files: {unsaved_files}."
        )

    def _collect_rotation_changes_by_file():
        rotation_map = {}
        for entry in page_entries:
            page_rotation = int(entry.get("rotation_degrees", 0)) % 360
            if page_rotation == 0:
                continue

            file_path = entry["file_path"]
            rotation_map.setdefault(file_path, {})[entry["page_index"]] = page_rotation
        return rotation_map

    def _update_row_header(file_path):
        row_state = row_state_by_file.get(file_path)
        if not row_state:
            return

        file_info = row_state["file_info"]
        page_count = int(file_info.get("page_count", 0))
        has_unsaved = _row_has_unsaved_changes(file_path)
        selected_pages_count = len(selected_pages_by_file.get(file_path, set()))

        suffix = " *" if has_unsaved else ""
        page_word = "page" if page_count == 1 else "pages"
        if selected_pages_count > 0:
            detail = f" ({page_count} {page_word}, selected pages: {selected_pages_count})"
        else:
            detail = f" ({page_count} {page_word})"

        row_state["header_label"].configure(text=f"{file_info['name']}{suffix}{detail}")

    def _update_row_visual(file_path):
        row_state = row_state_by_file.get(file_path)
        if not row_state:
            return

        is_selected_row = file_path in selected_row_paths
        is_active_row = file_path == active_row_path

        row_bg = row_bg_active if is_active_row else row_bg_normal
        row_border = row_border_selected if is_selected_row else row_border_normal

        row_state["row_frame"].configure(
            bg=row_bg,
            highlightbackground=row_border,
            highlightcolor=row_border,
        )
        row_state["header_label"].configure(bg=row_bg)
        row_state["pages_canvas"].configure(bg=row_bg)
        row_state["pages_inner"].configure(bg=row_bg)

    def _update_page_visual(entry):
        file_path = entry["file_path"]
        page_index = entry["page_index"]
        is_selected_page = page_index in selected_pages_by_file.get(file_path, set())

        if is_selected_page:
            page_bg = page_bg_selected
            page_border = page_border_selected
        else:
            page_bg = page_bg_normal
            page_border = page_border_normal

        entry["card_frame"].configure(
            bg=page_bg,
            highlightbackground=page_border,
            highlightcolor=page_border,
        )

        for widget in entry.get("bg_widgets", []):
            try:
                widget.configure(bg=page_bg)
            except tk.TclError:
                pass

    def _refresh_all_row_visuals():
        for info in selected_file_infos:
            file_path = info["path"]
            _update_row_header(file_path)
            _update_row_visual(file_path)
            for entry in page_entries_by_file.get(file_path, []):
                _update_page_visual(entry)

    def _select_all_rows():
        nonlocal active_row_path, row_selection_anchor_path, selection_context
        selected_row_paths.clear()
        ordered_row_paths = []
        for info in selected_file_infos:
            file_path = info["path"]
            if file_path in row_state_by_file:
                ordered_row_paths.append(file_path)
                selected_row_paths.add(file_path)
        active_row_path = ordered_row_paths[0] if ordered_row_paths else None
        row_selection_anchor_path = active_row_path
        selection_context = "rows"
        _refresh_all_row_visuals()
        _set_status("Selected all PDF rows for rotation.")

    def _handle_row_click(event, file_path):
        nonlocal active_row_path, row_selection_anchor_path, selection_context
        selection_context = "rows"
        ctrl_pressed = _is_ctrl_pressed(event)
        shift_pressed = _is_shift_pressed(event)

        if shift_pressed:
            anchor_file_path = row_selection_anchor_path
            if anchor_file_path not in row_state_by_file:
                anchor_file_path = active_row_path if active_row_path in row_state_by_file else file_path
            _apply_row_span_selection(anchor_file_path, file_path, additive=ctrl_pressed)
            active_row_path = file_path
        elif ctrl_pressed:
            if file_path in selected_row_paths:
                selected_row_paths.remove(file_path)
                if active_row_path == file_path:
                    ordered_rows = _ordered_selected_row_paths()
                    active_row_path = ordered_rows[0] if ordered_rows else None
            else:
                selected_row_paths.add(file_path)
                active_row_path = file_path
            row_selection_anchor_path = file_path
        else:
            selected_row_paths.clear()
            selected_row_paths.add(file_path)
            active_row_path = file_path
            row_selection_anchor_path = file_path

        _refresh_all_row_visuals()
        _set_status()
        return "break"

    def _handle_page_click(event, file_path, page_index):
        nonlocal active_row_path, row_selection_anchor_path, selection_context
        selection_context = "pages"
        active_row_path = file_path
        row_selection_anchor_path = file_path

        ctrl_pressed = _is_ctrl_pressed(event)
        shift_pressed = _is_shift_pressed(event)

        if ctrl_pressed or shift_pressed:
            selected_row_paths.add(file_path)
        else:
            selected_row_paths.clear()
            selected_row_paths.add(file_path)
            _clear_page_selections_except(file_path)

        if shift_pressed:
            anchor_page_index = page_selection_anchor_by_file.get(file_path, page_index)
            _apply_page_span_selection(file_path, anchor_page_index, page_index, additive=ctrl_pressed)
            if file_path not in page_selection_anchor_by_file:
                page_selection_anchor_by_file[file_path] = page_index
            if not ctrl_pressed:
                _clear_page_selections_except(file_path)
        else:
            selected_pages = selected_pages_by_file.setdefault(file_path, set())
            if ctrl_pressed:
                if page_index in selected_pages:
                    selected_pages.remove(page_index)
                else:
                    selected_pages.add(page_index)
            else:
                selected_pages.clear()
                selected_pages.add(page_index)
            page_selection_anchor_by_file[file_path] = page_index

        _refresh_all_row_visuals()
        _set_status()
        return "break"

    def _handle_ctrl_select_all(event=None):
        nonlocal selection_context

        if selection_context == "pages" and active_row_path in page_entries_by_file:
            selected_pages_by_file[active_row_path] = {
                entry["page_index"] for entry in page_entries_by_file[active_row_path]
            }
            if selected_pages_by_file[active_row_path]:
                page_selection_anchor_by_file[active_row_path] = min(selected_pages_by_file[active_row_path])
            _refresh_all_row_visuals()
            _set_status("Selected all pages in the active PDF row.")
            return "break"

        _select_all_rows()
        return "break"

    win.bind("<Control-a>", _handle_ctrl_select_all, add="+")
    win.bind("<Control-A>", _handle_ctrl_select_all, add="+")

    def _bind_row_shift_scroll(widget, row_canvas):
        def _on_shift_wheel(event=None):
            if event is None or not _is_shift_pressed(event):
                return None

            if hasattr(event, "delta") and event.delta:
                steps = int(-1 * (event.delta / 120))
                if steps != 0:
                    row_canvas.xview_scroll(steps, "units")
                return "break"

            event_num = getattr(event, "num", None)
            if event_num == 4:
                row_canvas.xview_scroll(-1, "units")
                return "break"
            if event_num == 5:
                row_canvas.xview_scroll(1, "units")
                return "break"
            return None

        widget.bind("<MouseWheel>", _on_shift_wheel, add="+")
        widget.bind("<Button-4>", _on_shift_wheel, add="+")
        widget.bind("<Button-5>", _on_shift_wheel, add="+")

    def _bind_preview_ctrl_zoom_scroll(widget):
        def _on_ctrl_zoom_wheel(event=None):
            if event is None or not _is_ctrl_pressed(event):
                return None

            zoom_direction = 0
            if hasattr(event, "delta") and event.delta:
                zoom_direction = 1 if event.delta > 0 else -1
            else:
                event_num = getattr(event, "num", None)
                if event_num == 4:
                    zoom_direction = 1
                elif event_num == 5:
                    zoom_direction = -1

            if zoom_direction == 0:
                return "break"

            try:
                current_width = int(float(thumbnail_size_var.get()))
            except Exception:
                current_width = ROTATION_PREVIEW_THUMB_MAX_WIDTH

            target_width = current_width + (10 * zoom_direction)
            target_width = max(ROTATION_PREVIEW_THUMB_MIN_WIDTH, target_width)
            target_width = min(ROTATION_PREVIEW_THUMB_MAX_LIMIT, target_width)

            if target_width != current_width:
                thumbnail_size_var.set(target_width)
                _refresh_all_thumbnail_sizes()

            return "break"

        widget.bind("<MouseWheel>", _on_ctrl_zoom_wheel, add="+")
        widget.bind("<Button-4>", _on_ctrl_zoom_wheel, add="+")
        widget.bind("<Button-5>", _on_ctrl_zoom_wheel, add="+")

    def _calculate_page_strip_canvas_height():
        try:
            thumb_width = int(float(thumbnail_size_var.get()))
        except Exception:
            thumb_width = ROTATION_PREVIEW_THUMB_MAX_WIDTH

        thumb_width = max(ROTATION_PREVIEW_THUMB_MIN_WIDTH, thumb_width)
        thumb_width = min(ROTATION_PREVIEW_THUMB_MAX_LIMIT, thumb_width)

        estimated_page_height = int(thumb_width * 1.45)
        return max(200, min(520, estimated_page_height + 80))

    def _refresh_page_entry_preview(entry):
        page_rotation = int(entry.get("rotation_degrees", 0)) % 360
        page_label = entry.get("page_label")
        if page_label is not None and page_label.winfo_exists():
            if page_rotation:
                page_label.configure(
                    text=f"Page {entry['page_index'] + 1} / {entry['page_count']} | {page_rotation}\N{DEGREE SIGN}"
                )
            else:
                page_label.configure(text=f"Page {entry['page_index'] + 1} / {entry['page_count']}")

        image_label = entry.get("image_label")
        base_thumbnail = entry.get("base_thumbnail_pil")
        if image_label is None or not image_label.winfo_exists() or base_thumbnail is None:
            return

        render_image = base_thumbnail
        if page_rotation:
            render_image = base_thumbnail.rotate(-page_rotation, expand=True)

        image_photo = _build_pdf_thumbnail_photo(
            render_image,
            max_width=max(90, int(float(thumbnail_size_var.get()))),
        )
        if image_photo is None:
            return

        entry["thumbnail_photo"] = image_photo
        image_label.configure(image=image_photo)

    def _refresh_all_thumbnail_sizes(*_args):
        global rotation_preview_thumb_width

        try:
            rotation_preview_thumb_width = int(float(thumbnail_size_var.get()))
            rotation_preview_thumb_width = max(ROTATION_PREVIEW_THUMB_MIN_WIDTH, rotation_preview_thumb_width)
            rotation_preview_thumb_width = min(ROTATION_PREVIEW_THUMB_MAX_LIMIT, rotation_preview_thumb_width)
        except Exception:
            pass

        page_size_value_label.configure(text=f"{int(float(thumbnail_size_var.get()))} px")

        target_canvas_height = _calculate_page_strip_canvas_height()
        for row_state in row_state_by_file.values():
            pages_canvas_widget = row_state.get("pages_canvas")
            if pages_canvas_widget is None:
                continue
            try:
                if pages_canvas_widget.winfo_exists():
                    pages_canvas_widget.configure(height=target_canvas_height)
            except tk.TclError:
                pass

        for entry in page_entries:
            _refresh_page_entry_preview(entry)
        _set_status()
        _update_preview_scrollregion()

    page_size_scale.configure(command=lambda _value: _refresh_all_thumbnail_sizes())

    def _apply_preview_rotation(degrees):
        ordered_rows = _ordered_selected_row_paths()
        if not ordered_rows:
            messagebox.showwarning(
                "Rotate PDF",
                "Select at least one PDF row to rotate.",
                parent=win,
            )
            return

        updated_count = 0
        for file_path in ordered_rows:
            row_entries = page_entries_by_file.get(file_path, [])
            selected_pages = selected_pages_by_file.get(file_path, set())

            for entry in row_entries:
                if selected_pages and entry["page_index"] not in selected_pages:
                    continue

                entry["rotation_degrees"] = (
                    int(entry.get("rotation_degrees", 0)) + int(degrees)
                ) % 360
                _refresh_page_entry_preview(entry)
                updated_count += 1

            _update_row_header(file_path)
            _update_row_visual(file_path)

        if updated_count > 0:
            pending_changes = _collect_rotation_changes_by_file()
            changed_files = len(pending_changes)
            status_var.set(
                f"Preview updated for {updated_count} page(s). Unsaved changes in {changed_files} file(s)."
            )
            _update_preview_scrollregion()

    def _clear_page_rows():
        for row in row_widgets:
            try:
                row.destroy()
            except tk.TclError:
                pass

        page_entries.clear()
        row_widgets.clear()

    def _populate_page_rows():
        _clear_page_rows()

        nonlocal active_row_path, row_selection_anchor_path, selection_context, total_loaded_pages, preview_error_count
        total_loaded_pages = 0
        preview_error_count = 0

        selected_row_paths.clear()
        selected_pages_by_file.clear()
        page_selection_anchor_by_file.clear()
        row_state_by_file.clear()
        page_entries_by_file.clear()

        displayable_files = 0

        for file_info in selected_file_infos:
            file_name = file_info["name"]
            file_path = file_info["path"]

            if not os.path.exists(file_path):
                preview_error_count += 1
                continue

            page_count = get_pdf_page_count(file_path)
            if not page_count:
                preview_error_count += 1
                continue

            file_info["page_count"] = page_count
            displayable_files += 1

            selected_row_paths.add(file_path)
            selected_pages_by_file[file_path] = set()

            file_row = tk.Frame(
                preview_cards_frame,
                bg=row_bg_normal,
                highlightthickness=2,
                highlightbackground=row_border_normal,
                highlightcolor=row_border_normal,
                bd=0,
                padx=8,
                pady=8,
            )
            file_row.grid(row=displayable_files - 1, column=0, sticky="ew", padx=6, pady=6)
            file_row.columnconfigure(0, weight=1)
            row_widgets.append(file_row)

            file_header = tk.Label(
                file_row,
                text=f"{file_name} ({page_count} page{'s' if page_count != 1 else ''})",
                bg=row_bg_normal,
                fg=TEXT_COLOR,
                font=("Segoe UI Semibold", 10),
                justify="left",
                wraplength=980,
                anchor="w",
            )
            file_header.grid(row=0, column=0, sticky="w")

            pages_row = tk.Frame(file_row, bg=row_bg_normal)
            pages_row.grid(row=1, column=0, sticky="we", pady=(6, 0))
            pages_row.columnconfigure(0, weight=1)

            pages_canvas = tk.Canvas(
                pages_row,
                bg=row_bg_normal,
                bd=0,
                highlightthickness=0,
                height=_calculate_page_strip_canvas_height(),
            )
            pages_canvas.pack(side="top", fill="x", expand=True)

            pages_hscrollbar = ttk.Scrollbar(pages_row, orient="horizontal", command=pages_canvas.xview)
            pages_hscrollbar.pack(side="top", fill="x", pady=(4, 0))
            pages_canvas.configure(xscrollcommand=pages_hscrollbar.set)

            pages_inner = tk.Frame(pages_canvas, bg=row_bg_normal)
            pages_canvas.create_window((0, 0), window=pages_inner, anchor="nw")

            pages_inner.bind(
                "<Configure>",
                lambda _event, target_canvas=pages_canvas: target_canvas.configure(scrollregion=target_canvas.bbox("all")),
                add="+",
            )

            row_state_by_file[file_path] = {
                "file_info": file_info,
                "row_frame": file_row,
                "header_label": file_header,
                "pages_canvas": pages_canvas,
                "pages_inner": pages_inner,
            }
            page_entries_by_file[file_path] = []

            for row_click_widget in (file_row, file_header, pages_row):
                row_click_widget.bind(
                    "<Button-1>",
                    lambda event, path=file_path: _handle_row_click(event, path),
                    add="+",
                )
                _bind_row_shift_scroll(row_click_widget, pages_canvas)

            _bind_row_shift_scroll(pages_canvas, pages_canvas)
            _bind_row_shift_scroll(pages_inner, pages_canvas)
            _bind_preview_ctrl_zoom_scroll(pages_canvas)
            _bind_preview_ctrl_zoom_scroll(pages_inner)

            pdf_document = None
            pdf_pages = None
            if pdfplumber is not None and Image is not None and ImageTk is not None:
                try:
                    pdf_document = pdfplumber.open(file_path)
                    pdf_pages = pdf_document.pages
                except Exception:
                    pdf_document = None
                    pdf_pages = None

            try:
                for page_index in range(page_count):
                    thumbnail_pil = None
                    if pdf_pages is not None and page_index < len(pdf_pages):
                        thumbnail_pil = _create_pdf_page_thumbnail(pdf_pages[page_index], max_width=420)

                    card = tk.Frame(
                        pages_inner,
                        bg=page_bg_normal,
                        highlightthickness=2,
                        highlightbackground=page_border_normal,
                        highlightcolor=page_border_normal,
                        bd=0,
                        padx=6,
                        pady=6,
                    )
                    card.pack(side="left", padx=(0, 8), pady=(0, 4), anchor="n")

                    page_title = tk.Label(
                        card,
                        text=f"Page {page_index + 1}",
                        bg=page_bg_normal,
                        fg=TEXT_COLOR,
                        font=("Segoe UI", 9),
                        anchor="w",
                    )
                    page_title.pack(anchor="w")

                    image_label = None
                    thumbnail_photo = None
                    bg_widgets = [card, page_title]
                    clickable_widgets = [card, page_title]

                    if thumbnail_pil is not None:
                        thumbnail_photo = _build_pdf_thumbnail_photo(
                            thumbnail_pil,
                            max_width=max(90, int(float(thumbnail_size_var.get()))),
                        )
                        image_label = tk.Label(
                            card,
                            image=thumbnail_photo,
                            bg=page_bg_normal,
                            fg=LISTBOX_TEXT,
                            bd=0,
                            padx=4,
                            pady=4,
                        )
                        image_label.pack(anchor="w", pady=(6, 4))
                        bg_widgets.append(image_label)
                        clickable_widgets.append(image_label)
                    else:
                        placeholder = tk.Label(
                            card,
                            text="Preview unavailable",
                            bg=page_bg_normal,
                            fg=LISTBOX_TEXT,
                            padx=8,
                            pady=18,
                            width=22,
                            anchor="center",
                            justify="center",
                        )
                        placeholder.pack(anchor="w", pady=(6, 4))
                        bg_widgets.append(placeholder)
                        clickable_widgets.append(placeholder)

                    page_label = tk.Label(
                        card,
                        text=f"Page {page_index + 1} / {page_count}",
                        bg=page_bg_normal,
                        fg=TEXT_COLOR,
                        font=("Segoe UI", 9),
                        anchor="w",
                    )
                    page_label.pack(anchor="w")
                    bg_widgets.append(page_label)
                    clickable_widgets.append(page_label)

                    for widget in clickable_widgets:
                        widget.bind(
                            "<Button-1>",
                            lambda event, path=file_path, idx=page_index: _handle_page_click(event, path, idx),
                            add="+",
                        )
                        _bind_row_shift_scroll(widget, pages_canvas)
                        _bind_preview_ctrl_zoom_scroll(widget)

                    entry = {
                        "file_name": file_name,
                        "file_path": file_path,
                        "page_index": page_index,
                        "page_count": page_count,
                        "page_label": page_label,
                        "image_label": image_label,
                        "base_thumbnail_pil": thumbnail_pil,
                        "thumbnail_photo": thumbnail_photo,
                        "rotation_degrees": 0,
                        "card_frame": card,
                        "bg_widgets": bg_widgets,
                    }
                    page_entries.append(entry)
                    page_entries_by_file[file_path].append(entry)
                    total_loaded_pages += 1
            finally:
                if pdf_document is not None:
                    try:
                        pdf_document.close()
                    except Exception:
                        pass

        if not row_widgets:
            empty_label = ttk.Label(
                preview_cards_frame,
                text="No pages are available for preview in the selected files.",
                style="Subheading.TLabel",
                justify="left",
                wraplength=620,
            )
            empty_label.grid(row=0, column=0, padx=6, pady=6, sticky="w")
            row_widgets.append(empty_label)

        ordered_rows = _ordered_selected_row_paths()
        active_row_path = ordered_rows[0] if ordered_rows else None
        row_selection_anchor_path = active_row_path
        selection_context = "rows"

        _refresh_all_row_visuals()
        _update_preview_scrollregion()
        _set_status(
            f"Loaded {total_loaded_pages} page(s) across {len(row_state_by_file)} selected PDF file(s). "
            f"Preview errors: {preview_error_count}. Rotate for preview, then click Save Rotation to write changes."
        )

    def _save_rotations_to_files():
        rotation_changes = _collect_rotation_changes_by_file()
        if not rotation_changes:
            messagebox.showinfo(
                "Rotate PDF",
                "No pending rotation changes to save.",
                parent=win,
            )
            return

        files_with_changes = len(rotation_changes)
        if not messagebox.askyesno(
            "Save Rotation",
            f"Save rotation changes for {files_with_changes} PDF file(s)?",
            parent=win,
        ):
            return

        rotated_files = 0
        rotated_pages = 0
        failed_items = []

        for file_info in selected_file_infos:
            file_path = file_info["path"]
            page_rotation_map = rotation_changes.get(file_path)
            if not page_rotation_map:
                continue

            try:
                if backup_before_rotate_var.get():
                    create_backup_file(file_path)

                rotated_count = _rotate_pdf_pages_in_place(
                    file_path,
                    page_rotation_map=page_rotation_map,
                )
                rotated_pages += max(0, int(rotated_count))
                rotated_files += 1
            except Exception as exc:
                failed_items.append(f"{file_info['name']}: {exc}")

        if callable(post_save_callback):
            try:
                post_save_callback()
            except Exception as exc:
                messagebox.showwarning(
                    "Rotate PDF",
                    f"Saved changes, but refresh callback failed: {exc}",
                    parent=win,
                )
        elif selected_file_infos_override is None:
            load_pending_files()
        if rotated_files > 0:
            status_var.set("Saved changes. Refreshing preview from updated PDFs...")
            win.update_idletasks()
            _populate_page_rows()

        if failed_items and rotated_files == 0:
            details = "\n".join(failed_items[:8])
            if len(failed_items) > 8:
                details += "\n..."
            messagebox.showerror(
                "Rotate PDF",
                "Saving rotation failed for all files.\n\n" + details,
                parent=win,
            )
            return

        if failed_items:
            details = "\n".join(failed_items[:6])
            if len(failed_items) > 6:
                details += "\n..."
            messagebox.showwarning(
                "Rotate PDF",
                f"Saved rotation for {rotated_files} file(s), {rotated_pages} page(s).\n"
                f"Failed files: {len(failed_items)}\n\n{details}",
                parent=win,
            )
            return

        _set_status(f"Saved rotation changes for {rotated_files} file(s), {rotated_pages} page(s).")
        messagebox.showinfo(
            "Rotate PDF",
            f"Rotation saved successfully for {rotated_files} file(s), {rotated_pages} page(s).",
            parent=win,
        )

    rotate_left_button.configure(command=lambda: _apply_preview_rotation(-90))
    rotate_right_button.configure(command=lambda: _apply_preview_rotation(90))
    save_rotation_button.configure(command=_save_rotations_to_files)

    status_var.set("Loading page previews...")
    win.update_idletasks()
    _populate_page_rows()


def get_selected_pending_files():
    return [name for name, var in pending_file_vars.items() if var.get()]


def preview_specific_pending_pdf(filename):
    folder = normalize_path(pending_folder.get())
    if not folder:
        messagebox.showwarning("Warning", "Select the pending folder first.")
        return

    file_path = normalize_path(os.path.join(folder, filename))
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"{filename} no longer exists at the pending location.")
        load_pending_files()
        return

    try:
        _launch_path(file_path)
    except RuntimeError as exc:
        messagebox.showerror("Error", f"Could not open PDF: {exc}")


def new_record_window(initial_filename=None, batch_context=None, on_complete=None):
    filename = initial_filename
    if not filename:
        selected = get_selected_pending_files()
        if not selected:
            messagebox.showwarning("Warning", "Select at least one pending PDF to create a new record.")
            return
        filename = selected[0]

    def abort_early():
        if on_complete:
            on_complete(False, filename)

    pending_dir = normalize_path(pending_folder.get())
    if not pending_dir:
        messagebox.showwarning("Warning", "Select the pending folder first.")
        abort_early()
        return

    pending_file_path = normalize_path(os.path.join(pending_dir, filename))
    if not os.path.exists(pending_file_path):
        messagebox.showerror("Error", "Pending file no longer exists. Refresh the list.")
        load_pending_files()
        abort_early()
        return

    win = tk.Toplevel(root)
    _apply_app_icon(win)
    win.title("New Record")
    configure_window_geometry(
        win,
        520,
        670,
        min_width=460,
        min_height=560,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(root)
    win.grab_set()
    win.focus_force()
    apply_theme(win)
    scroll_container, scroll_frame = create_scrollable_panel(win)
    scroll_container.pack(fill="both", expand=True)
    content = ttk.Frame(scroll_frame, padding=(18, 16, 18, 18), style="Shell.TFrame")
    content.pack(fill="both", expand=True)

    processed_successfully = False
    completion_reported = False

    def notify_completion():
        nonlocal completion_reported
        if on_complete and not completion_reported:
            on_complete(processed_successfully, filename)
            completion_reported = True

    def cancel_batch_processing_from_new_record():
        if batch_context is not None:
            batch_context["cancelled"] = True
        close_window()

    # Variables
    name_var = tk.StringVar()
    letter_var = tk.StringVar()
    status_var = tk.StringVar(value="Active")
    new_year_var = tk.StringVar()
    old_year_var = tk.StringVar()
    year_hint_var = tk.StringVar(value=_get_year_input_guidance("", ""))
    dest_path_var = tk.StringVar(value="Select root folder and enter the name to preview the destination path.")

    root_trace_id = None

    # Auto update letter based on name
    def refresh_destination_path(*args):
        root_path = normalize_path(root_folder.get().strip())
        name = name_var.get().strip()
        current_letter = letter_var.get().strip().upper()
        status_value = status_var.get()

        if not root_path:
            dest_path_var.set("Select Records Root Folder to calculate destination path.")
            return
        if not name:
            dest_path_var.set("Enter the employee name to calculate destination path.")
            return

        letter_segment = current_letter or (name[0].upper() if name else "#")
        preview_path = os.path.join(root_path, status_value, letter_segment, name)
        dest_path_var.set(normalize_path(preview_path))

    def update_letter(*args):
        name = name_var.get()
        if name:
            letter_var.set(name[0].upper())
        refresh_destination_path()

    def refresh_year_hint(*_args):
        year_hint_var.set(_get_year_input_guidance(new_year_var.get(), old_year_var.get()))

    name_var.trace_add("write", update_letter)
    letter_var.trace_add("write", refresh_destination_path)
    status_var.trace_add("write", refresh_destination_path)
    new_year_var.trace_add("write", refresh_year_hint)
    old_year_var.trace_add("write", refresh_year_hint)
    root_trace_id = root_folder.trace_add("write", refresh_destination_path)

    parsed_pending_metadata = parse_filename_metadata(filename)
    if parsed_pending_metadata:
        parsed_name = parsed_pending_metadata.get("name", "").strip()
        if parsed_name:
            name_var.set(parsed_name)
        parsed_latest = parsed_pending_metadata.get("latest", "").strip()
        parsed_earliest = parsed_pending_metadata.get("earliest", "").strip()
        if parsed_latest:
            new_year_var.set(parsed_latest)
        if parsed_earliest:
            old_year_var.set(parsed_earliest)

    header = ttk.Frame(content, style="HeaderCard.TFrame", padding=(14, 12))
    header.pack(fill="x", pady=(0, 12))
    ttk.Label(header, text="New Record Setup", style="HeaderSubheading.TLabel").pack(anchor="w")
    ttk.Label(header, text=f"Pending file: {filename}", style="HeaderSubheading.TLabel").pack(anchor="w", pady=(2, 0))
    header_actions = ttk.Frame(header, style="HeaderCard.TFrame")
    header_actions.pack(anchor="w", pady=(8, 0))
    ttk.Button(
        header_actions,
        text="Preview Pending",
        width=16,
        command=lambda: preview_specific_pending_pdf(filename),
        style="Subtle.TButton",
    ).pack(side="left")
    if batch_context:
        ttk.Label(
            header,
            text=f"Batch {batch_context['current']} of {batch_context['total']}",
            style="HeaderSubheading.TLabel",
            padding=4,
        ).pack(anchor="w", pady=(4, 0))
        ttk.Button(
            header_actions,
            text="View Batch Files",
            width=16,
            command=lambda: show_selected_batch_files_window(
                batch_files=batch_context.get("files"),
                current_filename=filename,
                parent_window=win,
                batch_context=batch_context,
                cancel_batch_callback=cancel_batch_processing_from_new_record,
            ),
            style="Subtle.TButton",
        ).pack(side="left", padx=(8, 0))

    metadata_card = ttk.Frame(content, style="Card.TFrame", padding=14)
    metadata_card.pack(fill="x", pady=(0, 10))
    metadata_card.columnconfigure(0, weight=1)
    metadata_card.columnconfigure(1, weight=1)

    ttk.Label(metadata_card, text="Record Metadata", style="SectionTitle.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w"
    )

    ttk.Label(metadata_card, text="Employee Name", style="FieldLabel.TLabel").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(10, 2)
    )
    name_field = ttk.Combobox(
        metadata_card,
        textvariable=name_var,
        values=employee_name_suggestions,
        state="normal",
    )
    _prevent_combobox_mousewheel_value_change(name_field)

    def _handle_name_key(event=None):
        # Keep user input untouched and only refresh/open suggestion list.
        _update_combobox_suggestions(name_field, name_var.get(), event)

    name_field.bind("<KeyRelease>", _handle_name_key)

    def _refresh_name_choices():
        name_field["values"] = get_filtered_name_suggestions(name_var.get())

    name_field.configure(postcommand=_refresh_name_choices)
    name_field.grid(row=2, column=0, columnspan=2, sticky="ew")

    ttk.Label(metadata_card, text="Surname Initial", style="FieldLabel.TLabel").grid(
        row=3, column=0, sticky="w", pady=(10, 2)
    )
    ttk.Entry(metadata_card, textvariable=letter_var, width=8).grid(
        row=4, column=0, sticky="ew", padx=(0, 8)
    )

    ttk.Label(metadata_card, text="Status", style="FieldLabel.TLabel").grid(
        row=3, column=1, sticky="w", pady=(10, 2)
    )
    status_field = ttk.Combobox(
        metadata_card,
        textvariable=status_var,
        values=("Active", "Retiree"),
        state="readonly",
    )
    _prevent_combobox_mousewheel_value_change(status_field)
    status_field.grid(row=4, column=1, sticky="ew")

    years_card = ttk.Frame(content, style="Card.TFrame", padding=14)
    years_card.pack(fill="x", pady=(0, 10))
    years_card.columnconfigure(0, weight=1)
    years_card.columnconfigure(1, weight=1)

    ttk.Label(years_card, text="Year Range", style="SectionTitle.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w"
    )
    ttk.Label(years_card, text="Latest Year", style="FieldLabel.TLabel").grid(
        row=1, column=0, sticky="w", pady=(10, 2)
    )
    ttk.Entry(years_card, textvariable=new_year_var).grid(row=2, column=0, sticky="ew", padx=(0, 8))

    ttk.Label(years_card, text="Oldest Year", style="FieldLabel.TLabel").grid(
        row=1, column=1, sticky="w", pady=(10, 2)
    )
    ttk.Entry(years_card, textvariable=old_year_var).grid(row=2, column=1, sticky="ew")
    ttk.Label(
        years_card,
        textvariable=year_hint_var,
        style="CardMuted.TLabel",
        wraplength=460,
        justify="left",
    ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

    preview_card = ttk.Frame(content, style="Card.TFrame", padding=14)
    preview_card.pack(fill="x")
    ttk.Label(preview_card, text="Destination Path Preview", style="SectionTitle.TLabel").pack(anchor="w")
    ttk.Label(
        preview_card,
        textvariable=dest_path_var,
        wraplength=460,
        justify="left",
        style="CardMuted.TLabel",
        anchor="w",
        padding=(0, 6),
    ).pack(fill="x", pady=(6, 0))

    refresh_destination_path()
    refresh_year_hint()

    def close_window():
        nonlocal root_trace_id
        if root_trace_id is not None:
            root_folder.trace_remove("write", root_trace_id)
            root_trace_id = None
        try:
            if win.grab_current() is win:
                win.grab_release()
        except tk.TclError:
            pass
        win.destroy()
        notify_completion()

    def save_record():
        nonlocal processed_successfully
        if not name_var.get().strip():
            messagebox.showerror("Error", "Name is required.")
            return

        root_path = normalize_path(root_folder.get())
        if not root_path:
            messagebox.showerror("Error", "Select the Records Root Folder first so the destination can be created.")
            return

        pending_path = normalize_path(pending_folder.get())
        if not pending_path:
            messagebox.showerror("Error", "Select the Pending Folder first so the source PDF can be located.")
            return

        name = name_var.get().strip()
        letter_value = letter_var.get().strip().upper() or (name[0].upper() if name else "")
        letter = letter_value if letter_value else "#"
        status = status_var.get()
        new_year_str = new_year_var.get().strip()
        old_year_str = old_year_var.get().strip()

        try:
            years = _normalize_record_year_inputs(new_year_str, old_year_str)
        except ValueError:
            messagebox.showerror("Year Validation", _get_year_input_guidance(new_year_str, old_year_str))
            return
        if years is None:
            messagebox.showerror("Year Validation", _get_year_input_guidance(new_year_str, old_year_str))
            return
        latest_year, earliest_year = years

        # Build path
        target_folder = normalize_path(os.path.join(root_path, status, letter, name))

        if os.path.exists(target_folder):
            choice = messagebox.askyesnocancel(
                "Folder Exists",
                f"The destination folder already exists:\n{target_folder}\n\n"
                "Yes = continue saving\nNo = open folder to review\nCancel = keep editing"
            )
            if choice is None:
                return
            if choice is False:
                try:
                    _launch_path(target_folder)
                except RuntimeError as exc:
                    messagebox.showerror("Error", f"Unable to open folder: {exc}")
                return
        else:
            os.makedirs(target_folder, exist_ok=True)

        # New filename
        new_filename = _build_record_filename(name, latest_year, earliest_year)
        new_path = normalize_path(os.path.join(target_folder, new_filename))

        # Check duplicate
        if os.path.exists(new_path):
            messagebox.showerror("Error", "File already exists.")
            return

        confirm = messagebox.askyesno(
            "Confirm Save",
            f"Convert {filename}\ninto {new_filename}\nand move it to:\n{target_folder}"
        )
        if not confirm:
            return

        # Save a copy to the destination, then archive the pending original.
        src_path = normalize_path(os.path.join(pending_path, filename))
        if not os.path.exists(src_path):
            messagebox.showerror("Error", "Pending file no longer exists. Refresh the list.")
            load_pending_files()
            return

        try:
            shutil.copy2(src_path, new_path)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to save file: {exc}")
            return

        try:
            archived_path = archive_pending_file(src_path)
            if not archived_path:
                raise OSError("Pending file was not found during archiving.")
        except OSError as exc:
            try:
                if os.path.exists(new_path):
                    os.remove(new_path)
            except OSError:
                pass
            messagebox.showerror(
                "Error",
                f"Unable to move pending PDF to the processed folder: {exc}",
            )
            return

        processed_successfully = True
        messagebox.showinfo("Success", "File saved successfully.")

        close_window()
        load_pending_files()

    action_frame = ttk.Frame(content, style="ActionBar.TFrame", padding=(12, 10))
    action_frame.pack(fill="x", pady=(14, 0))
    ttk.Label(action_frame, text="Save Action", style="ActionTitle.TLabel").pack(anchor="w")

    action_buttons = ttk.Frame(action_frame, style="ActionBar.TFrame")
    action_buttons.pack(fill="x", pady=(6, 0))

    ttk.Button(
        action_buttons,
        text="Close",
        command=close_window,
        style="SecondaryAction.TButton",
        width=14,
    ).pack(side="left")
    if batch_context:
        ttk.Button(
            action_buttons,
            text="Cancel Batch",
            command=cancel_batch_processing_from_new_record,
            style="Subtle.TButton",
            width=14,
        ).pack(side="left", padx=(8, 0))

    ttk.Button(
        action_buttons,
        text="Save New Record",
        command=save_record,
        style="PrimaryAction.TButton",
        width=18,
    ).pack(side="right")
    win.bind("<Escape>", lambda _event: (close_window(), "break")[1])
    win.after_idle(lambda: name_field.focus_set())
    win.protocol("WM_DELETE_WINDOW", close_window)


def merge_existing_window(pending_filename=None, batch_context=None, on_complete=None):
    filename = pending_filename
    if not filename:
        selected = get_selected_pending_files()
        if not selected:
            messagebox.showwarning("Warning", "Select at least one pending PDF to merge into an existing record.")
            return
        filename = selected[0]
    pending_filename = filename

    def abort_early():
        if on_complete:
            on_complete(False, pending_filename)

    if not pending_folder.get():
        messagebox.showerror("Error", "Select the pending folder first.")
        abort_early()
        return

    pending_path = normalize_path(os.path.join(pending_folder.get(), filename))
    if not os.path.exists(pending_path):
        messagebox.showerror("Error", "Pending file no longer exists. Refresh the list.")
        load_pending_files()
        abort_early()
        return

    root_path = normalize_path(root_folder.get())
    if not root_path:
        messagebox.showerror("Error", "Select the records root folder first.")
        abort_early()
        return

    win = tk.Toplevel(root)
    _apply_app_icon(win)
    win.title("Merge Existing Record")
    configure_window_geometry(
        win,
        640,
        900,
        min_width=560,
        min_height=700,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(root)
    win.lift()
    win.focus_force()
    win.grab_set()
    apply_theme(win)
    scroll_container, scroll_frame = create_scrollable_panel(win)
    scroll_container.pack(fill="both", expand=True)
    content = ttk.Frame(scroll_frame, padding=(18, 16, 18, 18), style="Shell.TFrame")
    content.pack(fill="both", expand=True)

    processed_successfully = False
    completion_reported = False

    def notify_completion():
        nonlocal completion_reported
        if on_complete and not completion_reported:
            on_complete(processed_successfully, pending_filename)
            completion_reported = True

    def cancel_batch_processing_from_merge():
        if batch_context is not None:
            batch_context["cancelled"] = True
        close_merge_window()

    def keep_merge_window_on_top():
        try:
            win.lift()
            win.attributes("-topmost", True)
            win.after(50, lambda: win.attributes("-topmost", False))
            win.focus_force()
        except tk.TclError:
            pass

    header = ttk.Frame(content, style="HeaderCard.TFrame", padding=(14, 12))
    header.pack(fill="x", pady=(0, 12))
    ttk.Label(header, text="Merge Existing Record", style="HeaderSubheading.TLabel").pack(anchor="w")
    ttk.Label(header, text=f"Pending file: {pending_filename}", style="HeaderSubheading.TLabel").pack(anchor="w", pady=(2, 0))
    header_actions = ttk.Frame(header, style="HeaderCard.TFrame")
    header_actions.pack(anchor="w", pady=(8, 0))
    ttk.Button(
        header_actions,
        text="Preview Pending",
        width=16,
        command=lambda: preview_specific_pending_pdf(pending_filename),
        style="Subtle.TButton",
    ).pack(side="left")
    if batch_context:
        ttk.Label(
            header,
            text=f"Batch {batch_context['current']} of {batch_context['total']}",
            style="HeaderSubheading.TLabel",
            padding=4,
        ).pack(anchor="w", pady=(4, 0))
        ttk.Button(
            header_actions,
            text="View Batch Files",
            width=16,
            command=lambda: show_selected_batch_files_window(
                batch_files=batch_context.get("files"),
                current_filename=pending_filename,
                parent_window=win,
                batch_context=batch_context,
                cancel_batch_callback=cancel_batch_processing_from_merge,
            ),
            style="Subtle.TButton",
        ).pack(side="left", padx=(8, 0))

    def close_merge_window():
        try:
            if win.grab_current() is win:
                win.grab_release()
        except tk.TclError:
            pass
        win.destroy()
        notify_completion()

    folder_var = tk.StringVar()
    existing_selection_var = tk.StringVar(value="No file selected")
    existing_selected_pdf_var = tk.StringVar(value="")
    name_var = tk.StringVar()
    letter_var = tk.StringVar()
    status_var = tk.StringVar(value="Active")
    new_year_var = tk.StringVar()
    old_year_var = tk.StringVar()
    year_hint_var = tk.StringVar(value=_get_year_input_guidance("", ""))
    dest_preview_var = tk.StringVar(value="Select destination folder and fill out fields to preview the final file path.")
    merge_summary_var = tk.StringVar(value="Select an existing PDF to view page counts.")
    existing_list_hint_var = tk.StringVar(value="Select a destination folder to load existing PDFs.")
    keep_backup_var = keep_backup_preference_var
    folder_path_suggestions = []
    folder_suggestion_entries = []
    folder_suggestion_label_to_path = {}
    existing_files_frame = None

    def update_merge_summary():
        if not os.path.exists(pending_path):
            merge_summary_var.set("Pending PDF not found; refresh the queue.")
            return

        pending_pages = get_pdf_page_count(pending_path)
        pending_text = (
            f"Pending: {pending_pages} pages" if pending_pages is not None else "Pending pages unavailable"
        )

        folder = normalize_path(folder_var.get().strip())
        existing_path_local = None
        selected_existing_filename = existing_selected_pdf_var.get().strip()
        if folder and selected_existing_filename:
            existing_candidate = normalize_path(os.path.join(folder, selected_existing_filename))
            if os.path.exists(existing_candidate):
                existing_path_local = existing_candidate

        if existing_path_local:
            existing_pages = get_pdf_page_count(existing_path_local)
            existing_text = (
                f"Existing: {existing_pages} pages" if existing_pages is not None else "Existing pages unavailable"
            )
        else:
            existing_pages = None
            existing_text = "Existing: not selected"

        total_text = "Merged total unknown"
        if pending_pages is not None and existing_pages is not None:
            total_text = f"After merge: {pending_pages + existing_pages} pages"

        final_preview = normalize_path(dest_preview_var.get().strip()) if dest_preview_var.get().strip() else ""
        overwrite_note = ""
        if final_preview and os.path.exists(final_preview):
            overwrite_note = "Final file already exists and will be replaced."

        summary_lines = [pending_text, existing_text, total_text]
        if overwrite_note:
            summary_lines.append(overwrite_note)

        merge_summary_var.set(" | ".join(summary_lines))

    def refresh_destination_preview(*args):
        folder = normalize_path(folder_var.get().strip())
        name = name_var.get().strip()
        letter = letter_var.get().strip().upper() or (name[0].upper() if name else "")
        latest = new_year_var.get().strip()
        earliest = old_year_var.get().strip()
        year_hint_var.set(_get_year_input_guidance(latest, earliest))

        if not folder:
            dest_preview_var.set("Select destination folder to preview final path.")
            update_merge_summary()
            return
        if not name:
            dest_preview_var.set("Enter the employee name to preview final path.")
            update_merge_summary()
            return

        try:
            years = _normalize_record_year_inputs(latest, earliest)
        except ValueError:
            dest_preview_var.set("Year fields must be numeric to preview final path.")
            update_merge_summary()
            return
        if years is None:
            dest_preview_var.set("Enter at least one year to preview final path.")
            update_merge_summary()
            return

        latest_val, earliest_val = years
        filename = _build_record_filename(name, latest_val, earliest_val)
        dest_preview_var.set(normalize_path(os.path.join(folder, filename)))

        update_merge_summary()

    def update_letter_from_name(*args):
        name = name_var.get()
        if name:
            letter_var.set(name[0].upper())
        refresh_destination_preview()

    name_var.trace_add("write", update_letter_from_name)
    letter_var.trace_add("write", refresh_destination_preview)
    new_year_var.trace_add("write", refresh_destination_preview)
    old_year_var.trace_add("write", refresh_destination_preview)
    folder_var.trace_add("write", refresh_destination_preview)

    parsed_pending_metadata = parse_filename_metadata(pending_filename)
    if parsed_pending_metadata:
        parsed_name = parsed_pending_metadata.get("name", "").strip()
        if parsed_name:
            name_var.set(parsed_name)
        parsed_latest = parsed_pending_metadata.get("latest", "").strip()
        parsed_earliest = parsed_pending_metadata.get("earliest", "").strip()
        if parsed_latest:
            new_year_var.set(parsed_latest)
        if parsed_earliest:
            old_year_var.set(parsed_earliest)

    def ensure_folder_under_root(chosen_folder):
        chosen_folder = normalize_path(chosen_folder)
        try:
            common = os.path.commonpath([os.path.abspath(chosen_folder), os.path.abspath(root_path)])
        except ValueError:
            return False
        return common == os.path.abspath(root_path)

    def prefill_from_folder(folder_path):
        folder_path = normalize_path(folder_path)
        relative = os.path.relpath(folder_path, root_path)
        parts = relative.split(os.sep)
        if parts:
            status_candidate = parts[0]
            if status_candidate in ("Active", "Retiree"):
                status_var.set(status_candidate)
        folder_name = os.path.basename(folder_path)
        if folder_name:
            name_var.set(folder_name)
            letter_var.set(folder_name[0].upper())

    def _validate_windows_path_component(raw_value, value_label):
        value = str(raw_value or "").strip()
        if not value:
            raise ValueError(f"{value_label} cannot be empty.")
        if value in {".", ".."}:
            raise ValueError(f"{value_label} cannot be '.' or '..'.")
        if value[-1] in {" ", "."}:
            raise ValueError(f"{value_label} cannot end with a space or period.")

        invalid_chars = '<>:"/\\|?*'
        bad_chars = [ch for ch in invalid_chars if ch in value]
        if bad_chars:
            raise ValueError(
                f"{value_label} contains invalid characters: {' '.join(bad_chars)}"
            )

        reserved_names = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9",
            "LPT1",
            "LPT2",
            "LPT3",
            "LPT4",
            "LPT5",
            "LPT6",
            "LPT7",
            "LPT8",
            "LPT9",
        }
        if value.split(".")[0].upper() in reserved_names:
            raise ValueError(f"{value_label} uses a reserved Windows name.")

        return value

    def _scan_employee_folder_paths():
        discovered = []
        try:
            status_entries = sorted(os.listdir(root_path))
        except OSError:
            return discovered

        for status_name in status_entries:
            status_path = normalize_path(os.path.join(root_path, status_name))
            if not os.path.isdir(status_path):
                continue

            try:
                letter_entries = sorted(os.listdir(status_path))
            except OSError:
                continue

            for letter_name in letter_entries:
                letter_path = normalize_path(os.path.join(status_path, letter_name))
                if not os.path.isdir(letter_path):
                    continue

                try:
                    employee_entries = sorted(os.listdir(letter_path))
                except OSError:
                    continue

                for employee_name in employee_entries:
                    employee_path = normalize_path(os.path.join(letter_path, employee_name))
                    if os.path.isdir(employee_path):
                        discovered.append(employee_path)

        return discovered

    def _refresh_folder_autocomplete_catalog():
        nonlocal folder_path_suggestions, folder_suggestion_entries, folder_suggestion_label_to_path
        scanned_paths = _scan_employee_folder_paths()
        folder_path_suggestions = list(dict.fromkeys(scanned_paths))

        entries = []
        label_to_path = {}

        for candidate_path in folder_path_suggestions:
            employee_name = os.path.basename(candidate_path)
            try:
                relative_path = os.path.relpath(candidate_path, root_path)
            except ValueError:
                relative_path = candidate_path

            base_label = f"{employee_name} | {relative_path}"
            unique_label = base_label
            suffix = 2
            while unique_label in label_to_path and label_to_path[unique_label] != candidate_path:
                unique_label = f"{base_label} ({suffix})"
                suffix += 1

            searchable = _normalize_folder_search_value(
                f"{employee_name} {relative_path} {candidate_path}"
            )
            entries.append(
                {
                    "label": unique_label,
                    "path": candidate_path,
                    "searchable": searchable,
                }
            )
            label_to_path[unique_label] = candidate_path

        folder_suggestion_entries = entries
        folder_suggestion_label_to_path = label_to_path

    def _normalize_folder_search_value(value):
        return " ".join(
            (value or "").lower().replace("\\", " ").replace("/", " ").split()
        )

    def _get_filtered_folder_suggestions(query):
        if not folder_suggestion_entries:
            return []

        normalized_query = _normalize_folder_search_value(query)
        if not normalized_query:
            return [entry["label"] for entry in folder_suggestion_entries]

        prefix_matches = []
        contains_matches = []

        for entry in folder_suggestion_entries:
            searchable = entry["searchable"]
            label = entry["label"]
            if searchable.startswith(normalized_query):
                prefix_matches.append(label)
            elif normalized_query in searchable:
                contains_matches.append(label)

        return prefix_matches + contains_matches

    def _update_folder_combobox_suggestions(combobox, query_text, event=None):
        suggestions = _get_filtered_folder_suggestions(query_text)
        combobox["values"] = suggestions

        keysym = getattr(event, "keysym", "") if event is not None else ""
        navigation_keys = {"Up", "Down", "Prior", "Next", "Return", "Tab", "Escape"}
        if keysym in navigation_keys:
            if keysym in {"Return", "Tab", "Escape"}:
                _hide_suggestion_popup(combobox)
            return

        if query_text.strip() and suggestions:
            combobox.after_idle(lambda: _show_suggestion_popup(combobox, suggestions))
        else:
            combobox.after_idle(lambda: _hide_suggestion_popup(combobox))

    def _apply_folder_input_selection(selected_value=None):
        raw_value = selected_value if selected_value is not None else folder_var.get()
        raw_value = (raw_value or "").strip()
        candidate = folder_suggestion_label_to_path.get(raw_value, "")
        if not candidate:
            candidate = normalize_path(raw_value)
        else:
            candidate = normalize_path(candidate)

        if not candidate:
            return

        if not os.path.isabs(candidate):
            candidate = normalize_path(os.path.join(root_path, candidate))

        if not os.path.isdir(candidate):
            return
        if not ensure_folder_under_root(candidate):
            return

        if folder_var.get().strip() != candidate:
            folder_var.set(candidate)

        prefill_from_folder(candidate)
        load_existing_pdfs()
        refresh_destination_preview()
        update_merge_summary()

    def select_existing_folder():
        chosen = filedialog.askdirectory(initialdir=root_path)
        if not chosen:
            keep_merge_window_on_top()
            return
        if not ensure_folder_under_root(chosen):
            messagebox.showerror("Error", "Please select a folder inside the Records Root Folder.")
            keep_merge_window_on_top()
            return

        _refresh_folder_autocomplete_catalog()
        folder_var.set(normalize_path(chosen))
        prefill_from_folder(chosen)
        load_existing_pdfs()
        refresh_destination_preview()
        keep_merge_window_on_top()
        update_merge_summary()

    def load_existing_pdfs():
        previous_selected = existing_selected_pdf_var.get().strip()
        existing_selected_pdf_var.set("")
        existing_selection_var.set("No file selected")
        existing_list_hint_var.set("Select a destination folder to load existing PDFs.")

        if existing_files_frame is None:
            update_merge_summary()
            return

        for child in existing_files_frame.winfo_children():
            child.destroy()

        folder = normalize_path(folder_var.get())
        if not folder:
            existing_list_hint_var.set("Select a destination employee folder to load existing PDFs.")
            update_merge_summary()
            return
        try:
            files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to list PDFs: {exc}")
            existing_list_hint_var.set("Unable to read PDF list. Check folder access and try again.")
            update_merge_summary()
            return

        if not files:
            existing_selection_var.set("No PDFs in this folder")
            existing_selected_pdf_var.set("")
            existing_list_hint_var.set("No PDFs found in this folder. Add files or choose another folder.")
            update_merge_summary()
            return

        existing_list_hint_var.set("Select one PDF to receive pending pages, then confirm Merge and Save.")

        if previous_selected in files:
            existing_selected_pdf_var.set(previous_selected)

        for file in files:
            row = ttk.Frame(existing_files_frame, style="PendingRow.TFrame", padding=(10, 6))
            row.pack(fill="x", padx=2, pady=3)

            checkbutton = ttk.Checkbutton(
                row,
                variable=existing_selected_pdf_var,
                onvalue=file,
                offvalue="",
                command=on_existing_select,
                style="PendingFile.TCheckbutton",
            )
            checkbutton.pack(side="left", padx=(2, 8))

            display_name = _format_pending_filename_for_display(file)
            name_label = ttk.Label(row, text=display_name, style="PendingFile.TLabel", anchor="w")
            name_label.pack(side="left", fill="x", expand=True, padx=(0, 8))

            preview_button = ttk.Button(
                row,
                style="ToolbarIcon.TButton",
                command=lambda target_file=file: preview_existing_pdf(target_file),
            )
            preview_button.pack(side="right", padx=(8, 0))
            _configure_icon_button(preview_button, "preview", TOOLBAR_ICON_PREVIEW, "Preview")
            _attach_hover_tooltip(preview_button, "Preview this existing PDF")
            preview_button.bind(
                "<Button-1>",
                lambda _event, target_button=preview_button: (target_button.invoke(), "break")[1],
                add="+",
            )

            edit_button = ttk.Button(
                row,
                style="ToolbarIcon.TButton",
                command=lambda target_file=file: edit_existing_pdf(target_file),
            )
            edit_button.pack(side="right", padx=(0, 8))
            _configure_icon_button(edit_button, "edit", TOOLBAR_ICON_EDIT, "Edit")
            _attach_hover_tooltip(edit_button, "Rename or rotate this existing PDF")
            edit_button.bind(
                "<Button-1>",
                lambda _event, target_button=edit_button: (target_button.invoke(), "break")[1],
                add="+",
            )

            if display_name != file:
                _attach_hover_tooltip(name_label, file)

            def _toggle_existing_selection(_event=None, target_file=file):
                current_value = existing_selected_pdf_var.get().strip()
                existing_selected_pdf_var.set("" if current_value == target_file else target_file)
                on_existing_select()
                return "break"

            def _on_existing_row_enter(
                _event=None,
                target_row=row,
                target_label=name_label,
                target_check=checkbutton,
            ):
                _set_pending_row_hover_state(target_row, target_label, target_check, True)

            def _on_existing_row_leave(
                _event=None,
                target_row=row,
                target_label=name_label,
                target_check=checkbutton,
            ):
                _set_pending_row_hover_state(target_row, target_label, target_check, False)

            name_label.bind("<Button-1>", _toggle_existing_selection, add="+")

            for hover_widget in (row, name_label, checkbutton):
                hover_widget.bind("<Enter>", _on_existing_row_enter, add="+")
                hover_widget.bind("<Leave>", _on_existing_row_leave, add="+")

            for action_button in (edit_button, preview_button):
                action_button.bind("<Enter>", _on_existing_row_enter, add="+")
                action_button.bind("<Leave>", _on_existing_row_leave, add="+")

            _set_pending_row_hover_state(row, name_label, checkbutton, False)

        if existing_selected_pdf_var.get().strip():
            on_existing_select()
        else:
            update_merge_summary()

    def on_existing_select(event=None):
        filename = existing_selected_pdf_var.get().strip()
        if not filename:
            existing_selection_var.set("No file selected")
            existing_list_hint_var.set("Select one PDF to receive pending pages, then confirm Merge and Save.")
            update_merge_summary()
            return

        existing_selection_var.set(filename)
        existing_list_hint_var.set("Selected PDF will receive pending pages first during merge.")

        metadata = parse_filename_metadata(filename)
        if metadata:
            name_var.set(metadata["name"])
            letter_var.set(metadata["name"][0].upper())
            new_year_var.set(metadata["latest"])
            old_year_var.set(metadata.get("earliest", ""))
        refresh_destination_preview()
        update_merge_summary()

    def preview_existing_pdf(target_filename=None):
        filename = (target_filename or existing_selected_pdf_var.get().strip()).strip()
        if not filename:
            messagebox.showwarning("Warning", "Select an existing PDF to preview.")
            return

        if target_filename and existing_selected_pdf_var.get().strip() != filename:
            existing_selected_pdf_var.set(filename)
            on_existing_select()

        folder = normalize_path(folder_var.get())
        file_path = normalize_path(os.path.join(folder, filename))
        try:
            _launch_path(file_path)
        except RuntimeError as exc:
            messagebox.showerror("Error", f"Could not open PDF: {exc}")

    def rename_existing_pdf(target_filename=None):
        filename = (target_filename or existing_selected_pdf_var.get().strip()).strip()
        if not filename:
            messagebox.showwarning("Warning", "Select an existing PDF to rename.", parent=win)
            return

        folder = normalize_path(folder_var.get())
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid employee folder first.", parent=win)
            return

        source_path = normalize_path(os.path.join(folder, filename))
        if not os.path.exists(source_path):
            messagebox.showerror("Error", "The selected PDF no longer exists.", parent=win)
            load_existing_pdfs()
            return

        name_root, _ext = os.path.splitext(filename)
        proposed_name = simpledialog.askstring(
            "Rename PDF",
            "Enter the new filename (.pdf optional):",
            initialvalue=name_root,
            parent=win,
        )
        if proposed_name is None:
            return

        proposed_name = proposed_name.strip()
        if not proposed_name:
            messagebox.showerror("Invalid Filename", "PDF filename cannot be empty.", parent=win)
            return

        if not proposed_name.lower().endswith(".pdf"):
            proposed_name = f"{proposed_name}.pdf"

        try:
            new_filename = _validate_windows_path_component(proposed_name, "PDF filename")
        except ValueError as exc:
            messagebox.showerror("Invalid Filename", str(exc), parent=win)
            return

        if not new_filename.lower().endswith(".pdf"):
            messagebox.showerror("Invalid Filename", "PDF filename must end with .pdf.", parent=win)
            return

        if new_filename == filename:
            return

        destination_path = normalize_path(os.path.join(folder, new_filename))
        source_norm = os.path.normcase(os.path.abspath(source_path))
        destination_norm = os.path.normcase(os.path.abspath(destination_path))
        if source_norm != destination_norm and os.path.exists(destination_path):
            messagebox.showerror(
                "Error",
                f"A file named '{new_filename}' already exists in this folder.",
                parent=win,
            )
            return

        confirm = messagebox.askyesno(
            "Confirm PDF Rename",
            f"Rename:\n{filename}\n\nto:\n{new_filename}",
            parent=win,
        )
        if not confirm:
            return

        try:
            os.rename(source_path, destination_path)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to rename PDF: {exc}", parent=win)
            return

        existing_selected_pdf_var.set(new_filename)
        load_existing_pdfs()
        refresh_destination_preview()
        update_merge_summary()

    def rotate_existing_pdf(target_filename=None):
        filename = (target_filename or existing_selected_pdf_var.get().strip()).strip()
        if not filename:
            messagebox.showwarning("Warning", "Select an existing PDF to rotate.", parent=win)
            return

        folder = normalize_path(folder_var.get())
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid employee folder first.", parent=win)
            return

        source_path = normalize_path(os.path.join(folder, filename))
        if not os.path.exists(source_path):
            messagebox.showerror("Error", "The selected PDF no longer exists.", parent=win)
            load_existing_pdfs()
            return

        def _refresh_existing_after_rotation_save():
            existing_selected_pdf_var.set(filename)
            load_existing_pdfs()
            refresh_destination_preview()
            update_merge_summary()

        rotate_selected_pending_pdfs(
            selected_file_infos_override=[
                {
                    "name": filename,
                    "path": source_path,
                    "page_count": 0,
                }
            ],
            window_title=f"Rotate Existing PDF - {filename}",
            post_save_callback=_refresh_existing_after_rotation_save,
            parent_window=win,
        )

    def edit_existing_pdf(target_filename=None):
        filename = (target_filename or existing_selected_pdf_var.get().strip()).strip()
        if not filename:
            messagebox.showwarning("Warning", "Select an existing PDF first.", parent=win)
            return

        action_choice = messagebox.askyesnocancel(
            "Edit Existing PDF",
            (
                f"Choose action for:\n{filename}\n\n"
                "Yes = Rename PDF\n"
                "No = Rotate PDF\n"
                "Cancel = Close"
            ),
            parent=win,
        )
        if action_choice is None:
            return

        if action_choice:
            rename_existing_pdf(filename)
        else:
            rotate_existing_pdf(filename)

    def validate_years():
        latest = new_year_var.get().strip()
        earliest = old_year_var.get().strip()
        try:
            years = _normalize_record_year_inputs(latest, earliest)
        except ValueError:
            messagebox.showerror("Year Validation", _get_year_input_guidance(latest, earliest), parent=win)
            return None
        if years is None:
            messagebox.showerror("Year Validation", _get_year_input_guidance(latest, earliest), parent=win)
            return None
        return years

    def perform_merge():
        nonlocal processed_successfully
        folder = normalize_path(folder_var.get())
        if not folder:
            messagebox.showerror("Error", "Select the employee folder first.")
            return
        if not ensure_folder_under_root(folder):
            messagebox.showerror("Error", "Selected folder must be inside the Records Root Folder.")
            return

        existing_filename = existing_selected_pdf_var.get().strip()
        if not existing_filename:
            messagebox.showerror("Error", "Select the existing PDF to merge with.")
            return

        employee_name = name_var.get().strip()
        if not employee_name:
            messagebox.showerror("Error", "Employee name is required.")
            return

        letter_value = letter_var.get().strip().upper() or (employee_name[0].upper() if employee_name else "")
        status_value = status_var.get()

        years = validate_years()
        if years is None:
            return
        latest_year, earliest_year = years
        final_filename = _build_record_filename(employee_name, latest_year, earliest_year)
        final_path = normalize_path(os.path.join(folder, final_filename))

        existing_path = normalize_path(os.path.join(folder, existing_filename))

        if not os.path.exists(existing_path):
            messagebox.showerror("Error", "The selected existing PDF was not found. Reload and try again.")
            load_existing_pdfs()
            return
        if not os.path.exists(pending_path):
            messagebox.showerror("Error", "Pending PDF no longer exists. Refresh the list.")
            load_pending_files()
            return

        if not employee_name:
            messagebox.showerror("Error", "Name is required before merging.")
            return

        confirm_text = (
            f"Merge '{pending_filename}' into '{existing_filename}' (new pages go first)\n"
            f"and save as '{final_filename}' in:\n{folder}"
        )
        summary_snapshot = merge_summary_var.get().strip()
        if summary_snapshot:
            confirm_text += f"\n\n{summary_snapshot}"
        if not messagebox.askyesno("Confirm Merge", confirm_text):
            return

        temp_path = None
        backup_paths = []
        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            temp_path = temp_file.name
            temp_file.close()

            merge_pdf_files(pending_path, existing_path, temp_path)

            if os.path.exists(final_path) and final_path != existing_path:
                replace = messagebox.askyesno(
                    "Replace File",
                    f"{final_filename} already exists. Do you want to replace it?"
                )
                if not replace:
                    return

            recycle_choice = messagebox.askyesno(
                "Recycle Original?",
                "Move the previously existing PDF to the Recycle Bin after saving?"
            )

            warned_recycle_missing = False

            def maybe_backup(target_path):
                if not keep_backup_var.get():
                    return
                if not target_path or not os.path.exists(target_path):
                    return
                try:
                    backup_path = create_backup_file(target_path)
                    if backup_path:
                        backup_paths.append(backup_path)
                except Exception as backup_error:
                    messagebox.showwarning(
                        "Backup Failed",
                        f"Unable to create backup for {os.path.basename(target_path)}: {backup_error}"
                    )

            def remove_original(target_path, force=False):
                nonlocal warned_recycle_missing
                if not os.path.exists(target_path):
                    return
                if recycle_choice:
                    if send2trash is not None:
                        send2trash(target_path)
                    else:
                        if not warned_recycle_missing:
                            messagebox.showwarning(
                                "Recycle Unavailable",
                                "send2trash is not installed; deleting the original file instead."
                            )
                            warned_recycle_missing = True
                        os.remove(target_path)
                elif force:
                    os.remove(target_path)

            same_target = os.path.abspath(final_path) == os.path.abspath(existing_path)
            if same_target:
                maybe_backup(existing_path)
                remove_original(existing_path, force=True)
            else:
                if recycle_choice:
                    maybe_backup(existing_path)
                    remove_original(existing_path)
                if os.path.exists(final_path):
                    maybe_backup(final_path)
                    os.remove(final_path)

            shutil.move(temp_path, final_path)
            temp_path = None

            try:
                archived_path = archive_pending_file(pending_path)
                if not archived_path:
                    raise OSError("Pending file was not found during archiving.")
            except OSError as exc:
                messagebox.showwarning(
                    "Warning",
                    f"Merged successfully but unable to archive pending file: {exc}"
                )

            processed_successfully = True
            messagebox.showinfo("Success", f"Merged file saved to:\n{final_path}")
            load_pending_files()
            close_merge_window()
        except Exception as exc:
            messagebox.showerror("Error", f"Merge failed: {exc}")
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)

    # Folder selection UI
    folder_card = ttk.Frame(content, style="Card.TFrame", padding=14)
    folder_card.pack(fill="x", pady=(0, 10))
    ttk.Label(folder_card, text="Destination Employee Folder", style="SectionTitle.TLabel").pack(anchor="w")
    ttk.Label(
        folder_card,
        text="Search by employee name or paste a path under the Records Root Folder.",
        style="CardMuted.TLabel",
    ).pack(anchor="w", pady=(2, 8))

    folder_row = ttk.Frame(folder_card, style="Card.TFrame")
    folder_row.pack(fill="x")

    folder_field = ttk.Combobox(folder_row, textvariable=folder_var, state="normal")
    folder_field.pack(side="left", expand=True, fill="x", padx=(0, 8))
    _prevent_combobox_mousewheel_value_change(folder_field)

    _refresh_folder_autocomplete_catalog()

    def _handle_folder_key(event=None):
        _update_folder_combobox_suggestions(folder_field, folder_var.get(), event)

    def _refresh_folder_choices():
        folder_field["values"] = _get_filtered_folder_suggestions(folder_var.get())

    folder_field.bind("<KeyRelease>", _handle_folder_key)
    folder_field.bind("<<ComboboxSelected>>", lambda _event: _apply_folder_input_selection())
    folder_field.bind("<Return>", lambda _event: (_apply_folder_input_selection(), "break")[1])
    folder_field.bind("<FocusOut>", lambda _event: _apply_folder_input_selection(), add="+")
    folder_field.configure(postcommand=_refresh_folder_choices)
    setattr(
        folder_field,
        "_on_suggestion_selected",
        lambda selected_value: _apply_folder_input_selection(selected_value),
    )

    ttk.Button(
        folder_row,
        text="Browse",
        command=select_existing_folder,
        style="Subtle.TButton",
        width=9,
    ).pack(side="left")

    existing_list_card = ttk.Frame(content, style="Card.TFrame", padding=10)
    existing_list_card.pack(fill="both", expand=True, pady=(0, 10))
    ttk.Label(existing_list_card, text="Existing PDFs in Folder", style="SectionTitle.TLabel").pack(anchor="w", padx=2)

    list_container = ttk.Frame(existing_list_card, style="Card.TFrame", padding=8)
    list_container.pack(fill="both", expand=True, pady=(6, 4))

    existing_canvas = tk.Canvas(
        list_container,
        highlightthickness=0,
        bg="#0b1220",
        bd=0,
    )
    existing_canvas.pack(side="left", fill="both", expand=True)
    existing_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=existing_canvas.yview)
    existing_scrollbar.pack(side="right", fill="y")
    existing_canvas.configure(yscrollcommand=existing_scrollbar.set)
    _mark_widget_as_scroll_canvas(existing_canvas)
    _ensure_global_mousewheel_binding()

    existing_files_frame = ttk.Frame(existing_canvas, style="Card.TFrame")
    existing_canvas_window = existing_canvas.create_window((0, 0), window=existing_files_frame, anchor="nw")

    def _resize_existing_canvas(event):
        existing_canvas.itemconfigure(existing_canvas_window, width=event.width)

    def _update_existing_scrollregion(_event=None):
        existing_canvas.configure(scrollregion=existing_canvas.bbox("all"))

    existing_canvas.bind("<Configure>", _resize_existing_canvas)
    existing_files_frame.bind("<Configure>", _update_existing_scrollregion)

    ttk.Label(existing_list_card, textvariable=existing_selection_var, style="CardMuted.TLabel").pack(
        anchor="w", pady=(6, 8), padx=2
    )
    ttk.Label(
        existing_list_card,
        textvariable=existing_list_hint_var,
        style="CardMuted.TLabel",
        wraplength=560,
        justify="left",
    ).pack(anchor="w", pady=(0, 8), padx=2)

    button_frame = ttk.Frame(existing_list_card, style="Card.TFrame")
    button_frame.pack(fill="x", pady=(0, 2))
    ttk.Button(
        button_frame,
        text="Reload List",
        command=load_existing_pdfs,
        style="Subtle.TButton",
    ).grid(row=0, column=0, padx=5)

    def open_folder():
        folder = normalize_path(folder_var.get())
        if not folder:
            messagebox.showwarning("Warning", "Select a folder first.")
            return
        try:
            _launch_path(folder)
        except RuntimeError as exc:
            messagebox.showerror("Error", f"Unable to open folder: {exc}")

    def rename_selected_folder():
        current_folder = normalize_path(folder_var.get().strip())
        if not current_folder:
            messagebox.showwarning("Warning", "Select a folder first.", parent=win)
            return
        if not os.path.isdir(current_folder):
            messagebox.showerror("Error", "Selected folder does not exist.", parent=win)
            return
        if not ensure_folder_under_root(current_folder):
            messagebox.showerror(
                "Error",
                "Selected folder must be inside the Records Root Folder.",
                parent=win,
            )
            return

        current_name = os.path.basename(current_folder)
        parent_folder = normalize_path(os.path.dirname(current_folder))
        proposed_name = simpledialog.askstring(
            "Rename Employee Folder",
            "Enter the corrected folder name:",
            initialvalue=current_name,
            parent=win,
        )
        if proposed_name is None:
            return

        try:
            new_folder_name = _validate_windows_path_component(proposed_name, "Folder name")
        except ValueError as exc:
            messagebox.showerror("Invalid Folder Name", str(exc), parent=win)
            return

        renamed_folder = normalize_path(os.path.join(parent_folder, new_folder_name))
        if renamed_folder == current_folder:
            return

        current_norm = os.path.normcase(os.path.abspath(current_folder))
        renamed_norm = os.path.normcase(os.path.abspath(renamed_folder))
        if current_norm != renamed_norm and os.path.exists(renamed_folder):
            messagebox.showerror(
                "Error",
                "A folder with that name already exists in this location.",
                parent=win,
            )
            return

        confirm = messagebox.askyesno(
            "Confirm Folder Rename",
            f"Rename folder:\n{current_name}\n\nto:\n{new_folder_name}",
            parent=win,
        )
        if not confirm:
            return

        try:
            os.rename(current_folder, renamed_folder)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to rename folder: {exc}", parent=win)
            return

        _refresh_folder_autocomplete_catalog()
        folder_var.set(renamed_folder)
        prefill_from_folder(renamed_folder)
        _refresh_folder_choices()
        load_existing_pdfs()
        refresh_destination_preview()
        update_merge_summary()

    ttk.Button(
        button_frame,
        text="View Folder",
        command=open_folder,
        style="Subtle.TButton",
    ).grid(row=0, column=1, padx=5)
    ttk.Button(
        button_frame,
        text="Rename Folder",
        command=rename_selected_folder,
        style="Subtle.TButton",
    ).grid(row=0, column=2, padx=5)

    # Metadata fields
    form_frame = ttk.Frame(content, style="Card.TFrame", padding=14)
    form_frame.pack(fill="x")
    form_frame.columnconfigure(0, weight=1)
    form_frame.columnconfigure(1, weight=1)

    ttk.Label(form_frame, text="Final Record Details", style="SectionTitle.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w"
    )

    ttk.Label(form_frame, text="Employee Name", style="FieldLabel.TLabel").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(10, 2)
    )
    merge_name_field = ttk.Combobox(
        form_frame,
        textvariable=name_var,
        values=employee_name_suggestions,
        state="normal",
    )
    _prevent_combobox_mousewheel_value_change(merge_name_field)

    def _handle_merge_name_key(event=None):
        # Keep user input untouched and only refresh/open suggestion list.
        _update_combobox_suggestions(merge_name_field, name_var.get(), event)

    merge_name_field.bind("<KeyRelease>", _handle_merge_name_key)

    def _refresh_merge_name_choices():
        merge_name_field["values"] = get_filtered_name_suggestions(name_var.get())

    merge_name_field.configure(postcommand=_refresh_merge_name_choices)
    merge_name_field.grid(row=2, column=0, columnspan=2, sticky="ew")

    ttk.Label(form_frame, text="Surname Initial", style="FieldLabel.TLabel").grid(
        row=3, column=0, sticky="w", pady=(10, 2)
    )
    ttk.Entry(form_frame, textvariable=letter_var).grid(row=4, column=0, sticky="ew", padx=(0, 8))

    ttk.Label(form_frame, text="Status", style="FieldLabel.TLabel").grid(
        row=3, column=1, sticky="w", pady=(10, 2)
    )
    status_field = ttk.Combobox(
        form_frame,
        textvariable=status_var,
        values=("Active", "Retiree"),
        state="readonly",
    )
    _prevent_combobox_mousewheel_value_change(status_field)
    status_field.grid(row=4, column=1, sticky="ew")

    ttk.Label(form_frame, text="Latest Year", style="FieldLabel.TLabel").grid(
        row=5, column=0, sticky="w", pady=(10, 2)
    )
    ttk.Entry(form_frame, textvariable=new_year_var).grid(row=6, column=0, sticky="ew", padx=(0, 8))

    ttk.Label(form_frame, text="Oldest Year", style="FieldLabel.TLabel").grid(
        row=5, column=1, sticky="w", pady=(10, 2)
    )
    ttk.Entry(form_frame, textvariable=old_year_var).grid(row=6, column=1, sticky="ew")
    ttk.Label(
        form_frame,
        textvariable=year_hint_var,
        style="CardMuted.TLabel",
        wraplength=560,
        justify="left",
    ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))

    ttk.Label(form_frame, text="Final File Preview", style="FieldLabel.TLabel").grid(
        row=8, column=0, columnspan=2, sticky="w", pady=(10, 2)
    )
    ttk.Label(
        form_frame,
        textvariable=dest_preview_var,
        wraplength=560,
        justify="left",
        style="CardMuted.TLabel",
        anchor="w",
        padding=(0, 6),
    ).grid(row=9, column=0, columnspan=2, sticky="ew")

    ttk.Label(form_frame, text="Merge Summary", style="FieldLabel.TLabel").grid(
        row=10, column=0, columnspan=2, sticky="w", pady=(10, 2)
    )
    ttk.Label(
        form_frame,
        textvariable=merge_summary_var,
        style="CardMuted.TLabel",
        wraplength=560,
        justify="left",
        anchor="w",
    ).grid(row=11, column=0, columnspan=2, sticky="ew")

    action_frame = ttk.Frame(content, style="ActionBar.TFrame", padding=(12, 10))
    action_frame.pack(fill="x", pady=(12, 0))
    ttk.Label(action_frame, text="Merge Action", style="ActionTitle.TLabel").pack(anchor="w")

    action_buttons = ttk.Frame(action_frame, style="ActionBar.TFrame")
    action_buttons.pack(fill="x", pady=(6, 0))

    ttk.Button(
        action_buttons,
        text="Close",
        command=close_merge_window,
        style="SecondaryAction.TButton",
        width=14,
    ).pack(side="left")
    if batch_context:
        ttk.Button(
            action_buttons,
            text="Cancel Batch",
            command=cancel_batch_processing_from_merge,
            style="Subtle.TButton",
            width=14,
        ).pack(side="left", padx=(8, 0))

    ttk.Button(
        action_buttons,
        text="Merge and Save",
        width=20,
        command=perform_merge,
        style="PrimaryAction.TButton",
    ).pack(side="right")
    win.bind("<Escape>", lambda _event: (close_merge_window(), "break")[1])
    win.after_idle(lambda: folder_field.focus_set())
    win.protocol("WM_DELETE_WINDOW", close_merge_window)

    refresh_destination_preview()

def _start_batch_processing(mode):
    selected = get_selected_pending_files()
    if not selected:
        messagebox.showwarning("Warning", "Select at least one pending PDF using the checkboxes first.")
        return

    files = selected.copy()
    total = len(files)
    action_label = "New Record" if mode == "new" else "Merge Existing"
    batch_state = {
        "files": files,
        "total": total,
        "current": 1,
        "processed": 0,
        "cancelled": False,
        "summary_shown": False,
    }

    def _show_batch_summary(cancelled=False):
        if batch_state["summary_shown"]:
            return

        if not cancelled and batch_state["total"] <= 1:
            batch_state["summary_shown"] = True
            return

        batch_state["summary_shown"] = True

        processed_count = batch_state["processed"]
        skipped_count = max(0, batch_state["total"] - processed_count)
        title = "Batch Cancelled" if cancelled else "Batch Complete"
        message = f"{action_label} batch finished. Processed {processed_count} of {batch_state['total']} pending PDFs."
        if skipped_count > 0:
            message += f" Skipped {skipped_count}."
        messagebox.showinfo(title, message)

    def launch_next():
        if batch_state["cancelled"]:
            _show_batch_summary(cancelled=True)
            return

        if not batch_state["files"]:
            _show_batch_summary(cancelled=False)
            return

        current_file = batch_state["files"][0]
        batch_state["current"] = batch_state["total"] - len(batch_state["files"]) + 1

        def handle_close(_success, _filename):
            if _success:
                batch_state["processed"] += 1

            if _filename in batch_state["files"]:
                batch_state["files"].remove(_filename)

            root.after(150, launch_next)

        if mode == "new":
            new_record_window(initial_filename=current_file, batch_context=batch_state, on_complete=handle_close)
        else:
            merge_existing_window(pending_filename=current_file, batch_context=batch_state, on_complete=handle_close)

    launch_next()


def start_new_record_batch():
    _start_batch_processing("new")


def start_merge_existing_batch():
    _start_batch_processing("merge")


def employee_details_editor_window():
    root_path = normalize_path(root_folder.get().strip())
    if not root_path:
        messagebox.showerror("Error", "Select the Records Root Folder first.")
        return
    if not os.path.isdir(root_path):
        messagebox.showerror("Error", "The selected Records Root Folder does not exist.")
        return

    win = tk.Toplevel(root)
    _apply_app_icon(win)
    win.title("Employee Details Editor")
    configure_window_geometry(
        win,
        1080,
        820,
        min_width=860,
        min_height=620,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(root)
    win.lift()
    win.focus_force()
    apply_theme(win)

    folder_var = tk.StringVar()
    employee_name_var = tk.StringVar(value="")
    current_status_var = tk.StringVar(value="Unknown")
    target_status_var = tk.StringVar(value="Active")
    folder_name_var = tk.StringVar()

    files_count_var = tk.StringVar(value="PDF files: 0")
    files_empty_hint_var = tk.StringVar(value="Select an employee folder to load PDF files.")
    selected_file_var = tk.StringVar(value="No PDF selected")
    file_name_var = tk.StringVar()
    created_var = tk.StringVar()
    modified_var = tk.StringVar()

    status_values = []
    folder_path_suggestions = []
    folder_suggestion_entries = []
    folder_suggestion_label_to_path = {}
    folder_watch_after_id = None
    folder_watch_snapshot = None

    def ensure_folder_under_root(chosen_folder):
        chosen_folder = normalize_path(chosen_folder)
        try:
            common = os.path.commonpath([os.path.abspath(chosen_folder), os.path.abspath(root_path)])
        except ValueError:
            return False
        return common == os.path.abspath(root_path)

    def _build_employee_folder_path(status_value, employee_name):
        status_name = str(status_value or "").strip()
        name = str(employee_name or "").strip()
        if not status_name or not name:
            return ""
        first_char = name[0].upper() if name else "#"
        letter_folder = first_char if first_char.isalpha() else "#"
        return normalize_path(os.path.join(root_path, status_name, letter_folder, name))

    def _normalize_folder_search_value(value):
        return " ".join((value or "").lower().replace("\\", " ").replace("/", " ").split())

    def _scan_employee_folder_paths():
        discovered = []
        try:
            status_entries = sorted(os.listdir(root_path))
        except OSError:
            return discovered

        for status_name in status_entries:
            status_path = normalize_path(os.path.join(root_path, status_name))
            if not os.path.isdir(status_path):
                continue

            try:
                letter_entries = sorted(os.listdir(status_path))
            except OSError:
                continue

            for letter_name in letter_entries:
                letter_path = normalize_path(os.path.join(status_path, letter_name))
                if not os.path.isdir(letter_path):
                    continue

                try:
                    employee_entries = sorted(os.listdir(letter_path))
                except OSError:
                    continue

                for employee_name in employee_entries:
                    employee_path = normalize_path(os.path.join(letter_path, employee_name))
                    if os.path.isdir(employee_path):
                        discovered.append(employee_path)

        return discovered

    def _refresh_folder_autocomplete_catalog():
        nonlocal folder_path_suggestions, folder_suggestion_entries, folder_suggestion_label_to_path

        scanned_paths = _scan_employee_folder_paths()
        folder_path_suggestions = list(dict.fromkeys(scanned_paths))

        entries = []
        label_to_path = {}

        for candidate_path in folder_path_suggestions:
            employee_name = os.path.basename(candidate_path)
            try:
                relative_path = os.path.relpath(candidate_path, root_path)
            except ValueError:
                relative_path = candidate_path

            base_label = f"{employee_name} | {relative_path}"
            unique_label = base_label
            suffix = 2
            while unique_label in label_to_path and label_to_path[unique_label] != candidate_path:
                unique_label = f"{base_label} ({suffix})"
                suffix += 1

            searchable = _normalize_folder_search_value(
                f"{employee_name} {relative_path} {candidate_path}"
            )
            entries.append(
                {
                    "label": unique_label,
                    "path": candidate_path,
                    "searchable": searchable,
                }
            )
            label_to_path[unique_label] = candidate_path

        folder_suggestion_entries = entries
        folder_suggestion_label_to_path = label_to_path

    def _get_filtered_folder_suggestions(query):
        if not folder_suggestion_entries:
            return []

        normalized_query = _normalize_folder_search_value(query)
        if not normalized_query:
            return [entry["label"] for entry in folder_suggestion_entries]

        prefix_matches = []
        contains_matches = []

        for entry in folder_suggestion_entries:
            searchable = entry["searchable"]
            label = entry["label"]
            if searchable.startswith(normalized_query):
                prefix_matches.append(label)
            elif normalized_query in searchable:
                contains_matches.append(label)

        return prefix_matches + contains_matches

    def _update_folder_combobox_suggestions(combobox, query_text, event=None):
        suggestions = _get_filtered_folder_suggestions(query_text)
        combobox["values"] = suggestions

        keysym = getattr(event, "keysym", "") if event is not None else ""
        navigation_keys = {"Up", "Down", "Prior", "Next", "Return", "Tab", "Escape"}
        if keysym in navigation_keys:
            if keysym in {"Return", "Tab", "Escape"}:
                _hide_suggestion_popup(combobox)
            return

        if query_text.strip() and suggestions:
            combobox.after_idle(lambda: _show_suggestion_popup(combobox, suggestions))
        else:
            combobox.after_idle(lambda: _hide_suggestion_popup(combobox))

    def _resolve_folder_input_value(raw_value):
        raw = str(raw_value or "").strip()
        candidate = folder_suggestion_label_to_path.get(raw, "")
        if not candidate:
            candidate = normalize_path(raw)
        else:
            candidate = normalize_path(candidate)

        if not candidate:
            return ""

        if not os.path.isabs(candidate):
            candidate = normalize_path(os.path.join(root_path, candidate))

        return candidate

    def _build_folder_pdf_snapshot(folder_path):
        folder = normalize_path(folder_path)
        if not folder or not os.path.isdir(folder):
            return ()

        try:
            pdf_names = sorted(
                filename for filename in os.listdir(folder) if filename.lower().endswith(".pdf")
            )
        except OSError:
            return None

        snapshot_items = []
        for filename in pdf_names:
            file_path = normalize_path(os.path.join(folder, filename))
            try:
                stats = os.stat(file_path)
                snapshot_items.append((filename, int(stats.st_mtime_ns), int(stats.st_size)))
            except OSError:
                snapshot_items.append((filename, -1, -1))

        return tuple(snapshot_items)

    def _refresh_status_choices():
        nonlocal status_values
        discovered = []
        try:
            entries = sorted(os.listdir(root_path))
        except OSError:
            entries = []

        for entry in entries:
            candidate = normalize_path(os.path.join(root_path, entry))
            if os.path.isdir(candidate) and entry not in discovered:
                discovered.append(entry)

        status_values = discovered
        status_field["values"] = status_values

        current_target = target_status_var.get().strip()
        current_status = current_status_var.get().strip()

        if current_target in status_values:
            return
        if current_status in status_values:
            target_status_var.set(current_status)
            return
        if status_values:
            target_status_var.set(status_values[0])
        else:
            target_status_var.set("")

    def _clear_file_editor_fields():
        selected_file_var.set("No PDF selected")
        file_name_var.set("")
        created_var.set("")
        modified_var.set("")

    def _get_selected_file_name():
        selection = files_listbox.curselection()
        if not selection:
            return ""
        return files_listbox.get(selection[0]).strip()

    def _refresh_selected_file_details():
        selected_name = _get_selected_file_name()
        if not selected_name:
            _clear_file_editor_fields()
            return

        folder = normalize_path(folder_var.get().strip())
        file_path = normalize_path(os.path.join(folder, selected_name))
        if not os.path.exists(file_path):
            _clear_file_editor_fields()
            return

        selected_file_var.set(selected_name)
        file_name_var.set(os.path.splitext(selected_name)[0])
        created_var.set(_format_editor_datetime(datetime.fromtimestamp(os.path.getctime(file_path))))
        modified_var.set(_format_editor_datetime(datetime.fromtimestamp(os.path.getmtime(file_path))))

    def _refresh_files_list(preferred_name="", show_errors=True):
        nonlocal folder_watch_snapshot

        current_name = preferred_name or _get_selected_file_name()
        files_listbox.delete(0, tk.END)

        folder = normalize_path(folder_var.get().strip())
        if not folder or not os.path.isdir(folder):
            files_count_var.set("PDF files: 0")
            files_empty_hint_var.set("Select an employee folder to load PDF files.")
            _clear_file_editor_fields()
            folder_watch_snapshot = ()
            return

        try:
            pdf_files = sorted(
                filename for filename in os.listdir(folder) if filename.lower().endswith(".pdf")
            )
        except OSError as exc:
            if show_errors:
                messagebox.showerror("Error", f"Unable to load PDFs: {exc}", parent=win)
            files_count_var.set("PDF files: 0")
            files_empty_hint_var.set("Unable to read this folder. Check permissions and try again.")
            _clear_file_editor_fields()
            folder_watch_snapshot = None
            return

        folder_watch_snapshot = _build_folder_pdf_snapshot(folder)

        for filename in pdf_files:
            files_listbox.insert(tk.END, filename)

        files_count_var.set(f"PDF files: {len(pdf_files)}")

        if not pdf_files:
            files_empty_hint_var.set("No PDF files found in this folder.")
            _clear_file_editor_fields()
            return

        files_empty_hint_var.set("Select a PDF to edit filename or timestamp metadata.")

        target_name = current_name if current_name in pdf_files else pdf_files[0]
        target_index = pdf_files.index(target_name)
        files_listbox.selection_clear(0, tk.END)
        files_listbox.selection_set(target_index)
        files_listbox.activate(target_index)
        files_listbox.see(target_index)
        _refresh_selected_file_details()

    def _load_employee_folder_details(target_folder=None, show_errors=True):
        source_value = target_folder if target_folder is not None else folder_var.get()
        folder = _resolve_folder_input_value(source_value)
        if not folder:
            if show_errors:
                messagebox.showwarning("Warning", "Select an employee folder first.", parent=win)
            return
        if not os.path.isdir(folder):
            if show_errors:
                messagebox.showerror("Error", "The selected folder does not exist.", parent=win)
            return
        if not ensure_folder_under_root(folder):
            if show_errors:
                messagebox.showerror(
                    "Error",
                    "Please choose a folder inside the Records Root Folder.",
                    parent=win,
                )
            return

        folder_var.set(folder)

        employee_name = os.path.basename(folder)
        employee_name_var.set(employee_name)
        folder_name_var.set(employee_name)

        try:
            relative = os.path.relpath(folder, root_path)
            parts = relative.split(os.sep)
        except ValueError:
            parts = []

        detected_status = parts[0] if parts else ""
        current_status_var.set(detected_status or "Unknown")

        if detected_status and detected_status in status_values:
            target_status_var.set(detected_status)

        _refresh_files_list()

    def _apply_folder_input_selection(selected_value=None):
        source_value = selected_value if selected_value is not None else folder_var.get()
        folder = _resolve_folder_input_value(source_value)
        if not folder or not os.path.isdir(folder):
            return
        if not ensure_folder_under_root(folder):
            return
        _load_employee_folder_details(folder, show_errors=False)

    def _browse_employee_folder():
        chosen = filedialog.askdirectory(initialdir=root_path)
        if not chosen:
            return
        _refresh_folder_autocomplete_catalog()
        _load_employee_folder_details(chosen, show_errors=True)

    def _poll_folder_file_changes():
        nonlocal folder_watch_after_id, folder_watch_snapshot

        folder_watch_after_id = None
        try:
            if not win.winfo_exists():
                return
        except tk.TclError:
            return

        folder = normalize_path(folder_var.get().strip())
        if folder and os.path.isdir(folder) and ensure_folder_under_root(folder):
            current_snapshot = _build_folder_pdf_snapshot(folder)
            if current_snapshot is not None and current_snapshot != folder_watch_snapshot:
                selected_name = _get_selected_file_name()
                _refresh_files_list(preferred_name=selected_name, show_errors=False)
                _refresh_selected_file_details()
                folder_watch_snapshot = _build_folder_pdf_snapshot(folder)
        else:
            if folder_watch_snapshot not in (None, ()):
                _refresh_files_list(show_errors=False)
            folder_watch_snapshot = _build_folder_pdf_snapshot(folder)

        try:
            folder_watch_after_id = win.after(1200, _poll_folder_file_changes)
        except tk.TclError:
            folder_watch_after_id = None

    def _open_selected_folder():
        folder = normalize_path(folder_var.get().strip())
        if not folder:
            messagebox.showwarning("Warning", "Select an employee folder first.", parent=win)
            return
        try:
            _launch_path(folder)
        except RuntimeError as exc:
            messagebox.showerror("Error", f"Unable to open folder: {exc}", parent=win)

    def _rename_employee_folder():
        current_folder = normalize_path(folder_var.get().strip())
        if not current_folder or not os.path.isdir(current_folder):
            messagebox.showerror("Error", "Select a valid employee folder first.", parent=win)
            return

        try:
            new_folder_name = _validate_filesystem_component_name(folder_name_var.get(), "Folder name")
        except ValueError as exc:
            messagebox.showerror("Invalid Folder Name", str(exc), parent=win)
            return

        current_status = current_status_var.get().strip()
        target_status = current_status if current_status and current_status != "Unknown" else target_status_var.get().strip()
        destination_folder = _build_employee_folder_path(target_status, new_folder_name)
        if not destination_folder:
            messagebox.showerror("Error", "Unable to build destination folder path.", parent=win)
            return

        if os.path.normcase(os.path.abspath(destination_folder)) == os.path.normcase(os.path.abspath(current_folder)):
            return

        if os.path.exists(destination_folder):
            messagebox.showerror(
                "Error",
                "A folder with that employee name already exists in the destination.",
                parent=win,
            )
            return

        confirm = messagebox.askyesno(
            "Confirm Folder Rename",
            f"Rename employee folder to:\n{new_folder_name}\n\nDestination:\n{destination_folder}",
            parent=win,
        )
        if not confirm:
            return

        try:
            os.makedirs(os.path.dirname(destination_folder), exist_ok=True)
            os.rename(current_folder, destination_folder)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to rename employee folder: {exc}", parent=win)
            return

        _refresh_status_choices()
        _refresh_folder_autocomplete_catalog()
        _load_employee_folder_details(destination_folder, show_errors=False)
        messagebox.showinfo(
            "Success",
            f"Employee folder renamed successfully to:\n{new_folder_name}",
            parent=win,
        )

    def _apply_status_change():
        current_folder = normalize_path(folder_var.get().strip())
        if not current_folder or not os.path.isdir(current_folder):
            messagebox.showerror("Error", "Select a valid employee folder first.", parent=win)
            return

        new_status = target_status_var.get().strip()
        if not new_status:
            messagebox.showerror("Error", "Select a target status.", parent=win)
            return

        employee_name = os.path.basename(current_folder)
        destination_folder = _build_employee_folder_path(new_status, employee_name)
        if not destination_folder:
            messagebox.showerror("Error", "Unable to build destination folder path.", parent=win)
            return

        if os.path.normcase(os.path.abspath(destination_folder)) == os.path.normcase(os.path.abspath(current_folder)):
            return

        if os.path.exists(destination_folder):
            messagebox.showerror(
                "Error",
                "An employee folder with this status and name already exists.",
                parent=win,
            )
            return

        confirm = messagebox.askyesno(
            "Confirm Status Change",
            f"Move employee to status '{new_status}'?\n\nDestination:\n{destination_folder}",
            parent=win,
        )
        if not confirm:
            return

        try:
            os.makedirs(os.path.dirname(destination_folder), exist_ok=True)
            os.rename(current_folder, destination_folder)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to change employee status: {exc}", parent=win)
            return

        _refresh_status_choices()
        _refresh_folder_autocomplete_catalog()
        _load_employee_folder_details(destination_folder, show_errors=False)
        messagebox.showinfo(
            "Success",
            f"Employee status updated to '{new_status}'.",
            parent=win,
        )

    def _rename_selected_pdf():
        folder = normalize_path(folder_var.get().strip())
        filename = _get_selected_file_name()
        if not folder or not filename:
            messagebox.showwarning("Warning", "Select a PDF file first.", parent=win)
            return

        new_name_input = file_name_var.get().strip()
        if not new_name_input:
            messagebox.showerror("Error", "File name cannot be empty.", parent=win)
            return

        if not new_name_input.lower().endswith(".pdf"):
            new_name_input = f"{new_name_input}.pdf"

        try:
            new_filename = _validate_filesystem_component_name(new_name_input, "PDF filename")
        except ValueError as exc:
            messagebox.showerror("Invalid Filename", str(exc), parent=win)
            return

        if not new_filename.lower().endswith(".pdf"):
            messagebox.showerror("Invalid Filename", "PDF filename must end with .pdf.", parent=win)
            return

        source_path = normalize_path(os.path.join(folder, filename))
        destination_path = normalize_path(os.path.join(folder, new_filename))

        if os.path.normcase(os.path.abspath(source_path)) == os.path.normcase(os.path.abspath(destination_path)):
            return

        if os.path.exists(destination_path):
            messagebox.showerror(
                "Error",
                f"A PDF named '{new_filename}' already exists in this folder.",
                parent=win,
            )
            return

        try:
            os.rename(source_path, destination_path)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to rename PDF: {exc}", parent=win)
            return

        _refresh_files_list(preferred_name=new_filename)
        messagebox.showinfo(
            "Success",
            f"PDF renamed successfully to:\n{new_filename}",
            parent=win,
        )

    def _apply_selected_pdf_timestamps():
        folder = normalize_path(folder_var.get().strip())
        filename = _get_selected_file_name()
        if not folder or not filename:
            messagebox.showwarning("Warning", "Select a PDF file first.", parent=win)
            return

        file_path = normalize_path(os.path.join(folder, filename))
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "The selected PDF no longer exists.", parent=win)
            _refresh_files_list()
            return

        try:
            created_dt = _parse_editor_datetime_input(created_var.get(), "Created date")
            modified_dt = _parse_editor_datetime_input(modified_var.get(), "Modified date")
        except ValueError as exc:
            messagebox.showerror("Invalid Date", str(exc), parent=win)
            return

        try:
            _set_file_creation_and_modified_time(file_path, created_dt, modified_dt)
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to update file dates: {exc}", parent=win)
            return

        _refresh_selected_file_details()
        messagebox.showinfo(
            "Success",
            f"File dates updated successfully for:\n{filename}",
            parent=win,
        )

    def _parse_bulk_pdf_paths(raw_text):
        parsed_paths = []
        seen_paths = set()

        for raw_line in str(raw_text or "").replace("\r", "\n").split("\n"):
            line = raw_line.strip()
            if not line:
                continue

            raw_candidates = []

            if '"' in line:
                raw_candidates.extend(re.findall(r'"([^"]+)"', line))
                line = re.sub(r'"[^"]+"', ' ', line)

            if "'" in line:
                raw_candidates.extend(re.findall(r"'([^']+)'", line))
                line = re.sub(r"'[^']+'", ' ', line)

            for chunk in re.split(r";|\t", line):
                chunk = chunk.strip()
                if chunk:
                    raw_candidates.append(chunk)

            for candidate in raw_candidates:
                normalized_candidate = str(candidate or "").strip().strip('"').strip("'").strip()
                if not normalized_candidate:
                    continue

                candidate_path = normalize_path(normalized_candidate)
                if not os.path.isabs(candidate_path):
                    candidate_path = normalize_path(os.path.join(root_path, candidate_path))

                normalized_key = os.path.normcase(os.path.abspath(candidate_path))
                if normalized_key in seen_paths:
                    continue

                seen_paths.add(normalized_key)
                parsed_paths.append(candidate_path)

        return parsed_paths

    def _apply_bulk_pdf_timestamps(
        raw_paths_text=None,
        parent_window=None,
        created_text=None,
        modified_text=None,
    ):
        dialog_parent = parent_window if parent_window is not None else win
        if raw_paths_text is None:
            raw_paths_text = ""

        file_paths = _parse_bulk_pdf_paths(raw_paths_text)
        if not file_paths:
            messagebox.showwarning(
                "Warning",
                "Paste at least one PDF path first (quoted paths are supported).",
                parent=dialog_parent,
            )
            return

        created_input = created_var.get() if created_text is None else created_text
        modified_input = modified_var.get() if modified_text is None else modified_text

        try:
            created_dt = _parse_editor_datetime_input(created_input, "Created date")
            modified_dt = _parse_editor_datetime_input(modified_input, "Modified date")
        except ValueError as exc:
            messagebox.showerror("Invalid Date", str(exc), parent=dialog_parent)
            return

        updated_count = 0
        failed_items = []

        for path in file_paths:
            if not path.lower().endswith(".pdf"):
                failed_items.append(f"{path} (not a PDF)")
                continue
            if not os.path.exists(path):
                failed_items.append(f"{path} (not found)")
                continue

            try:
                _set_file_creation_and_modified_time(path, created_dt, modified_dt)
                updated_count += 1
            except OSError as exc:
                failed_items.append(f"{path} ({exc})")

        _refresh_selected_file_details()

        if updated_count > 0 and not failed_items:
            messagebox.showinfo(
                "Success",
                f"Updated dates for {updated_count} PDF file(s).",
                parent=dialog_parent,
            )
            return

        if updated_count > 0:
            details = "\n".join(failed_items[:5])
            if len(failed_items) > 5:
                details += "\n..."
            messagebox.showwarning(
                "Partial Success",
                f"Updated dates for {updated_count} PDF file(s).\n"
                f"Failed: {len(failed_items)}\n\n{details}",
                parent=dialog_parent,
            )
            return

        details = "\n".join(failed_items[:8])
        if len(failed_items) > 8:
            details += "\n..."
        messagebox.showerror(
            "Error",
            "No PDF file dates were updated.\n\n" + details,
            parent=dialog_parent,
        )

    def _open_bulk_date_editor_window():
        bulk_win = tk.Toplevel(win)
        _apply_app_icon(bulk_win)
        bulk_win.title("Bulk PDF Date Update")
        configure_window_geometry(
            bulk_win,
            760,
            520,
            min_width=620,
            min_height=420,
            margin_x=DEFAULT_MARGIN_X,
            margin_y=DEFAULT_MARGIN_Y,
        )
        bulk_win.transient(win)
        bulk_win.lift()
        bulk_win.focus_force()
        apply_theme(bulk_win)

        now_text = _format_editor_datetime(datetime.now())
        popup_created_var = tk.StringVar(value=created_var.get().strip() or now_text)
        popup_modified_var = tk.StringVar(value=modified_var.get().strip() or now_text)

        bulk_content = ttk.Frame(bulk_win, padding=(14, 12, 14, 14), style="Shell.TFrame")
        bulk_content.pack(fill="both", expand=True)

        header_card = ttk.Frame(bulk_content, style="HeaderCard.TFrame", padding=(12, 10))
        header_card.pack(fill="x", pady=(0, 10))
        ttk.Label(
            header_card,
            text="Bulk Date Update by PDF Paths",
            style="HeaderSubheading.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            header_card,
            text=(
                "Edit Created and Modified date/time below, then paste one path per line. "
                "Quoted paths are accepted."
            ),
            style="HeaderSubheading.TLabel",
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        fields_card = ttk.Frame(bulk_content, style="Card.TFrame", padding=12)
        fields_card.pack(fill="x", pady=(0, 8))
        fields_card.columnconfigure(0, weight=1)
        fields_card.columnconfigure(1, weight=1)

        ttk.Label(fields_card, text="Created Date", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(fields_card, textvariable=popup_created_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(fields_card, text="Modified Date", style="FieldLabel.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Entry(fields_card, textvariable=popup_modified_var).grid(row=1, column=1, sticky="ew")

        ttk.Label(
            fields_card,
            text="Date format: YYYY-MM-DD HH:MM[:SS]",
            style="CardMuted.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        bulk_paths_container = ttk.Frame(bulk_content, style="Card.TFrame")
        bulk_paths_container.pack(fill="both", expand=True)

        bulk_paths_scrollbar = ttk.Scrollbar(bulk_paths_container, orient="vertical")
        bulk_paths_scrollbar.pack(side="right", fill="y")

        popup_bulk_paths_text = tk.Text(
            bulk_paths_container,
            height=12,
            wrap="none",
            yscrollcommand=bulk_paths_scrollbar.set,
        )
        popup_bulk_paths_text.configure(
            bg=LISTBOX_BG,
            fg=LISTBOX_TEXT,
            insertbackground=TEXT_COLOR,
            selectbackground=ACCENT_COLOR,
            selectforeground="white",
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=LISTBOX_BORDER,
            highlightcolor=FOCUS_RING_COLOR,
            font=("Segoe UI", 9),
            padx=6,
            pady=6,
        )
        popup_bulk_paths_text.pack(side="left", fill="both", expand=True)
        bulk_paths_scrollbar.configure(command=popup_bulk_paths_text.yview)
        _mark_widget_as_scroll_list(popup_bulk_paths_text)

        bulk_actions = ttk.Frame(bulk_content, style="ActionBar.TFrame", padding=(12, 10))
        bulk_actions.pack(fill="x", pady=(10, 0))
        ttk.Label(bulk_actions, text="Bulk Update Action", style="ActionTitle.TLabel").pack(anchor="w")

        bulk_action_buttons = ttk.Frame(bulk_actions, style="ActionBar.TFrame")
        bulk_action_buttons.pack(fill="x", pady=(6, 0))

        ttk.Button(
            bulk_action_buttons,
            text="Apply Dates to Listed Paths",
            style="PrimaryAction.TButton",
            command=lambda: _apply_bulk_pdf_timestamps(
                raw_paths_text=popup_bulk_paths_text.get("1.0", tk.END),
                parent_window=bulk_win,
                created_text=popup_created_var.get(),
                modified_text=popup_modified_var.get(),
            ),
        ).pack(side="left")
        ttk.Button(
            bulk_action_buttons,
            text="Close",
            command=bulk_win.destroy,
            style="SecondaryAction.TButton",
        ).pack(side="right")

        bulk_win.bind("<Escape>", lambda _event: (bulk_win.destroy(), "break")[1])
        bulk_win.after_idle(lambda: popup_bulk_paths_text.focus_set())

    def _preview_selected_pdf():
        folder = normalize_path(folder_var.get().strip())
        filename = _get_selected_file_name()
        if not folder or not filename:
            messagebox.showwarning("Warning", "Select a PDF file first.", parent=win)
            return

        target_path = normalize_path(os.path.join(folder, filename))
        try:
            _launch_path(target_path)
        except RuntimeError as exc:
            messagebox.showerror("Error", f"Could not open PDF: {exc}", parent=win)

    content = ttk.Frame(win, style="Shell.TFrame", padding=(16, 14, 16, 16))
    content.pack(fill="both", expand=True)

    header_card = ttk.Frame(content, style="HeaderCard.TFrame", padding=(14, 12))
    header_card.pack(fill="x", pady=(0, 12))
    ttk.Label(header_card, text="Employee Details Editor", style="HeaderSubheading.TLabel").pack(anchor="w")
    ttk.Label(
        header_card,
        text=(
            "Manage employee folders and PDF metadata in one screen. "
            "Choose a folder on the left, then edit PDF details on the right."
        ),
        style="HeaderSubheading.TLabel",
        wraplength=980,
        justify="left",
    ).pack(anchor="w", pady=(4, 0))

    body_grid = ttk.Frame(content, style="Shell.TFrame")
    body_grid.pack(fill="both", expand=True)
    body_grid.columnconfigure(0, weight=4)
    body_grid.columnconfigure(1, weight=6)
    body_grid.rowconfigure(0, weight=1)

    left_column = ttk.Frame(body_grid, style="Shell.TFrame")
    left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    left_column.columnconfigure(0, weight=1)

    workspace_card = ttk.Frame(left_column, style="Card.TFrame", padding=12)
    workspace_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    workspace_card.columnconfigure(0, weight=1)
    ttk.Label(workspace_card, text="Step 1: Find Employee", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(
        workspace_card,
        text="Search an employee folder under the Records Root Folder.",
        style="CardMuted.TLabel",
    ).grid(row=1, column=0, sticky="w", pady=(2, 8))

    folder_picker_row = ttk.Frame(workspace_card, style="Card.TFrame")
    folder_picker_row.grid(row=2, column=0, sticky="ew")

    folder_field = ttk.Combobox(folder_picker_row, textvariable=folder_var, state="normal")
    folder_field.pack(side="left", fill="x", expand=True, padx=(0, 8))
    _prevent_combobox_mousewheel_value_change(folder_field)
    ttk.Button(
        folder_picker_row,
        text="Browse",
        command=_browse_employee_folder,
        style="Subtle.TButton",
    ).pack(side="left", padx=(0, 8))
    ttk.Button(
        folder_picker_row,
        text="Open Folder",
        command=_open_selected_folder,
        style="Subtle.TButton",
    ).pack(side="left")

    identity_row = ttk.Frame(workspace_card, style="Card.TFrame")
    identity_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
    identity_row.columnconfigure(0, weight=1)
    identity_row.columnconfigure(1, weight=1)

    employee_block = ttk.Frame(identity_row, style="Card.TFrame")
    employee_block.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Label(employee_block, text="Employee Name", style="FieldLabel.TLabel").pack(anchor="w")
    ttk.Label(employee_block, textvariable=employee_name_var, style="StatusStrong.TLabel").pack(anchor="w", pady=(2, 0))

    status_block = ttk.Frame(identity_row, style="Card.TFrame")
    status_block.grid(row=0, column=1, sticky="ew")
    ttk.Label(status_block, text="Current Status", style="FieldLabel.TLabel").pack(anchor="w")
    ttk.Label(status_block, textvariable=current_status_var, style="StatusStrong.TLabel").pack(anchor="w", pady=(2, 0))

    folder_ops_card = ttk.Frame(left_column, style="Card.TFrame", padding=12)
    folder_ops_card.grid(row=1, column=0, sticky="ew")
    folder_ops_card.columnconfigure(0, weight=1)
    ttk.Label(folder_ops_card, text="Step 2: Folder Operations", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

    ttk.Label(folder_ops_card, text="Rename Employee Folder", style="FieldLabel.TLabel").grid(
        row=1, column=0, sticky="w", pady=(8, 2)
    )
    rename_row = ttk.Frame(folder_ops_card, style="Card.TFrame")
    rename_row.grid(row=2, column=0, sticky="ew")

    folder_name_field = ttk.Combobox(rename_row, textvariable=folder_name_var, state="normal")
    folder_name_field.pack(side="left", fill="x", expand=True, padx=(0, 8))
    _prevent_combobox_mousewheel_value_change(folder_name_field)
    ttk.Button(
        rename_row,
        text="Rename Folder",
        command=_rename_employee_folder,
        style="SecondaryAction.TButton",
    ).pack(side="left")

    ttk.Label(folder_ops_card, text="Move Employee Status", style="FieldLabel.TLabel").grid(
        row=3, column=0, sticky="w", pady=(10, 2)
    )
    status_row = ttk.Frame(folder_ops_card, style="Card.TFrame")
    status_row.grid(row=4, column=0, sticky="ew")
    status_field = ttk.Combobox(status_row, textvariable=target_status_var, state="readonly")
    status_field.pack(side="left", fill="x", expand=True, padx=(0, 8))
    _prevent_combobox_mousewheel_value_change(status_field)
    ttk.Button(
        status_row,
        text="Apply Status",
        command=_apply_status_change,
        style="SecondaryAction.TButton",
    ).pack(side="left")

    def _handle_folder_key(event=None):
        _update_folder_combobox_suggestions(folder_field, folder_var.get(), event)

    def _handle_folder_name_key(event=None):
        _update_combobox_suggestions(folder_name_field, folder_name_var.get(), event)

    def _refresh_folder_choices():
        folder_field["values"] = _get_filtered_folder_suggestions(folder_var.get())

    def _refresh_folder_name_choices():
        folder_name_field["values"] = get_filtered_name_suggestions(folder_name_var.get())

    folder_field.bind("<KeyRelease>", _handle_folder_key)
    folder_field.bind("<<ComboboxSelected>>", lambda _event: _apply_folder_input_selection())
    folder_field.bind("<Return>", lambda _event: (_apply_folder_input_selection(), "break")[1])
    folder_field.bind("<FocusOut>", lambda _event: _apply_folder_input_selection(), add="+")
    folder_field.configure(postcommand=_refresh_folder_choices)
    setattr(
        folder_field,
        "_on_suggestion_selected",
        lambda selected_value: _apply_folder_input_selection(selected_value),
    )

    folder_name_field.bind("<KeyRelease>", _handle_folder_name_key)
    folder_name_field.configure(postcommand=_refresh_folder_name_choices)

    right_column = ttk.Frame(body_grid, style="Shell.TFrame")
    right_column.grid(row=0, column=1, sticky="nsew")

    files_card = ttk.Frame(right_column, style="Card.TFrame", padding=12)
    files_card.pack(fill="both", expand=True)

    ttk.Label(files_card, text="Step 3: PDF Metadata Editor", style="SectionTitle.TLabel").pack(anchor="w")
    ttk.Label(files_card, textvariable=files_count_var, style="CardMuted.TLabel").pack(anchor="w", pady=(2, 0))
    ttk.Label(
        files_card,
        text="Select a PDF from the list, then edit filename or timestamp metadata.",
        style="CardMuted.TLabel",
    ).pack(anchor="w", pady=(0, 6))

    files_area = ttk.Frame(files_card, style="Card.TFrame")
    files_area.pack(fill="both", expand=True, pady=(8, 0))
    files_area.columnconfigure(0, weight=5)
    files_area.columnconfigure(1, weight=6)
    files_area.rowconfigure(0, weight=1)

    list_area = ttk.Frame(files_area, style="Card.TFrame", padding=(10, 8))
    list_area.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    ttk.Label(list_area, text="PDF List", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 6))
    ttk.Label(
        list_area,
        textvariable=files_empty_hint_var,
        style="CardMuted.TLabel",
        wraplength=340,
        justify="left",
    ).pack(anchor="w", pady=(0, 6))

    list_body = ttk.Frame(list_area, style="Card.TFrame")
    list_body.pack(fill="both", expand=True)

    files_scrollbar = ttk.Scrollbar(list_body, orient="vertical")
    files_scrollbar.pack(side="right", fill="y")

    files_listbox = tk.Listbox(list_body, yscrollcommand=files_scrollbar.set)
    _apply_modern_listbox_style(files_listbox, compact=True, export_selection=False)
    files_listbox.pack(side="left", fill="both", expand=True)
    files_scrollbar.configure(command=files_listbox.yview)
    _mark_widget_as_scroll_list(files_listbox)

    editor_area = ttk.Frame(files_area, style="Card.TFrame", padding=(10, 8))
    editor_area.grid(row=0, column=1, sticky="nsew")
    ttk.Label(editor_area, text="Selected PDF Details", style="FieldLabel.TLabel").pack(anchor="w")
    ttk.Label(editor_area, textvariable=selected_file_var, style="StatusStrong.TLabel", wraplength=420).pack(anchor="w", pady=(2, 0))

    ttk.Label(editor_area, text="File Name", style="FieldLabel.TLabel").pack(anchor="w", pady=(10, 2))
    ttk.Entry(editor_area, textvariable=file_name_var).pack(fill="x")

    file_actions = ttk.Frame(editor_area, style="Card.TFrame")
    file_actions.pack(fill="x", pady=(8, 0))
    file_actions.columnconfigure(0, weight=1)
    file_actions.columnconfigure(1, weight=1)

    ttk.Button(
        file_actions,
        text="Rename File",
        command=_rename_selected_pdf,
        style="Subtle.TButton",
    ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
    ttk.Button(
        file_actions,
        text="Preview",
        command=_preview_selected_pdf,
        style="Subtle.TButton",
    ).grid(row=0, column=1, sticky="ew")
    ttk.Button(
        file_actions,
        text="Bulk Date Update",
        command=_open_bulk_date_editor_window,
        style="Subtle.TButton",
    ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    ttk.Separator(editor_area, orient="horizontal").pack(fill="x", pady=12)

    ttk.Label(editor_area, text="Created Date", style="FieldLabel.TLabel").pack(anchor="w")
    ttk.Entry(editor_area, textvariable=created_var).pack(fill="x")

    ttk.Label(editor_area, text="Modified Date", style="FieldLabel.TLabel").pack(anchor="w", pady=(10, 0))
    ttk.Entry(editor_area, textvariable=modified_var).pack(fill="x")

    ttk.Label(
        editor_area,
        text="Date format: YYYY-MM-DD HH:MM[:SS]",
        style="CardMuted.TLabel",
    ).pack(anchor="w", pady=(6, 0))

    ttk.Button(
        editor_area,
        text="Apply File Dates",
        command=_apply_selected_pdf_timestamps,
        style="PrimaryAction.TButton",
    ).pack(anchor="w", pady=(12, 0))

    def _close_employee_details_window():
        nonlocal folder_watch_after_id

        if folder_watch_after_id is not None:
            try:
                win.after_cancel(folder_watch_after_id)
            except tk.TclError:
                pass
            folder_watch_after_id = None

        if win.winfo_exists():
            win.destroy()

    bottom_actions = ttk.Frame(content, style="Shell.TFrame")
    bottom_actions.pack(fill="x", pady=(8, 0))
    ttk.Button(
        bottom_actions,
        text="Close",
        command=_close_employee_details_window,
        style="SecondaryAction.TButton",
    ).pack(side="right")

    files_listbox.bind("<<ListboxSelect>>", lambda _event=None: _refresh_selected_file_details())

    _refresh_status_choices()
    _refresh_folder_autocomplete_catalog()
    _refresh_folder_choices()
    folder_watch_snapshot = _build_folder_pdf_snapshot(folder_var.get().strip())
    try:
        folder_watch_after_id = win.after(1200, _poll_folder_file_changes)
    except tk.TclError:
        folder_watch_after_id = None
    win.bind("<Escape>", lambda _event: (_close_employee_details_window(), "break")[1])
    win.after_idle(lambda: folder_field.focus_set())
    win.protocol("WM_DELETE_WINDOW", _close_employee_details_window)


def show_selected_batch_files_window(
    batch_files=None,
    current_filename="",
    parent_window=None,
    batch_context=None,
    cancel_batch_callback=None,
):
    if batch_context is not None and isinstance(batch_context.get("files"), list):
        selected_files = batch_context["files"]
    else:
        selected_files = [
            str(name).strip()
            for name in (list(batch_files) if batch_files is not None else get_selected_pending_files())
            if str(name).strip()
        ]

    owner = parent_window if (parent_window is not None and parent_window.winfo_exists()) else root
    restore_parent_grab = False
    if parent_window is not None and parent_window.winfo_exists():
        try:
            if parent_window.grab_current() is parent_window:
                parent_window.grab_release()
                restore_parent_grab = True
        except tk.TclError:
            pass

    win = tk.Toplevel(owner)
    _apply_app_icon(win)
    win.title("Selected Batch Files")
    configure_window_geometry(
        win,
        560,
        620,
        min_width=420,
        min_height=360,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(owner)
    win.lift()
    win.focus_force()
    win.grab_set()
    apply_theme(win)

    def _close_batch_window():
        try:
            if win.grab_current() is win:
                win.grab_release()
        except tk.TclError:
            pass

        if win.winfo_exists():
            win.destroy()

        if restore_parent_grab and parent_window is not None and parent_window.winfo_exists():
            try:
                parent_window.grab_set()
                parent_window.lift()
                parent_window.focus_force()
            except tk.TclError:
                pass

    win.protocol("WM_DELETE_WINDOW", _close_batch_window)
    win.bind("<Escape>", lambda _event: (_close_batch_window(), "break")[1])

    content = ttk.Frame(win, padding=16, style="TFrame")
    content.pack(fill="both", expand=True)

    count_text_var = tk.StringVar(value="Selected for batch: 0 files")

    ttk.Label(
        content,
        textvariable=count_text_var,
        style="Card.TLabel",
    ).pack(anchor="w", pady=(0, 10))

    empty_hint_label = ttk.Label(
        content,
        text="No pending files are selected. Use the checkboxes in Pending Files first.",
        style="Subheading.TLabel",
        justify="left",
        wraplength=500,
    )

    if current_filename:
        ttk.Label(
            content,
            text=f"Current in progress: {current_filename}",
            style="Subheading.TLabel",
            justify="left",
            wraplength=500,
        ).pack(anchor="w", pady=(0, 8))

    list_container = ttk.Frame(content, style="Card.TFrame", padding=8)
    list_container.pack(fill="both", expand=True)

    files_scrollbar = ttk.Scrollbar(list_container, orient="vertical")
    files_scrollbar.pack(side="right", fill="y")

    files_listbox = tk.Listbox(
        list_container,
        selectmode="extended",
        yscrollcommand=files_scrollbar.set,
    )
    _apply_modern_listbox_style(files_listbox, compact=True, export_selection=False)
    files_listbox.pack(side="left", fill="both", expand=True)
    files_scrollbar.configure(command=files_listbox.yview)

    actions = ttk.Frame(content, style="TFrame")
    actions.pack(fill="x", pady=(12, 0))

    preview_button = ttk.Button(actions, text="Preview Selected")
    preview_button.pack(side="left")

    remove_button = ttk.Button(actions, text="Remove Selected")
    remove_button.pack(side="left", padx=(8, 0))

    cancel_button = ttk.Button(actions, text="Cancel All Batch", style="Accent.TButton")
    cancel_button.pack(side="left", padx=(8, 0))

    ttk.Button(actions, text="Close", command=_close_batch_window).pack(side="right")

    def _refresh_files_list():
        selected_count = len(selected_files)
        file_text = "file" if selected_count == 1 else "files"
        count_text_var.set(f"Selected for batch: {selected_count} {file_text}")

        files_listbox.delete(0, tk.END)
        for file_name in selected_files:
            files_listbox.insert(tk.END, file_name)

        if selected_count == 0:
            if not empty_hint_label.winfo_ismapped():
                empty_hint_label.pack(anchor="w", pady=(0, 8))
        else:
            if empty_hint_label.winfo_ismapped():
                empty_hint_label.pack_forget()

        initial_index = 0
        if current_filename and current_filename in selected_files:
            initial_index = selected_files.index(current_filename)

        if files_listbox.size() > 0:
            files_listbox.selection_set(initial_index)
            files_listbox.activate(initial_index)
            files_listbox.see(initial_index)

        if selected_count > 0:
            preview_button.state(["!disabled"])
            remove_button.state(["!disabled"])
        else:
            preview_button.state(["disabled"])
            remove_button.state(["disabled"])

        if batch_context is not None and not batch_context.get("cancelled", False):
            cancel_button.state(["!disabled"])
        else:
            cancel_button.state(["disabled"])

    def _preview_selected_file(_event=None):
        selection = files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Select one file to preview.", parent=win)
            return "break"

        target_file = files_listbox.get(selection[0])
        preview_specific_pending_pdf(target_file)
        return "break"

    def _remove_selected_batch_files():
        selection = files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Select one or more files to remove from the batch.", parent=win)
            return

        chosen_files = [files_listbox.get(idx) for idx in selection]
        removable_files = []
        blocked_current = False

        for file_name in chosen_files:
            if current_filename and file_name == current_filename:
                blocked_current = True
                continue
            removable_files.append(file_name)

        if blocked_current:
            messagebox.showwarning(
                "Current File Locked",
                "The file currently being processed cannot be removed. Use Cancel All Batch to stop processing.",
                parent=win,
            )

        if not removable_files:
            return

        if not messagebox.askyesno(
            "Remove Selected",
            f"Remove {len(removable_files)} selected file(s) from this batch queue?",
            parent=win,
        ):
            return

        for file_name in removable_files:
            try:
                selected_files.remove(file_name)
            except ValueError:
                pass

        _refresh_files_list()

    def _cancel_all_batch_processing():
        if batch_context is None:
            messagebox.showwarning(
                "Not Available",
                "Batch cancellation is only available while running an active batch process.",
                parent=win,
            )
            return

        if not messagebox.askyesno(
            "Cancel Batch Processing",
            "Cancel processing for all remaining batch files?",
            parent=win,
        ):
            return

        batch_context["cancelled"] = True
        _close_batch_window()

        if callable(cancel_batch_callback):
            cancel_batch_callback()

    files_listbox.bind("<Double-Button-1>", _preview_selected_file)
    files_listbox.bind("<Return>", _preview_selected_file)
    preview_button.configure(command=_preview_selected_file)
    remove_button.configure(command=_remove_selected_batch_files)
    cancel_button.configure(command=_cancel_all_batch_processing)

    _refresh_files_list()


def _create_startup_loading_window(
    title="Loading PDF Record Manager",
    heading="Starting PDF Record Manager",
    initial_status="Preparing startup...",
):
    loading_win = tk.Toplevel(root)
    _apply_app_icon(loading_win)
    loading_win.title(title)
    loading_win.resizable(False, False)
    loading_win.transient(root)
    # Keep this dialog above the app window only, not above all apps.
    loading_win.lift(root)
    loading_win.protocol("WM_DELETE_WINDOW", lambda: None)

    container = ttk.Frame(loading_win, style="Card.TFrame", padding=16)
    container.pack(fill="both", expand=True, padx=12, pady=12)

    ttk.Label(container, text=heading, style="Card.TLabel").pack(anchor="w")
    status_var = tk.StringVar(value=initial_status)
    detail_var = tk.StringVar(value="")

    ttk.Label(container, textvariable=status_var, style="Subheading.TLabel", anchor="w").pack(fill="x", pady=(6, 4))
    progress_bar = ttk.Progressbar(
        container,
        mode="determinate",
        maximum=1,
        value=0,
        length=380,
        style="Success.Horizontal.TProgressbar",
    )
    progress_bar.pack(fill="x")
    ttk.Label(container, textvariable=detail_var, style="Subheading.TLabel", anchor="w").pack(fill="x", pady=(4, 0))

    loading_win.update_idletasks()
    width = loading_win.winfo_reqwidth()
    height = loading_win.winfo_reqheight()
    work_x, work_y, work_w, work_h = _get_display_work_area(root)
    x_pos = work_x + max(0, (work_w - width) // 2)
    y_pos = work_y + max(0, (work_h - height) // 2)
    loading_win.geometry(f"{width}x{height}+{x_pos}+{y_pos}")

    return loading_win, status_var, detail_var, progress_bar


def _update_startup_loading_ui(loading_win, status_var, detail_var, progress_bar, message, current=None, total=None):
    status_var.set(message)

    if total is None or total <= 0:
        if str(progress_bar.cget("mode")) != "indeterminate":
            progress_bar.configure(mode="indeterminate")
        if not getattr(progress_bar, "_is_running", False):
            progress_bar.start(12)
            progress_bar._is_running = True
        detail_var.set("Working...")
    else:
        if getattr(progress_bar, "_is_running", False):
            progress_bar.stop()
            progress_bar._is_running = False
        progress_bar.configure(mode="determinate", maximum=max(1, total), value=max(0, min(current or 0, total)))
        detail_var.set(f"{current or 0}/{total} source(s)")

    try:
        loading_win.update_idletasks()
        loading_win.update()
    except tk.TclError:
        pass

    try:
        root.update_idletasks()
    except tk.TclError:
        pass


def initialize_settings(progress_callback=None):
    if progress_callback is not None:
        progress_callback("Loading saved settings...", 0, 0)

    load_settings(progress_callback=progress_callback)
    _refresh_update_status()

    if progress_callback is not None:
        progress_callback("Loading pending queue...", 0, 0)

    if pending_folder.get():
        load_pending_files()

# ------------------------
# UI Layout
# ------------------------

scroll_container, scroll_frame = create_scrollable_panel(root)
scroll_container.pack(fill="both", expand=True)

main_container = ttk.Frame(scroll_frame, padding=(24, 20, 24, 24), style="Shell.TFrame")
main_container.pack(fill="both", expand=True)

header_card = ttk.Frame(main_container, style="HeaderCard.TFrame", padding=(18, 14))
header_card.pack(fill="x", pady=(0, 14))

header_top_row = ttk.Frame(header_card, style="HeaderCard.TFrame")
header_top_row.pack(fill="x")
ttk.Label(header_top_row, text="PDF Record Manager", style="HeaderTitle.TLabel").pack(side="left", anchor="w")

header_status_row = ttk.Frame(header_card, style="HeaderCard.TFrame")
header_status_row.pack(fill="x", pady=(6, 0))
ttk.Label(header_status_row, text="Update Status:", style="HeaderSubheading.TLabel").pack(side="left", anchor="w")
header_status_badge = ttk.Frame(header_status_row, style="StatusBadge.TFrame", padding=(8, 3))
header_status_badge.pack(side="left", anchor="w", padx=(8, 0))
ttk.Label(
    header_status_badge,
    textvariable=update_status_var,
    style="StatusBadge.TLabel",
).pack(anchor="w")

ttk.Label(
    header_card,
    text="Review pending scans, create new folders, and merge updates without leaving this window.",
    style="HeaderSubheading.TLabel",
    wraplength=760,
    justify="left",
).pack(anchor="w", pady=(4, 0))

content_grid = ttk.Frame(main_container, style="Shell.TFrame")
content_grid.pack(fill="both", expand=True)
content_grid.columnconfigure(0, weight=4)
content_grid.columnconfigure(1, weight=6)
content_grid.rowconfigure(0, weight=1)

left_column = ttk.Frame(content_grid, style="Shell.TFrame")
left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

right_column = ttk.Frame(content_grid, style="Shell.TFrame")
right_column.grid(row=0, column=1, sticky="nsew")

paths_card = ttk.Frame(left_column, style="Card.TFrame", padding=16)
paths_card.pack(fill="x", pady=(0, 12))
ttk.Label(paths_card, text="Workspace Paths", style="SectionTitle.TLabel").pack(anchor="w")

pending_path_row = ttk.Frame(paths_card, style="Card.TFrame")
pending_path_row.pack(fill="x", pady=(10, 6))
pending_path_input = ttk.Frame(pending_path_row, style="Card.TFrame")
pending_path_input.pack(side="left", fill="x", expand=True, padx=(0, 8))
ttk.Label(pending_path_input, text="Pending Folder", style="FieldLabel.TLabel").pack(anchor="w")
ttk.Entry(pending_path_input, textvariable=pending_folder).pack(fill="x", pady=(4, 0))
ttk.Button(
    pending_path_row,
    text="Browse",
    command=select_pending_folder,
    style="Subtle.TButton",
    width=9,
).pack(side="right", anchor="s")

root_path_row = ttk.Frame(paths_card, style="Card.TFrame")
root_path_row.pack(fill="x")
root_path_input = ttk.Frame(root_path_row, style="Card.TFrame")
root_path_input.pack(side="left", fill="x", expand=True, padx=(0, 8))
ttk.Label(root_path_input, text="Records Root Folder", style="FieldLabel.TLabel").pack(anchor="w")
ttk.Entry(root_path_input, textvariable=root_folder).pack(fill="x", pady=(4, 0))
ttk.Button(
    root_path_row,
    text="Browse",
    command=select_root_folder,
    style="Subtle.TButton",
    width=9,
).pack(side="right", anchor="s")

names_card = ttk.Frame(left_column, style="Card.TFrame", padding=16)
names_card.pack(fill="both", expand=True)
ttk.Label(names_card, text="Employee Name Sources", style="SectionTitle.TLabel").pack(anchor="w")
ttk.Label(
    names_card,
    text="Supports PDF, Excel, CSV, and TXT source files.",
    style="CardMuted.TLabel",
).pack(anchor="w", pady=(2, 8))

employee_sources_container = ttk.Frame(names_card, style="Card.TFrame")
employee_sources_container.pack(fill="both", expand=True, pady=(6, 4))
employee_sources_scrollbar = ttk.Scrollbar(employee_sources_container, orient="vertical")
employee_sources_scrollbar.pack(side="right", fill="y")

employee_sources_listbox = tk.Listbox(
    employee_sources_container,
    height=6,
    selectmode="extended",
    yscrollcommand=employee_sources_scrollbar.set,
)
_apply_modern_listbox_style(employee_sources_listbox)
employee_sources_listbox.pack(side="left", fill="x", expand=True)
employee_sources_scrollbar.configure(command=employee_sources_listbox.yview)
_mark_widget_as_scroll_list(employee_sources_listbox)


def _select_all_employee_sources(_event=None):
    if employee_sources_listbox is None or not employee_sources_listbox.winfo_exists():
        return "break"

    list_size = employee_sources_listbox.size()
    if list_size <= 0:
        return "break"

    employee_sources_listbox.selection_set(0, tk.END)
    employee_sources_listbox.activate(0)
    employee_sources_listbox.see(0)
    _refresh_employee_sources_selection_count()
    return "break"


employee_sources_listbox.bind("<Control-a>", _select_all_employee_sources, add="+")
employee_sources_listbox.bind("<Control-A>", _select_all_employee_sources, add="+")
employee_sources_listbox.bind("<<ListboxSelect>>", _refresh_employee_sources_selection_count, add="+")
_refresh_employee_sources_listbox()

ttk.Label(
    names_card,
    textvariable=employee_sources_selection_count_var,
    style="CardMuted.TLabel",
    padding=(0, 2),
    anchor="w",
).pack(fill="x")

ui_icon_images = _build_pending_toolbar_icon_images()

name_buttons = ttk.Frame(names_card, style="Card.TFrame")
name_buttons.pack(fill="x", pady=(0, 4))
name_buttons_container = name_buttons

add_sources_button = ttk.Button(name_buttons, style="ToolbarIcon.TButton", command=select_employee_sources)
add_sources_button.pack(side="left", padx=(0, 6))
_configure_icon_button(add_sources_button, "source_add", TOOLBAR_ICON_SOURCE_ADD, "Add Sources")
_attach_hover_tooltip(add_sources_button, "Add employee source files")

remove_sources_button = ttk.Button(name_buttons, style="ToolbarIcon.TButton", command=remove_selected_employee_source)
remove_sources_button.pack(side="left", padx=6)
_configure_icon_button(remove_sources_button, "source_remove", TOOLBAR_ICON_SOURCE_REMOVE, "Remove Selected")
_attach_hover_tooltip(remove_sources_button, "Remove selected source entries")

clear_sources_button = ttk.Button(name_buttons, style="ToolbarIcon.TButton", command=clear_employee_sources)
clear_sources_button.pack(side="left", padx=6)
_configure_icon_button(clear_sources_button, "clear_selection", TOOLBAR_ICON_SOURCE_CLEAR, "Clear All")
_attach_hover_tooltip(clear_sources_button, "Clear all employee source entries")

view_parsed_names_button = ttk.Button(name_buttons, style="ToolbarIcon.TButton", command=show_parsed_names_window)
view_parsed_names_button.pack(side="left", padx=6)
_configure_icon_button(view_parsed_names_button, "preview", TOOLBAR_ICON_PREVIEW, "View Parsed")
_attach_hover_tooltip(view_parsed_names_button, "View all parsed employee names")

employee_sources_status_surface = ttk.Frame(names_card, style="StatusSurface.TFrame", padding=(8, 6))
employee_sources_status_surface.pack(fill="x", pady=(6, 0))
ttk.Label(
    employee_sources_status_surface,
    textvariable=employee_list_status_var,
    style="StatusSurface.TLabel",
    anchor="w",
    justify="left",
    wraplength=420,
).pack(fill="x")

list_card = ttk.Frame(right_column, style="Card.TFrame", padding=16)
list_card.pack(fill="both", expand=True)
header_row = ttk.Frame(list_card, style="Card.TFrame")
header_row.pack(fill="x")
ttk.Label(header_row, text="Pending Queue", style="SectionTitle.TLabel").pack(side="left")
ttk.Label(header_row, textvariable=pending_files_count_var, style="CardMuted.TLabel").pack(side="left", padx=(8, 0))

listbox_container = ttk.Frame(list_card, style="Card.TFrame")
listbox_container.pack(fill="both", expand=True, pady=(10, 0))

pending_master_row = ttk.Frame(listbox_container, style="Card.TFrame", padding=(6, 2))
pending_master_row.pack(fill="x", padx=2, pady=(0, 3))

pending_master_toggle_button = ttk.Button(
    pending_master_row,
    style="ToolbarIcon.TButton",
    command=_on_pending_master_toggle_clicked,
)
pending_master_toggle_button.pack(side="left", anchor="w")
_attach_hover_tooltip(pending_master_toggle_button, "Toggle all pending selections")

pending_master_actions = ttk.Frame(pending_master_row, style="Card.TFrame")
pending_master_actions.pack(side="right")

pending_rotate_button = ttk.Button(
    pending_master_actions,
    style="ToolbarIcon.TButton",
    command=rotate_selected_pending_pdfs,
)
pending_rotate_button.pack(side="left", padx=4)
_attach_hover_tooltip(pending_rotate_button, "Rotate pages in selected pending PDFs")

pending_preview_button = ttk.Button(
    pending_master_actions,
    style="ToolbarIcon.TButton",
    command=preview_selected_pdf,
)
pending_preview_button.pack(side="left", padx=4)
_attach_hover_tooltip(pending_preview_button, "Preview the currently selected PDF")

_update_icon_button_labels()

ttk.Separator(listbox_container, orient="horizontal").pack(fill="x", padx=2, pady=(0, 2))
_refresh_pending_master_toggle_state()

pending_canvas = tk.Canvas(
    listbox_container,
    highlightthickness=0,
    bg="#0b1220",
    bd=0,
)
pending_canvas.pack(side="left", fill="both", expand=True)
pending_scrollbar = ttk.Scrollbar(listbox_container, orient="vertical", command=pending_canvas.yview)
pending_scrollbar.pack(side="right", fill="y")
pending_canvas.configure(yscrollcommand=pending_scrollbar.set)
pending_canvas_widget = pending_canvas
pending_canvas.configure(takefocus=1)
_mark_widget_as_scroll_canvas(pending_canvas)
_ensure_global_mousewheel_binding()

pending_canvas.bind("<Button-1>", lambda _event: _focus_pending_selection_surface(), add="+")
pending_canvas.bind("<Control-a>", _on_pending_ctrl_select_all, add="+")
pending_canvas.bind("<Control-A>", _on_pending_ctrl_select_all, add="+")

pending_items_frame = ttk.Frame(pending_canvas, style="Card.TFrame")
pending_canvas_window = pending_canvas.create_window((0, 0), window=pending_items_frame, anchor="nw")
pending_items_frame.bind("<Button-1>", lambda _event: _focus_pending_selection_surface(), add="+")
pending_items_frame.bind("<Control-a>", _on_pending_ctrl_select_all, add="+")
pending_items_frame.bind("<Control-A>", _on_pending_ctrl_select_all, add="+")

def _resize_pending_canvas(event):
    pending_canvas.itemconfigure(pending_canvas_window, width=event.width)


def _update_pending_scrollregion(_event=None):
    pending_canvas.configure(scrollregion=pending_canvas.bbox("all"))


pending_canvas.bind("<Configure>", _resize_pending_canvas)
pending_items_frame.bind("<Configure>", _update_pending_scrollregion)

actions_card = ttk.Frame(main_container, style="ActionBar.TFrame", padding=(16, 12))
actions_card.pack(fill="x", pady=(14, 0))
ttk.Label(actions_card, text="Core Operations", style="ActionTitle.TLabel").pack(anchor="w")

btn_frame = ttk.Frame(actions_card, style="ActionBar.TFrame")
btn_frame.pack(fill="x", pady=(8, 0))
btn_frame.columnconfigure(0, weight=1)
btn_frame.columnconfigure(1, weight=1)
btn_frame.columnconfigure(2, weight=1)
ttk.Button(
    btn_frame,
    text="New Record",
    width=20,
    command=start_new_record_batch,
    style="PrimaryAction.TButton",
).grid(row=0, column=0, padx=8)
ttk.Button(
    btn_frame,
    text="Merge Existing",
    width=20,
    command=start_merge_existing_batch,
    style="SecondaryAction.TButton",
).grid(row=0, column=1, padx=8)
ttk.Button(
    btn_frame,
    text="Edit Employee Details",
    width=22,
    command=employee_details_editor_window,
    style="SecondaryAction.TButton",
).grid(row=0, column=2, padx=8)

def _run_startup_sequence():
    _center_window_to_current_size(root)
    loading_win, loading_status_var, loading_detail_var, loading_progress_bar = _create_startup_loading_window()
    try:
        _update_startup_loading_ui(
            loading_win,
            loading_status_var,
            loading_detail_var,
            loading_progress_bar,
            "Starting application...",
            0,
            0,
        )

        initialize_settings(
            progress_callback=lambda message, current=None, total=None: _update_startup_loading_ui(
                loading_win,
                loading_status_var,
                loading_detail_var,
                loading_progress_bar,
                message,
                current,
                total,
            )
        )
        _ensure_auto_refresh_job()
        root.after(1200, lambda: check_for_updates(manual=False))
    finally:
        if getattr(loading_progress_bar, "_is_running", False):
            loading_progress_bar.stop()
            loading_progress_bar._is_running = False
        if loading_win.winfo_exists():
            loading_win.destroy()
        _center_window_to_current_size(root)


_center_window_to_current_size(root)
root.after(10, _run_startup_sequence)

# Run App
root.mainloop()