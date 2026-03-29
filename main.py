import os
import sys
import json
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
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    Image = None
    ImageDraw = None
    ImageTk = None

PDF_IMPORT_ERROR = ""
PdfReader = None
try:
    from pypdf import PdfMerger, PdfReader
except ImportError as pypdf_error:
    try:
        from PyPDF2 import PdfMerger, PdfReader  # fallback for older installs
    except ImportError as pypdf2_error:
        PdfMerger = None
        PdfReader = None
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
PENDING_ROW_BG = "#121d31"
PENDING_ROW_HOVER_BG = "#1a2a45"
PENDING_ROW_TEXT = "#e5edff"
LISTBOX_BG = "#101a2b"
LISTBOX_BORDER = "#2f405d"
LISTBOX_TEXT = "#e5ecf8"

AUTO_REFRESH_INTERVAL_MS = 1000
APP_ICON_PREFERRED_NAMES = ("app.ico", "application.ico", "icon.ico")
APP_VERSION = "1.0.0"
APP_BUILD_COMMIT = os.environ.get("PDF_AUTOTOOL_COMMIT", "unknown")
APP_BUILD_DATE = os.environ.get("PDF_AUTOTOOL_BUILD_DATE", "unknown")
DEFAULT_UPDATE_MANIFEST_URL = ""
UPDATE_CHECK_TIMEOUT_SEC = 8
TOOLBAR_ICON_SIZE = 20
TOOLBAR_ICON_COLOR = "#FFFFFF"
TOOLBAR_ICON_STROKE_MULTIPLIER = 1.55
TOOLBAR_ICON_RENDER_SCALE = 6
TOOLBAR_ICON_REFRESH = "\u21bb"
TOOLBAR_ICON_PREVIEW = "\u25c9"
TOOLBAR_ICON_SELECT_ALL = "\u2611"
TOOLBAR_ICON_SELECT_NONE = "\u2610"
TOOLBAR_ICON_SELECT_PARTIAL = "\u25a3"
TOOLBAR_ICON_SOURCE_ADD = "\u2795"
TOOLBAR_ICON_SOURCE_REMOVE = "\u2796"
TOOLBAR_ICON_SOURCE_CLEAR = "\u2715"


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
    window.configure(bg=BG_COLOR)
    style = ttk.Style(window)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "TFrame",
        background=BG_COLOR,
    )
    style.configure(
        "Card.TFrame",
        background=SURFACE_COLOR,
        relief="flat",
    )
    style.configure(
        "Title.TLabel",
        background=BG_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI Semibold", 20),
    )
    style.configure(
        "Subheading.TLabel",
        background=BG_COLOR,
        foreground=SUBTEXT_COLOR,
        font=("Segoe UI", 10),
    )
    style.configure(
        "Card.TLabel",
        background=SURFACE_COLOR,
        foreground=TEXT_COLOR,
        font=("Segoe UI", 11),
    )
    style.configure(
        "TLabel",
        background=BG_COLOR,
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
        fieldbackground="#131c2f",
        foreground=TEXT_COLOR,
        insertcolor=TEXT_COLOR,
        bordercolor="#334155",
        relief="flat",
    )
    style.configure(
        "TCombobox",
        fieldbackground="#131c2f",
        foreground=TEXT_COLOR,
        arrowcolor=ACCENT_COLOR,
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
        background="#334155",
        troughcolor="#0b1220",
        bordercolor="#0b1220",
        lightcolor="#334155",
        darkcolor="#334155",
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
        background="#334155",
        troughcolor="#0b1220",
        bordercolor="#0b1220",
        lightcolor="#334155",
        darkcolor="#334155",
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
        return None

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
        "highlightcolor": ACCENT_COLOR,
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
employee_name_suggestions = []
name_filter_mode = tk.StringVar(value="strict")
_suppress_name_filter_refresh = False

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
            candidate = _extract_comma_prefix_candidate(base_candidate, max_given_tokens=6)
        else:
            candidate = base_candidate

        if _line_passes_filter(candidate):
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
                candidate = _extract_comma_prefix_candidate(normalized, max_given_tokens=6)
            else:
                candidate = normalized

            if _line_passes_filter(candidate):
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

        # After a middle initial, additional tokens are usually from the next column.
        if given_tokens and _is_initial_token(given_tokens[-1]) and token.lower() not in _NAME_SUFFIXES:
            break

        # In many table exports the nickname column repeats a prior name token.
        if given_tokens and token.lower() in {existing.lower() for existing in given_tokens}:
            break

        # Stop before compact uppercase codes from following columns.
        if given_tokens and token.isupper() and len(token) <= 3:
            break

        given_tokens.append(token)

        if max_given_tokens and len(given_tokens) >= max_given_tokens:
            break
        if len(given_tokens) >= 6:
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
            if _line_passes_filter(candidate) and not _is_header_row(candidate):
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
        if _line_passes_filter(candidate) and not _is_header_row(candidate):
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
        return 3
    counts.sort()
    percentile_index = int((len(counts) - 1) * 0.9)
    inferred = counts[percentile_index]
    return max(2, min(6, inferred))


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
        if _line_passes_filter(candidate):
            candidates.append(candidate)
    return candidates


def _line_passes_filter(normalized_line):
    if not normalized_line:
        return False
    mode = name_filter_mode.get()
    if mode == "strict":
        return "," in normalized_line and any(ch.isalpha() for ch in normalized_line)
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
        if _line_passes_filter(candidate):
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
                if _line_passes_filter(candidate):
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
        return
    employee_sources_listbox.delete(0, tk.END)
    for path in employee_source_paths:
        employee_sources_listbox.insert(tk.END, os.path.basename(path))


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
        state = {"popup": None, "listbox": None}
        setattr(combobox, "_suggestion_popup_state", state)
    return state


def _hide_suggestion_popup(combobox):
    state = _get_suggestion_popup_state(combobox)
    popup = state.get("popup")
    if popup is not None and popup.winfo_exists():
        popup.destroy()
    state["popup"] = None
    state["listbox"] = None


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
        if popup is None or not popup.winfo_exists():
            return

        focus_widget = combobox.focus_get()
        if focus_widget is combobox:
            return
        if listbox is not None and focus_widget is listbox:
            return
        if focus_widget is not None and str(focus_widget).startswith(str(popup)):
            return

        _hide_suggestion_popup(combobox)

    combobox.after(delay_ms, _close_if_focus_left)


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
        popup_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=popup_scrollbar.set)

        listbox.pack(side="left", fill="both", expand=True)
        popup_scrollbar.pack(side="right", fill="y")

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

    combobox.update_idletasks()
    popup.update_idletasks()

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


def _get_about_date_text():
    if APP_BUILD_DATE and APP_BUILD_DATE != "unknown":
        return APP_BUILD_DATE

    target_path = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    try:
        modified_dt = datetime.fromtimestamp(os.path.getmtime(target_path)).astimezone()
        return modified_dt.isoformat(timespec="seconds")
    except Exception:
        return "unknown"


def show_about_dialog():
    os_name = os.environ.get("OS") or platform.system() or "unknown"
    architecture = platform.machine() or "unknown"
    os_version = platform.version() or "unknown"
    update_feed_value = DEFAULT_UPDATE_MANIFEST_URL.strip() or "Unavailable in this build"

    details = [
        "App: PDF Record Manager",
        f"Version: {APP_VERSION}",
        f"Install Type: {_get_installation_scope()}",
        f"Build Commit: {APP_BUILD_COMMIT}",
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
    if len(parts) < 3:
        return None

    last_two = parts[-2:]
    if not all(segment.isdigit() for segment in last_two):
        return None

    latest_year, earliest_year = last_two
    employee_name = "_".join(parts[:-2])
    return {
        "name": employee_name,
        "latest": latest_year,
        "earliest": earliest_year,
    }


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
        
pending_items_frame = None
pending_file_vars = {}
pending_snapshot = set()
auto_refresh_job_id = None
ui_icon_images = {}
pending_row_preview_buttons = []
add_sources_button = None
remove_sources_button = None
clear_sources_button = None
pending_refresh_button = None
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


def _update_icon_button_labels():
    button_specs = (
        (add_sources_button, "source_add", TOOLBAR_ICON_SOURCE_ADD, "Add Sources"),
        (remove_sources_button, "source_remove", TOOLBAR_ICON_SOURCE_REMOVE, "Remove Selected"),
        (clear_sources_button, "clear_selection", TOOLBAR_ICON_SOURCE_CLEAR, "Clear All"),
        (pending_refresh_button, "refresh", TOOLBAR_ICON_REFRESH, "Refresh"),
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
        "none": "Select All",
        "partial": "Partially Selected",
        "all": "All Selected",
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


def _refresh_pending_master_toggle_state():
    global pending_master_selection_state

    if pending_master_toggle_button is None or not pending_master_toggle_button.winfo_exists():
        pending_master_selection_state = "none"
        return

    total_count, selected_count = _get_pending_selection_counts()

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
    if hovered:
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


def load_pending_files():
    global pending_file_vars, pending_row_preview_buttons
    if pending_items_frame is None:
        pending_row_preview_buttons = []
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    pending_row_preview_buttons = []

    for child in pending_items_frame.winfo_children():
        child.destroy()

    folder = normalize_path(pending_folder.get())
    if not folder:
        pending_file_vars = {}
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    try:
        files = sorted(file for file in os.listdir(folder) if file.lower().endswith(".pdf"))
    except OSError as exc:
        messagebox.showerror("Error", f"Unable to load pending PDFs: {exc}")
        pending_file_vars = {}
        _refresh_pending_master_toggle_state()
        _set_pending_snapshot()
        return

    previous_state = {name: var.get() for name, var in pending_file_vars.items()}
    pending_file_vars = {}

    if not files:
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

        display_name = _format_pending_filename_for_display(filename)
        name_label = ttk.Label(row, text=display_name, style="PendingFile.TLabel", anchor="w")
        name_label.pack(side="left", fill="x", expand=True, padx=(0, 8))

        if display_name != filename:
            _attach_hover_tooltip(name_label, filename)

        def _toggle_row_selection(_event=None, target_var=var):
            target_var.set(not target_var.get())
            return "break"

        def _on_row_enter(_event=None, target_row=row, target_label=name_label, target_check=checkbutton):
            _set_pending_row_hover_state(target_row, target_label, target_check, True)

        def _on_row_leave(_event=None, target_row=row, target_label=name_label, target_check=checkbutton):
            _set_pending_row_hover_state(target_row, target_label, target_check, False)

        name_label.bind("<Button-1>", _toggle_row_selection, add="+")

        for hover_widget in (row, name_label, checkbutton):
            hover_widget.bind("<Enter>", _on_row_enter, add="+")
            hover_widget.bind("<Leave>", _on_row_leave, add="+")

        _set_pending_row_hover_state(row, name_label, checkbutton, False)

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

    _refresh_pending_master_toggle_state()

def preview_selected_pdf():
    selected = get_selected_pending_files()
    if not selected:
        messagebox.showwarning("Warning", "Select at least one pending PDF using the checkboxes.")
        return

    for filename in selected:
        preview_specific_pending_pdf(filename)


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
        440,
        560,
        min_width=400,
        min_height=480,
        margin_x=DEFAULT_MARGIN_X,
        margin_y=DEFAULT_MARGIN_Y,
    )
    win.transient(root)
    win.grab_set()
    win.focus_force()
    apply_theme(win)
    scroll_container, scroll_frame = create_scrollable_panel(win)
    scroll_container.pack(fill="both", expand=True)
    content = ttk.Frame(scroll_frame, padding=20, style="TFrame")
    content.pack(fill="both", expand=True)

    processed_successfully = False
    completion_reported = False

    def notify_completion():
        nonlocal completion_reported
        if on_complete and not completion_reported:
            on_complete(processed_successfully, filename)
            completion_reported = True

    # Variables
    name_var = tk.StringVar()
    letter_var = tk.StringVar()
    status_var = tk.StringVar(value="Active")
    new_year_var = tk.StringVar()
    old_year_var = tk.StringVar()
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

    name_var.trace_add("write", update_letter)
    letter_var.trace_add("write", refresh_destination_path)
    status_var.trace_add("write", refresh_destination_path)
    root_trace_id = root_folder.trace_add("write", refresh_destination_path)

    header = ttk.Frame(content)
    header.pack(fill="x", pady=(0, 12))
    ttk.Label(header, text=f"Pending file: {filename}", style="Subheading.TLabel").pack(anchor="w")
    ttk.Button(
        header,
        text="Preview Pending",
        command=lambda: preview_specific_pending_pdf(filename),
    ).pack(anchor="w", pady=(4, 0))
    if batch_context:
        ttk.Label(
            header,
            text=f"Batch {batch_context['current']} of {batch_context['total']}",
            style="Card.TLabel",
            padding=4,
        ).pack(anchor="w", pady=(4, 0))

    # UI
    ttk.Label(content, text="Name:").pack(anchor="w")
    name_field = ttk.Combobox(
        content,
        textvariable=name_var,
        values=employee_name_suggestions,
        state="normal",
    )

    def _handle_name_key(event=None):
        # Keep user input untouched and only refresh/open suggestion list.
        _update_combobox_suggestions(name_field, name_var.get(), event)

    name_field.bind("<KeyRelease>", _handle_name_key)

    def _refresh_name_choices():
        name_field["values"] = get_filtered_name_suggestions(name_var.get())

    name_field.configure(postcommand=_refresh_name_choices)
    name_field.pack(fill="x")

    ttk.Label(content, text="Surname First Letter:").pack(anchor="w", pady=(10, 0))
    ttk.Entry(content, textvariable=letter_var).pack(fill="x")

    ttk.Label(content, text="Status:").pack(anchor="w", pady=(10, 0))
    ttk.OptionMenu(content, status_var, status_var.get(), "Active", "Retiree").pack(fill="x")

    ttk.Label(content, text="Latest Year (most recent):").pack(anchor="w", pady=(10, 0))
    ttk.Entry(content, textvariable=new_year_var).pack(fill="x")

    ttk.Label(content, text="Oldest Year (first record):").pack(anchor="w", pady=(10, 0))
    ttk.Entry(content, textvariable=old_year_var).pack(fill="x")

    ttk.Label(content, text="Destination Path Preview:").pack(anchor="w", pady=(14, 0))
    ttk.Label(
        content,
        textvariable=dest_path_var,
        wraplength=360,
        justify="left",
        style="Card.TLabel",
        anchor="w",
        padding=6,
    ).pack(fill="x")

    refresh_destination_path()

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
            messagebox.showerror("Error", "Select root folder first.")
            return

        pending_path = normalize_path(pending_folder.get())
        if not pending_path:
            messagebox.showerror("Error", "Select pending folder first.")
            return

        name = name_var.get().strip()
        letter_value = letter_var.get().strip().upper() or (name[0].upper() if name else "")
        letter = letter_value if letter_value else "#"
        status = status_var.get()
        new_year_str = new_year_var.get()
        old_year_str = old_year_var.get()

        if not (new_year_str and old_year_str):
            messagebox.showerror("Error", "Both year fields are required.")
            return
        if not (new_year_str.isdigit() and old_year_str.isdigit()):
            messagebox.showerror("Error", "Year fields must be numeric.")
            return
        latest_year = max(int(new_year_str), int(old_year_str))
        earliest_year = min(int(new_year_str), int(old_year_str))

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
        new_filename = f"{name}_{latest_year}_{earliest_year}.pdf"
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

    ttk.Button(content, text="Save", command=save_record, style="Accent.TButton").pack(pady=20)
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
        560,
        820,
        min_width=520,
        min_height=640,
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
    content = ttk.Frame(scroll_frame, padding=20, style="TFrame")
    content.pack(fill="both", expand=True)

    processed_successfully = False
    completion_reported = False

    def notify_completion():
        nonlocal completion_reported
        if on_complete and not completion_reported:
            on_complete(processed_successfully, pending_filename)
            completion_reported = True

    def keep_merge_window_on_top():
        try:
            win.lift()
            win.attributes("-topmost", True)
            win.after(50, lambda: win.attributes("-topmost", False))
            win.focus_force()
        except tk.TclError:
            pass

    header = ttk.Frame(content)
    header.pack(fill="x", pady=(0, 12))
    ttk.Label(header, text=f"Pending file: {pending_filename}", style="Subheading.TLabel").pack(anchor="w")
    ttk.Button(
        header,
        text="Preview Pending",
        command=lambda: preview_specific_pending_pdf(pending_filename),
    ).pack(anchor="w", pady=(4, 0))
    if batch_context:
        ttk.Label(
            header,
            text=f"Batch {batch_context['current']} of {batch_context['total']}",
            style="Card.TLabel",
            padding=4,
        ).pack(anchor="w", pady=(4, 0))

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
    dest_preview_var = tk.StringVar(value="Select destination folder and fill out fields to preview the final file path.")
    merge_summary_var = tk.StringVar(value="Select an existing PDF to view page counts.")
    keep_backup_var = keep_backup_preference_var
    folder_path_suggestions = []
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

        if not folder:
            dest_preview_var.set("Select destination folder to preview final path.")
            update_merge_summary()
            return
        if not name:
            dest_preview_var.set("Enter the employee name to preview final path.")
            update_merge_summary()
            return
        if not (latest and earliest):
            dest_preview_var.set("Enter both year values to preview final path.")
            update_merge_summary()
            return
        if not (latest.isdigit() and earliest.isdigit()):
            dest_preview_var.set("Year fields must be numeric to preview final path.")
            update_merge_summary()
            return

        latest_val = max(int(latest), int(earliest))
        earliest_val = min(int(latest), int(earliest))
        filename = f"{name}_{latest_val}_{earliest_val}.pdf"
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
        nonlocal folder_path_suggestions
        scanned_paths = _scan_employee_folder_paths()
        folder_path_suggestions = list(dict.fromkeys(scanned_paths))

    def _normalize_folder_search_value(value):
        return " ".join(
            (value or "").lower().replace("\\", " ").replace("/", " ").split()
        )

    def _get_filtered_folder_suggestions(query):
        if not folder_path_suggestions:
            return []

        normalized_query = _normalize_folder_search_value(query)
        if not normalized_query:
            return folder_path_suggestions

        prefix_matches = []
        contains_matches = []

        for candidate_path in folder_path_suggestions:
            employee_name = os.path.basename(candidate_path)
            try:
                relative_path = os.path.relpath(candidate_path, root_path)
            except ValueError:
                relative_path = candidate_path

            searchable = _normalize_folder_search_value(
                f"{employee_name} {relative_path} {candidate_path}"
            )
            if searchable.startswith(normalized_query):
                prefix_matches.append(candidate_path)
            elif normalized_query in searchable:
                contains_matches.append(candidate_path)

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
        candidate = normalize_path((raw_value or "").strip())
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

        if existing_files_frame is None:
            update_merge_summary()
            return

        for child in existing_files_frame.winfo_children():
            child.destroy()

        folder = normalize_path(folder_var.get())
        if not folder:
            update_merge_summary()
            return
        try:
            files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        except OSError as exc:
            messagebox.showerror("Error", f"Unable to list PDFs: {exc}")
            update_merge_summary()
            return

        if not files:
            existing_selection_var.set("No PDFs in this folder")
            existing_selected_pdf_var.set("")
            update_merge_summary()
            return

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

            preview_button.bind("<Enter>", _on_existing_row_enter, add="+")
            preview_button.bind("<Leave>", _on_existing_row_leave, add="+")

            _set_pending_row_hover_state(row, name_label, checkbutton, False)

        if existing_selected_pdf_var.get().strip():
            on_existing_select()
        else:
            update_merge_summary()

    def on_existing_select(event=None):
        filename = existing_selected_pdf_var.get().strip()
        if not filename:
            existing_selection_var.set("No file selected")
            update_merge_summary()
            return

        existing_selection_var.set(filename)

        metadata = parse_filename_metadata(filename)
        if metadata:
            name_var.set(metadata["name"])
            letter_var.set(metadata["name"][0].upper())
            new_year_var.set(metadata["latest"])
            old_year_var.set(metadata["earliest"])
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

    def validate_years():
        latest = new_year_var.get().strip()
        earliest = old_year_var.get().strip()
        if not (latest and earliest):
            messagebox.showerror("Error", "Both year fields are required.")
            return None
        if not (latest.isdigit() and earliest.isdigit()):
            messagebox.showerror("Error", "Year fields must be numeric.")
            return None
        return max(int(latest), int(earliest)), min(int(latest), int(earliest))

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
        final_filename = f"{employee_name}_{latest_year}_{earliest_year}.pdf"
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
    folder_frame = ttk.Frame(content, padding=(0, 10))
    folder_frame.pack(fill="x")
    ttk.Label(folder_frame, text="Employee Folder:").pack(anchor="w")

    folder_field = ttk.Combobox(folder_frame, textvariable=folder_var, state="normal")
    folder_field.pack(side="left", expand=True, fill="x", padx=(0, 8))

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

    ttk.Button(folder_frame, text="Browse", command=select_existing_folder).pack(side="left")

    ttk.Label(content, text="Existing PDFs in folder:").pack(anchor="w")
    list_container = ttk.Frame(content, style="Card.TFrame", padding=8)
    list_container.pack(fill="both", expand=True, pady=(4, 4))

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

    ttk.Label(content, textvariable=existing_selection_var, style="Subheading.TLabel").pack(anchor="w", pady=(4, 10))

    button_frame = ttk.Frame(content)
    button_frame.pack(pady=5)
    ttk.Button(button_frame, text="Reload", command=load_existing_pdfs).grid(row=0, column=0, padx=5)

    def open_folder():
        folder = normalize_path(folder_var.get())
        if not folder:
            messagebox.showwarning("Warning", "Select a folder first.")
            return
        try:
            _launch_path(folder)
        except RuntimeError as exc:
            messagebox.showerror("Error", f"Unable to open folder: {exc}")

    ttk.Button(button_frame, text="Open Folder", command=open_folder).grid(row=0, column=1, padx=5)

    # Metadata fields
    form_frame = ttk.Frame(content, padding=(0, 10))
    form_frame.pack(fill="x")

    ttk.Label(form_frame, text="Name:").pack(anchor="w")
    merge_name_field = ttk.Combobox(
        form_frame,
        textvariable=name_var,
        values=employee_name_suggestions,
        state="normal",
    )

    def _handle_merge_name_key(event=None):
        # Keep user input untouched and only refresh/open suggestion list.
        _update_combobox_suggestions(merge_name_field, name_var.get(), event)

    merge_name_field.bind("<KeyRelease>", _handle_merge_name_key)

    def _refresh_merge_name_choices():
        merge_name_field["values"] = get_filtered_name_suggestions(name_var.get())

    merge_name_field.configure(postcommand=_refresh_merge_name_choices)
    merge_name_field.pack(fill="x")

    ttk.Label(form_frame, text="Surname First Letter:").pack(anchor="w", pady=(10, 0))
    ttk.Entry(form_frame, textvariable=letter_var).pack(fill="x")

    ttk.Label(form_frame, text="Status:").pack(anchor="w", pady=(10, 0))
    ttk.OptionMenu(form_frame, status_var, status_var.get(), "Active", "Retiree").pack(fill="x")

    ttk.Label(form_frame, text="Latest Year (most recent):").pack(anchor="w", pady=(10, 0))
    ttk.Entry(form_frame, textvariable=new_year_var).pack(fill="x")

    ttk.Label(form_frame, text="Oldest Year (first record):").pack(anchor="w", pady=(10, 0))
    ttk.Entry(form_frame, textvariable=old_year_var).pack(fill="x")

    ttk.Label(form_frame, text="Final File Preview:").pack(anchor="w", pady=(14,0))
    ttk.Label(
        form_frame,
        textvariable=dest_preview_var,
        wraplength=480,
        justify="left",
        style="Card.TLabel",
        anchor="w",
        padding=6,
    ).pack(fill="x")

    summary_frame = ttk.Frame(form_frame, padding=(0, 10))
    summary_frame.pack(fill="x")
    ttk.Label(summary_frame, text="Merge Summary:").pack(anchor="w")
    ttk.Label(
        summary_frame,
        textvariable=merge_summary_var,
        style="Subheading.TLabel",
        wraplength=480,
        justify="left",
    ).pack(fill="x", pady=(4, 4))

    action_frame = ttk.Frame(content)
    action_frame.pack(pady=15)
    ttk.Button(action_frame, text="Merge & Save", width=20, command=perform_merge, style="Accent.TButton").pack()
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
    processed_count = 0

    def launch_next(index=0):
        nonlocal processed_count
        if index >= total:
            if total > 1 and processed_count > 0:
                messagebox.showinfo(
                    "Batch Complete",
                    f"{action_label} batch finished. Processed {processed_count} of {total} pending PDFs.",
                )
            return

        current_file = files[index]

        def handle_close(_success, _filename):
            nonlocal processed_count
            if _success:
                processed_count += 1
            root.after(150, lambda: launch_next(index + 1))

        batch_context = {"current": index + 1, "total": total}
        if mode == "new":
            new_record_window(initial_filename=current_file, batch_context=batch_context, on_complete=handle_close)
        else:
            merge_existing_window(pending_filename=current_file, batch_context=batch_context, on_complete=handle_close)

    launch_next()


def start_new_record_batch():
    _start_batch_processing("new")


def start_merge_existing_batch():
    _start_batch_processing("merge")


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

main_container = ttk.Frame(scroll_frame, padding=24, style="TFrame")
main_container.pack(fill="both", expand=True)

ttk.Label(main_container, text="PDF Record Manager", style="Title.TLabel").pack(anchor="w")
ttk.Label(
    main_container,
    text="Review pending scans, create new folders, and merge updates without leaving this window.",
    style="Subheading.TLabel",
).pack(anchor="w", pady=(4, 12))

folders_frame = ttk.Frame(main_container, style="TFrame")
folders_frame.pack(fill="x", pady=(0, 20))

pending_card = ttk.Frame(folders_frame, style="Card.TFrame", padding=16)
pending_card.pack(side="left", fill="x", expand=True, padx=(0, 12))
ttk.Label(pending_card, text="Pending Folder", style="Card.TLabel").pack(anchor="w")
ttk.Entry(pending_card, textvariable=pending_folder).pack(fill="x", pady=(6, 4))
ttk.Button(pending_card, text="Browse", command=select_pending_folder).pack(anchor="e")

root_card = ttk.Frame(folders_frame, style="Card.TFrame", padding=16)
root_card.pack(side="left", fill="x", expand=True)
ttk.Label(root_card, text="Records Root Folder", style="Card.TLabel").pack(anchor="w")
ttk.Entry(root_card, textvariable=root_folder).pack(fill="x", pady=(6, 4))
ttk.Button(root_card, text="Browse", command=select_root_folder).pack(anchor="e")

names_card = ttk.Frame(main_container, style="Card.TFrame", padding=16)
names_card.pack(fill="x", pady=(0, 20))
ttk.Label(names_card, text="Employee Name Sources (PDF / Excel / CSV / TXT)", style="Card.TLabel").pack(anchor="w")

employee_sources_container = ttk.Frame(names_card, style="Card.TFrame")
employee_sources_container.pack(fill="x", pady=(6, 4))
employee_sources_scrollbar = ttk.Scrollbar(employee_sources_container, orient="vertical")
employee_sources_scrollbar.pack(side="right", fill="y")

employee_sources_listbox = tk.Listbox(
    employee_sources_container,
    height=4,
    yscrollcommand=employee_sources_scrollbar.set,
)
_apply_modern_listbox_style(employee_sources_listbox)
employee_sources_listbox.pack(side="left", fill="x", expand=True)
employee_sources_scrollbar.configure(command=employee_sources_listbox.yview)
_mark_widget_as_scroll_list(employee_sources_listbox)
_refresh_employee_sources_listbox()

ui_icon_images = _build_pending_toolbar_icon_images()

name_buttons = ttk.Frame(names_card, style="Card.TFrame")
name_buttons.pack(fill="x", pady=(0, 4))
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

ttk.Label(
    names_card,
    textvariable=employee_list_status_var,
    style="Subheading.TLabel",
    padding=(0, 4),
    anchor="w",
).pack(fill="x", pady=(4, 0))

list_card = ttk.Frame(main_container, style="Card.TFrame", padding=18)
list_card.pack(fill="both", expand=True)
header_row = ttk.Frame(list_card, style="Card.TFrame")
header_row.pack(fill="x")
ttk.Label(header_row, text="Pending Files", style="Card.TLabel").pack(side="left")

listbox_container = ttk.Frame(list_card, style="Card.TFrame")
listbox_container.pack(fill="both", expand=True, pady=(12, 0))

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

pending_refresh_button = ttk.Button(
    pending_master_actions,
    style="ToolbarIcon.TButton",
    command=load_pending_files,
)
pending_refresh_button.pack(side="left", padx=4)
_attach_hover_tooltip(pending_refresh_button, "Refresh pending files")

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
_mark_widget_as_scroll_canvas(pending_canvas)
_ensure_global_mousewheel_binding()

pending_items_frame = ttk.Frame(pending_canvas, style="Card.TFrame")
pending_canvas_window = pending_canvas.create_window((0, 0), window=pending_items_frame, anchor="nw")

def _resize_pending_canvas(event):
    pending_canvas.itemconfigure(pending_canvas_window, width=event.width)


def _update_pending_scrollregion(_event=None):
    pending_canvas.configure(scrollregion=pending_canvas.bbox("all"))


pending_canvas.bind("<Configure>", _resize_pending_canvas)
pending_items_frame.bind("<Configure>", _update_pending_scrollregion)

btn_frame = ttk.Frame(main_container)
btn_frame.pack(pady=20)
ttk.Button(btn_frame, text="New Record", width=20, command=start_new_record_batch, style="Accent.TButton").grid(row=0, column=0, padx=8)
ttk.Button(btn_frame, text="Merge Existing", width=20, command=start_merge_existing_batch).grid(row=0, column=1, padx=8)

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