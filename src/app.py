import threading
import json
import multiprocessing
import importlib
import inspect
import queue as _queue
from collections import OrderedDict
from typing import TYPE_CHECKING
import customtkinter as ctk
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog
import datetime
import os
import ctypes
from pathlib import Path
import sys
import time
import traceback
import subprocess
import re

from PIL import Image, ImageDraw, ImageTk

try:
    import winsound  # Windows only: system sounds for the completion alert
except ImportError:
    winsound = None

if TYPE_CHECKING:  # never runs; lets PyInstaller discover the cleaner modules that the worker imports by name
    import cadence_cleaner, brand_performance_cleaner, existing_db_cleaner, current_sales_cleaner, database_cleaner  # noqa: F401

# --- Fix Taskbar Icon Grouping on Windows ---
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("StrategyDivision.StratDataCleaner.3.8")
except Exception:
    pass

# --- Setup CustomTkinter Theme ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ======================================================================
# Palette: Midnight Enterprise (now with gradient accents)
# ======================================================================
NAVY_BLUE = "#0A1020"
CARD_BG = "#101A2E"
SURFACE_BG = "#1B2740"
BORDER_COLOR = "#24365A"
WHITE = "#F1F5F9"
SECTION_TEXT = "#9DB7E8"

CLOSE_HOVER = "#7F1D1D"
ACCENT_ORANGE = "#2563EB"      # primary accent (name kept for compatibility)
HOVER_ORANGE = "#3B82F6"
ACCENT_GRADIENT_END = "#7C3AED"

# Header banner gradient (left -> right)
HEADER_STOPS = ["#14275E", "#2B2A7D", "#6A2A86", "#B03A66"]
HEADER_BORDER = "#3A4A94"
HEADER_SUBTITLE = "#E2E8FF"
RETAIL_AMBER = "#F59E0B"
VIOLET_ACCENT = "#A78BFA"

# Each cleaner keeps its own identity through an accent border + gradient tint.
CADENCE_COLOR = "#172A46"
CADENCE_HOVER = "#203B60"
CADENCE_ACCENT = "#3B82F6"
DB_COLOR = "#15363A"
DB_HOVER = "#1B494C"
DB_ACCENT = "#14B8A6"
DB_TEXT_COLOR = WHITE
BRAND_COLOR = "#302044"
BRAND_HOVER = "#432D5E"
BRAND_ACCENT = "#A78BFA"
EXISTING_DB_COLOR = "#3B2B24"
EXISTING_DB_HOVER = "#52392D"
EXISTING_DB_ACCENT = "#F59E0B"
CURRENT_SALES_COLOR = "#3A2030"
CURRENT_SALES_HOVER = "#4F2B42"
CURRENT_SALES_ACCENT = "#F472B6"

CANCEL_COLOR = "#7F1D1D"
CANCEL_HOVER = "#991B1B"

OPEN_BTN_COLOR = "#0F766E"
OPEN_BTN_HOVER = "#14958A"
OPEN_BTN_GRADIENT_END = "#0891B2"
OPEN_BTN_DISABLED = "#16213A"
DISABLED_BORDER = "#1F2D47"
DISABLED_TEXT = "#8091AD"

MUTED_TEXT = "#94A3B8"
LOG_BG = "#060C18"
LOG_TEXT = "#CBD5E1"
SUCCESS_COLOR = "#34D399"
WARNING_COLOR = "#FBBF24"
ERROR_COLOR = "#F87171"
INFO_COLOR = "#60A5FA"

# Operation-card border follows the current status (a quiet highlight).
STATE_BORDERS = {
    INFO_COLOR: "#2B5FA8",
    SUCCESS_COLOR: "#1F7A5C",
    WARNING_COLOR: "#8A6A12",
    ERROR_COLOR: "#8F3A3A",
}
FILE_LOADED_BORDER = "#1F7A5C"

# Step badges (1 Select -> 2 Choose -> 3 Review)
STEP_TODO_BG = "#1B2740"
STEP_TODO_TEXT = "#8091AD"
STEP_ACTIVE_BG = "#2563EB"
STEP_DONE_BG = "#0F5F49"
STEP_DONE_TITLE = "#8FD9B8"

# --- Layout constants (unscaled units) ---
MIN_WIDTH = 520
DUAL_BREAKPOINT = 980         # at/above this width: controls on the left, activity log on the right
MIN_LEFT = 520                # minimum width of the controls column in two-pane mode
MIN_RIGHT = 320               # minimum width of the log column in two-pane mode
PANE_GAP = 12
WIDE_COLS_BREAKPOINT = 860    # controls column wide enough for 4 cleaner buttons across
FOOTER_STACK_INNER = 800      # below this inner width the footer wraps to two rows
MAIN_PADX = 25
MAIN_PADY = 18
MAIN_PADY_COMPACT = 10
MAX_LOG_LINES = 1500
LOG_MIN_HEIGHT = 96
LOG_MIN_HEIGHT_COMPACT = 84
LOG_ROW_MIN = 70              # the log row never collapses below this in single-column mode
SCREEN_RESERVE_PX = 100       # taskbar + window title bar, in real pixels (fallback when the monitor is unknown)
TITLEBAR_RESERVE_PX = 56      # title bar + borders, in real pixels (the monitor work area already excludes the taskbar)
APP_VERSION = "v3.12.22"
CURRENT_YEAR = datetime.datetime.now().year
BASE_TITLE = "Strat Data Cleaner"
CONFIG_PATH = Path.home() / ".strat_data_cleaner.json"

SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
SPINNER_INTERVAL_MS = 100     # ~10 spinner frames per second (was 1 per second)
LAYOUT_DEBOUNCE_MS = 40       # normal resize debounce
SCALE_SETTLE_MS = 220         # extra quiet time after the window lands on a monitor with a different DPI


# ======================================================================
# Scaling helpers
# ======================================================================
try:  # imported once (this is called on every redraw and every layout pass)
    from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker as _ScalingTracker
except Exception:
    _ScalingTracker = None


def _monitor_work_height_px():
    """Usable height (real pixels, taskbar excluded) of the monitor the window is on right now.
    Tk's winfo_screenheight() only describes the primary monitor, so it is wrong after the window is
    moved to another one. Returns None when unknown (non-Windows or any failure)."""
    if sys.platform != "win32":
        return None
    try:
        from ctypes import wintypes

        class _MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]

        user32 = ctypes.windll.user32
        user32.MonitorFromWindow.restype = ctypes.c_void_p
        user32.MonitorFromWindow.argtypes = [ctypes.c_void_p, wintypes.DWORD]
        user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.POINTER(_MONITORINFO)]
        hwnd = user32.GetParent(app.winfo_id()) or app.winfo_id()
        handle = user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
        info = _MONITORINFO()
        info.cbSize = ctypes.sizeof(_MONITORINFO)
        if handle and user32.GetMonitorInfoW(handle, ctypes.byref(info)):
            return int(info.rcWork.bottom - info.rcWork.top)
    except Exception:
        pass
    return None


def widget_scale(widget):
    """Current DPI / CustomTkinter scaling factor for a widget."""
    if _ScalingTracker is not None:
        try:
            return _ScalingTracker.get_widget_scaling(widget)
        except Exception:
            pass
    try:
        return widget.winfo_fpixels("1i") / 96.0
    except Exception:
        return 1.0


def px(value, widget=None):
    """Scale an unscaled value to real pixels (for plain tk widgets like Canvas)."""
    return int(round(value * widget_scale(widget if widget is not None else app)))


# ======================================================================
# Gradient rendering (Pillow) - cached for speed
# ======================================================================
_IMG_CACHE = OrderedDict()   # PhotoImages, least-recently-used eviction (never a "clear everything" stall)
_IMG_CACHE_MAX = 400
_MASK_CACHE = {}             # rounded-corner masks, shared by every widget of the same size
_STRIP_CACHE = {}            # 256-sample gradient strips, one per colour set
_SHEEN_CACHE = {}
_SS = 3  # supersampling factor for smooth rounded edges


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _mix(c1, c2, t):
    a, b = _hex_to_rgb(c1), _hex_to_rgb(c2)
    return "#%02x%02x%02x" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _cache_get(key):
    img = _IMG_CACHE.get(key)
    if img is not None:
        _IMG_CACHE.move_to_end(key)
    return img


def _cache_put(key, img):
    _IMG_CACHE[key] = img
    while len(_IMG_CACHE) > _IMG_CACHE_MAX:
        _IMG_CACHE.popitem(last=False)


def _rounded_mask(w, h, radius, inset=0):
    """Anti-aliased rounded rectangle mask. Cached: the four cleaner buttons (and every hover/pressed
    state of them) have identical geometry, so they all share one mask instead of rebuilding it."""
    key = (w, h, radius, inset)
    mask = _MASK_CACHE.get(key)
    if mask is not None:
        return mask
    big = Image.new("L", (w * _SS, h * _SS), 0)
    d = ImageDraw.Draw(big)
    box = [inset * _SS, inset * _SS, (w - inset) * _SS - 1, (h - inset) * _SS - 1]
    r = max(radius - inset, 0) * _SS
    try:
        d.rounded_rectangle(box, radius=r, fill=255)
    except AttributeError:  # very old Pillow
        d.rectangle(box, fill=255)
    try:
        mask = big.reduce(_SS)                      # box filter: exact area average, several times faster than LANCZOS
    except AttributeError:                          # Pillow < 7
        mask = big.resize((w, h), Image.BOX)
    if len(_MASK_CACHE) > 96:
        _MASK_CACHE.clear()
    _MASK_CACHE[key] = mask
    return mask


def _gradient_strip(stops):
    """256-sample horizontal gradient. Width-independent, so it is built once per colour set and just stretched."""
    key = tuple(stops)
    strip = _STRIP_CACHE.get(key)
    if strip is None:
        rgb = [_hex_to_rgb(c) for c in stops]
        segs = len(rgb) - 1
        row = []
        for x in range(256):
            t = x / 255.0
            if segs == 0:
                row.append(rgb[0])
                continue
            i = min(int(t * segs), segs - 1)
            lt = t * segs - i
            a, b = rgb[i], rgb[i + 1]
            row.append(tuple(int(a[k] + (b[k] - a[k]) * lt) for k in range(3)))
        strip = Image.new("RGB", (256, 1))
        strip.putdata(row)
        _STRIP_CACHE[key] = strip
    return strip


def _sheen_ramp(sheen):
    ramp = _SHEEN_CACHE.get(sheen)
    if ramp is None:
        ramp = Image.linear_gradient("L").point(lambda v: int(max(0, 255 - 2 * v) * sheen))
        _SHEEN_CACHE[sheen] = ramp
    return ramp


def render_gradient(w, h, stops, radius, bg, border=None, border_w=1, sheen=0.0):
    """Horizontal multi-stop gradient with rounded corners, optional border and top sheen."""
    w, h = max(int(w), 2), max(int(h), 2)
    key = (w, h, tuple(stops), radius, bg, border, border_w, round(sheen, 3))
    cached = _cache_get(key)
    if cached is not None:
        return cached

    grad = _gradient_strip(stops).resize((w, h), Image.BILINEAR)
    if sheen > 0:
        grad.paste((255, 255, 255), (0, 0), _sheen_ramp(round(sheen, 3)).resize((w, h), Image.BILINEAR))

    img = Image.new("RGB", (w, h), _hex_to_rgb(bg))
    if border:
        img.paste(_hex_to_rgb(border), (0, 0), _rounded_mask(w, h, radius, 0))
        img.paste(grad, (0, 0), _rounded_mask(w, h, radius, border_w))
    else:
        img.paste(grad, (0, 0), _rounded_mask(w, h, radius, 0))

    photo = ImageTk.PhotoImage(img)
    _cache_put(key, photo)
    return photo


# ======================================================================
# Custom widgets
# ======================================================================
class GradientButton(tk.Canvas):
    """
    A gradient, rounded button that mirrors the CTkButton API used by this app
    (configure with text / state / fg_color / hover_color / border_color / command).
    Dimensions are given in unscaled units and follow DPI scaling.
    """

    _OWN_KEYS = ("text", "fg_color", "hover_color", "border_color", "text_color",
                 "gradient_to", "command", "state", "font", "corner_radius")

    def __init__(self, master, text="", font=("Segoe UI", 12, "bold"), fg_color=ACCENT_ORANGE,
                 hover_color=HOVER_ORANGE, border_color=None, text_color=WHITE, gradient_to=None,
                 corner_radius=10, width=100, height=44, command=None, state="normal",
                 bg=NAVY_BLUE, **_ignored):
        scale = widget_scale(master)
        super().__init__(master, width=int(width * scale), height=int(height * scale), bg=bg,
                         highlightthickness=0, bd=0, takefocus=1,
                         cursor="hand2" if state == "normal" else "arrow")
        self._gb_w_u = width      # logical (unscaled) size, kept so the button can follow a DPI change
        self._gb_h_u = height
        self._gb_text = text
        self._gb_font = font
        self._gb_fg = fg_color
        self._gb_hover_color = hover_color
        self._gb_border = border_color
        self._gb_text_color = text_color
        self._gb_gradient_to = gradient_to
        self._gb_radius = corner_radius
        self._gb_command = command
        self._gb_state = state
        self._gb_bg = bg
        self._gb_hover = False
        self._gb_pressed = False
        self._gb_focused = False
        self._gb_size = (0, 0)
        self._gb_job = None
        self._gb_img = None
        self._gb_txt = None
        self._gb_photo = None

        self.bind("<Configure>", self._on_configure)
        self.bind("<Enter>", lambda e: self._set_flag("_gb_hover", True))
        self.bind("<Leave>", lambda e: self._set_flag("_gb_hover", False, also_unpress=True))
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<FocusIn>", lambda e: self._set_flag("_gb_focused", True))
        self.bind("<FocusOut>", lambda e: self._set_flag("_gb_focused", False))
        self.bind("<Return>", lambda e: self._invoke())
        self.bind("<space>", lambda e: self._invoke())

    # ---- CTkButton-compatible API ----
    def configure(self, **kw):
        changed = False
        mapping = {"text": "_gb_text", "fg_color": "_gb_fg", "hover_color": "_gb_hover_color",
                   "border_color": "_gb_border", "text_color": "_gb_text_color",
                   "gradient_to": "_gb_gradient_to", "command": "_gb_command",
                   "font": "_gb_font", "corner_radius": "_gb_radius"}
        for key, attr in mapping.items():
            if key in kw:
                setattr(self, attr, kw.pop(key))
                changed = True
        if "state" in kw:
            self._gb_state = kw.pop("state")
            tk.Canvas.configure(self, cursor="hand2" if self._gb_state == "normal" else "arrow")
            changed = True
        if kw:
            tk.Canvas.configure(self, **kw)
        if changed:
            self._schedule(immediate=True)

    config = configure

    def set_height(self, height):
        """Sets the button height in unscaled units."""
        self._gb_h_u = height
        tk.Canvas.configure(self, height=int(height * widget_scale(self)))

    def set_text(self, text):
        """Fast path for animated labels (spinner): swaps the text item only, no image re-render."""
        if text == self._gb_text:
            return
        self._gb_text = text
        if self._gb_txt is not None:
            self.itemconfigure(self._gb_txt, text=text)
        elif self._gb_job is None:
            self._schedule(immediate=True)

    def rescale(self):
        """Re-applies the pixel size for the current DPI. Raw Tk canvases are not rescaled by CustomTkinter,
        so this is called when the window lands on a monitor with a different scaling factor."""
        s = widget_scale(self)
        tk.Canvas.configure(self, width=int(self._gb_w_u * s), height=int(self._gb_h_u * s))
        self._schedule(immediate=False)

    def cget(self, key):
        mapping = {"text": self._gb_text, "state": self._gb_state, "fg_color": self._gb_fg,
                   "hover_color": self._gb_hover_color, "border_color": self._gb_border}
        if key in mapping:
            return mapping[key]
        return tk.Canvas.cget(self, key)

    # ---- internals ----
    def _style(self):
        if self._gb_state == "disabled":
            return {"stops": [OPEN_BTN_DISABLED, OPEN_BTN_DISABLED], "border": DISABLED_BORDER,
                    "text": DISABLED_TEXT, "sheen": 0.0}
        active = self._gb_hover or self._gb_pressed
        base = self._gb_hover_color if active else self._gb_fg
        flat = self._gb_fg == self._gb_hover_color
        if flat:
            stops = [base, base]
        else:
            if self._gb_gradient_to:
                end = _mix(self._gb_gradient_to, "#FFFFFF", 0.12) if active else self._gb_gradient_to
            elif self._gb_border:
                end = _mix(base, self._gb_border, 0.55 if active else 0.40)
            else:
                end = _mix(base, "#FFFFFF", 0.15)
            stops = [base, end]
        if self._gb_pressed:
            stops = [_mix(c, "#000000", 0.2) for c in stops]
        border = self._gb_border
        if self._gb_focused:
            border = _mix(border or "#FFFFFF", "#FFFFFF", 0.6)
        return {"stops": stops, "border": border, "text": self._gb_text_color,
                "sheen": 0.0 if flat else 0.07}

    def _scaled_font(self):
        family, size = self._gb_font[0], self._gb_font[1]
        weight = self._gb_font[2] if len(self._gb_font) > 2 else "normal"
        if size > 0:
            size = -int(round(size * widget_scale(self)))
        return (family, size, weight)

    def _set_flag(self, name, value, also_unpress=False):
        setattr(self, name, value)
        if also_unpress:
            self._gb_pressed = False
        self._schedule(immediate=True)

    def _on_press(self, _e):
        if self._gb_state != "normal":
            return
        self.focus_set()
        self._gb_pressed = True
        self._schedule(immediate=True)

    def _on_release(self, e):
        was_pressed = self._gb_pressed
        self._gb_pressed = False
        self._schedule(immediate=True)
        if was_pressed and 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height():
            self._invoke()

    def _invoke(self):
        if self._gb_state == "normal" and self._gb_command:
            self._gb_command()

    def _on_configure(self, e):
        size = (e.width, e.height)
        if size == self._gb_size:
            return
        self._gb_size = size
        if self._gb_txt is not None:
            self.coords(self._gb_txt, e.width / 2, e.height / 2)
        self._schedule(immediate=False)

    def _schedule(self, immediate):
        if self._gb_job is not None:
            try:
                self.after_cancel(self._gb_job)
            except Exception:
                pass
        # Resizing is debounced; state/text changes are coalesced into one idle redraw.
        self._gb_job = self.after_idle(self._redraw) if immediate else self.after(25, self._redraw)

    def _redraw(self):
        self._gb_job = None
        w, h = self.winfo_width(), self.winfo_height()
        if w < 8 or h < 8:
            return
        st = self._style()
        scale = widget_scale(self)
        radius = min(int(round(self._gb_radius * scale)), h // 2)
        photo = render_gradient(w, h, st["stops"], radius, self._gb_bg, st["border"],
                                max(1, int(round(scale))), st["sheen"])
        self._gb_photo = photo
        if self._gb_img is None:
            self._gb_img = self.create_image(0, 0, image=photo, anchor="nw")
            self._gb_txt = self.create_text(w / 2, h / 2, text=self._gb_text, fill=st["text"],
                                            font=self._scaled_font())
        else:
            self.itemconfigure(self._gb_img, image=photo)
            self.itemconfigure(self._gb_txt, text=self._gb_text, fill=st["text"], font=self._scaled_font())
            self.coords(self._gb_txt, w / 2, h / 2)


def _tag_icon(size):
    """White retail price-tag glyph (supersampled for smooth edges), cached."""
    key = ("tag", size)
    cached = _IMG_CACHE.get(key)
    if cached is not None:
        return cached
    n = size * 4
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pts = [(0.08, 0.50), (0.36, 0.14), (0.94, 0.14), (0.94, 0.86), (0.36, 0.86)]
    d.polygon([(x * n, y * n) for x, y in pts], fill=(255, 255, 255, 235))
    r = 0.075 * n
    cx, cy = 0.34 * n, 0.50 * n
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(0, 0, 0, 0))          # string hole
    d.rectangle([0.54 * n, 0.38 * n, 0.82 * n, 0.45 * n], fill=(0, 0, 0, 0))  # "price" lines
    d.rectangle([0.54 * n, 0.55 * n, 0.74 * n, 0.62 * n], fill=(0, 0, 0, 0))
    photo = ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))
    _IMG_CACHE[key] = photo
    return photo


_LOGO_SRC = None  # decoded app logo (PIL image); False once loading has failed so we only try once


def _logo_icon(size):
    """The app logo (assets/strat_data_cleaner.ico) scaled to `size` px, cached.
    Falls back to the price-tag glyph if the file is missing or unreadable."""
    global _LOGO_SRC
    key = ("logo", size)
    cached = _IMG_CACHE.get(key)
    if cached is not None:
        return cached
    if _LOGO_SRC is None:
        try:
            ico = Image.open(resource_path(os.path.join("assets", "strat_data_cleaner.ico")))
            try:  # .ico files hold several sizes: start from the largest for the sharpest downscale
                best = max(ico.ico.sizes())
                ico.size = best
            except Exception:
                pass
            _LOGO_SRC = ico.convert("RGBA")
        except Exception:
            _LOGO_SRC = False
    if not _LOGO_SRC:
        return _tag_icon(size)
    photo = ImageTk.PhotoImage(_LOGO_SRC.resize((size, size), Image.LANCZOS))
    _IMG_CACHE[key] = photo
    return photo


class HeaderBanner(tk.Canvas):
    """Retail gradient banner with the app logo on the left. Title and subtitle sit side by side, or stack when space is tight."""

    def __init__(self, master, title, subtitle, stops, bg=NAVY_BLUE, on_height_change=None):
        scale = widget_scale(master)
        super().__init__(master, width=10, height=int(72 * scale), bg=bg, highlightthickness=0, bd=0)
        self._hb_title = title
        self._hb_subtitle = subtitle
        self._hb_stops = stops
        self._hb_bg = bg
        self._hb_cb = on_height_change
        self._hb_heights = (72, 98)  # (side-by-side, stacked) in unscaled units
        self._hb_size = (0, 0)
        self._hb_job = None
        self._hb_photo = None
        self._hb_icon = None
        self._hb_items = None
        self._hb_font_key = None
        self._hb_fonts = None
        self.bind("<Configure>", self._on_configure)

    def set_heights(self, side, stacked):
        self._hb_heights = (side, stacked)
        self._hb_size = (0, 0)
        self._redraw()

    def refresh(self):
        """Forces a redraw at the current DPI (used after the window moves to another monitor)."""
        self._hb_size = (0, 0)
        if self._hb_job is not None:
            try:
                self.after_cancel(self._hb_job)
            except Exception:
                pass
        self._hb_job = self.after(25, self._redraw)

    def _fonts_for(self, scale):
        """Title/subtitle fonts are built once per scale factor instead of on every redraw."""
        key = round(scale, 3)
        changed = key != self._hb_font_key
        if changed:
            self._hb_font_key = key
            self._hb_fonts = (
                tkfont.Font(root=self, family="Segoe UI", size=-int(round(25 * scale)), weight="bold"),
                tkfont.Font(root=self, family="Segoe UI", size=-int(round(10 * scale)), weight="bold"),
            )
        return self._hb_fonts, changed

    def _on_configure(self, e):
        if (e.width, e.height) == self._hb_size:
            return
        self._hb_size = (e.width, e.height)
        if self._hb_job is not None:
            self.after_cancel(self._hb_job)
        self._hb_job = self.after(25, self._redraw)

    def _redraw(self):
        self._hb_job = None
        w = self.winfo_width()
        if w < 40:
            return
        s = widget_scale(self)
        (title_font, sub_font), fonts_changed = self._fonts_for(s)
        gap = int(round(22 * s))
        icon_px = int(round(40 * s))
        text_x = gap + icon_px + int(round(12 * s))
        side_by_side = (text_x + title_font.measure(self._hb_title)
                        + sub_font.measure(self._hb_subtitle) + gap * 2) <= w
        h = int(round(self._hb_heights[0 if side_by_side else 1] * s))

        height_changed = int(self.cget("height")) != h
        if height_changed:
            tk.Canvas.configure(self, height=h)

        photo = render_gradient(w, h, self._hb_stops, int(round(14 * s)), self._hb_bg,
                                HEADER_BORDER, max(1, int(round(s))), 0.08)
        self._hb_photo = photo
        icon = _logo_icon(icon_px)
        self._hb_icon = icon  # keep a reference: the LRU cache may evict it, and Tk would then drop the image

        if side_by_side:
            t_pos, s_pos, s_anchor = (text_x, h / 2), (w - gap, h / 2 + 2 * s), "e"
        else:
            t_pos, s_pos, s_anchor = (text_x, h * 0.37), (text_x, h * 0.74), "w"

        if self._hb_items is None:
            img = self.create_image(0, 0, image=photo, anchor="nw")
            tag = self.create_image(gap, h / 2, image=icon, anchor="w")
            title = self.create_text(*t_pos, text=self._hb_title, anchor="w", fill=WHITE, font=title_font)
            sub = self.create_text(*s_pos, text=self._hb_subtitle, anchor=s_anchor, fill=HEADER_SUBTITLE, font=sub_font)
            self._hb_items = (img, tag, title, sub)
        else:
            img, tag, title, sub = self._hb_items
            self.itemconfigure(img, image=photo)
            self.itemconfigure(tag, image=icon)
            if fonts_changed:
                self.itemconfigure(title, font=title_font)
                self.itemconfigure(sub, font=sub_font)
            self.itemconfigure(sub, anchor=s_anchor)
            self.coords(tag, gap, h / 2)
            self.coords(title, *t_pos)
            self.coords(sub, *s_pos)

        if height_changed and self._hb_cb:
            self._hb_cb()



class SmoothProgressBar(tk.Canvas):
    """
    Drop-in replacement for CTkProgressBar (same slim rounded look) built for smoothness:
      * one canvas item is moved per frame, instead of redrawing the whole rounded widget
      * the position comes from the clock, not from a frame counter, so a slow frame (dragging the
        window, switching monitors) never slows the motion down - it simply catches up
      * nothing is drawn while the window is minimized
    Supports configure(mode=, progress_color=, fg_color=), start(), stop() and set(value).
    """
    _PERIOD = 1.25   # seconds for one sweep
    _FRAME_MS = 33   # ~30 fps (Windows timers tick at ~15.6 ms, so 33 ms is a stable cadence)

    def __init__(self, master, height=4, fg_color=SURFACE_BG, progress_color=INFO_COLOR,
                 bg=CARD_BG, mode="determinate"):
        super().__init__(master, width=10, height=max(2, int(round(height * widget_scale(master)))),
                         bg=bg, highlightthickness=0, bd=0)
        self._pb_h_u = height
        self._pb_track_color = fg_color
        self._pb_color = progress_color
        self._pb_mode = mode
        self._pb_value = 0.0
        self._pb_running = False
        self._pb_job = None
        self._pb_t0 = 0.0
        self._pb_size = (0, 0)
        self._pb_track = None
        self._pb_bar = None
        self.bind("<Configure>", self._on_configure)

    # ---- CTkProgressBar-compatible API ----
    def configure(self, **kw):
        if "mode" in kw:
            self._pb_mode = kw.pop("mode")
        if "progress_color" in kw:
            self._pb_color = kw.pop("progress_color")
        if "fg_color" in kw:
            self._pb_track_color = kw.pop("fg_color")
        if kw:
            tk.Canvas.configure(self, **kw)
        if self._pb_track is not None:
            self.itemconfigure(self._pb_track, fill=self._pb_track_color)
            self.itemconfigure(self._pb_bar, fill=self._pb_color)
        self._draw()

    config = configure

    def set(self, value):
        self._pb_value = max(0.0, min(1.0, float(value)))
        self._draw()

    def start(self):
        if self._pb_running:
            return
        self._pb_running = True
        self._pb_t0 = time.perf_counter()
        self._tick()

    def stop(self):
        self._pb_running = False
        if self._pb_job is not None:
            try:
                self.after_cancel(self._pb_job)
            except Exception:
                pass
            self._pb_job = None

    def rescale(self):
        """Re-applies the pixel height for the current DPI (see GradientButton.rescale)."""
        tk.Canvas.configure(self, height=max(2, int(round(self._pb_h_u * widget_scale(self)))))

    # ---- internals ----
    def _on_configure(self, e):
        if (e.width, e.height) == self._pb_size:
            return
        self._pb_size = (e.width, e.height)
        w, h = e.width, e.height
        r, y = h / 2.0, h / 2.0
        if self._pb_track is None:
            self._pb_track = self.create_line(r, y, w - r, y, width=h, capstyle="round", fill=self._pb_track_color)
            self._pb_bar = self.create_line(r, y, r, y, width=h, capstyle="round", fill=self._pb_color, state="hidden")
        else:
            self.coords(self._pb_track, r, y, w - r, y)
            self.itemconfigure(self._pb_track, width=h)
            self.itemconfigure(self._pb_bar, width=h)
        self._draw()

    def _tick(self):
        self._pb_job = None
        if not self._pb_running:
            return
        try:
            if self.winfo_viewable():
                self._draw()
            self._pb_job = self.after(self._FRAME_MS, self._tick)
        except tk.TclError:  # window is being destroyed
            self._pb_running = False

    def _draw(self):
        if self._pb_bar is None:
            return
        w, h = self._pb_size
        if w < 8 or h < 2:
            return
        r, y = h / 2.0, h / 2.0
        span = max(w - h, 1)                        # travel of the pill's centre line
        if self._pb_mode == "indeterminate":
            seg = max(span * 0.35, float(h))        # length of the moving segment
            phase = ((time.perf_counter() - self._pb_t0) / self._PERIOD) % 1.0
            head = phase * (span + seg)             # slides in from the left, out through the right
            a, b = max(head - seg, 0.0), min(head, float(span))
        else:
            a, b = 0.0, span * self._pb_value
        if b - a < 0.5:
            self.itemconfigure(self._pb_bar, state="hidden")
            return
        self.coords(self._pb_bar, r + a, y, r + b, y)
        self.itemconfigure(self._pb_bar, state="normal")


# ======================================================================
# ======================================================================
# Persistent settings (remembers the last folder you browsed)
# ======================================================================
def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass


# ======================================================================
# Activity-log tools: level filter + warning badge
# ======================================================================
LOG_FILTERS = {"All": None, "Warnings": "warning", "Errors": "error"}
_LOG_LEVELS = ("normal", "success", "warning", "notice", "error", "processing")
_log_state = {"filter": "All", "warn_idx": -1, "flash_job": None, "badge_shown": False}


def _apply_log_filter(mode=None):
    """Shows only the chosen level of log lines (the full log is kept, and Copy/Save still export all of it)."""
    if mode is not None:
        _log_state["filter"] = mode
    keep = LOG_FILTERS.get(_log_state["filter"])
    tb = status_box._textbox
    for level in _LOG_LEVELS:
        tb.tag_configure(level, elide=(keep is not None and level != keep))
    tb.see("end")


def _refresh_warn_badge():
    """Keeps the 'N warnings' badge in step with the warning lines currently in the log."""
    try:
        count = len(status_box._textbox.tag_ranges("warnline")) // 2
        if count:
            warn_badge.configure(text=f"⚠ {count} warning{'s' if count != 1 else ''}")
            if not _log_state["badge_shown"]:
                warn_badge.pack(side="left", padx=(8, 0))
                _log_state["badge_shown"] = True
        else:
            if _log_state["badge_shown"]:
                warn_badge.pack_forget()
                _log_state["badge_shown"] = False
            _log_state["warn_idx"] = -1
    except Exception:
        pass


def jump_to_warning(_event=None):
    """Scrolls to the newest warning, then to the next older one on every further click."""
    try:
        tb = status_box._textbox
        ranges = tb.tag_ranges("warnline")
        count = len(ranges) // 2
        if not count:
            return
        if _log_state["filter"] == "Errors":  # warnings are hidden under the Errors filter
            log_filter.set("All")
            _apply_log_filter("All")
        idx = _log_state["warn_idx"]
        idx = count - 1 if idx < 0 or idx >= count else (idx - 1) % count
        _log_state["warn_idx"] = idx
        start = str(ranges[2 * idx])
        tb.see(start)
        tb.tag_remove("flash", "1.0", "end")
        tb.tag_add("flash", start, f"{start} lineend +1c")
        if _log_state["flash_job"] is not None:
            app.after_cancel(_log_state["flash_job"])
        _log_state["flash_job"] = app.after(1400, lambda: tb.tag_remove("flash", "1.0", "end"))
    except Exception:
        pass


# ======================================================================
# Logging, status banner, window title, progress
# ======================================================================
def log_message(message, level=None):
    """Inserts a timestamped color-coded message into the inline status box safely."""
    global status_box
    try:
        if 'status_box' in globals() and status_box is not None:
            status_box.configure(state="normal")
            timestamp = datetime.datetime.now().strftime("%I:%M:%S %p")

            tag = "normal"
            if any(w in message for w in ["Error", "Missing", "KeyError"]):
                tag = "error"
            elif any(w in message for w in ["Success", "Opened", "loaded", "Saved", "consolidated", "copied"]):
                tag = "success"
            elif any(w in message for w in ["Warning", "cancelled", "cleared", "Tip:"]):
                tag = "warning"
            elif any(w in message for w in ["Processing", "Consolidating", "Running"]):
                tag = "processing"

            if level in _LOG_LEVELS:  # the caller knows best (e.g. every line of a traceback is an error)
                tag = level

            if tag == "warning" and "Warning" not in message:
                tag = "notice"  # same orange, but not counted as a warning (e.g. "log cleared", "cancelled")

            tb = status_box._textbox
            tags = (tag, "warnline") if tag == "warning" else (tag,)
            tb.insert("end", f"[{timestamp}] {message}\n", tags)

            # Keep the log bounded so long sessions never slow the UI down.
            extra = int(tb.index("end-1c").split(".")[0]) - MAX_LOG_LINES
            if extra > 0:
                tb.delete("1.0", f"{extra + 1}.0")

            status_box.see("end")
            status_box.configure(state="disabled")
            _refresh_warn_badge()
            return
    except Exception:
        pass

    print(f"[Startup Log] {message}")


_banner_border = None


def set_banner(text, color):
    """Updates the live status banner and tints the card border to match the state."""
    global _banner_border
    live_status_banner.configure(text=text, text_color=color)
    border = STATE_BORDERS.get(color, BORDER_COLOR)
    if border != _banner_border:
        _banner_border = border
        operation_card.configure(border_color=border)


def set_title(suffix=None):
    """Shows the current state in the window title (visible in the taskbar even when minimized)."""
    app.title(BASE_TITLE if not suffix else f"{BASE_TITLE}  —  {suffix}")


def set_progress(state):
    """idle | running | success | error"""
    try:
        if state == "running":
            progress.configure(mode="indeterminate", progress_color=INFO_COLOR)
            progress.start()
            return
        progress.stop()
        progress.configure(mode="determinate")
        if state == "success":
            progress.configure(progress_color=SUCCESS_COLOR)
            progress.set(1)
        elif state == "error":
            progress.configure(progress_color=ERROR_COLOR)
            progress.set(1)
        else:
            progress.set(0)
    except Exception:
        pass


# --- 3-step guide: 1 Select file -> 2 Choose cleaner -> 3 Review results ---
_STEP_STYLE = {
    "todo": (STEP_TODO_BG, STEP_TODO_TEXT, MUTED_TEXT),
    "active": (STEP_ACTIVE_BG, WHITE, WHITE),
    "done": (STEP_DONE_BG, SUCCESS_COLOR, STEP_DONE_TITLE),
}


def refresh_steps():
    """Derives the step states and the helper hint from the current app state."""
    has_file = bool(getattr(app, "selected_path", None))
    busy = getattr(app, "is_processing", False)
    finished = getattr(app, "run_done", False)

    if not has_file:
        states, hint, hint_color = ("active", "todo", "todo"), "Select a file first to unlock", MUTED_TEXT
    elif busy:
        states, hint, hint_color = ("done", "active", "todo"), "Running — press Esc to cancel", INFO_COLOR
    elif finished:
        states, hint, hint_color = ("done", "done", "active"), "Done — see results below", SUCCESS_COLOR
    else:
        states, hint, hint_color = ("done", "active", "todo"), "Pick the cleaner that fits your file", INFO_COLOR

    for (badge, title, number), state in zip(step_widgets, states):
        bg, fg, title_color = _STEP_STYLE[state]
        badge.configure(fg_color=bg, text_color=fg, text="✓" if state == "done" else str(number))
        title.configure(text_color=title_color)
    step2_hint.configure(text=hint, text_color=hint_color)


animation_job = None
timer_job = None
dot_count = 0
start_time = 0
_banner_secs = -1
_stage_text = ""


def set_stage(text):
    """Shows the cleaner's current stage in the banner and the activity log."""
    global _stage_text, _banner_secs
    _stage_text = text
    _banner_secs = -1  # refresh the banner on the next spinner frame
    log_message(f"Running: {text}")


def animate_process_state(button, base_text):
    """
    Spinner at ~10 fps. Each frame only swaps the button's text item (no image re-render, no layout),
    and the elapsed-time banner is refreshed only when the displayed second actually changes.
    """
    global animation_job, dot_count, _banner_secs
    animation_job = None
    if getattr(app, 'is_processing', False) and not getattr(app, 'cancel_requested', False):
        button.set_text(f"{SPINNER_FRAMES[dot_count % len(SPINNER_FRAMES)]} Processing...")

        elapsed = int(time.time() - start_time)
        if elapsed != _banner_secs:
            _banner_secs = elapsed
            m, s = divmod(elapsed, 60)
            time_str = f"{m} min {s} sec" if m > 0 else f"{s} sec"
            stage = f"  ·  {_stage_text}" if _stage_text else ""
            set_banner(f"Running {base_text}{stage}  ·  {time_str}", INFO_COLOR)

        dot_count += 1
        animation_job = app.after(SPINNER_INTERVAL_MS, lambda: animate_process_state(button, base_text))


# ======================================================================
# Background processing: a separate PROCESS (not a thread)
#   * own interpreter + own GIL  -> the UI never competes with heavy pandas/openpyxl work
#   * can be terminated instantly -> Cancel really stops the operation
#   * runs at lower priority      -> the UI stays smooth even when every core is busy
# ======================================================================
_ctx = multiprocessing.get_context("spawn")
_worker = {"proc": None, "queue": None, "job": None}


def _lower_priority():
    """Runs inside the worker: yield CPU to the UI."""
    try:
        if sys.platform == "win32":
            k32 = ctypes.windll.kernel32
            k32.GetCurrentProcess.restype = ctypes.c_void_p
            k32.SetPriorityClass.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
            k32.SetPriorityClass(k32.GetCurrentProcess(), 0x00004000)  # BELOW_NORMAL_PRIORITY_CLASS
        else:
            os.nice(5)
    except Exception:
        pass


def _worker_entry(module_name, func_name, file_path, result_q):
    """Entry point of the worker process. Imports the cleaner here, so the UI process never loads pandas."""
    try:
        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w")
        if sys.stderr is None:
            sys.stderr = open(os.devnull, "w")
        _lower_priority()
        cleaner_func = getattr(importlib.import_module(module_name), func_name)
        kwargs = {}
        try:  # cleaners that accept `progress` stream stage messages; older ones simply run as before
            if "progress" in inspect.signature(cleaner_func).parameters:
                kwargs["progress"] = lambda text: result_q.put(("progress", str(text)))
        except (TypeError, ValueError):
            pass
        result = cleaner_func(file_path, **kwargs)
        stats = result[2] if len(result) > 2 else None   # optional run summary
        result_q.put(("ok", result[0], result[1], stats))
    except BaseException:
        result_q.put(("err", traceback.format_exc()))


def _stop_worker():
    """Kills the worker process (if any) and clears bookkeeping. Safe to call at any time."""
    job = _worker.get("job")
    if job is not None:
        try:
            app.after_cancel(job)
        except Exception:
            pass
    proc, q = _worker.get("proc"), _worker.get("queue")
    _worker.update(proc=None, queue=None, job=None)
    if proc is not None:
        try:
            if proc.is_alive():
                proc.terminate()
        except Exception:
            pass
        # Reap in the background so the UI never waits on it.
        threading.Thread(target=lambda: proc.join(5), daemon=True).start()
    if q is not None:
        try:
            q.close()
        except Exception:
            pass


def poll_worker(button, base_text, default_color, hover_color):
    """Checks the worker every 120 ms from the UI thread (very cheap) and reports its result."""
    proc, q = _worker.get("proc"), _worker.get("queue")
    if proc is None or q is None or getattr(app, 'cancel_requested', False):
        return

    msg = None
    while msg is None:  # drain stage messages; stop at the final result
        try:
            item = q.get_nowait()
        except _queue.Empty:
            break
        if item[0] == "progress":
            set_stage(item[1])
        else:
            msg = item
    if msg is None and not proc.is_alive():
        try:
            while msg is None:  # the result may have landed just before the process exited
                item = q.get(timeout=0.3)
                if item[0] == "progress":
                    set_stage(item[1])
                else:
                    msg = item
        except _queue.Empty:
            msg = ("err", f"The worker process ended unexpectedly (exit code {proc.exitcode}).\n"
                          "This usually means the file was too large for the available memory.")

    if msg is None:
        _worker["job"] = app.after(120, lambda: poll_worker(button, base_text, default_color, hover_color))
        return

    elapsed = time.time() - start_time
    _stop_worker()
    if msg[0] == "ok":
        handle_success(msg[1], msg[2], elapsed, button, base_text, default_color, hover_color,
                       msg[3] if len(msg) > 3 else None)
    else:
        handle_error(msg[1], button, base_text, default_color, hover_color)


def set_cleaners_enabled(enabled, except_btn=None):
    """Enables/disables all cleaner buttons, restoring their own colors and labels when enabling."""
    for c in CLEANERS:
        b = c["btn"]
        if b is except_btn:
            continue
        if enabled:
            b.configure(state="normal", fg_color=c["color"], hover_color=c["hover"],
                        border_color=c["accent"], text=c["label"])
        else:
            b.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)


def _finish_common():
    """Shared 'operation ended' UI reset used by success, error and cancel."""
    global animation_job
    app.is_processing = False
    if animation_job:
        try:
            app.after_cancel(animation_job)
        except Exception:
            pass
        animation_job = None
    set_cleaners_enabled(True)
    upload_btn.configure(state="normal")
    cancel_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    if getattr(app, "selected_path", None):
        clear_btn.configure(state="normal")


def _clean_key(file_path, cleaner_label):
    """Identity of one clean: the file NAME (case-insensitive) plus the cleaner used."""
    return (Path(file_path).name.casefold(), cleaner_label)


def already_cleaned(file_path, cleaner_label):
    """True when this file name was already cleaned successfully with this cleaner in this session."""
    return _clean_key(file_path, cleaner_label) in getattr(app, "cleaned_log", set())


def warn_already_cleaned(cleaner_label):
    """Logs the duplicate-clean warning and mirrors it in the banner."""
    log_message(f"Warning: already cleaned this file with {cleaner_label}")
    set_banner(f"Already cleaned this file with {cleaner_label}", WARNING_COLOR)


def run_cleaner_process(cleaner_ref, button, base_text, default_color, hover_color, force=False):
    """Runs a cleaner ((module_name, function_name)) in a worker process with live updates and real cancellation."""
    global start_time, dot_count, _banner_secs, _stage_text
    if getattr(app, 'is_processing', False):
        return

    if not getattr(app, 'selected_path', None):
        log_message("Warning: Please upload a retail file before running the cleaner.")
        return

    if not force and already_cleaned(app.selected_path, base_text):
        warn_already_cleaned(base_text)
        return

    file_name = Path(app.selected_path).name
    app.current_run = (file_name, base_text)
    log_message(f"Processing file on '{file_name}' using {base_text}...")

    app.is_processing = True
    app.cancel_requested = False
    app.run_done = False
    start_time = time.time()

    upload_btn.configure(state="disabled")
    clear_btn.configure(state="disabled")
    cancel_btn.configure(state="normal", fg_color=CANCEL_COLOR, hover_color=CANCEL_HOVER)
    set_banner(f"Starting {base_text}...", INFO_COLOR)
    set_progress("running")
    set_title(f"Running {base_text}…")

    set_cleaners_enabled(False, except_btn=button)
    button.configure(fg_color=SURFACE_BG, hover_color=SURFACE_BG)
    refresh_steps()

    dot_count = 0
    _banner_secs = -1
    _stage_text = ""
    animate_process_state(button, base_text)

    try:
        result_q = _ctx.Queue()
        proc = _ctx.Process(target=_worker_entry, args=(cleaner_ref[0], cleaner_ref[1], app.selected_path, result_q), daemon=True)
        proc.start()
    except Exception:
        handle_error(traceback.format_exc(), button, base_text, default_color, hover_color)
        return

    _worker.update(proc=proc, queue=result_q, job=None)
    poll_worker(button, base_text, default_color, hover_color)


def cancel_operation():
    """Immediately kills the worker process and restores the UI."""
    if getattr(app, 'is_processing', False):
        app.cancel_requested = True
        _stop_worker()  # terminates the worker process immediately
        log_message("Warning: Operation cancelled by user.")
        _finish_common()
        set_progress("idle")
        set_title(None)
        set_banner("Operation cancelled. Ready for the next task.", WARNING_COLOR)
        refresh_steps()
        run = getattr(app, "current_run", None)
        if run:
            add_history("cancel", run[0], run[1])


def request_cancel(_event=None):
    """Cancel button / Esc. Long-running jobs ask first so a stray keypress can't discard minutes of work."""
    if not getattr(app, 'is_processing', False):
        return
    if time.time() - start_time > 20:
        if not custom_askyesno("Stop operation?", "Stop the running operation now? Progress so far will be discarded.",
                               allow_remember=False):
            return
    cancel_operation()


def format_duration(seconds):
    """Formats elapsed seconds into a clean human-readable string."""
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    m = int(seconds // 60)
    s = seconds % 60
    return f"{m} min {s:.1f} sec"


def short_duration(seconds):
    if seconds < 60:
        return f"{seconds:.1f}s"
    return f"{int(seconds // 60)}m {int(seconds % 60):02d}s"


def _summary_text():
    stats = app.stats
    if not stats["files"]:
        return "No files cleaned yet"
    last = f"  ·  last run {short_duration(stats['last'])}" if stats.get("last") is not None else ""
    return f"{stats['files']:,} file{'s' if stats['files'] != 1 else ''}  ·  {stats['rows']:,} rows{last}"


def update_stats(rows, elapsed):
    stats = app.stats
    stats["files"] += 1
    stats["rows"] += int(rows)
    stats["last"] = elapsed
    session_summary.configure(text=_summary_text())


# --- Session run history ---
HISTORY_MAX = 50
_HIST_ICONS = {"ok": ("✓", "ok"), "warn": ("⚠", "warn"), "err": ("✗", "err"), "cancel": ("■", "cancel")}


def _current_file_name():
    path = getattr(app, "selected_path", None)
    return Path(path).name if path else "(no file)"


def _render_history():
    t = hist_text
    t.configure(state="normal")
    t.delete("1.0", "end")
    if not app.history:
        t.insert("end", "No runs yet — finished runs will be listed here.", "muted")
    else:
        last = len(app.history) - 1
        for i, e in enumerate(app.history):
            icon, tag = _HIST_ICONS[e["status"]]
            t.insert("end", icon + "  ", tag)
            parts = [e["time"]]
            if e["status"] in ("ok", "warn"):
                parts += [short_duration(e["elapsed"]), f"{e['rows']:,} rows"]
                if e["warns"]:
                    parts.append(f"{e['warns']} warning{'s' if e['warns'] != 1 else ''}")
            elif e["status"] == "err":
                parts.append("failed")
            else:
                parts.append("cancelled")
            parts += [e["cleaner"], e["file"]]
            t.insert("end", "  ·  ".join(parts) + ("\n" if i < last else ""), "rest")
    t.configure(state="disabled")
    t.yview_moveto(0)


def add_history(status, file_name, cleaner, rows=None, elapsed=None, warns=0):
    """Records one run (newest first) in the session strip."""
    app.history.insert(0, {"status": status, "file": file_name, "cleaner": cleaner, "rows": rows,
                           "elapsed": elapsed, "warns": warns,
                           "time": datetime.datetime.now().strftime("%I:%M %p").lstrip("0")})
    del app.history[HISTORY_MAX:]
    _render_history()


# --- Completion alert (taskbar flash + system sound) ---
class _FLASHWINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("hwnd", ctypes.c_void_p), ("dwFlags", ctypes.c_uint),
                ("uCount", ctypes.c_uint), ("dwTimeout", ctypes.c_uint)]


def notify_done(kind, elapsed):
    """When a run ends while the window is in the background, flash the taskbar button until it is
    focused and play a system sound. A run you were watching only gets the sound, and only if it was long."""
    try:
        away = app.state() == "iconic" or app.focus_displayof() is None
        if not away and elapsed < 20:
            return
        if sys.platform == "win32":
            if away:
                user32 = ctypes.windll.user32
                user32.GetParent.argtypes = [ctypes.c_void_p]
                user32.GetParent.restype = ctypes.c_void_p
                info = _FLASHWINFO(ctypes.sizeof(_FLASHWINFO), user32.GetParent(app.winfo_id()),
                                   0x2 | 0xC, 0, 0)  # FLASHW_TRAY | FLASHW_TIMERNOFG
                user32.FlashWindowEx(ctypes.byref(info))
            if winsound is not None:
                winsound.MessageBeep({"success": winsound.MB_ICONASTERISK,
                                      "warning": winsound.MB_ICONEXCLAMATION,
                                      "error": winsound.MB_ICONHAND}.get(kind, winsound.MB_OK))
        else:
            app.bell()
    except Exception:
        pass


def log_run_summary(stats):
    """Logs what the cleaner did and returns how many warnings it raised."""
    if not stats:
        return 0
    warnings = 0
    rows_in, rows_out = stats.get("rows_in"), stats.get("rows_out")
    if rows_in is not None and rows_out is not None:
        log_message(f"Summary: {rows_in:,} rows in, {rows_out:,} rows out (net {rows_out - rows_in:+,})")
    for label, count in stats.get("removed", []):
        log_message(f"   {label}: {count:,} row(s)")
    if stats.get("pairs"):
        log_message(f"   Sale/return pairs matched: {stats['pairs']:,}")
    unmapped = stats.get("unmapped_stores") or []
    if unmapped:
        warnings += 1
        shown = ", ".join(unmapped[:8]) + (f" … (+{len(unmapped) - 8} more)" if len(unmapped) > 8 else "")
        log_message(f"Warning: {len(unmapped)} store name(s) not in the store map, kept as-is: {shown}")
    if stats.get("interest_unapplied"):
        warnings += 1
        log_message(f"Warning: {stats['interest_unapplied']} interest line(s) had no sale line on their receipt (see the Removal Audit Log)")
    if stats.get("interest_mismatch"):
        warnings += 1
        log_message(f"Warning: {stats['interest_mismatch']} interest line(s) are marked MISMATCH in the Removal Audit Log")
    return warnings


def handle_success(row_count, output_path, elapsed_seconds, button, base_text, default_color, hover_color, stats=None):
    _finish_common()
    app.run_done = True
    if getattr(app, "selected_path", None):
        app.cleaned_log.add(_clean_key(app.selected_path, base_text))
    duration_str = format_duration(elapsed_seconds)

    set_progress("success")
    log_message(f"Success! Cleaned {row_count} rows in {duration_str}.")
    warn_count = log_run_summary(stats)
    if warn_count:
        set_banner(f"Completed with {warn_count} warning(s) — {int(row_count):,} rows cleaned in {duration_str}. See the Activity Log.", WARNING_COLOR)
    else:
        set_banner(f"Completed — {int(row_count):,} rows cleaned in {duration_str}", SUCCESS_COLOR)
    set_title("Completed")
    update_stats(row_count, elapsed_seconds)

    app.last_output_path = output_path
    log_message(f"Saved to: {output_path}")
    open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)
    refresh_steps()
    add_history("warn" if warn_count else "ok", _current_file_name(), base_text,
                rows=int(row_count), elapsed=elapsed_seconds, warns=warn_count)
    notify_done("warning" if warn_count else "success", elapsed_seconds)


def handle_error(error_message, button, base_text, default_color, hover_color):
    _finish_common()
    set_progress("error")
    set_banner("Something went wrong — details are in the Activity Log", ERROR_COLOR)
    set_title("Failed")

    log_message("Error during cleaning process:", level="error")
    for line in error_message.strip().split("\n"):
        log_message(line, level="error")
    log_message("Tip: make sure the selected file matches the cleaner you chose, then try again.")
    refresh_steps()
    add_history("err", _current_file_name(), base_text)
    notify_done("error", time.time() - start_time)


# --- File Handlers ---
_COUNTABLE = (".csv", ".txt", ".xlsx", ".xlsm")


def _format_modified(ts):
    dt = datetime.datetime.fromtimestamp(ts)
    days = (datetime.date.today() - dt.date()).days
    clock = dt.strftime("%I:%M %p").lstrip("0")
    if days == 0:
        return f"today, {clock}"
    if days == 1:
        return f"yesterday, {clock}"
    return dt.strftime("%b %d, %Y").replace(" 0", " ")


def _count_rows(path):
    """Row count text for the file card (e.g. '48,210 rows'), or None when it can't be told cheaply.
    Runs in a background thread so big files never freeze the window."""
    ext = Path(path).suffix.lower()
    try:
        if ext in (".csv", ".txt"):
            lines, last = 0, b""
            with open(path, "rb") as f:
                while True:
                    chunk = f.read(1 << 20)
                    if not chunk:
                        break
                    lines += chunk.count(b"\n")
                    last = chunk[-1:]
            if last and last != b"\n":
                lines += 1
            if ext == ".csv":
                return f"{max(lines - 1, 0):,} rows"
            return f"{lines:,} lines"
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                n = wb.active.max_row
                if not n or n >= 1048576:  # no usable dimension info in the file
                    return None
                text = f"{max(n - 1, 0):,} rows"
                if len(wb.sheetnames) > 1:
                    text += f" (active sheet of {len(wb.sheetnames)})"
                return text
            finally:
                wb.close()
    except Exception:
        return None
    return None


def _render_file_label(count_text, pending=False):
    name, size_text, modified = app.file_meta
    parts = []
    if pending:
        parts.append("counting rows…")
    elif count_text:
        parts.append(count_text)
    if modified:
        parts.append(f"modified {modified}")
    text = f"{name}  •  {size_text}"
    if parts:
        text += "\n" + "  ·  ".join(parts)
    file_label.configure(text=text, text_color=SUCCESS_COLOR)


def _start_file_info(filename):
    """Counts rows off the UI thread and fills in the file card when done (ignored if the file changed)."""
    token = object()
    app._info_token = token
    box = {"text": None, "done": False}

    def work():
        try:
            box["text"] = _count_rows(filename)
        finally:
            box["done"] = True

    threading.Thread(target=work, daemon=True).start()

    def poll():
        if getattr(app, "_info_token", None) is not token:
            return
        if not box["done"]:
            app.after(150, poll)
            return
        _render_file_label(box["text"])
        schedule_layout()

    app.after(150, poll)

def process_selected_file(filename):
    if filename:
        short_name = Path(filename).name
        file_size = os.path.getsize(filename)
        size_text = f"{file_size / (1024 * 1024):.2f} MB" if file_size >= 1024 * 1024 else f"{file_size / 1024:.1f} KB"
        try:
            modified = _format_modified(os.path.getmtime(filename))
        except Exception:
            modified = ""
        app.file_meta = (short_name, size_text, modified)
        can_count = Path(filename).suffix.lower() in _COUNTABLE
        _render_file_label(None, pending=can_count)
        if can_count:
            _start_file_info(filename)
        else:
            app._info_token = None
        upload_card.configure(border_color=FILE_LOADED_BORDER, border_width=1)
        app.selected_path = filename
        app.run_done = False
        clear_btn.grid(row=0, column=2, padx=(5, 15))

        set_cleaners_enabled(True)

        CONFIG["last_dir"] = str(Path(filename).parent)
        save_config(CONFIG)

        set_progress("idle")
        set_title(None)
        set_banner("File ready — choose a cleaner in step 2.", MUTED_TEXT)
        log_message(f"Retail file loaded: {short_name}")
        refresh_steps()
        schedule_layout()  # the file name may wrap onto a new line


def upload_file(_event=None):
    if getattr(app, 'is_processing', False):
        return
    initial = CONFIG.get("last_dir")
    filename = filedialog.askopenfilename(
        title="Select Retail File",
        initialdir=initial if initial and os.path.isdir(initial) else None,
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if filename:
        process_selected_file(filename)
    else:
        log_message("Retail file selection cancelled.")


def clear_file():
    app.selected_path = None
    app._info_token = None
    app.run_done = False
    file_label.configure(text=FILE_PLACEHOLDER, text_color=MUTED_TEXT)
    upload_card.configure(border_color=BORDER_COLOR, border_width=1)
    clear_btn.grid_forget()

    set_cleaners_enabled(False)
    if not getattr(app, 'last_output_path', None):
        open_output_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)

    set_progress("idle")
    set_title(None)
    set_banner("Ready — start by selecting a retail file.", MUTED_TEXT)
    log_message("Retail file selection cleared.")
    refresh_steps()
    schedule_layout()


def _open_path(path):
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        os.system(f"open '{path}'")
    else:
        os.system(f"xdg-open '{path}'")


def open_cleaned_file():
    if getattr(app, 'last_output_path', None) and os.path.exists(app.last_output_path):
        try:
            _open_path(app.last_output_path)
            log_message(f"Opened file: {Path(app.last_output_path).name}")
        except Exception as e:
            log_message(f"Error opening file: {str(e)}")
    else:
        log_message("No generated file found to open.")


def _reveal_path(path):
    """Opens the file's folder with the file highlighted (Explorer / Finder)."""
    path = str(path)
    if sys.platform == "win32":
        subprocess.Popen(["explorer", f"/select,{os.path.normpath(path)}"])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", str(Path(path).parent)])


def open_destination_folder():
    last = getattr(app, "last_output_path", None)
    if last and os.path.exists(last):
        try:
            _reveal_path(last)
            log_message(f"Opened destination folder with {Path(last).name} selected")
            return
        except Exception:
            pass  # fall through to plain folder open
    output_dir = Path.home() / "Desktop" / "Cleaned Data"
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        _open_path(output_dir)
        log_message(f"Opened destination folder: {output_dir.name}")
    except Exception as e:
        log_message(f"Error opening folder: {str(e)}")


def clear_log(_event=None):
    """Clears the activity log."""
    status_box.configure(state="normal")
    status_box.delete("1.0", "end")
    status_box.configure(state="disabled")
    _refresh_warn_badge()
    log_message("Activity log cleared.")


def copy_log():
    """Copies the complete activity log to the clipboard."""
    try:
        content = status_box.get("1.0", "end-1c")
        app.clipboard_clear()
        app.clipboard_append(content)
        log_message("Activity log copied to clipboard.")
    except Exception as e:
        log_message(f"Error copying activity log: {e}")


def save_log():
    """Saves the complete activity log (every level, whatever the filter shows) to a text file."""
    try:
        content = status_box.get("1.0", "end-1c")
        out_dir = Path.home() / "Desktop" / "Cleaned Data"
        last_dir = CONFIG.get("last_dir")
        initial = out_dir if out_dir.is_dir() else (last_dir if last_dir and os.path.isdir(last_dir) else Path.home())
        path = filedialog.asksaveasfilename(
            title="Save Activity Log", initialdir=str(initial), defaultextension=".txt",
            initialfile=f"activity_log_{datetime.datetime.now():%Y-%m-%d_%H%M}.txt",
            filetypes=(("Text Files", "*.txt"), ("All Files", "*.*")))
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(content + "\n")
        log_message(f"Saved activity log to: {Path(path).name}")
    except Exception as e:
        log_message(f"Error saving activity log: {e}")


def custom_askyesno(title, message, allow_remember=True):
    """Modal confirmation. Enter = Yes, Esc = No. Optionally offers 'don't ask again this session'."""
    dialog = ctk.CTkToplevel(app)
    dialog.title(title)
    dialog_width, dialog_height = 400, (232 if allow_remember else 214)
    dialog.geometry(f"{dialog_width}x{dialog_height}")
    dialog.resizable(False, False)
    dialog.configure(fg_color=NAVY_BLUE)

    try:
        icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
        dialog.iconbitmap(str(icon_path))
    except Exception:
        pass

    dialog.transient(app)
    dialog.grab_set()

    app.update_idletasks()
    dialog.update_idletasks()
    scale = widget_scale(app)
    x = app.winfo_x() + (app.winfo_width() - int(dialog_width * scale)) // 2
    y = app.winfo_y() + (app.winfo_height() - int(dialog_height * scale)) // 2
    dialog.geometry(f"{dialog_width}x{dialog_height}+{max(x, 0)}+{max(y, 0)}")

    result = [False]
    closed = [False]
    remember = tk.BooleanVar(master=dialog, value=False)

    def _close(value):
        if closed[0]:  # Enter can reach both the focused button and the window; only act once
            return
        closed[0] = True
        result[0] = value
        dialog.destroy()

    def on_yes(_e=None):
        _close(True)

    def on_no(_e=None):
        _close(False)

    ctk.CTkLabel(dialog, text=title, font=("Segoe UI", 14, "bold"), text_color=WHITE,
                 anchor="w").pack(fill="x", padx=24, pady=(20, 4))
    ctk.CTkLabel(dialog, text=message, font=("Segoe UI", 12), text_color=LOG_TEXT, wraplength=352,
                 justify="left", anchor="w").pack(fill="x", padx=24, pady=(0, 10))

    if allow_remember:
        ctk.CTkCheckBox(dialog, text="Don't ask me again this session", variable=remember, font=("Segoe UI", 11),
                        text_color=MUTED_TEXT, fg_color=ACCENT_ORANGE, hover_color=HOVER_ORANGE,
                        border_color=BORDER_COLOR, checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=24, pady=(0, 6))

    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(side="bottom", fill="x", padx=24, pady=(0, 20))
    btn_frame.grid_columnconfigure(0, weight=1)
    btn_frame.grid_columnconfigure(1, weight=1)

    yes_btn = GradientButton(btn_frame, text="Yes, continue", font=("Segoe UI", 12, "bold"), fg_color=ACCENT_ORANGE,
                             hover_color=HOVER_ORANGE, gradient_to=ACCENT_GRADIENT_END, corner_radius=8,
                             height=38, command=on_yes, bg=NAVY_BLUE)
    yes_btn.grid(row=0, column=0, sticky="ew", padx=(0, px(6)))

    no_btn = ctk.CTkButton(btn_frame, text="No", font=("Segoe UI", 12, "bold"), fg_color=SURFACE_BG,
                           hover_color=CANCEL_COLOR, text_color=WHITE, corner_radius=8, height=38, command=on_no)
    no_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))

    dialog.bind("<Return>", on_yes)
    dialog.bind("<Escape>", on_no)
    dialog.protocol("WM_DELETE_WINDOW", on_no)
    try:
        dialog.after(120, yes_btn.focus_set)
    except Exception:
        pass

    dialog.wait_window()
    if result[0] and allow_remember and remember.get():
        app.skip_confirm = True
        log_message("Confirmations turned off for this session.")
    return result[0]


def start_cleaner(c):
    """Confirms (unless turned off) and starts the chosen cleaner."""
    if getattr(app, "is_processing", False):
        return
    if not getattr(app, "selected_path", None):
        log_message("Warning: Please upload a retail file before running the cleaner.")
        return
    name = Path(app.selected_path).name
    force = False
    if already_cleaned(app.selected_path, c["label"]):
        if not custom_askyesno("Already cleaned",
                               f"“{name}” was already cleaned with the {c['label']} this session. "
                               "Run it again anyway? A new output file is created; nothing is overwritten.",
                               allow_remember=False):
            warn_already_cleaned(c["label"])
            return
        force = True  # they just confirmed, so no second prompt
    elif not app.skip_confirm:
        if not custom_askyesno("Confirm Action", f"Run the {c['label']} on “{name}”?"):
            log_message(f"{c['label']} cancelled before starting.")
            return
    run_cleaner_process((c["module"], c["func"]), c["btn"], c["label"], c["color"], c["hover"], force=force)


# ======================================================================
# Loading splash: one small window that covers the build of the real UI
# ======================================================================
class LoadingSplash(tk.Toplevel):
    """
    Borderless loading window. The real window is built invisibly behind it (alpha 0) and then swapped
    in fully laid out, so the user never watches the components appear one by one.
    """
    W, H = 400, 180  # unscaled units

    def __init__(self, master):
        super().__init__(master)
        self.withdraw()
        self.overrideredirect(True)
        self.configure(bg=BORDER_COLOR)  # shows through the 1 px canvas padding as a border
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        def s(value):
            return px(value, master)

        w, h = s(self.W), s(self.H)
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2 - s(20)
        self.geometry(f"{w}x{h}+{max(x, 0)}+{max(y, 0)}")

        cw = w - 2
        self._canvas = tk.Canvas(self, bg=NAVY_BLUE, highlightthickness=0, bd=0)
        self._canvas.pack(fill="both", expand=True, padx=1, pady=1)
        c = self._canvas

        c.create_rectangle(0, 0, cw, s(3), fill=ACCENT_ORANGE, outline="")
        c.create_text(cw / 2, s(58), text=BASE_TITLE, fill=WHITE, font=("Segoe UI", -s(24), "bold"))
        c.create_text(cw / 2, s(88), text="RETAIL STRATEGY DIVISION  /  DATA ANALYTICS", fill=SECTION_TEXT,
                      font=("Segoe UI", -s(10), "bold"))
        self._status = c.create_text(cw / 2, s(124), text="Starting…", fill=MUTED_TEXT, font=("Segoe UI", -s(12)))

        self._x0, self._x1 = s(40), cw - s(40)
        self._y0, self._y1 = s(146), s(146) + max(s(4), 3)
        c.create_rectangle(self._x0, self._y0, self._x1, self._y1, fill=SURFACE_BG, outline="")
        self._fill = c.create_rectangle(self._x0, self._y0, self._x0, self._y1, fill=ACCENT_ORANGE, outline="")
        c.create_text(cw / 2, s(166), text=APP_VERSION, fill=MUTED_TEXT, font=("Segoe UI", -s(9)))

        self.deiconify()
        self.lift()
        self.update()  # first paint; safe here because nothing else has been built yet

    def step(self, text, fraction):
        """Advance the bar. update_idletasks() only repaints, so no half-built UI callbacks fire."""
        try:
            self._canvas.itemconfigure(self._status, text=text)
            self._canvas.coords(self._fill, self._x0, self._y0,
                                self._x0 + (self._x1 - self._x0) * max(0.0, min(1.0, fraction)), self._y1)
            self.update_idletasks()
        except Exception:
            pass

    def close(self):
        try:
            self.destroy()
        except Exception:
            pass


# ======================================================================
# Main App Initialization
# ======================================================================
FILE_PLACEHOLDER = "Click here to choose an Excel or CSV file  (Ctrl+O)"

if __name__ == "__main__":
    multiprocessing.freeze_support()

    app = ctk.CTk()

    # Keep the real window invisible while the UI is built. It stays fully "alive" (mapped, measurable),
    # so every size/DPI calculation below behaves exactly as before; only the pixels are hidden.
    try:
        app.attributes("-alpha", 0.0)
    except Exception:
        pass
    try:
        splash = LoadingSplash(app)
    except Exception:
        splash = None

    def _splash_step(text, fraction):
        if splash is not None:
            splash.step(text, fraction)

    app.title(BASE_TITLE)

    icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
    try:
        app.iconbitmap(str(icon_path))
    except Exception:
        pass

    app.resizable(True, True)
    app.minsize(MIN_WIDTH, 560)  # refined from real content size once the UI is built
    app.configure(fg_color=NAVY_BLUE)

    app.is_processing = False
    app.cancel_requested = False
    app.selected_path = None
    app.last_output_path = None
    app.run_done = False
    app.skip_confirm = False
    app.stats = {"files": 0, "rows": 0, "last": None}
    app.history = []          # this session's runs, newest first
    app.current_run = None    # (file name, cleaner label) of the run in progress
    app.file_meta = None
    app.cleaned_log = set()  # (file name, cleaner label) pairs that already finished successfully this session
    CONFIG = load_config()

    # Grid: row0 header | row1 controls (+ log in two-pane mode) | row2 log (single column) | row3 footer
    main_frame = ctk.CTkFrame(app, fg_color=NAVY_BLUE, corner_radius=0)
    main_frame.pack(fill="both", expand=True, padx=MAIN_PADX, pady=MAIN_PADY)

    _splash_step("Building header…", 0.12)
    # --- Header (retail gradient banner with the app logo) ---
    header_banner = HeaderBanner(
        main_frame, "Strat Data Cleaner", "RETAIL STRATEGY DIVISION  /  DATA ANALYTICS",
        HEADER_STOPS, bg=NAVY_BLUE, on_height_change=lambda: schedule_layout(),
    )
    header_banner.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, px(16)))

    left_col = ctk.CTkFrame(main_frame, fg_color="transparent")
    log_col = ctk.CTkFrame(main_frame, fg_color="transparent")
    footer_frame = ctk.CTkFrame(main_frame, fg_color="transparent")

    # ------------------------------------------------------------------
    # Controls column
    # ------------------------------------------------------------------
    def make_step_row(parent, number, title):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.grid_columnconfigure(2, weight=1)
        badge = ctk.CTkLabel(row, text=str(number), width=22, height=22, corner_radius=11, fg_color=STEP_TODO_BG,
                             text_color=STEP_TODO_TEXT, font=("Segoe UI", 11, "bold"))
        badge.grid(row=0, column=0)
        label = ctk.CTkLabel(row, text=title, font=("Segoe UI", 12, "bold"), text_color=MUTED_TEXT)
        label.grid(row=0, column=1, padx=(8, 0))
        hint = ctk.CTkLabel(row, text="", font=("Segoe UI", 10), text_color=MUTED_TEXT, anchor="e", justify="right")
        hint.grid(row=0, column=2, sticky="e")
        return row, badge, label, hint

    _splash_step("Building controls…", 0.30)
    # Step 1 - select file
    step1_row, step1_badge, step1_title, _ = make_step_row(left_col, 1, "SELECT RETAIL FILE")
    step1_row.pack(fill="x", pady=(0, 6))

    upload_card = ctk.CTkFrame(left_col, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
    upload_card.pack(fill="x", pady=(0, 12))
    upload_card.grid_columnconfigure(1, weight=1)

    upload_btn = GradientButton(upload_card, text="Browse File", font=("Segoe UI", 12, "bold"),
                                fg_color=ACCENT_ORANGE, hover_color=HOVER_ORANGE, gradient_to=ACCENT_GRADIENT_END,
                                text_color=WHITE, corner_radius=8, width=132, height=42, command=upload_file, bg=CARD_BG)
    upload_btn.grid(row=0, column=0, padx=px(14), pady=px(14))

    file_label = ctk.CTkLabel(upload_card, text=FILE_PLACEHOLDER, font=("Segoe UI", 11),
                              text_color=MUTED_TEXT, anchor="w", justify="left", wraplength=200)
    file_label.grid(row=0, column=1, sticky="ew", padx=10)
    file_label.bind("<Button-1>", upload_file)  # the whole label is a big click target
    try:
        file_label._label.configure(cursor="hand2")
    except Exception:
        pass

    clear_btn = ctk.CTkButton(upload_card, text="×", font=("Segoe UI", 17, "bold"), fg_color=SURFACE_BG, hover_color=CLOSE_HOVER,
                              text_color=WHITE, corner_radius=7, width=34, height=34, command=clear_file)

    # Step 2 - choose cleaner
    step2_row, step2_badge, step2_title, step2_hint = make_step_row(left_col, 2, "CHOOSE A CLEANER")
    step2_row.pack(fill="x", pady=(0, 6))

    btn_frame = ctk.CTkFrame(left_col, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(0, 8))

    button_style = {
        "font": ("Segoe UI", 12, "bold"),
        "text_color": WHITE,
        "corner_radius": 10,
        "height": 48,
        "bg": NAVY_BLUE,
    }
    _splash_step("Rendering buttons…", 0.50)
    cadence_btn = GradientButton(btn_frame, text="Cadence Cleaner", fg_color=CADENCE_COLOR, hover_color=CADENCE_HOVER,
                                 border_color=CADENCE_ACCENT, state="disabled", **button_style)
    database_btn = GradientButton(btn_frame, text="Database Cleaner", fg_color=DB_COLOR, hover_color=DB_HOVER,
                                  border_color=DB_ACCENT, state="disabled", **button_style)
    brand_performance_btn = GradientButton(btn_frame, text="Brand Performance Cleaner", fg_color=BRAND_COLOR, hover_color=BRAND_HOVER,
                                           border_color=BRAND_ACCENT, state="disabled", **button_style)
    existing_db_btn = GradientButton(btn_frame, text="Existing DB Cleaner", fg_color=EXISTING_DB_COLOR, hover_color=EXISTING_DB_HOVER,
                                     border_color=EXISTING_DB_ACCENT, state="disabled", **button_style)
    current_sales_btn = GradientButton(btn_frame, text="Current/Daily Sales Cleaner", fg_color=CURRENT_SALES_COLOR, hover_color=CURRENT_SALES_HOVER,
                                       border_color=CURRENT_SALES_ACCENT, state="disabled", **button_style)
    cleaner_buttons = [cadence_btn, database_btn, brand_performance_btn, existing_db_btn, current_sales_btn]

    # The module/function pairs are imported by name inside the worker process.
    CLEANERS = [
        {"btn": cadence_btn, "label": "Cadence Cleaner", "color": CADENCE_COLOR, "hover": CADENCE_HOVER,
         "accent": CADENCE_ACCENT, "module": "cadence_cleaner", "func": "process_cadence", "desc": "Cadence report. Adds CATEGORY, CHANNEL and PRODUCT_TAG (Raptor, EcoFlow, Insta360, Oura). Offsets returns and rolls up interest and accessories."},
        {"btn": database_btn, "label": "Database Cleaner", "color": DB_COLOR, "hover": DB_HOVER,
         "accent": DB_ACCENT, "module": "database_cleaner", "func": "process_database", "desc": "Database export with standard store names (e.g. Altitude Alabang Town Center). Offsets returns and rolls up interest and accessories."},
        {"btn": brand_performance_btn, "label": "Brand Performance Cleaner", "color": BRAND_COLOR, "hover": BRAND_HOVER,
         "accent": BRAND_ACCENT, "module": "brand_performance_cleaner", "func": "process_brand_performance", "desc": "Brand performance report. Adds CHANNEL_SEGMENT, TRACKED_BRAND and QUARTER, and removes store-days with a negative total."},
        {"btn": existing_db_btn, "label": "Existing DB Cleaner", "color": EXISTING_DB_COLOR, "hover": EXISTING_DB_HOVER,
         "accent": EXISTING_DB_ACCENT, "module": "existing_db_cleaner", "func": "process_existing_database", "desc": "Re-cleans a database file you already cleaned (reads its DatabaseCleanTable): refreshes month/week/quarter, offsets returns, rolls up interest and accessories."},
        {"btn": current_sales_btn, "label": "Current/Daily Sales Cleaner", "color": CURRENT_SALES_COLOR, "hover": CURRENT_SALES_HOVER,
         "accent": CURRENT_SALES_ACCENT, "module": "current_sales_cleaner", "func": "process_current_sales", "desc": "Current/daily sales export with all-caps store names (e.g. ALTITUDE ALABANG TOWN CENTER). Same cleaning as the Database Cleaner."},
    ]
    for _c in CLEANERS:
        _c["btn"].configure(command=lambda c=_c: start_cleaner(c))

    # One-line description of whichever cleaner the pointer (or keyboard focus) is on.
    DESC_DEFAULT = f"Hover over a cleaner to see what it is for.  Press 1–{len(CLEANERS)} to run one."
    # The label sits in a frame whose height is reserved for the LONGEST description at the current width
    # (see _fit_desc_box), so hovering from button to button can never resize or reflow the window.
    desc_frame = ctk.CTkFrame(left_col, fg_color="transparent", height=32)
    desc_frame.pack(fill="x", pady=(0, 10))
    try:
        desc_frame.pack_propagate(False)
    except Exception:
        pass
    cleaner_desc = ctk.CTkLabel(desc_frame, text=DESC_DEFAULT, font=("Segoe UI", 10), text_color=MUTED_TEXT,
                                anchor="nw", justify="left", wraplength=400)
    cleaner_desc.pack(fill="both", expand=True)

    def _show_desc(c=None):
        cleaner_desc.configure(text=f"{c['label']}:  {c['desc']}" if c else DESC_DEFAULT)

    for _c in CLEANERS:
        _c["btn"].bind("<Enter>", lambda e, c=_c: _show_desc(c), add="+")
        _c["btn"].bind("<Leave>", lambda e: _show_desc(), add="+")
        _c["btn"].bind("<FocusIn>", lambda e, c=_c: _show_desc(c), add="+")
        _c["btn"].bind("<FocusOut>", lambda e: _show_desc(), add="+")

    # Operation status + indeterminate progress bar + cancel
    operation_card = ctk.CTkFrame(left_col, fg_color=CARD_BG, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
    operation_card.pack(fill="x", pady=(0, 12))
    operation_card.grid_columnconfigure(0, weight=1)

    live_status_banner = ctk.CTkLabel(operation_card, text="Ready — start by selecting a retail file.", font=("Segoe UI", 11),
                                      text_color=MUTED_TEXT, anchor="w", justify="left", wraplength=300, height=34)
    live_status_banner.grid(row=0, column=0, sticky="ew", padx=14, pady=(8, 2))

    cancel_btn = ctk.CTkButton(operation_card, text="Cancel", font=("Segoe UI", 11, "bold"), fg_color=CANCEL_COLOR,
                               hover_color=CANCEL_HOVER, text_color=WHITE, corner_radius=7, height=32, width=92,
                               state="disabled", command=request_cancel)
    cancel_btn.grid(row=0, column=1, padx=12, pady=(8, 2))

    progress = SmoothProgressBar(operation_card, height=4, fg_color=SURFACE_BG, progress_color=INFO_COLOR,
                                 bg=CARD_BG, mode="determinate")
    progress.set(0)
    # A raw Tk canvas is not auto-scaled by CustomTkinter, so its padding goes through px().
    progress.grid(row=1, column=0, columnspan=2, sticky="ew", padx=px(14), pady=(0, px(10)))

    # Step 3 - review results
    step3_row, step3_badge, step3_title, _ = make_step_row(left_col, 3, "REVIEW RESULTS")
    step3_row.pack(fill="x", pady=(0, 6))
    step_widgets = [(step1_badge, step1_title, 1), (step2_badge, step2_title, 2), (step3_badge, step3_title, 3)]

    output_frame = ctk.CTkFrame(left_col, fg_color="transparent")
    output_frame.pack(fill="x", pady=(0, 8))
    output_frame.grid_columnconfigure((0, 1), weight=1, uniform="outputs")

    open_output_btn = GradientButton(output_frame, text="Open Cleaned File", font=("Segoe UI", 11, "bold"),
                                     fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER, gradient_to=OPEN_BTN_GRADIENT_END,
                                     text_color=WHITE, corner_radius=8, height=40, state="disabled",
                                     command=open_cleaned_file, bg=NAVY_BLUE)
    open_output_btn.grid(row=0, column=0, sticky="ew", padx=(0, px(6)))

    open_folder_btn = GradientButton(output_frame, text="Open Destination Folder", font=("Segoe UI", 11, "bold"),
                                     fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER, gradient_to=OPEN_BTN_GRADIENT_END,
                                     text_color=WHITE, corner_radius=8, height=40,
                                     command=open_destination_folder, bg=NAVY_BLUE)
    open_folder_btn.grid(row=0, column=1, sticky="ew", padx=(px(6), 0))

    # Session strip: one-line totals plus the most recent runs (scroll for older ones)
    stats_card = ctk.CTkFrame(left_col, fg_color=CARD_BG, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
    stats_card.pack(fill="x")

    hist_head = ctk.CTkFrame(stats_card, fg_color="transparent")
    hist_head.pack(fill="x", padx=12, pady=(8, 3))
    ctk.CTkLabel(hist_head, text="THIS SESSION", font=("Segoe UI", 9, "bold"), text_color=MUTED_TEXT).pack(side="left")
    session_summary = ctk.CTkLabel(hist_head, text="No files cleaned yet", font=("Segoe UI", 10, "bold"),
                                   text_color=INFO_COLOR, anchor="e")
    session_summary.pack(side="right")

    # A raw Tk text keeps the history light; its font is rescaled with the DPI like the other raw pieces.
    hist_text = tk.Text(stats_card, height=3, bg=CARD_BG, fg=LOG_TEXT, bd=0, highlightthickness=0, wrap="none",
                        cursor="arrow", font=("Segoe UI", -px(11)), padx=0, pady=1, takefocus=0, state="disabled")
    hist_text.pack(fill="x", padx=12, pady=(0, 8))
    hist_text.tag_configure("ok", foreground=SUCCESS_COLOR)
    hist_text.tag_configure("warn", foreground=WARNING_COLOR)
    hist_text.tag_configure("err", foreground=ERROR_COLOR)
    hist_text.tag_configure("cancel", foreground=MUTED_TEXT)
    hist_text.tag_configure("muted", foreground=MUTED_TEXT)
    hist_text.tag_configure("rest", foreground=LOG_TEXT)
    _render_history()

    # ------------------------------------------------------------------
    # Activity log column
    # ------------------------------------------------------------------
    _splash_step("Building activity log…", 0.72)
    log_col.grid_columnconfigure(0, weight=1)
    log_col.grid_rowconfigure(1, weight=1)

    log_header = ctk.CTkFrame(log_col, fg_color="transparent")
    log_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))

    log_top = ctk.CTkFrame(log_header, fg_color="transparent")
    log_top.pack(fill="x")

    status_label = ctk.CTkLabel(log_top, text="ACTIVITY LOG", font=("Segoe UI", 12, "bold"), text_color=SECTION_TEXT)
    status_label.pack(side="left")

    log_actions = ctk.CTkFrame(log_top, fg_color="transparent")
    log_actions.pack(side="right")

    def _log_button(text, command, padx):
        btn = ctk.CTkButton(log_actions, text=text, font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG, hover_color="#334155",
                            text_color=WHITE, corner_radius=6, height=27, width=54, command=command)
        btn.pack(side="left", padx=padx)
        return btn

    copy_log_btn = _log_button("Copy", copy_log, (0, 6))
    save_log_btn = _log_button("Save", save_log, (0, 6))
    clear_log_btn = _log_button("Clear", clear_log, 0)

    # Second row: level filter + a badge that jumps to warnings (shown only when there are some)
    log_tools = ctk.CTkFrame(log_header, fg_color="transparent")
    log_tools.pack(fill="x", pady=(6, 0))

    log_filter = ctk.CTkSegmentedButton(log_tools, values=list(LOG_FILTERS), command=_apply_log_filter,
                                        font=("Segoe UI", 10, "bold"), height=24, corner_radius=6,
                                        fg_color=SURFACE_BG, unselected_color=SURFACE_BG,
                                        unselected_hover_color="#334155", selected_color=ACCENT_ORANGE,
                                        selected_hover_color=HOVER_ORANGE, text_color=WHITE)
    log_filter.set("All")
    log_filter.pack(side="left")

    warn_badge = ctk.CTkButton(log_tools, text="", font=("Segoe UI", 10, "bold"), fg_color="#3A2E0C", hover_color="#4D3C0F",
                               text_color=WARNING_COLOR, border_width=1, border_color="#8A6A12", corner_radius=6,
                               height=24, width=96, command=jump_to_warning)  # packed by _refresh_warn_badge

    status_box = ctk.CTkTextbox(log_col, height=LOG_MIN_HEIGHT, corner_radius=9, fg_color=LOG_BG, text_color=LOG_TEXT,
                                font=("Consolas", 11), border_width=1, border_color=BORDER_COLOR, wrap="word")
    status_box.grid(row=1, column=0, sticky="nsew")

    status_box._textbox.tag_config("normal", foreground=LOG_TEXT)
    status_box._textbox.tag_config("success", foreground=SUCCESS_COLOR)
    status_box._textbox.tag_config("warning", foreground=WARNING_COLOR)
    status_box._textbox.tag_config("error", foreground=ERROR_COLOR)
    status_box._textbox.tag_config("notice", foreground=WARNING_COLOR)
    status_box._textbox.tag_config("processing", foreground=INFO_COLOR)
    status_box._textbox.tag_config("flash", background="#3B2F0B")  # brief highlight when jumping to a warning

    status_box.insert("end", "[System] Application ready. Browse a retail file to begin.\n", "normal")
    status_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # Footer
    # ------------------------------------------------------------------
    footer_frame.grid_columnconfigure(1, weight=1)
    footer_label = ctk.CTkLabel(footer_frame, text=f"STRAT DATA CLEANER  •  {APP_VERSION}", font=("Segoe UI", 10, "bold"), text_color=MUTED_TEXT)
    footer_hint = ctk.CTkLabel(footer_frame, text=f"Ctrl+O  Browse    ·    1–{len(CLEANERS)}  Run cleaner    ·    Esc  Cancel    ·    Ctrl+L  Clear log",
                               font=("Segoe UI", 10), text_color=MUTED_TEXT)
    footer_credit = ctk.CTkLabel(footer_frame, text=f"Retail Strategy Division · Data Analytics · {CURRENT_YEAR}", font=("Segoe UI", 10),
                                 text_color=MUTED_TEXT)


    # ======================================================================
    # Responsive layout engine
    # ======================================================================
    _splash_step("Fitting to your screen…", 0.88)
    _layout = {"mode": None, "cols": None, "stacked_footer": None, "wrap": {}, "compact": False,
               "main_pady": MAIN_PADY, "auto_compact": True, "dual_w": None,
               "scale": None, "settle_until": 0.0, "scale_job": None}
    _resize_job = None
    _last_width = None
    _min_height = 560

    # (widget, normal pady, compact pady) - compact trims spacing only, never text size
    _PAD_TABLE = [
        (step1_row, (0, 6), (0, 4)),
        (upload_card, (0, 12), (0, 8)),
        (step2_row, (0, 6), (0, 4)),
        (btn_frame, (0, 8), (0, 6)),
        (desc_frame, (0, 10), (0, 6)),
        (operation_card, (0, 12), (0, 8)),
        (step3_row, (0, 6), (0, 4)),
        (output_frame, (0, 8), (0, 6)),
    ]

    def _app_width_px():
        """Real window width in pixels (falls back to the requested geometry before first map)."""
        w = app.winfo_width()
        if w <= 1:
            try:
                w = int(app.geometry().split("x")[0])
            except Exception:
                w = 1
        return w

    def _set_wrap(label, key, value):
        """Updates a label's wraplength only if it actually changed (avoids pointless relayouts)."""
        value = max(int(value), 80)
        if _layout["wrap"].get(key) != value:
            _layout["wrap"][key] = value
            label.configure(wraplength=value)

    def _fit_desc_box(wrap_u):
        """Reserves room for the longest cleaner description at this width, so hovering never changes the layout."""
        scale = widget_scale(app)
        key = (int(wrap_u), round(scale, 3))
        if _layout.get("desc_key") == key:
            return
        _layout["desc_key"] = key
        font = tkfont.Font(root=app, family="Segoe UI", size=-int(round(10 * scale)))
        limit = max(wrap_u * scale - 16, 40)  # small safety margin so the real label never wraps one line more
        longest = 1
        for text in [DESC_DEFAULT] + [f"{c['label']}:  {c['desc']}" for c in CLEANERS]:
            lines, current = 1, ""
            for word in text.split():
                trial = f"{current} {word}".strip()
                if current and font.measure(trial) > limit:
                    lines += 1
                    current = word
                else:
                    current = trial
            longest = max(longest, lines)
        desc_frame.configure(height=int((longest * font.metrics("linespace") + 6) / scale) + 1)

    def set_density(compact):
        """Normal vs compact spacing. Compact is used automatically only when the screen is too short."""
        if _layout["compact"] == compact:
            return
        _layout["compact"] = compact
        _layout["main_pady"] = MAIN_PADY_COMPACT if compact else MAIN_PADY
        main_frame.pack(pady=_layout["main_pady"])
        header_banner.grid_configure(pady=(0, px(10 if compact else 16)))
        for widget, normal, tight in _PAD_TABLE:
            widget.pack(pady=tight if compact else normal)
        header_banner.set_heights(*((58, 84) if compact else (72, 98)))
        for b in cleaner_buttons:
            b.set_height(42 if compact else 48)
        for b in (open_output_btn, open_folder_btn):
            b.set_height(36 if compact else 40)
        upload_btn.set_height(38 if compact else 42)
        upload_btn.grid_configure(pady=px(8 if compact else 14))
        status_box.configure(height=LOG_MIN_HEIGHT_COMPACT if compact else LOG_MIN_HEIGHT)
        place_footer(_layout["stacked_footer"], force=True)

    def place_panes(mode):
        """Single column (log underneath) or two panes (controls left, log right)."""
        if _layout["mode"] == mode:
            return
        _layout["mode"] = mode
        if mode == "dual":
            main_frame.grid_columnconfigure(0, weight=3, minsize=px(MIN_LEFT))
            main_frame.grid_columnconfigure(1, weight=2, minsize=px(MIN_RIGHT))
            main_frame.grid_rowconfigure(1, weight=1)
            main_frame.grid_rowconfigure(2, weight=0, minsize=0)
            left_col.grid(row=1, column=0, columnspan=1, sticky="nsew", padx=(0, PANE_GAP // 2), pady=(0, 0))
            log_col.grid(row=1, column=1, columnspan=1, sticky="nsew", padx=(PANE_GAP // 2, 0), pady=(0, 0))
        else:
            main_frame.grid_columnconfigure(0, weight=1, minsize=0)
            main_frame.grid_columnconfigure(1, weight=0, minsize=0)
            main_frame.grid_rowconfigure(1, weight=0)
            main_frame.grid_rowconfigure(2, weight=1, minsize=px(LOG_ROW_MIN))
            left_col.grid(row=1, column=0, columnspan=2, sticky="new", padx=0, pady=(0, 12))
            log_col.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=0, pady=(0, 0))
        place_footer(_layout["stacked_footer"], force=True)

    def place_footer(stacked, force=False):
        """Footer in one row when there is room, otherwise the shortcut hint drops to a second row."""
        if stacked is None:
            stacked = False
        if _layout["stacked_footer"] == stacked and not force and footer_frame.winfo_manager():
            return
        _layout["stacked_footer"] = stacked
        top_pad = 3 if _layout["compact"] else 6
        footer_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(top_pad, 0))
        footer_label.grid(row=0, column=0, sticky="w")
        footer_credit.grid(row=0, column=2, sticky="e")
        if stacked:
            footer_hint.grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 0))
        else:
            footer_hint.grid(row=0, column=1, columnspan=1, sticky="", pady=(0, 0))

    def _need_height_u():
        """Height (unscaled) the content needs in order to show everything."""
        app.update_idletasks()
        return main_frame.winfo_reqheight() / widget_scale(app) + 2 * _layout["main_pady"] + 2

    def _avail_height_u():
        """Usable window height (unscaled) on the monitor the window is on, after taskbar + title bar."""
        work = _monitor_work_height_px()
        if work:
            return (work - TITLEBAR_RESERVE_PX) / widget_scale(app)
        return (app.winfo_screenheight() - SCREEN_RESERVE_PX) / widget_scale(app)

    def update_min_size():
        """Keeps the minimum window size in sync with what the content really needs."""
        global _min_height
        scale = widget_scale(app)
        avail = _avail_height_u()
        need = _need_height_u()
        if _layout["auto_compact"] and not _layout["compact"] and need > avail:
            set_density(True)
            need = _need_height_u()
        need_h = int(min(need, avail))
        if need_h != _min_height:
            _min_height = need_h
            app.minsize(MIN_WIDTH, need_h)
        if app.winfo_height() > 1 and app.winfo_height() / scale < need_h - 1:
            app.geometry(f"{int(app.winfo_width() / scale)}x{need_h}")

    def apply_layout(width_u=None):
        """Reflows the UI for the current width. Control sizes stay fixed; layout adapts around them."""
        global _resize_job
        _resize_job = None
        scale = widget_scale(app)
        if width_u is None:
            width_u = _app_width_px() / scale

        # 1) Panes
        mode = "dual" if width_u >= DUAL_BREAKPOINT else "single"
        place_panes(mode)
        inner_u = width_u - 2 * MAIN_PADX
        col_w = max((inner_u - PANE_GAP) * 0.6, MIN_LEFT) if mode == "dual" else inner_u

        # 2) Cleaner buttons: 2 across normally, 4 across when the controls column is very wide
        cols = 4 if col_w >= WIDE_COLS_BREAKPOINT else 2
        if cols != _layout["cols"]:
            _layout["cols"] = cols
            for c in range(4):
                btn_frame.grid_columnconfigure(c, weight=1 if c < cols else 0, uniform="cleaners" if c < cols else "")
            last_row = (len(cleaner_buttons) - 1) // cols
            for i, b in enumerate(cleaner_buttons):
                r, c = divmod(i, cols)
                left = 0 if c == 0 else 7
                right = 0 if c == cols - 1 else 7
                bottom = 8 if r < last_row else 0
                b.grid(row=r, column=c, sticky="ew", padx=(px(left), px(right)), pady=(0, px(bottom)))

        # 3) Footer
        place_footer(inner_u < FOOTER_STACK_INNER)

        # 4) Wrap text to the room that is actually available (never shrink the text itself)
        _set_wrap(file_label, "file", col_w - 234 - 24)
        _set_wrap(live_status_banner, "banner", col_w - 154)
        _set_wrap(cleaner_desc, "desc", col_w - 8)
        _fit_desc_box(_layout["wrap"]["desc"])

        # 5) Keep the minimum size in sync with whatever the layout now needs
        update_min_size()

    def schedule_layout():
        """
        Debounced layout pass so dragging the window edge stays smooth. Right after the window lands on a
        monitor with a different DPI the pass is held back until things have settled, so the layout is
        computed once with the final scale instead of several times with a half-updated one.
        """
        global _resize_job
        if _resize_job is not None:
            try:
                app.after_cancel(_resize_job)
            except Exception:
                pass
        hold_ms = int((_layout["settle_until"] - time.monotonic()) * 1000)
        _resize_job = app.after(max(LAYOUT_DEBOUNCE_MS, hold_ms), apply_layout)

    def _on_app_configure(event):
        global _last_width
        if event.widget is not app or event.width == _last_width:
            return
        _last_width = event.width
        schedule_layout()

    app.bind("<Configure>", _on_app_configure, add="+")

    # ======================================================================
    # Moving the window to a monitor with a different DPI scale
    # CustomTkinter rescales its own widgets, but the gradient buttons, the header, the progress bar and
    # every px()-based padding are raw Tk pieces that would otherwise keep the old monitor's pixel sizes.
    # ======================================================================
    _pixel_widgets = [upload_btn, open_output_btn, open_folder_btn, *cleaner_buttons]

    def _apply_scaled_paddings():
        compact = _layout["compact"]
        header_banner.grid_configure(pady=(0, px(10 if compact else 16)))
        upload_btn.grid_configure(padx=px(14), pady=px(8 if compact else 14))
        open_output_btn.grid_configure(padx=(0, px(6)))
        open_folder_btn.grid_configure(padx=(px(6), 0))
        progress.grid_configure(padx=px(14), pady=(0, px(10)))
        hist_text.configure(font=("Segoe UI", -px(11)))
        _layout["cols"] = None  # rebuilds the cleaner grid (and its px() gaps) on the next layout pass

    def _refit_after_move(step=0):
        """
        Re-fits the layout after the window lands on another monitor. Run a few times on purpose: right after
        the move CustomTkinter is still rescaling its widgets, so a single pass can measure half-updated sizes
        and leave the log or footer clipped (which a manual full-screen toggle used to fix).
        """
        _layout["refit_job"] = None
        try:
            _layout["mode"] = None       # grid minsizes are px() values of the old scale: rebuild them
            _layout["cols"] = None
            _layout["desc_key"] = None
            _layout["wrap"].clear()
            if _layout["compact"]:       # the new monitor may have room again; update_min_size re-compacts if not
                set_density(False)
            apply_layout()

            app.update_idletasks()
            scale = widget_scale(app)
            win_bottom = app.winfo_rooty() + app.winfo_height()
            footer_bottom = footer_frame.winfo_rooty() + footer_frame.winfo_height()
            if footer_bottom > win_bottom + 1 or status_box.winfo_height() < px(60):
                if not _layout["compact"]:
                    set_density(True)
                height = int(min(_need_height_u(), _avail_height_u()))
                app.geometry(f"{int(app.winfo_width() / scale)}x{height}")
            if step < 2:
                _layout["refit_job"] = app.after((900, 1800)[step], lambda: _refit_after_move(step + 1))
        except Exception:
            pass

    def _schedule_refit():
        job = _layout.get("refit_job")
        if job is not None:
            try:
                app.after_cancel(job)
            except Exception:
                pass
        _layout["refit_job"] = app.after(350, _refit_after_move)

    def _check_scale():
        """Cheap 120 ms watcher: when the DPI scale changes, refresh the raw-canvas pieces once, then re-layout."""
        _layout["scale_job"] = None
        try:
            scale = widget_scale(app)
            known = _layout["scale"]
            if known is None:
                _layout["scale"] = scale
            elif abs(scale - known) > 0.001:
                _layout["scale"] = scale
                _layout["settle_until"] = time.monotonic() + SCALE_SETTLE_MS / 1000.0
                for w in _pixel_widgets:
                    w.rescale()
                progress.rescale()
                _apply_scaled_paddings()
                header_banner.refresh()
                schedule_layout()
                _schedule_refit()
            # Same DPI scale but a different monitor (e.g. a bigger one): refit as well.
            work = _monitor_work_height_px()
            if work is not None:
                if _layout.get("work_h") not in (None, work):
                    _schedule_refit()
                _layout["work_h"] = work
        except Exception:
            pass
        try:
            _layout["scale_job"] = app.after(120, _check_scale)
        except Exception:
            pass

    # ======================================================================
    # Guaranteed fit on spawn
    # ======================================================================
    def fit_and_place():
        """
        Picks the first layout that fully fits this screen, then sizes + centers the window:
          1) single column  ->  2) two panes (log on the right)  ->  3) two panes, compact
          ->  4) single column, compact
        """
        app.update_idletasks()
        scale = widget_scale(app)
        sw_u = app.winfo_screenwidth() / scale
        sh_u = app.winfo_screenheight() / scale
        avail = _avail_height_u()

        single_w = 760
        dual_w = int(min(sw_u * 0.82, 1240))
        _layout["dual_w"] = dual_w if dual_w >= DUAL_BREAKPOINT + 20 else None
        candidates = [(single_w, False)]
        if _layout["dual_w"]:
            candidates += [(dual_w, False), (dual_w, True)]
        candidates.append((single_w, True))

        _layout["auto_compact"] = False
        chosen, need = candidates[-1], None
        for width, compact in candidates:
            set_density(compact)
            apply_layout(width_u=width)
            need = _need_height_u()
            chosen = (width, compact)
            if need <= avail:
                break
        _layout["auto_compact"] = True

        width, _ = chosen
        roomy = max(need + 80, sh_u * 0.80) if width == single_w else need + 24
        height = int(max(min(avail, roomy, 880), min(need, avail)))
        x = int((app.winfo_screenwidth() - width * scale) / 2)
        y = int((app.winfo_screenheight() - height * scale - 40) / 2)
        app.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")
        app.after(350, verify_fit)

    def verify_fit(attempt=0):
        """After the window is on screen: confirm the footer and log are visible, escalating if they are not."""
        try:
            app.update_idletasks()
            scale = widget_scale(app)
            win_bottom = app.winfo_rooty() + app.winfo_height()
            footer_bottom = footer_frame.winfo_rooty() + footer_frame.winfo_height()
            clipped = footer_bottom > win_bottom + 1 or status_box.winfo_height() < px(60)

            screen_h = app.winfo_screenheight()
            if win_bottom > screen_h - 8:  # window hangs off the bottom of the screen: nudge it up
                app.geometry(f"+{max(app.winfo_x(), 0)}+{max(0, screen_h - app.winfo_height() - 56)}")

            if clipped and attempt < 3:
                if attempt == 0 and not _layout["compact"]:
                    set_density(True)
                    update_min_size()
                elif attempt <= 1 and _layout["mode"] == "single" and _layout["dual_w"]:
                    app.geometry(f"{_layout['dual_w']}x{int(app.winfo_height() / scale)}")
                else:
                    status_box.configure(height=60)
                app.after(350, lambda: verify_fit(attempt + 1))
        except Exception:
            pass

    # ======================================================================
    # Keyboard shortcuts + close handling
    # ======================================================================
    app.bind("<Control-o>", upload_file)
    app.bind("<Control-O>", upload_file)
    app.bind("<Control-l>", clear_log)
    app.bind("<Control-L>", clear_log)
    app.bind("<Escape>", request_cancel)

    def _cleaner_shortcut(index):
        def handler(_event=None):
            if getattr(app, "is_processing", False):
                return
            if not getattr(app, "selected_path", None):
                log_message("Warning: Please upload a retail file before running the cleaner.")
                return
            start_cleaner(CLEANERS[index])
        return handler

    for _i in range(len(CLEANERS)):
        app.bind(f"<Key-{_i + 1}>", _cleaner_shortcut(_i))

    def _save_window_state():
        """Remembers the window's size (DPI-independent units) and position for the next launch."""
        try:
            if app.state() != "normal":
                return
            m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", app.geometry())
            if not m:
                return
            CONFIG["window"] = {"w": int(m.group(1)), "h": int(m.group(2)),
                                "x": int(m.group(3)), "y": int(m.group(4))}
            save_config(CONFIG)
        except Exception:
            pass

    def restore_window_state():
        """Re-applies the saved size and position, unless that would put the window off-screen
        (e.g. a monitor that is no longer connected); then the automatic fit-and-center stands."""
        saved = CONFIG.get("window")
        if not isinstance(saved, dict):
            return
        try:
            w, h, x, y = (int(saved[k]) for k in ("w", "h", "x", "y"))
        except Exception:
            return
        app.update_idletasks()
        scale = widget_scale(app)
        vx, vy = app.winfo_vrootx(), app.winfo_vrooty()
        vw, vh = app.winfo_vrootwidth(), app.winfo_vrootheight()
        pw = w * scale
        visible_w = min(x + pw, vx + vw) - max(x, vx)
        if visible_w < 160 or not (vy - 10 <= y <= vy + vh - 80):
            return
        min_w, min_h = tk.Wm.wm_minsize(app)
        w = int(max(min(w, (vw - 40) / scale), min_w / scale))
        h = int(max(min(h, (vh - 80) / scale), min_h / scale))
        app.geometry(f"{w}x{h}+{x}+{y}")
        app.update_idletasks()

    def on_close():
        """Never leave a worker process running after the window closes."""
        _save_window_state()
        try:
            if _layout["scale_job"] is not None:
                app.after_cancel(_layout["scale_job"])
            _stop_worker()
        finally:
            app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_close)

    # ======================================================================
    # Start
    # ======================================================================
    def _reveal_app():
        """One-step handover: the finished, fully laid-out window appears and the splash goes away."""
        try:
            app.attributes("-alpha", 1.0)
            app.lift()
        except Exception:
            pass
        if splash is not None:
            splash.close()

    _splash_step("Finishing up…", 0.97)
    refresh_steps()
    set_progress("idle")
    try:
        fit_and_place()
    except Exception:
        app.geometry("1100x700")
    try:
        restore_window_state()
    except Exception:
        pass
    _layout["scale"] = widget_scale(app)
    _layout["scale_job"] = app.after(120, _check_scale)
    schedule_layout()

    # Settle everything while the window is still invisible: realise the final geometry, run the
    # responsive layout pass and the clipping check now (instead of 40 ms / 350 ms after it is on screen),
    # then reveal. This is what removes the visible re-flow.
    try:
        app.update()
        apply_layout()
        verify_fit()
        app.update_idletasks()
    except Exception:
        pass
    _splash_step("Ready", 1.0)
    _reveal_app()

    # Optional diagnostics: set STRAT_PERF_LOG=1 to log every moment the UI loop was blocked for 80 ms or more.
    if os.environ.get("STRAT_PERF_LOG"):
        _beat = {"t": time.perf_counter()}

        def _heartbeat():
            now = time.perf_counter()
            late_ms = (now - _beat["t"]) * 1000.0 - 50.0
            if late_ms > 80:
                log_message(f"Perf: UI loop was blocked for ~{late_ms:.0f} ms")
            _beat["t"] = now
            try:
                app.after(50, _heartbeat)
            except Exception:
                pass

        app.after(50, _heartbeat)

    app.mainloop()