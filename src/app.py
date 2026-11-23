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
import math
import csv
import shutil
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk

try:
    import winsound  # Windows only: system sounds for the completion alert
except ImportError:
    winsound = None

if TYPE_CHECKING:  # never runs; lets PyInstaller discover the cleaner modules that the worker imports by name
    import cadence_cleaner, brand_performance_cleaner, existing_db_cleaner, current_sales_cleaner, database_cleaner, meta_ads_cleaner  # noqa: F401

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
META_ADS_COLOR = "#12303F"
META_ADS_HOVER = "#1A435A"
META_ADS_ACCENT = "#38BDF8"

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
LOG_STAMP = "#6E7F99"          # timestamps in the activity log
LOG_FRESH_BG = "#16294D"       # background a new log line fades out of

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
FIXED_WIDTH = 1020            # the window's fixed width (unscaled): two panes, fits a 1366 px laptop even at 125 % scale
FADE_IN_MS = 220              # fade-in of the real window after the loading screen (0 = appear instantly)
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
APP_VERSION = "v3.14.22"
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


def lock_fixed_scale():
    """One UI scale for the whole run, on every monitor. CustomTkinter's per-monitor DPI tracking is switched off
    (it is what re-scaled every widget, and flashed the window, each time it met a different monitor); the app
    uses the system scale once and keeps it. Windows itself scales the finished window if it is shown on a
    monitor with another DPI, so the window is the same size and layout everywhere and nothing is re-laid out."""
    ctk.deactivate_automatic_dpi_awareness()
    scale = 1.0
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)      # system-DPI aware (a no-op error if already set)
        except Exception:
            pass
        try:
            scale = ctypes.windll.user32.GetDpiForSystem() / 96.0
        except Exception:
            scale = 1.0
    scale = max(1.0, min(scale, 3.0))
    ctk.set_widget_scaling(scale)
    ctk.set_window_scaling(scale)
    return scale


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
# Motion: one shared clock for every short UI animation
#   * a single ~30 fps timer that only runs while something is animating
#   * time comes from the clock (a slow frame catches up instead of slowing the motion down)
#   * while the window is minimized every animation jumps to its end state
#   * before the first reveal (`_motion["on"]` is False) changes apply instantly, so nothing plays at startup
# ======================================================================
_motion = {"on": False}
_anim = {"items": {}, "job": None}
ANIM_FRAME_MS = 33


def _ease_out(t):
    return 1.0 - (1.0 - t) ** 3


def _linear(t):
    return t


def animate(key, duration_ms, step, done=None, ease=_ease_out):
    """Runs step(eased_t) from 0 to 1 on the shared clock. Starting a key again replaces the running
    animation with that key. With motion off, or a zero duration, it jumps straight to the end state."""
    _anim["items"].pop(key, None)
    if not _motion["on"] or duration_ms <= 0:
        step(1.0)
        if done:
            done()
        return
    _anim["items"][key] = {"t0": time.perf_counter(), "dur": duration_ms / 1000.0,
                           "step": step, "done": done, "ease": ease}
    if _anim["job"] is None:
        _anim["job"] = app.after(ANIM_FRAME_MS, _anim_tick)


def animate_cancel(key):
    _anim["items"].pop(key, None)


def _anim_tick():
    _anim["job"] = None
    try:
        minimized = app.state() == "iconic"
    except Exception:
        minimized = False
    now = time.perf_counter()
    for key, item in list(_anim["items"].items()):
        if _anim["items"].get(key) is not item:   # replaced or cancelled by an earlier step this frame
            continue
        t = 1.0 if minimized else min(1.0, (now - item["t0"]) / item["dur"])
        try:
            item["step"](item["ease"](t))
            finished = t >= 1.0
        except Exception:                          # e.g. the widget is gone: finish the animation quietly
            finished = True
        if finished:
            if _anim["items"].get(key) is item:
                del _anim["items"][key]
            if item["done"]:
                try:
                    item["done"]()
                except Exception:
                    pass
    if _anim["items"]:
        try:
            _anim["job"] = app.after(ANIM_FRAME_MS, _anim_tick)
        except Exception:
            pass


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
        self._gb_lvl = 0.0            # hover amount 0..1 (animated, quantised to steps of 0.2 so redraws stay cached)
        self._gb_key = ("gb", id(self))
        self._gb_cap = ""             # optional keycap hint in the corner (e.g. "1")
        self._gb_cap_items = None
        self._gb_cap_key = None
        self._gb_cap_ok = True

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
        lvl = 1.0 if self._gb_pressed else self._gb_lvl   # press is instant; hover eases in and out
        base = _mix(self._gb_fg, self._gb_hover_color, lvl)
        flat = self._gb_fg == self._gb_hover_color
        if flat:
            stops = [base, base]
        else:
            if self._gb_gradient_to:
                end = _mix(self._gb_gradient_to, "#FFFFFF", 0.12 * lvl)
            elif self._gb_border:
                end = _mix(base, self._gb_border, 0.40 + 0.15 * lvl)
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
        if name == "_gb_hover":
            self._tween_hover()
        else:
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
        self._draw_keycap(w, h, st, scale)

    # ---- motion + keycap ----
    def _tween_hover(self, duration=120):
        """Eases the hover amount toward 1 (pointer over) or 0 (pointer away)."""
        target = 1.0 if self._gb_hover else 0.0
        start = self._gb_lvl
        if start == target:
            animate_cancel(self._gb_key)
            self._schedule(immediate=True)
            return

        def step(t):
            lvl = round((start + (target - start) * t) * 5) / 5.0
            if lvl != self._gb_lvl:
                self._gb_lvl = lvl
                self._schedule(immediate=True)

        animate(self._gb_key, duration, step)

    def pulse(self):
        """A brief brighten that fades back; used when a button becomes available."""
        if self._gb_state != "normal" or not _motion["on"]:
            return
        self._gb_lvl = 1.0
        self._schedule(immediate=True)
        self._tween_hover(duration=380)

    def set_keycap(self, text):
        """Shows a small key hint (e.g. the 1-6 shortcut) in the button's right corner."""
        self._gb_cap = str(text)
        self._gb_cap_key = None
        self._schedule(immediate=True)

    def _draw_keycap(self, w, h, st, scale):
        if not self._gb_cap:
            return
        side = max(int(round(18 * scale)), 12)
        pad = max(int(round(10 * scale)), 6)
        x1 = w - pad
        x0 = x1 - side
        y0 = (h - side) // 2
        y1 = y0 + side
        base = st["stops"][0]
        fill = _mix(base, "#000000", 0.28)
        outline = _mix(base, "#FFFFFF", 0.22)
        ink = _mix(st["text"], base, 0.30)
        font = (self._gb_font[0], -max(int(round(side * 0.62)), 8), "bold")

        key = (w, self._gb_text, self._gb_font, scale)
        if key != self._gb_cap_key:   # hide the hint when the label is so long that it would collide with it
            self._gb_cap_key = key
            try:
                text_w = tkfont.Font(root=self, font=self._scaled_font()).measure(self._gb_text)
                self._gb_cap_ok = text_w <= w - 2 * (side + 2 * pad)
            except Exception:
                self._gb_cap_ok = True

        if self._gb_cap_items is None:
            rect = self.create_rectangle(x0, y0, x1, y1, fill=fill, outline=outline)
            txt = self.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=self._gb_cap, fill=ink, font=font)
            self._gb_cap_items = (rect, txt)
        else:
            rect, txt = self._gb_cap_items
            self.coords(rect, x0, y0, x1, y1)
            self.itemconfigure(rect, fill=fill, outline=outline)
            self.coords(txt, (x0 + x1) / 2, (y0 + y1) / 2)
            self.itemconfigure(txt, text=self._gb_cap, fill=ink, font=font)
        visible = "normal" if self._gb_cap_ok else "hidden"
        self.itemconfigure(rect, state=visible)
        self.itemconfigure(txt, state=visible)


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
        self._pb_key = ("progress", id(self))
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
        animate_cancel(self._pb_key)
        self._pb_value = max(0.0, min(1.0, float(value)))
        self._draw()

    def animate_to(self, value, ms=320, from_value=None):
        """Eases the determinate fill to `value` (used for the finish sweep)."""
        value = max(0.0, min(1.0, float(value)))
        start = self._pb_value if from_value is None else max(0.0, min(1.0, float(from_value)))
        self._pb_value = start

        def step(t):
            self._pb_value = start + (value - start) * t
            self._draw()

        animate(self._pb_key, ms, step)

    def start(self):
        if self._pb_running:
            return
        animate_cancel(self._pb_key)
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
            phase += (phase * phase * (3.0 - 2.0 * phase) - phase) * 0.7   # ease in/out: lingers off-screen, glides through
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
# Output settings (saved in the config file, so they apply to every future run, not just this session)
# ======================================================================
OUTPUT_TOKENS = ("name", "source", "cleaner", "date", "time")
DEFAULT_OUTPUT = {"mode": "default", "folder": "", "template": "{name}"}
_BAD_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def get_output_settings():
    """The saved output settings, with anything missing or invalid replaced by the defaults."""
    raw = CONFIG.get("output") if isinstance(CONFIG.get("output"), dict) else {}
    out = dict(DEFAULT_OUTPUT)
    if raw.get("mode") == "custom" and str(raw.get("folder") or "").strip():
        out["mode"], out["folder"] = "custom", str(raw["folder"]).strip()
    template = str(raw.get("template") or "").strip()
    if template:
        out["template"] = template
    return out


def output_is_customised(cfg=None):
    cfg = cfg or get_output_settings()
    return cfg["mode"] == "custom" or cfg["template"] != DEFAULT_OUTPUT["template"]


def unknown_tokens(template):
    return sorted({t for t in re.findall(r"\{(\w*)\}", template) if t not in OUTPUT_TOKENS})


def render_output_name(template, name, source, cleaner, now=None):
    """Fills the tokens of a file-name template and makes the result a legal Windows file name (no extension)."""
    now = now or datetime.datetime.now()
    values = {"name": name, "source": source, "cleaner": cleaner,
              "date": now.strftime("%Y-%m-%d"), "time": now.strftime("%H%M")}
    text = re.sub(r"\{(\w*)\}", lambda m: values.get(m.group(1), m.group(0)), template or "{name}")
    text = _BAD_NAME_CHARS.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip().rstrip(". ")     # a trailing dot or space is not allowed on Windows
    return text[:150].rstrip(". ") or _BAD_NAME_CHARS.sub("_", name)


def _unique_path(path):
    """The path itself if free, otherwise 'name (2).ext', 'name (3).ext'... so nothing is ever overwritten."""
    path = Path(path)
    if not path.exists():
        return path
    for n in range(2, 1000):
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
    return path.with_name(f"{path.stem} ({int(time.time())}){path.suffix}")


def apply_output_settings(output_path, cleaner_label, source_path=None):
    """Moves / renames a freshly cleaned file according to the saved output settings.
    Returns the final path. With the default settings (or on any problem) the file stays where it is."""
    cfg = get_output_settings()
    if not output_is_customised(cfg):
        return output_path
    try:
        src = Path(output_path)
        if not src.is_file():
            return output_path
        stem = render_output_name(cfg["template"], src.stem, Path(source_path).stem if source_path else "", cleaner_label)
        if src.suffix and stem.lower().endswith(src.suffix.lower()):    # the template already ends in ".xlsx"
            stem = stem[:-len(src.suffix)].rstrip(". ") or src.stem
        dest_dir = Path(cfg["folder"]) if cfg["mode"] == "custom" else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / (stem + src.suffix)
        try:
            if dest.exists() and os.path.samefile(dest, src):
                return output_path                                      # already named and placed as requested
        except OSError:
            pass
        dest = _unique_path(dest)
        shutil.move(str(src), str(dest))
        return str(dest)
    except Exception as e:
        log_message(f"Warning: Output settings could not be applied ({e}). "
                    "The cleaned file stays where the cleaner saved it.")
        return output_path


# ======================================================================
# Activity-log search: highlights every match, Enter / Shift+Enter jump between them
# ======================================================================
_log_search = {"entry": None, "count": None, "query": "", "matches": [], "cur": -1}
LOG_HIT_BG = "#4B3A0E"
LOG_CUR_BG = "#F59E0B"
LOG_MAX_HITS = 2000


def _set_search_count(text, color=MUTED_TEXT):
    lbl = _log_search["count"]
    if lbl is not None:
        try:
            lbl.configure(text=text, text_color=color)
        except Exception:
            pass


def _goto_log_match(i, scroll=True):
    matches = _log_search["matches"]
    if not matches:
        return
    i %= len(matches)
    _log_search["cur"] = i
    tb = status_box._textbox
    tb.tag_remove("searchcur", "1.0", "end")
    start, end = matches[i]
    tb.tag_add("searchcur", start, end)
    if scroll:
        tb.see(start)
    _set_search_count(f"{i + 1}/{len(matches)}", WHITE)


def _refresh_log_search(jump=False):
    """Re-finds the matches of the search text. Only visible lines count (the level filter is respected),
    and the [time] stamps are skipped so that searching 'AM' or '5' does not light up every line."""
    if _log_search["entry"] is None:
        return
    try:
        tb = status_box._textbox
        query = _log_search["query"]
        tb.tag_remove("searchhit", "1.0", "end")
        tb.tag_remove("searchcur", "1.0", "end")
        matches = []
        if query:
            length = tk.IntVar(master=tb)
            idx = "1.0"
            while len(matches) < LOG_MAX_HITS:
                pos = tb.search(query, idx, stopindex="end", nocase=True, count=length)
                if not pos or length.get() <= 0:
                    break
                if "stamp" in tb.tag_names(pos):
                    idx = f"{pos}+1c"
                    continue
                end = f"{pos}+{length.get()}c"
                tb.tag_add("searchhit", pos, end)
                matches.append((pos, end))
                idx = end
        prev = _log_search["cur"]
        _log_search["matches"] = matches
        if not query:
            _log_search["cur"] = -1
            _set_search_count("")
        elif not matches:
            _log_search["cur"] = -1
            _set_search_count("No match", WARNING_COLOR)
        else:
            cur = len(matches) - 1 if (jump or prev < 0) else min(prev, len(matches) - 1)
            _goto_log_match(cur, scroll=jump)
    except Exception:
        pass


def _on_log_search_key(_event=None):
    entry = _log_search["entry"]
    if entry is None:
        return
    query = entry.get()
    if query != _log_search["query"]:
        _log_search["query"] = query
        _refresh_log_search(jump=True)


def _log_search_step(delta):
    def handler(_event=None):
        if _log_search["matches"]:
            _goto_log_match(_log_search["cur"] + delta)
        return "break"
    return handler


def clear_log_search(_event=None):
    entry = _log_search["entry"]
    if entry is not None:
        entry.delete(0, "end")
        _log_search["query"] = ""
        _refresh_log_search()
        app.focus_set()
    return "break"


def focus_log_search(_event=None):
    entry = _log_search["entry"]
    if entry is not None:
        entry.focus_set()
        try:
            entry.select_range(0, "end")
        except Exception:
            pass
    return "break"


# ======================================================================
# File preview: the first rows of the loaded file, read without pandas (so the UI process stays light)
# ======================================================================
PREVIEW_ROWS = 20
PREVIEW_COLS = 40
PREVIEW_CELL_CHARS = 60
PREVIEW_SCAN_ROWS = 500       # how far down to look for non-empty rows


def _preview_cell(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d") if not (v.hour or v.minute or v.second) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else format(v, ".10g")
    s = str(v).replace("\r", " ").replace("\n", " ").strip()
    return s if len(s) <= PREVIEW_CELL_CHARS else s[:PREVIEW_CELL_CHARS - 1] + "…"


def _col_letter(j):
    s = ""
    j += 1
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


def _read_preview(path, sheet=None):
    """Returns {"rows": [(row number, [cells])], "ncols", "shown_cols", "sheets", "sheet", "total_text", "error"}.
    The first row in 'rows' is the header. Runs in a thread; never raises."""
    ext = Path(path).suffix.lower()
    out = {"rows": [], "ncols": 0, "shown_cols": 0, "sheets": [], "sheet": None, "total_text": None, "error": None}
    raw = []
    try:
        if ext in (".csv", ".txt"):
            with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
                sample = f.read(8192)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except csv.Error:
                    dialect = csv.excel
                for n, row in enumerate(csv.reader(f, dialect), start=1):
                    if any(str(c).strip() for c in row):
                        raw.append((n, row))
                    if len(raw) > PREVIEW_ROWS or n >= PREVIEW_SCAN_ROWS:
                        break
            out["total_text"] = _count_rows(path)
        elif ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                out["sheets"] = list(wb.sheetnames)
                name = sheet if sheet in out["sheets"] else wb.active.title
                out["sheet"] = name
                ws = wb[name]
                for n, row in enumerate(ws.iter_rows(max_row=PREVIEW_SCAN_ROWS, values_only=True), start=1):
                    if any(v is not None and str(v).strip() for v in row):
                        raw.append((n, list(row)))
                    if len(raw) > PREVIEW_ROWS:
                        break
                total = ws.max_row
                if total and total < 1048576:
                    out["total_text"] = f"{max(total - 1, 0):,} rows"
            finally:
                wb.close()
        elif ext == ".xls":
            out["error"] = "Preview is not available for legacy .xls files. Open the file in Excel, or save it as .xlsx first."
            return out
        else:
            out["error"] = "This file type cannot be previewed."
            return out
    except Exception as e:
        out["error"] = f"The file could not be read for the preview ({e})."
        return out

    if not raw:
        out["error"] = "No data was found in the first part of this file."
        return out
    width = max(len(r) for _, r in raw)
    cells = [[_preview_cell(c) for c in r] + [""] * (width - len(r)) for _, r in raw]
    last = max((j for row in cells for j, c in enumerate(row) if c), default=-1)
    out["ncols"] = last + 1
    out["shown_cols"] = min(last + 1, PREVIEW_COLS)
    out["rows"] = [(raw[i][0], cells[i][:out["shown_cols"]]) for i in range(len(raw))]
    return out


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
    _refresh_log_search()
    _refresh_warn_badge()


def _refresh_warn_badge():
    """Shows the number of warnings / errors inside the 'Warnings' and 'Errors' filters themselves (no separate badge).
    An error counts once per failure, not once per line of its traceback."""
    for tag, name, icon, accent in (("warnline", "Warnings", "⚠", WARNING_COLOR), ("errline", "Errors", "✗", ERROR_COLOR)):
        try:
            count = len(status_box._textbox.tag_ranges(tag)) // 2
            btn = log_filter._buttons_dict[name]
            if count:
                colour = WHITE if _log_state["filter"] == name else accent
                btn.configure(text=f"{icon} {name} {count}", text_color=colour)
            else:
                btn.configure(text=name, text_color=WHITE)
        except Exception:
            pass


# ======================================================================
# Toast: a small message pill that fades in, waits and fades out (one at a time)
# ======================================================================
TOAST_KINDS = {"info": INFO_COLOR, "success": SUCCESS_COLOR, "warning": WARNING_COLOR, "error": ERROR_COLOR}
_toast = {"label": None}


def show_toast(text, kind="info"):
    """Brief confirmation at the bottom-right of the window (instead of adding a line to the log)."""
    try:
        accent = TOAST_KINDS.get(kind, INFO_COLOR)
        lbl = _toast["label"]
        if lbl is None:
            lbl = tk.Label(app, bd=0, padx=px(14), pady=px(8), highlightthickness=1,
                           font=("Segoe UI", -px(12), "bold"))
            _toast["label"] = lbl
        lbl.configure(text=text)
        total, fade_in, fade_out = 2000.0, 160.0, 320.0

        def step(t):                       # t is linear over the toast's whole life
            ms = t * total
            if ms < fade_in:
                a = ms / fade_in
            elif ms > total - fade_out:
                a = max(0.0, (total - ms) / fade_out)
            else:
                a = 1.0
            lbl.configure(bg=_mix(NAVY_BLUE, SURFACE_BG, a), fg=_mix(NAVY_BLUE, WHITE, a),
                          highlightbackground=_mix(NAVY_BLUE, accent, a), highlightcolor=_mix(NAVY_BLUE, accent, a))
            rise = int((1.0 - min(1.0, ms / fade_in)) * px(8))      # slides up while it fades in
            lbl.place(relx=1.0, rely=1.0, x=-px(22), y=-px(50) + rise, anchor="se")

        step(0.0)
        lbl.lift()
        animate("toast", total, step, lbl.place_forget, ease=_linear)
    except Exception:
        pass


# ======================================================================
# Logging, status banner, window title, progress
# ======================================================================
def _pulse_log_line():
    """The newest log line starts with a soft highlight that fades out."""
    try:
        tb = status_box._textbox

        def step(t):
            tb.tag_configure("fresh", background=_mix(LOG_FRESH_BG, LOG_BG, t))

        step(0.0)
        animate("logfresh", 750, step, lambda: tb.tag_remove("fresh", "1.0", "end"))
    except Exception:
        pass


def log_message(message, level=None, continued=False):
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
            if tag == "warning":
                tags = (tag, "warnline")
            elif tag == "error" and not continued:
                tags = (tag, "errline")
            else:
                tags = (tag,)
            # The timestamp is muted; the line break carries only the level tag (not "warnline"), so two
            # warnings in a row stay two separate ranges and the warning badge counts them correctly.
            tb.insert("end", f"[{timestamp}] ", tags + ("stamp",))
            tb.insert("end", message, tags)
            tb.insert("end", "\n", (tag,))

            # Keep the log bounded so long sessions never slow the UI down.
            extra = int(tb.index("end-1c").split(".")[0]) - MAX_LOG_LINES
            if extra > 0:
                tb.delete("1.0", f"{extra + 1}.0")

            tb.tag_remove("fresh", "1.0", "end")
            tb.tag_add("fresh", "end-2c linestart", "end-1c")
            _pulse_log_line()
            _refresh_log_search()

            status_box.see("end")
            status_box.configure(state="disabled")
            _refresh_warn_badge()
            return
    except Exception:
        pass

    print(f"[Startup Log] {message}")


# --- Fixed-height text -------------------------------------------------------------------------------
# A label that re-wraps (a timer ticking in the banner, a long file name) changes its height, and everything
# below it jumps. These labels are instead cut to a fixed number of lines with an ellipsis, so their height
# never changes. The full text is kept and re-fitted whenever the width or the DPI scale changes.
_clamped = {}          # key -> (label, full text, point size, max lines, max lines per paragraph, extra configure args)
_clamped_shown = {}    # key -> text currently on the label
_measure_fonts = {}


def _measure_font(size, scale):
    key = (size, round(scale, 3))
    font = _measure_fonts.get(key)
    if font is None:
        font = tkfont.Font(root=app, family="Segoe UI", size=-int(round(size * scale)))
        _measure_fonts[key] = font
    return font


def _ellipsize(text, font, limit):
    text = text.rstrip()
    if font.measure(text) <= limit:
        return text
    while text and font.measure(text + "…") > limit:
        text = text[:-1]
    return text.rstrip() + "…"


def _clamp_lines(text, size, wrap_u, max_lines, para_lines=None):
    """Word-wraps `text` the way the label will and cuts it to `max_lines` lines (ending in an ellipsis)."""
    scale = widget_scale(app)
    font = _measure_font(size, scale)
    limit = max(wrap_u * scale - 16, 40)      # same safety margin as _fit_desc_box
    lines = []
    for para in text.split("\n"):
        if para_lines == 1:                    # one line only: no wrapping, just cut with an ellipsis
            lines.append(_ellipsize(para, font, limit))
            continue
        para_out, current = [], ""
        for word in para.split(" "):
            trial = f"{current} {word}" if current else word
            if current and font.measure(trial) > limit:
                para_out.append(current)
                current = word
            else:
                current = trial
        para_out.append(current)
        para_out = [_ellipsize(l, font, limit) for l in para_out]
        if para_lines is not None and len(para_out) > para_lines:
            para_out = para_out[:para_lines]
            para_out[-1] = _ellipsize(para_out[-1] + " …", font, limit)
        lines.extend(para_out)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = _ellipsize(lines[-1] + " …", font, limit)
    return "\n".join(lines)


def _render_clamped(key):
    label, text, size, max_lines, para_lines, extra = _clamped[key]
    layout = globals().get("_layout") or {}
    wrap_u = layout.get("wrap", {}).get(key)
    try:
        shown = _clamp_lines(text, size, wrap_u, max_lines, para_lines) if wrap_u else text
        if extra or _clamped_shown.get(key) != shown:
            _clamped_shown[key] = shown
            label.configure(text=shown, **extra)
    except Exception:
        label.configure(text=text, **extra)


def set_clamped_text(label, key, text, size, max_lines, para_lines=None, **extra):
    _clamped[key] = (label, text, size, max_lines, para_lines, extra)
    _render_clamped(key)


_banner_fx = {"text": MUTED_TEXT, "border": BORDER_COLOR,           # what is on screen right now
              "to_text": MUTED_TEXT, "to_border": BORDER_COLOR}    # where the current fade is heading


def set_banner(text, color):
    """Updates the live status banner; its colour and the card border fade to the new state."""
    set_clamped_text(live_status_banner, "banner", text, 11, 2)
    border = STATE_BORDERS.get(color, BORDER_COLOR)
    fx = _banner_fx
    if color == fx["to_text"] and border == fx["to_border"]:
        return
    fx["to_text"], fx["to_border"] = color, border
    from_text, from_border = fx["text"], fx["border"]

    def step(t):
        fx["text"] = _mix(from_text, color, t)
        fx["border"] = _mix(from_border, border, t)
        live_status_banner.configure(text_color=fx["text"])
        operation_card.configure(border_color=fx["border"])

    animate("banner", 240, step)


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
            progress.set(0)
            progress.animate_to(1.0)       # the bar sweeps full instead of snapping
        elif state == "error":
            progress.configure(progress_color=ERROR_COLOR)
            progress.set(0)
            progress.animate_to(1.0)
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


_step_goal = {}   # step index -> (badge bg, badge text, title colour) it is fading toward
_step_cur = {}    # ... and what is on screen right now


def _fade_step(i, badge, title, target, pop):
    """Crossfades one step badge/title to its new colours; badges that light up get a brief highlight."""
    if _step_goal.get(i) == target:
        return
    _step_goal[i] = target
    start = _step_cur.get(i, target)

    def step(t):
        bg = _mix(start[0], target[0], t)
        fg = _mix(start[1], target[1], t)
        tc = _mix(start[2], target[2], t)
        _step_cur[i] = (bg, fg, tc)
        shown = _mix(bg, "#FFFFFF", 0.30 * math.sin(math.pi * t)) if (pop and t < 1.0) else bg
        badge.configure(fg_color=shown, text_color=fg)
        title.configure(text_color=tc)

    animate(("step", i), 260, step)


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
        states = ("done", "active", "todo")
        if getattr(app, "retail_path", None):
            if getattr(app, "pair_ok", False):
                hint, hint_color = "Two files ready — run the Meta Ads Cleaner", INFO_COLOR
            elif getattr(app, "pair_msg", None):
                hint, hint_color = "The two files don't match — see the Activity Log", WARNING_COLOR
            else:
                hint, hint_color = "Checking the two files…", MUTED_TEXT
        else:
            hint, hint_color = "Pick the cleaner that fits your file", INFO_COLOR

    for i, ((badge, title, number), state) in enumerate(zip(step_widgets, states)):
        bg, fg, title_color = _STEP_STYLE[state]
        badge.configure(text="✓" if state == "done" else str(number))
        _fade_step(i, badge, title, (bg, fg, title_color), pop=(state != "todo"))
    step2_hint.configure(text=hint, text_color=hint_color)


animation_job = None
timer_job = None
dot_count = 0
start_time = 0
_banner_secs = -1
_stage_text = ""
_typical_text = ""   # "about 14 s" while a run is in progress and enough past runs are known


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
            typical = f"  ·  usually ~{_typical_text}" if _typical_text else ""
            set_banner(f"Running {base_text}{stage}  ·  {time_str}{typical}", INFO_COLOR)

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


def _worker_entry(module_name, func_name, file_path, result_q, extra_path=None):
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
            params = inspect.signature(cleaner_func).parameters
            if "progress" in params:
                kwargs["progress"] = lambda text: result_q.put(("progress", str(text)))
            if extra_path and "retail_path" in params:   # two-file cleaners (Meta Ads + retail sales)
                kwargs["retail_path"] = extra_path
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


_unlock_jobs = []


def _cleaner_available(c):
    """One file loaded -> the single-file cleaners. Two files loaded and recognised as Meta Ads + retail -> only the
    Meta Ads Cleaner."""
    has_first = bool(getattr(app, "selected_path", None))
    has_second = bool(getattr(app, "retail_path", None))
    if c.get("two_file"):
        return has_first and has_second and bool(getattr(app, "pair_ok", False))
    return has_first and not has_second


def _enable_cleaner(b, c, pulse=False):
    b.configure(state="normal", fg_color=c["color"], hover_color=c["hover"],
                border_color=c["accent"], text=c["label"])
    if pulse:
        b.pulse()


def set_cleaners_enabled(enabled, except_btn=None, stagger=False):
    """Enables/disables all cleaner buttons, restoring their own colors and labels when enabling.
    With stagger=True, buttons that were disabled wake up one after another with a brief glow."""
    for job in _unlock_jobs:
        try:
            app.after_cancel(job)
        except Exception:
            pass
    _unlock_jobs.clear()
    for i, c in enumerate(CLEANERS):
        b = c["btn"]
        if b is except_btn:
            continue
        if enabled and _cleaner_available(c):
            if stagger and _motion["on"] and b.cget("state") == "disabled":
                _unlock_jobs.append(app.after(i * 45, lambda b=b, c=c: _enable_cleaner(b, c, pulse=True)))
            else:
                _enable_cleaner(b, c)
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
    _reset_cancel_slot()
    if getattr(app, "selected_path", None):
        clear_btn.configure(state="normal")


def _clean_key(file_path, cleaner_label, extra_path=None):
    """Identity of one clean: the file NAME (case-insensitive), the second file's name for two-file cleaners,
    plus the cleaner used."""
    return (Path(file_path).name.casefold(), cleaner_label, Path(extra_path).name.casefold() if extra_path else "")


def already_cleaned(file_path, cleaner_label, extra_path=None):
    """True when this file (pair) was already cleaned successfully with this cleaner in this session."""
    return _clean_key(file_path, cleaner_label, extra_path) in getattr(app, "cleaned_log", set())


def warn_already_cleaned(cleaner_label):
    """Logs the duplicate-clean warning and mirrors it in the banner."""
    log_message(f"Warning: already cleaned this file with {cleaner_label}")
    set_banner(f"Already cleaned this file with {cleaner_label}", WARNING_COLOR)


def run_cleaner_process(cleaner_ref, button, base_text, default_color, hover_color, force=False, extra_path=None):
    """Runs a cleaner ((module_name, function_name)) in a worker process with live updates and real cancellation."""
    global start_time, dot_count, _banner_secs, _stage_text, _typical_text
    if getattr(app, 'is_processing', False):
        return

    if not getattr(app, 'selected_path', None):
        log_message("Warning: Please upload a retail file before running the cleaner.")
        return

    if not force and already_cleaned(app.selected_path, base_text, extra_path):
        warn_already_cleaned(base_text)
        return

    file_name = Path(app.selected_path).name
    app.current_run = (file_name, base_text)
    app.current_extra = extra_path
    if extra_path:
        log_message(f"Processing '{file_name}' + '{Path(extra_path).name}' using {base_text}...")
    else:
        log_message(f"Processing file on '{file_name}' using {base_text}...")

    app.current_size_mb = _input_size_mb([app.selected_path, extra_path])
    _est = estimate_seconds(base_text, app.current_size_mb)
    _typical_text = _approx_duration(_est) if _est else ""
    if _typical_text:
        log_message(f"Usually takes about {_typical_text} for a file this size.")

    app.is_processing = True
    app.cancel_requested = False
    app.run_done = False
    start_time = time.time()

    upload_btn.configure(state="disabled")
    clear_btn.configure(state="disabled")
    _reset_cancel_slot()
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
        proc = _ctx.Process(target=_worker_entry, args=(cleaner_ref[0], cleaner_ref[1], app.selected_path, result_q, extra_path), daemon=True)
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


def _hist_hover(tag_name, on):
    hist_text.tag_configure(tag_name, underline=on)
    hist_text.configure(cursor="hand2" if on else "arrow")


def _open_history_file(path):
    """Click on a finished run: show its cleaned file in Explorer / Finder."""
    if path and os.path.exists(path):
        try:
            _reveal_path(path)
            return
        except Exception:
            pass
    show_toast("That cleaned file is no longer there", "warning")


def _render_history():
    """Render newest-first, single-line entries; long entries are clipped, never wrapped."""
    t = hist_text
    t.configure(state="normal")
    t.delete("1.0", "end")
    for name in t.tag_names():
        if name.startswith("hist"):
            t.tag_configure(name, underline=False)
    t.configure(cursor="arrow")

    if not app.history:
        t.insert("end", "No runs yet — finished runs will be listed here.", "muted")
    else:
        scale = widget_scale(t)
        font = tkfont.Font(root=app, family="Segoe UI", size=-int(round(11 * scale)))
        # The scrollbar occupies its own column, so only the text widget's inner width is relevant.
        available = max(t.winfo_width() - px(26, t), px(40, t))
        for i, e in enumerate(app.history):
            icon, status_tag = _HIST_ICONS[e["status"]]
            clickable = e["status"] in ("ok", "warn") and bool(e.get("output"))
            if clickable:
                tag_name = f"hist{i}"
                t.tag_configure(tag_name, underline=False)
                t.tag_bind(tag_name, "<Button-1>", lambda ev, p=e["output"]: _open_history_file(p))
                t.tag_bind(tag_name, "<Enter>", lambda ev, n=tag_name: _hist_hover(n, True))
                t.tag_bind(tag_name, "<Leave>", lambda ev, n=tag_name: _hist_hover(n, False))
            else:
                tag_name = None

            if e["status"] in ("ok", "warn"):
                details = "  ·  ".join((e["time"], short_duration(e["elapsed"]),
                                        f"{int(e['rows'] or 0):,} rows", e["cleaner"],
                                        e.get("output_name") or "(output name unavailable)"))
            elif e["status"] == "err":
                details = "  ·  ".join((e["time"], "failed", e["cleaner"]))
            else:
                details = "  ·  ".join((e["time"], "cancelled", e["cleaner"]))
            details = _ellipsize(details, font, available - font.measure(icon + "  "))
            row_start = t.index("end-1c")
            t.insert("end", icon + "  ", status_tag)
            t.insert("end", details, "rest")
            if tag_name:
                t.tag_add(tag_name, row_start, "end-1c")
            if i < len(app.history) - 1:
                t.insert("end", "\n")

    t.configure(state="disabled")
    if len(app.history) > 3:
        hist_scroll.grid()
    else:
        hist_scroll.grid_remove()
    t.yview_moveto(0)


def add_history(status, file_name, cleaner, rows=None, elapsed=None, warns=0, output=None):
    """Records one run (newest first); only the cleaned output name is shown in history."""
    app.history.insert(0, {"status": status, "cleaner": cleaner, "rows": rows,
                           "elapsed": elapsed, "warns": warns, "output": output,
                           "output_name": Path(output).name if output else None,
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
    retail = stats.get("retail")
    if retail:
        log_message(f"Retail: {retail['rows_in']:,} rows read, {retail['retail_rows_out']:,} Retail Data rows, "
                    f"{retail['online_receipts']:,} distinct online transactions")
        if retail.get("interest_lines") or retail.get("zz_removed"):
            log_message(f"   Interest lines set aside: {retail.get('interest_lines', 0):,}  ·  "
                        f"ZZ paper-bag lines removed: {retail.get('zz_removed', 0):,}  (see Audit)")
        if retail.get("stores_masked"):
            masks = retail["stores_masked"]
            more = f" … (+{len(masks) - 6} more)" if len(masks) > 6 else ""
            log_message("   Retail store masks applied: " + "; ".join(masks[:6]) + more)
        if retail.get("online_returns_excluded"):
            log_message(f"   Online return lines not counted: {retail['online_returns_excluded']:,}")
        if retail.get("online_no_receipt"):
            warnings += 1
            log_message(f"Warning: {retail['online_no_receipt']:,} online line(s) had no receipt number and were not counted")
        if retail.get("unmapped_stores"):
            warnings += 1
            shown = ", ".join(retail["unmapped_stores"][:8]) + (f" … (+{len(retail['unmapped_stores']) - 8} more)" if len(retail["unmapped_stores"]) > 8 else "")
            log_message(f"Warning: {len(retail['unmapped_stores'])} retail store name(s) not in the recognized store list, kept as-is: {shown}")
        if retail.get("unmapped_brands"):
            warnings += 1
            log_message(f"Warning: {len(retail['unmapped_brands'])} retail brand(s) do not match a Meta Ads brand name, kept as-is: {', '.join(retail['unmapped_brands'][:8])}")
        if retail.get("other_tags"):
            warnings += 1
            log_message(f"Warning: retail TAG values other than Online/Offline found: {', '.join(retail['other_tags'][:8])}")
        if retail.get("interest_unmatched"):
            warnings += 1
            log_message(f"Warning: {retail['interest_unmatched']:,} retail interest line(s) had no matching sale line (see Audit)")
    if stats.get("interest_unapplied"):
        warnings += 1
        log_message(f"Warning: {stats['interest_unapplied']} interest line(s) had no sale line on their receipt (see the Removal Audit Log)")
    if stats.get("interest_mismatch"):
        warnings += 1
        log_message(f"Warning: {stats['interest_mismatch']} interest line(s) are marked MISMATCH in the Removal Audit Log")
    return warnings


def handle_success(row_count, output_path, elapsed_seconds, button, base_text, default_color, hover_color, stats=None):
    _finish_common()
    _final_path = apply_output_settings(output_path, base_text, getattr(app, "selected_path", None))
    relocated = _final_path != output_path
    output_path = _final_path
    app.run_done = True
    if getattr(app, "selected_path", None):
        app.cleaned_log.add(_clean_key(app.selected_path, base_text, getattr(app, "current_extra", None)))
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
    record_timing(base_text, getattr(app, "current_size_mb", None), elapsed_seconds)

    app.last_output_path = output_path
    log_message(f"Saved to: {output_path}")
    if relocated:
        log_message("Output settings applied to the cleaned file.")
    open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)
    preview_result_btn.configure(state="normal")
    refresh_steps()
    add_history("warn" if warn_count else "ok", _current_file_name(), base_text,
                rows=int(row_count), elapsed=elapsed_seconds, warns=warn_count, output=output_path)
    notify_done("warning" if warn_count else "success", elapsed_seconds)


# ======================================================================
# Typical run time: learned from this machine's own past runs (kept per cleaner in the config file)
# ======================================================================
TIMING_KEEP = 12


def _input_size_mb(paths):
    try:
        return sum(os.path.getsize(p) for p in paths if p) / (1024 * 1024)
    except Exception:
        return None


def estimate_seconds(label, size_mb):
    """Median time of the (up to) three past runs whose input size was closest to this one, or None when there are
    fewer than two runs on record for this cleaner."""
    if size_mb is None:
        return None
    runs = CONFIG.get("timings", {}).get(label, []) if isinstance(CONFIG.get("timings"), dict) else []
    good = []
    for r in runs:
        try:
            good.append((float(r[0]), float(r[1])))
        except Exception:
            continue
    if len(good) < 2:
        return None
    nearest = sorted(good, key=lambda r: abs(r[0] - size_mb))[:3]
    times = sorted(t for _, t in nearest)
    return times[len(times) // 2]


def record_timing(label, size_mb, seconds):
    """Remembers a successful run (cancelled and failed runs say nothing about how long a clean takes)."""
    if size_mb is None or not seconds or seconds <= 0:
        return
    try:
        timings = CONFIG.get("timings") if isinstance(CONFIG.get("timings"), dict) else {}
        runs = list(timings.get(label, []))
        runs.append([round(size_mb, 3), round(seconds, 2)])
        timings[label] = runs[-TIMING_KEEP:]
        CONFIG["timings"] = timings
        save_config(CONFIG)
    except Exception:
        pass


def _approx_duration(seconds):
    s = max(1, int(round(seconds)))
    return f"{s} s" if s < 60 else f"{s // 60} min {s % 60:02d} s"


# ======================================================================
# Failed runs: a plain-language summary, and a one-click error report (the idle Cancel button becomes "Copy error")
# ======================================================================
def explain_error(tb_text):
    """One plain sentence for the usual failures, from the last line of the traceback."""
    lines = [ln.strip() for ln in str(tb_text).strip().splitlines() if ln.strip()]
    last = lines[-1] if lines else ""
    name, _, detail = last.partition(":")
    name, detail = name.split(".")[-1].strip(), detail.strip()
    low = last.lower()
    if "permissionerror" in low:
        return "A file is locked or cannot be written. If the output or source file is open in Excel, close it and run again."
    if "memory" in low:
        return "The file was too large for the available memory."
    if name == "KeyError":
        return f"A column the cleaner expects was not found ({detail or 'unknown'}). This file may not match the cleaner you chose."
    if name in ("FileNotFoundError", "NotADirectoryError"):
        return "A file or folder could not be found. It may have been moved or renamed after it was loaded."
    if name in ("BadZipFile",) or "not a zip file" in low:
        return "The file looks damaged or is not a real .xlsx workbook."
    if name in ("ModuleNotFoundError", "ImportError"):
        return "A library or cleaner module is missing from this build of the app."
    if name in ("ValueError", "ParserError", "EmptyDataError", "UnicodeDecodeError") or "could not convert" in low:
        return "Some values are not in the format this cleaner expects. This file may not match the cleaner you chose."
    return f"The cleaner stopped with {name or 'an unexpected error'}." if last else "The cleaner stopped unexpectedly."


def _reset_cancel_slot():
    """Puts the Cancel button back (it doubles as 'Copy error' after a failed run)."""
    app.last_error = None
    try:
        cancel_btn.configure(text="Cancel", command=request_cancel)
        if not getattr(app, "is_processing", False):
            cancel_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    except Exception:
        pass


def _store_last_error(tb_text, cleaner_label, plain):
    names = [Path(p).name for p in (getattr(app, "selected_path", None), getattr(app, "current_extra", None)) if p]
    size = _input_size_mb([getattr(app, "selected_path", None), getattr(app, "current_extra", None)])
    app.last_error = "\n".join([
        f"Strat Data Cleaner {APP_VERSION} - error report",
        f"Time: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Cleaner: {cleaner_label}",
        "File: " + "  +  ".join(names) + (f"  ({size:.2f} MB)" if size is not None else ""),
        f"System: {sys.platform}, Python {sys.version.split()[0]}",
        f"Summary: {plain}",
        "",
        "Traceback:",
        str(tb_text).strip(),
    ])


def _show_copy_error_slot():
    try:
        cancel_btn.configure(text="Copy error", command=copy_last_error, state="normal",
                             fg_color=SURFACE_BG, hover_color="#334155")
    except Exception:
        pass


def copy_last_error():
    report = getattr(app, "last_error", None)
    if not report:
        return
    try:
        app.clipboard_clear()
        app.clipboard_append(report)
        show_toast("Error details copied to clipboard", "success")
    except Exception as e:
        log_message(f"Error copying the error details: {e}")


def handle_error(error_message, button, base_text, default_color, hover_color):
    _finish_common()
    set_progress("error")
    plain = explain_error(error_message)
    set_banner(f"Failed — {plain}", ERROR_COLOR)
    set_title("Failed")

    log_message("Error during cleaning process:", level="error")
    log_message(f"In short: {plain}", level="error", continued=True)
    for line in error_message.strip().split("\n"):
        log_message(line, level="error", continued=True)
    log_message("Tip: make sure the selected file matches the cleaner you chose, then try again. "
                "“Copy error” (next to the status bar) copies the full details.")
    _store_last_error(error_message, base_text, plain)
    _show_copy_error_slot()
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


# Header cells that identify the two kinds of file (kept in sync with meta_ads_cleaner.py).
_META_HEADER = {"date", "campaign name", "ad set name", "cost"}
_RETAIL_HEADER = {"store_name", "rcpt_date", "receipt_no", "brand", "tag"}


def _kind_of_rows(rows):
    for row in rows:
        cells = {str(v).strip().lower() for v in row if v is not None and str(v).strip()}
        if _META_HEADER <= cells:
            return "meta"
        if _RETAIL_HEADER <= cells:
            return "retail"
    return None


def _sniff_kind(path):
    """'meta', 'retail', 'other' (readable, neither) or 'unreadable', from the first rows only. Runs in a thread."""
    ext = Path(path).suffix.lower()
    try:
        if ext in (".csv", ".txt"):
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                rows = [re.split(r"[,\t;|]", f.readline().replace('"', "")) for _ in range(15)]
            return _kind_of_rows(rows) or "other"
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                for ws in wb.worksheets:
                    kind = _kind_of_rows(ws.iter_rows(min_row=1, max_row=15, values_only=True))
                    if kind:
                        return kind
            finally:
                wb.close()
            return "other"
    except Exception:
        return "unreadable"
    return "unreadable"          # legacy .xls and anything else: columns cannot be checked here


def _pair_assignment(paths, kinds):
    """(meta_path, retail_path, None) for a valid pair in either order, else (None, None, what is wrong)."""
    if sorted(kinds) == ["meta", "retail"]:
        i = kinds.index("meta")
        return paths[i], paths[1 - i], None
    if kinds[0] == kinds[1] and kinds[0] in ("meta", "retail"):
        what = "Meta Ads exports" if kinds[0] == "meta" else "retail sales files"
        return None, None, f"Both files look like {what}. Select one Meta Ads export and one retail sales file."
    problems = []
    for p, k in zip(paths, kinds):
        name = Path(p).name
        if k == "other":
            problems.append(f"“{name}” is neither a Meta Ads export (Date, Campaign name, Ad set name, Cost) "
                            "nor a retail sales file (STORE_NAME, RCPT_DATE, RECEIPT_NO, BRAND, TAG)")
        elif k == "unreadable":
            problems.append(f"“{name}” could not be read to check its columns (use .xlsx or .csv)")
    return None, None, "; ".join(problems) + ". The Meta Ads Cleaner stays locked."


def _short_name(path, limit=24):
    name = Path(path).name
    return name if len(name) <= limit else name[:limit // 2 - 1] + "…" + name[-(limit // 2):]


def _render_file_label(count_text, pending=False):
    name, size_text, modified = app.file_meta
    parts = []
    if getattr(app, "file_tag", None):
        parts.append(app.file_tag)
    if pending:
        parts.append("counting rows…")
    elif count_text:
        parts.append(count_text)
    if modified:
        parts.append(f"modified {modified}")
    text = f"{name}  •  {size_text}"
    if parts:
        text += "\n" + "  ·  ".join(parts)
    set_clamped_text(file_label, "file", text, 11, 2, para_lines=1,
                     text_color=WARNING_COLOR if getattr(app, "file_warn", False) else SUCCESS_COLOR)


def _start_file_info(paths):
    """Counts rows off the UI thread and fills in the file card when done (ignored if the files changed)."""
    token = object()
    app._info_token = token
    box = {"text": None, "done": False}

    def work():
        try:
            texts = [_count_rows(p) for p in paths]
            if len(paths) == 1:
                box["text"] = texts[0]
            else:
                box["text"] = " + ".join((t.split()[0] if t else "?") for t in texts) + " rows"
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


def _show_files(paths, tag=None, warn=False, count=True):
    """Fills the file card for one file, or for a pair (names side by side, total size, row counts)."""
    sizes = [os.path.getsize(p) for p in paths]
    total = sum(sizes)
    size_text = f"{total / (1024 * 1024):.2f} MB" if total >= 1024 * 1024 else f"{total / 1024:.1f} KB"
    if len(paths) == 1:
        name = Path(paths[0]).name
        try:
            modified = _format_modified(os.path.getmtime(paths[0]))
        except Exception:
            modified = ""
    else:
        name, modified = "  +  ".join(_short_name(p) for p in paths), ""
    app.file_meta = (name, size_text, modified)
    app.file_tag, app.file_warn = tag, warn
    can_count = count and all(Path(p).suffix.lower() in _COUNTABLE for p in paths)
    _render_file_label(None, pending=can_count)
    if can_count:
        _start_file_info(paths)
    else:
        app._info_token = None


def _start_file_check(paths):
    """Recognises the loaded file(s) by their columns, off the UI thread, then applies the result."""
    token = object()
    app._check_token = token
    box = {"kinds": None}

    def work():
        try:
            box["kinds"] = [_sniff_kind(p) for p in paths]
        except Exception:
            box["kinds"] = ["unreadable"] * len(paths)

    threading.Thread(target=work, daemon=True).start()

    def poll():
        if getattr(app, "_check_token", None) is not token:
            return
        if box["kinds"] is None:
            app.after(100, poll)
            return
        _apply_file_check(paths, box["kinds"])

    app.after(100, poll)


def _apply_file_check(paths, kinds):
    if len(paths) == 1:
        if kinds[0] == "meta":    # a heads-up only: the single-file cleaners stay available
            msg = (f"“{Path(paths[0]).name}” looks like a Meta Ads export. The Meta Ads Cleaner needs two files: "
                   "Browse again and Ctrl+click the Meta Ads export together with a retail sales file.")
            log_message(f"Warning: {msg}")
            set_banner("This looks like a Meta Ads export — select it together with a retail sales file.", WARNING_COLOR)
        return
    meta, retail, problem = _pair_assignment(paths, kinds)
    if problem:
        app.pair_ok, app.pair_msg = False, problem
        log_message(f"Warning: {problem}")
        set_banner("The two files do not match the Meta Ads Cleaner — see the Activity Log.", WARNING_COLOR)
        _show_files(paths, tag="⚠ files do not match", warn=True, count=False)
    else:
        app.selected_path, app.retail_path = meta, retail
        app.pair_ok, app.pair_msg = True, None
        log_message(f"Files loaded: Meta Ads = {Path(meta).name}  |  Retail = {Path(retail).name}")
        set_banner("Two files ready — run the Meta Ads Cleaner.", MUTED_TEXT)
        _show_files([meta, retail], tag="Meta Ads + retail")
        set_cleaners_enabled(True, stagger=True)
    refresh_steps()
    schedule_layout()


def process_selected_file(filenames):
    """Loads one file (single-file cleaners) or two files (Meta Ads + retail)."""
    paths = [filenames] if isinstance(filenames, str) else list(filenames)
    if not paths:
        return
    app.selected_path = paths[0]
    app.retail_path = paths[1] if len(paths) == 2 else None
    app.pair_ok, app.pair_msg = False, None
    short = Path(paths[0]).name
    _show_files(paths, tag="checking the two files…" if len(paths) == 2 else None, count=len(paths) == 1)
    _refresh_upload_border()
    app.run_done = False
    clear_btn.grid(row=0, column=2, padx=(5, 15))
    _close_preview()
    _reset_cancel_slot()
    preview_btn.grid(row=0, column=2, sticky="e")

    set_cleaners_enabled(True, stagger=True)   # one file: the single-file cleaners; two files: locked until checked

    CONFIG["last_dir"] = str(Path(paths[0]).parent)
    save_config(CONFIG)

    set_progress("idle")
    set_title(None)
    if len(paths) == 1:
        set_banner("File ready — choose a cleaner in step 2.", MUTED_TEXT)
        log_message(f"Retail file loaded: {short}")
    else:
        set_banner("Checking the two files…", MUTED_TEXT)
        log_message(f"Two files selected: {short}  +  {Path(paths[1]).name}")
    _start_file_check(paths)
    refresh_steps()
    schedule_layout()  # the file name may wrap onto a new line


def upload_file(_event=None):
    if getattr(app, 'is_processing', False):
        return
    initial = CONFIG.get("last_dir")
    filenames = filedialog.askopenfilenames(
        title="Select Retail File  (Ctrl+click two files for the Meta Ads Cleaner)",
        initialdir=initial if initial and os.path.isdir(initial) else None,
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if isinstance(filenames, str):          # some Tk builds return one string
        filenames = app.tk.splitlist(filenames)
    filenames = [f for f in filenames if f]
    if not filenames:
        log_message("Retail file selection cancelled.")
    elif len(filenames) > 2:
        log_message(f"Warning: {len(filenames)} files were selected, but the limit is two (one Meta Ads export and one "
                    "retail sales file). Nothing was loaded.")
        set_banner("Too many files — select one file, or two for the Meta Ads Cleaner.", WARNING_COLOR)
    else:
        process_selected_file(filenames)


def clear_file():
    app.selected_path = None
    app.retail_path = None
    app.pair_ok, app.pair_msg = False, None
    app.file_tag, app.file_warn = None, False
    app._check_token = None
    app._info_token = None
    app.run_done = False
    set_clamped_text(file_label, "file", FILE_PLACEHOLDER, 11, 2, text_color=MUTED_TEXT)
    clear_btn.grid_forget()
    preview_btn.grid_forget()
    _close_preview()
    _reset_cancel_slot()
    _refresh_upload_border()

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
        # Passed as ONE string with the path quoted: as a list, subprocess wraps the whole "/select,..." argument
        # in quotes, and Explorer then ignores it (it opens Documents) whenever the path contains a space.
        subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')
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
        show_toast("Activity log copied to clipboard", "success")
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
        show_toast(f"Log saved: {Path(path).name}", "success")
    except Exception as e:
        log_message(f"Error saving activity log: {e}")


def _show_dialog_settled(dialog, w_u, h_u):
    """Centres the dialog over the window at its final size and shows it in one step. The UI scale is the same on
    every monitor, so there is nothing to wait for or re-fit."""
    app.update_idletasks()
    s_ = widget_scale(app)
    x = app.winfo_x() + (app.winfo_width() - int(w_u * s_)) // 2
    y = app.winfo_y() + (app.winfo_height() - int(h_u * s_)) // 2
    try:
        dialog.geometry(f"{w_u}x{h_u}+{x}+{y}")
        dialog.update_idletasks()
    except Exception:
        pass
    try:
        dialog.deiconify()
        tk.Wm.wm_attributes(dialog, "-alpha", 1.0)
        dialog.lift()
        dialog.focus_force()
    except Exception:
        pass


def custom_askyesno(title, message, allow_remember=True):
    """Modal confirmation. Enter = Yes, Esc = No. Optionally offers 'don't ask again this session'."""
    if getattr(app, "_dialog_open", False):
        return False
    app._dialog_open = True
    dialog = ctk.CTkToplevel(app)
    dialog.withdraw()                                     # built invisibly, shown by _show_dialog_settled
    try:
        tk.Wm.wm_attributes(dialog, "-alpha", 0.0)
    except Exception:
        pass
    dialog.title(title)
    dialog_width, dialog_height = 400, (232 if allow_remember else 214)
    dialog.resizable(False, False)
    dialog.configure(fg_color=NAVY_BLUE)

    try:
        icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
        dialog.iconbitmap(str(icon_path))
    except Exception:
        pass

    dialog.transient(app)

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

    _show_dialog_settled(dialog, dialog_width, dialog_height)
    try:
        dialog.grab_set()                                 # needs a visible window, so only after it is shown
    except Exception:
        pass
    try:
        dialog.wait_window()
    finally:
        app._dialog_open = False
    if result[0] and allow_remember and remember.get():
        app.skip_confirm = True
        log_message("Confirmations turned off for this session.")
    return result[0]


# ======================================================================
# File preview window (read-only, first rows of the loaded file(s))
# ======================================================================
def _close_preview():
    win = getattr(app, "_preview_win", None)
    app._preview_win = None
    app._preview_key = None
    if win is not None:
        try:
            win.destroy()
        except Exception:
            pass


def _style_preview_table(win):
    style = ttk.Style(win)
    try:
        style.theme_use("clam")          # the native Windows theme ignores row / header colours
    except Exception:
        pass
    style.configure("Preview.Treeview", background=CARD_BG, fieldbackground=CARD_BG, foreground=LOG_TEXT,
                    rowheight=px(24), borderwidth=0, font=("Segoe UI", -px(11)))
    style.configure("Preview.Treeview.Heading", background=SURFACE_BG, foreground=WHITE, relief="flat",
                    borderwidth=0, font=("Segoe UI", -px(11), "bold"), padding=(px(6), px(5)))
    style.map("Preview.Treeview", background=[("selected", ACCENT_ORANGE)], foreground=[("selected", WHITE)])
    style.map("Preview.Treeview.Heading", background=[("active", "#334155")])


def open_result_preview(_event=None):
    """Preview of the cleaned output file of the last successful run."""
    path = getattr(app, "last_output_path", None)
    if not path or not os.path.exists(path):
        show_toast("No cleaned file to preview yet", "warning")
        return
    open_preview(paths=[path], title="Cleaned file preview")


def open_preview(_event=None, paths=None, title="File preview"):
    """Shows the first rows of the loaded file (or either file of a Meta Ads pair), or of the given file(s)."""
    if paths is None:
        paths = [p for p in (getattr(app, "selected_path", None), getattr(app, "retail_path", None)) if p]
    if not paths:
        show_toast("Select a file first to preview it", "warning")
        return
    existing = getattr(app, "_preview_win", None)
    if existing is not None:
        if getattr(app, "_preview_key", None) == tuple(paths):
            try:
                existing.deiconify()
                existing.lift()
                existing.focus_force()
                return
            except Exception:
                pass
        _close_preview()          # a different file was asked for: replace the window
    app._preview_key = tuple(paths)

    win = ctk.CTkToplevel(app)
    win.withdraw()
    try:
        tk.Wm.wm_attributes(win, "-alpha", 0.0)
    except Exception:
        pass
    win.title(title)
    win.configure(fg_color=NAVY_BLUE)
    win.minsize(640, 380)
    try:
        win.iconbitmap(str(resource_path(os.path.join("assets", "strat_data_cleaner.ico"))))
    except Exception:
        pass
    win.transient(app)
    app._preview_win = win
    _style_preview_table(win)

    st = {"file": paths[0], "sheet": None, "token": None}
    labels = {f"{i + 1}. {Path(p).name}": p for i, p in enumerate(paths)}

    # --- header: title + file / sheet pickers ---
    head = ctk.CTkFrame(win, fg_color="transparent")
    head.pack(fill="x", padx=18, pady=(16, 2))
    ctk.CTkLabel(head, text=title, font=("Segoe UI", 14, "bold"), text_color=WHITE).pack(side="left")

    def _menu(values, command):
        return ctk.CTkOptionMenu(head, values=values, command=command, font=("Segoe UI", 11), height=26, width=190,
                                 dynamic_resizing=False, fg_color=SURFACE_BG, button_color=SURFACE_BG,
                                 button_hover_color="#334155", dropdown_fg_color=CARD_BG,
                                 dropdown_hover_color=SURFACE_BG, text_color=WHITE, dropdown_text_color=WHITE,
                                 dropdown_font=("Segoe UI", 11), corner_radius=6)

    sheet_menu = _menu(["-"], lambda choice: _pick_sheet(choice))
    if len(paths) > 1:
        file_menu = _menu(list(labels), lambda choice: _pick_file(choice))
        file_menu.set(next(iter(labels)))
        file_menu.pack(side="right", padx=(8, 0))

    info = ctk.CTkLabel(win, text="Loading preview…", font=("Segoe UI", 11), text_color=MUTED_TEXT, anchor="w",
                        justify="left", wraplength=780)
    info.pack(fill="x", padx=18, pady=(2, 6))

    # --- table ---
    table_box = ctk.CTkFrame(win, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
    table_box.pack(fill="both", expand=True, padx=18, pady=(0, 8))
    table_box.grid_columnconfigure(0, weight=1)
    table_box.grid_rowconfigure(0, weight=1)
    tree = ttk.Treeview(table_box, style="Preview.Treeview", show="headings", selectmode="browse")
    vbar = ctk.CTkScrollbar(table_box, orientation="vertical", command=tree.yview, width=12)
    hbar = ctk.CTkScrollbar(table_box, orientation="horizontal", command=tree.xview, height=12)
    tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
    tree.grid(row=0, column=0, sticky="nsew", padx=(px(4), 0), pady=(px(4), 0))
    vbar.grid(row=0, column=1, sticky="ns", padx=px(2), pady=px(4))
    hbar.grid(row=1, column=0, sticky="ew", padx=px(4), pady=(0, px(2)))
    tree.tag_configure("odd", background="#0D1628")
    tree.tag_configure("even", background=CARD_BG)

    # --- footer ---
    foot = ctk.CTkFrame(win, fg_color="transparent")
    foot.pack(fill="x", padx=18, pady=(0, 16))
    ctk.CTkLabel(foot, text="Read-only. The first row is shown as the column header; “#” is the row number in the file.",
                 font=("Segoe UI", 10), text_color=MUTED_TEXT, anchor="w").pack(side="left")
    ctk.CTkButton(foot, text="Close", font=("Segoe UI", 11, "bold"), fg_color=SURFACE_BG, hover_color="#334155",
                  text_color=WHITE, corner_radius=7, height=30, width=84, command=_close_preview).pack(side="right")

    meas_font = tkfont.Font(root=win, family="Segoe UI", size=-px(11))
    bold_font = tkfont.Font(root=win, family="Segoe UI", size=-px(11), weight="bold")

    def show(res):
        tree.delete(*tree.get_children())
        if res["error"]:
            tree["columns"] = ()
            info.configure(text=res["error"], text_color=WARNING_COLOR)
            sheet_menu.pack_forget()
            return
        rows = res["rows"]
        header, body = rows[0][1], rows[1:]
        cols = ["rn"] + [f"c{j}" for j in range(len(header))]
        tree["columns"] = cols
        tree.heading("rn", text="#", anchor="e")
        tree.column("rn", width=px(52), minwidth=px(40), stretch=False, anchor="e")
        for j, h in enumerate(header):
            title = h or f"({_col_letter(j)})"
            samples = [title] + [r[1][j] for r in body]
            wid = max(bold_font.measure(title), max((meas_font.measure(s) for s in samples[1:]), default=0))
            wid = max(px(56), min(px(280), wid + px(26)))
            tree.heading(f"c{j}", text=title, anchor="w")
            tree.column(f"c{j}", width=wid, minwidth=px(40), stretch=False, anchor="w")
        for i, (rn, cells) in enumerate(body):
            tree.insert("", "end", values=[rn] + cells, tags=("odd" if i % 2 else "even",))
        tree.xview_moveto(0)
        tree.yview_moveto(0)
        parts = [Path(st["file"]).name]
        if res["sheet"] and len(res["sheets"]) > 1:
            parts.append(f"sheet “{res['sheet']}” of {len(res['sheets'])}")
        if res["total_text"]:
            parts.append(res["total_text"])
        cols_text = f"{res['ncols']} columns"
        if res["shown_cols"] < res["ncols"]:
            cols_text += f" (first {res['shown_cols']} shown)"
        parts.append(cols_text)
        parts.append(f"first {len(body)} rows shown")
        info.configure(text="  ·  ".join(parts), text_color=SUCCESS_COLOR)
        if len(res["sheets"]) > 1:
            sheet_menu.configure(values=res["sheets"])
            sheet_menu.set(res["sheet"])
            sheet_menu.pack(side="right", padx=(8, 0))
        else:
            sheet_menu.pack_forget()

    def load():
        token = object()
        st["token"] = token
        info.configure(text="Loading preview…", text_color=MUTED_TEXT)
        box = {}

        def work():
            try:
                box["res"] = _read_preview(st["file"], st["sheet"])
            except Exception as e:
                box["res"] = {"error": f"The file could not be read for the preview ({e})."}

        threading.Thread(target=work, daemon=True).start()

        def poll():
            try:
                if st["token"] is not token or not win.winfo_exists():
                    return
                if "res" not in box:
                    win.after(80, poll)
                    return
                show(box["res"])
            except Exception:
                pass

        win.after(80, poll)

    def _pick_file(choice):
        st["file"], st["sheet"] = labels[choice], None
        load()

    def _pick_sheet(choice):
        st["sheet"] = choice
        load()

    win.bind("<Escape>", lambda e: _close_preview())
    win.protocol("WM_DELETE_WINDOW", _close_preview)
    _show_dialog_settled(win, 900, 520)
    load()


# ======================================================================
# Output settings dialog (where cleaned files go and what they are called; saved permanently)
# ======================================================================
def _refresh_output_btn():
    try:
        if output_is_customised():
            output_btn.configure(text="● Output settings", text_color=INFO_COLOR)
        else:
            output_btn.configure(text="Output settings", text_color=MUTED_TEXT)
    except Exception:
        pass


def open_output_settings(_event=None):
    if getattr(app, "_dialog_open", False):
        return
    app._dialog_open = True
    cfg = get_output_settings()

    dialog = ctk.CTkToplevel(app)
    dialog.withdraw()
    try:
        tk.Wm.wm_attributes(dialog, "-alpha", 0.0)
    except Exception:
        pass
    dialog.title("Output settings")
    dialog.resizable(False, False)
    dialog.configure(fg_color=NAVY_BLUE)
    try:
        dialog.iconbitmap(str(resource_path(os.path.join("assets", "strat_data_cleaner.ico"))))
    except Exception:
        pass
    dialog.transient(app)

    mode_var = tk.StringVar(master=dialog, value=cfg["mode"])
    folder_var = tk.StringVar(master=dialog, value=cfg["folder"])
    tmpl_var = tk.StringVar(master=dialog, value=cfg["template"])
    closed = [False]
    saved = [False]

    def _close(ok=False):
        if closed[0]:
            return
        closed[0] = True
        saved[0] = ok
        dialog.destroy()

    def section(text, top):
        ctk.CTkLabel(dialog, text=text, font=("Segoe UI", 10, "bold"), text_color=SECTION_TEXT,
                     anchor="w").pack(fill="x", padx=24, pady=(top, 4))

    ctk.CTkLabel(dialog, text="Output settings", font=("Segoe UI", 14, "bold"), text_color=WHITE,
                 anchor="w").pack(fill="x", padx=24, pady=(20, 0))
    ctk.CTkLabel(dialog, text="Saved permanently. They apply to every cleaned file from now on.",
                 font=("Segoe UI", 11), text_color=MUTED_TEXT, anchor="w").pack(fill="x", padx=24, pady=(0, 2))

    radio_style = {"font": ("Segoe UI", 12), "text_color": LOG_TEXT, "fg_color": ACCENT_ORANGE,
                   "hover_color": HOVER_ORANGE, "border_color": BORDER_COLOR, "radiobutton_width": 18,
                   "radiobutton_height": 18}

    section("SAVE LOCATION", 12)
    ctk.CTkRadioButton(dialog, text="Where each cleaner saves it (default)", variable=mode_var, value="default",
                       command=lambda: refresh(), **radio_style).pack(anchor="w", padx=24, pady=(0, 6))
    ctk.CTkRadioButton(dialog, text="A folder I choose", variable=mode_var, value="custom",
                       command=lambda: refresh(), **radio_style).pack(anchor="w", padx=24, pady=(0, 6))

    folder_row = ctk.CTkFrame(dialog, fg_color="transparent")
    folder_row.pack(fill="x", padx=24, pady=(0, 2))
    folder_row.grid_columnconfigure(0, weight=1)
    entry_style = {"font": ("Segoe UI", 11), "height": 32, "corner_radius": 7, "fg_color": SURFACE_BG,
                   "border_color": BORDER_COLOR, "border_width": 1, "text_color": WHITE}
    folder_entry = ctk.CTkEntry(folder_row, textvariable=folder_var, **entry_style)
    folder_entry.grid(row=0, column=0, sticky="ew")

    def browse():
        start = folder_var.get().strip()
        if not os.path.isdir(start):
            start = CONFIG.get("last_dir") if CONFIG.get("last_dir") and os.path.isdir(CONFIG.get("last_dir")) else str(Path.home())
        chosen = filedialog.askdirectory(parent=dialog, title="Choose the folder for cleaned files", initialdir=start)
        if chosen:
            folder_var.set(os.path.normpath(chosen))
            mode_var.set("custom")
            refresh()

    browse_btn = ctk.CTkButton(folder_row, text="Browse…", font=("Segoe UI", 11, "bold"), fg_color=SURFACE_BG,
                               hover_color="#334155", text_color=WHITE, corner_radius=7, height=32, width=84,
                               command=browse)
    browse_btn.grid(row=0, column=1, padx=(8, 0))

    section("FILE NAME", 14)
    tmpl_entry = ctk.CTkEntry(dialog, textvariable=tmpl_var, **entry_style)
    tmpl_entry.pack(fill="x", padx=24, pady=(0, 6))

    chips = ctk.CTkFrame(dialog, fg_color="transparent")
    chips.pack(fill="x", padx=24, pady=(0, 4))

    def insert_token(token):
        inner = tmpl_entry._entry
        inner.insert(inner.index("insert"), "{" + token + "}")
        tmpl_entry.focus_set()

    for token in OUTPUT_TOKENS:
        ctk.CTkButton(chips, text="{" + token + "}", font=("Consolas", 10, "bold"), fg_color=SURFACE_BG,
                      hover_color="#334155", text_color=SECTION_TEXT, corner_radius=6, height=24, width=64,
                      command=lambda t=token: insert_token(t)).pack(side="left", padx=(0, 6))

    ctk.CTkLabel(dialog, text="{name} the cleaner's own file name  ·  {source} the loaded file  ·  {cleaner} the cleaner "
                              "used  ·  {date} 2026-11-22  ·  {time} 1530",
                 font=("Segoe UI", 10), text_color=MUTED_TEXT, wraplength=412, justify="left",
                 anchor="w").pack(fill="x", padx=24, pady=(0, 6))

    example_lbl = ctk.CTkLabel(dialog, text="", font=("Segoe UI", 11), text_color=SUCCESS_COLOR, wraplength=412,
                               justify="left", anchor="w")
    example_lbl.pack(fill="x", padx=24, pady=(0, 2))
    error_lbl = ctk.CTkLabel(dialog, text="", font=("Segoe UI", 11), text_color=ERROR_COLOR, wraplength=412,
                             justify="left", anchor="w")
    error_lbl.pack(fill="x", padx=24, pady=(0, 0))

    last_out = getattr(app, "last_output_path", None)
    sample_name = Path(last_out).stem if last_out else "Cleaned_Data"
    sample_ext = (Path(last_out).suffix if last_out else "") or ".xlsx"
    sample_source = Path(app.selected_path).stem if getattr(app, "selected_path", None) else "SourceFile"
    sample_cleaner = app.current_run[1] if getattr(app, "current_run", None) else "Cadence Cleaner"

    def check():
        template = tmpl_var.get().strip()
        if not template:
            return "The file name can't be empty."
        bad = unknown_tokens(template)
        if bad:
            return ("Unknown " + ("tokens " if len(bad) > 1 else "token ") + ", ".join("{" + b + "}" for b in bad)
                    + ". Use: " + "  ".join("{" + t + "}" for t in OUTPUT_TOKENS))
        if mode_var.get() == "custom":
            folder = folder_var.get().strip()
            if not folder:
                return "Choose a folder, or switch back to the default location."
            if os.path.exists(folder) and not os.path.isdir(folder):
                return "That path is a file, not a folder."
        return None

    def refresh(*_):
        custom = mode_var.get() == "custom"
        state = "normal" if custom else "disabled"
        folder_entry.configure(state=state)
        browse_btn.configure(state=state)
        err = check()
        error_lbl.configure(text=err or "")
        if err and err.startswith("The file name"):
            example_lbl.configure(text="")
            return
        stem = render_output_name(tmpl_var.get().strip(), sample_name, sample_source, sample_cleaner)
        if sample_ext and stem.lower().endswith(sample_ext.lower()):
            stem = stem[:-len(sample_ext)].rstrip(". ") or sample_name
        text = f"Example:  {stem}{sample_ext}"
        if custom and folder_var.get().strip():
            text += f"\nin  {folder_var.get().strip()}"
        example_lbl.configure(text=text)

    tmpl_var.trace_add("write", refresh)
    folder_var.trace_add("write", refresh)

    def on_save(_e=None):
        err = check()
        if err:
            error_lbl.configure(text=err)
            return
        mode = mode_var.get()
        folder = os.path.normpath(folder_var.get().strip()) if mode == "custom" else ""
        if mode == "custom":
            try:
                os.makedirs(folder, exist_ok=True)
            except OSError as e:
                error_lbl.configure(text=f"That folder can't be used ({e}).")
                return
        CONFIG["output"] = {"mode": mode, "folder": folder, "template": tmpl_var.get().strip()}
        save_config(CONFIG)
        _close(True)

    def on_reset():
        mode_var.set(DEFAULT_OUTPUT["mode"])
        folder_var.set("")
        tmpl_var.set(DEFAULT_OUTPUT["template"])
        refresh()

    btns = ctk.CTkFrame(dialog, fg_color="transparent")
    btns.pack(side="bottom", fill="x", padx=24, pady=(0, 20))
    btns.grid_columnconfigure(0, weight=2)
    btns.grid_columnconfigure((1, 2), weight=1)
    save_btn = GradientButton(btns, text="Save", font=("Segoe UI", 12, "bold"), fg_color=ACCENT_ORANGE,
                              hover_color=HOVER_ORANGE, gradient_to=ACCENT_GRADIENT_END, corner_radius=8,
                              height=38, command=on_save, bg=NAVY_BLUE)
    save_btn.grid(row=0, column=0, sticky="ew", padx=(0, px(6)))
    ctk.CTkButton(btns, text="Reset", font=("Segoe UI", 12, "bold"), fg_color=SURFACE_BG, hover_color="#334155",
                  text_color=WHITE, corner_radius=8, height=38, command=on_reset).grid(row=0, column=1, sticky="ew", padx=6)
    ctk.CTkButton(btns, text="Cancel", font=("Segoe UI", 12, "bold"), fg_color=SURFACE_BG, hover_color=CANCEL_COLOR,
                  text_color=WHITE, corner_radius=8, height=38, command=_close).grid(row=0, column=2, sticky="ew", padx=(6, 0))

    dialog.bind("<Return>", on_save)
    dialog.bind("<Escape>", lambda e: _close(False))
    dialog.protocol("WM_DELETE_WINDOW", _close)
    refresh()

    _show_dialog_settled(dialog, 460, 540)
    try:
        dialog.grab_set()
    except Exception:
        pass
    try:
        dialog.wait_window()
    finally:
        app._dialog_open = False
    if saved[0]:
        _refresh_output_btn()
        new = get_output_settings()
        where = "default location" if new["mode"] == "default" else new["folder"]
        log_message(f"Saved output settings: {where}, file name {new['template']}")
        show_toast("Output settings saved", "success")


def _explain_unavailable(c):
    """Why a cleaner cannot run right now (the 1-6 keys bypass the disabled buttons, so this is checked here too)."""
    first, second = getattr(app, "selected_path", None), getattr(app, "retail_path", None)
    label = c["label"]
    if not first:
        msg = "Please upload a retail file before running the cleaner."
    elif c.get("two_file") and not second:
        msg = (f"The {label} needs two files: one Meta Ads export and one retail sales file. "
               "Browse again and Ctrl+click both.")
    elif c.get("two_file"):
        msg = getattr(app, "pair_msg", None) or "The two files are still being checked. Try again in a moment."
    else:
        msg = f"The {label} takes one file, but two are loaded. Browse again and pick a single file."
    log_message(f"Warning: {msg}")
    set_banner(msg, WARNING_COLOR)


def start_cleaner(c):
    """Confirms (unless turned off) and starts the chosen cleaner."""
    if getattr(app, "is_processing", False) or getattr(app, "_dialog_open", False):
        return
    if not _cleaner_available(c):
        _explain_unavailable(c)
        return
    extra = app.retail_path if c.get("two_file") else None
    name = Path(app.selected_path).name
    if extra:
        name = f"{name}” + “{Path(extra).name}"
    force = False
    if already_cleaned(app.selected_path, c["label"], extra):
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
    run_cleaner_process((c["module"], c["func"]), c["btn"], c["label"], c["color"], c["hover"], force=force,
                        extra_path=extra)


# ======================================================================
# Loading splash: one small window that covers the build of the real UI
# ======================================================================
class LoadingSplash(tk.Toplevel):
    """
    Borderless loading window. The real window is built invisibly behind it (alpha 0) and then swapped
    in fully laid out, so the user never watches the components appear one by one.
    """
    W, H = 400, 180  # unscaled units

    def __init__(self, master, center=None):
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
        if center:                                   # same screen as the window will open on
            x, y = center[0] - w // 2, center[1] - h // 2 - s(20)
        else:
            x = (self.winfo_screenwidth() - w) // 2
            y = (self.winfo_screenheight() - h) // 2 - s(20)
        self.geometry(f"{w}x{h}+{x}+{y}")

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

    def set_alpha(self, alpha):
        try:
            self.attributes("-alpha", max(0.0, min(1.0, alpha)))
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
FILE_PLACEHOLDER = "Click here to choose an Excel or CSV file  (Ctrl+O)\nMeta Ads? Ctrl+click two files: Meta Ads + retail"

if __name__ == "__main__":
    multiprocessing.freeze_support()

    lock_fixed_scale()
    app = ctk.CTk()

    # Keep the real window invisible while the UI is built. It stays fully "alive" (mapped, measurable),
    # so every size/DPI calculation below behaves exactly as before; only the pixels are hidden.
    try:
        app.attributes("-alpha", 0.0)
    except Exception:
        pass

    def _saved_window_center():
        """Centre of the spot where the window will open (its saved position), so the loading window shows on the
        same screen. None on first launch or when that spot is no longer on any monitor."""
        try:
            saved = load_config().get("window")
            s_ = widget_scale(app)
            cx = int(saved["x"]) + FIXED_WIDTH * s_ / 2
            cy = int(saved["y"]) + int(saved.get("h", 700)) * s_ / 2
            vx, vy = app.winfo_vrootx(), app.winfo_vrooty()
            if vx <= cx <= vx + app.winfo_vrootwidth() and vy <= cy <= vy + app.winfo_vrootheight():
                return int(cx), int(cy)
        except Exception:
            pass
        return None

    try:
        splash = LoadingSplash(app, center=_saved_window_center())
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
    app.selected_path = None   # the only file, or the Meta Ads file when two are loaded
    app.retail_path = None     # second file (retail sales) of a Meta Ads pair
    app.pair_ok = False        # True once the two files are recognised as Meta Ads + retail
    app.pair_msg = None        # what is wrong with the pair, when it is not usable
    app.file_tag = None
    app.file_warn = False
    app.current_extra = None
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
    preview_btn = ctk.CTkButton(step1_row, text="Preview file", font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG,
                                hover_color="#334155", text_color=WHITE, corner_radius=6, height=22, width=92,
                                command=open_preview)   # shown (step 1 row, right) once a file is loaded

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
    _clamped["file"] = (file_label, FILE_PLACEHOLDER, 11, 2, None, {})
    file_label.bind("<Button-1>", upload_file)  # the whole label is a big click target
    try:
        file_label._label.configure(cursor="hand2")
    except Exception:
        pass

    clear_btn = ctk.CTkButton(upload_card, text="×", font=("Segoe UI", 17, "bold"), fg_color=SURFACE_BG, hover_color=CLOSE_HOVER,
                              text_color=WHITE, corner_radius=7, width=34, height=34, command=clear_file)

    # The whole card is a click target: it lights up under the pointer and a click opens the file dialog.
    _upload_fx = {"cur": BORDER_COLOR, "goal": None, "hover": False}

    def _refresh_upload_border():
        """Fades the card border toward its resting colour (grey / green when a file is loaded) or its hover colour."""
        base = FILE_LOADED_BORDER if getattr(app, "selected_path", None) else BORDER_COLOR
        goal = _mix(base, HOVER_ORANGE, 0.75) if _upload_fx["hover"] else base
        if goal == _upload_fx["goal"]:
            return
        _upload_fx["goal"] = goal
        start = _upload_fx["cur"]

        def step(t):
            _upload_fx["cur"] = _mix(start, goal, t)
            upload_card.configure(border_color=_upload_fx["cur"])

        animate("upload_border", 160, step)

    def _upload_hover_check(_event=None):
        def check():
            try:
                x, y = app.winfo_pointerxy()
                rx, ry = upload_card.winfo_rootx(), upload_card.winfo_rooty()
                inside = rx <= x < rx + upload_card.winfo_width() and ry <= y < ry + upload_card.winfo_height()
            except Exception:
                inside = False
            if inside != _upload_fx["hover"]:
                _upload_fx["hover"] = inside
                _refresh_upload_border()
        app.after(12, check)  # let Tk settle Enter/Leave between the card and its children first

    for _w in (upload_card, file_label, upload_btn, clear_btn):
        _w.bind("<Enter>", _upload_hover_check, add="+")
        _w.bind("<Leave>", _upload_hover_check, add="+")
    upload_card.bind("<Button-1>", upload_file, add="+")
    try:
        upload_card.configure(cursor="hand2")
    except Exception:
        pass

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
    meta_ads_btn = GradientButton(btn_frame, text="Meta Ads Cleaner", fg_color=META_ADS_COLOR, hover_color=META_ADS_HOVER,
                                  border_color=META_ADS_ACCENT, state="disabled", **button_style)
    cleaner_buttons = [cadence_btn, database_btn, brand_performance_btn, existing_db_btn, current_sales_btn, meta_ads_btn]

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
        {"btn": meta_ads_btn, "label": "Meta Ads Cleaner", "color": META_ADS_COLOR, "hover": META_ADS_HOVER,
         "accent": META_ADS_ACCENT, "module": "meta_ads_cleaner", "func": "process_meta_ads", "two_file": True, "desc": "Needs TWO files: Ctrl+click the Meta Ads (Supermetrics) export and a retail sales file in the browse window. Output: Meta Cleaned Data, Retail Data (with distinct online transactions), Reconciliation and Audit sheets."},
    ]
    for _c in CLEANERS:
        _c["btn"].configure(command=lambda c=_c: start_cleaner(c))
    for _i, _c in enumerate(CLEANERS):
        _c["btn"].set_keycap(str(_i + 1))      # the 1-6 shortcut, shown on the button itself

    # One-line description of whichever cleaner the pointer (or keyboard focus) is on.
    DESC_DEFAULT = "Hover over a cleaner to see what it is for."   # the 1–6 keys are on the buttons and in the footer
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
    _clamped["banner"] = (live_status_banner, "Ready — start by selecting a retail file.", 11, 2, None, {})

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
    step3_tools = ctk.CTkFrame(step3_row, fg_color="transparent")
    step3_tools.grid(row=0, column=2, sticky="e")
    preview_result_btn = ctk.CTkButton(step3_tools, text="Preview result", font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG,
                                       hover_color="#334155", text_color=WHITE, text_color_disabled=DISABLED_TEXT,
                                       corner_radius=6, height=22, width=100, state="disabled",
                                       command=open_result_preview)   # enabled after the first successful run
    preview_result_btn.pack(side="left", padx=(0, 6))
    output_btn = ctk.CTkButton(step3_tools, text="Output settings", font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG,
                               hover_color="#334155", text_color=MUTED_TEXT, corner_radius=6, height=22, width=112,
                               command=open_output_settings)
    output_btn.pack(side="left")
    _refresh_output_btn()
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

    # Session history sits immediately below the review-results controls and has a fixed three-row viewport.
    stats_card = ctk.CTkFrame(left_col, fg_color=CARD_BG, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
    stats_card.pack(fill="x", pady=(0, 0))

    hist_head = ctk.CTkFrame(stats_card, fg_color="transparent")
    hist_head.pack(fill="x", padx=12, pady=(7, 2))
    ctk.CTkLabel(hist_head, text="THIS SESSION", font=("Segoe UI", 9, "bold"), text_color=MUTED_TEXT).pack(side="left")
    session_summary = ctk.CTkLabel(hist_head, text="No files cleaned yet", font=("Segoe UI", 10, "bold"),
                                   text_color=INFO_COLOR, anchor="e")
    session_summary.pack(side="right")

    hist_body = ctk.CTkFrame(stats_card, fg_color="transparent")
    hist_body.pack(fill="x", padx=12, pady=(0, 7))
    hist_body.grid_columnconfigure(0, weight=1)
    hist_text = tk.Text(hist_body, height=3, bg=CARD_BG, fg=LOG_TEXT, bd=0, highlightthickness=0, wrap="none",
                        cursor="arrow", font=("Segoe UI", -px(11)), padx=0, pady=1, takefocus=0, state="disabled")
    hist_text.grid(row=0, column=0, sticky="ew")
    hist_scroll = tk.Scrollbar(hist_body, orient="vertical", command=hist_text.yview, width=px(8, hist_body),
                               bd=0, highlightthickness=0, relief="flat", bg=CARD_BG, troughcolor=CARD_BG,
                               activebackground=SURFACE_BG)
    hist_scroll.grid(row=0, column=1, sticky="ns", padx=(px(5, hist_body), 0))
    hist_text.configure(yscrollcommand=hist_scroll.set)
    hist_text.bind("<MouseWheel>", lambda e: hist_text.yview_scroll(int(-1 * (e.delta / 120)), "units"))
    hist_text.bind("<Configure>", lambda e: app.after_idle(_render_history))
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
    # The log never dictates its own size: whatever it holds (long lines, a scrollbar appearing) stays inside it.
    log_col.grid_propagate(False)

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

    # Second row: level filter (the Warnings button carries the warning count)
    log_tools = ctk.CTkFrame(log_header, fg_color="transparent")
    log_tools.pack(fill="x", pady=(6, 0))

    log_filter = ctk.CTkSegmentedButton(log_tools, values=list(LOG_FILTERS), command=_apply_log_filter,
                                        font=("Segoe UI", 10, "bold"), height=24, corner_radius=6,
                                        fg_color=SURFACE_BG, unselected_color=SURFACE_BG,
                                        unselected_hover_color="#334155", selected_color=ACCENT_ORANGE,
                                        selected_hover_color=HOVER_ORANGE, text_color=WHITE)
    log_filter.set("All")
    log_filter.pack(side="left")

    log_search = ctk.CTkEntry(log_tools, placeholder_text="Search log…  (Ctrl+F)", height=24, width=90, corner_radius=6,
                              font=("Segoe UI", 10), fg_color=SURFACE_BG, border_color=BORDER_COLOR, border_width=1,
                              text_color=WHITE, placeholder_text_color=MUTED_TEXT)
    log_search.pack(side="left", fill="x", expand=True, padx=(8, 0))
    log_search_count = ctk.CTkLabel(log_tools, text="", width=52, font=("Segoe UI", 10, "bold"), text_color=MUTED_TEXT)
    log_search_count.pack(side="left", padx=(4, 0))
    _log_search["entry"], _log_search["count"] = log_search, log_search_count
    log_search.bind("<KeyRelease>", _on_log_search_key, add="+")
    log_search.bind("<Return>", _log_search_step(1), add="+")
    log_search.bind("<Shift-Return>", _log_search_step(-1), add="+")
    log_search.bind("<Escape>", clear_log_search, add="+")

    status_box = ctk.CTkTextbox(log_col, height=LOG_MIN_HEIGHT, corner_radius=9, fg_color=LOG_BG, text_color=LOG_TEXT,
                                font=("Consolas", 11), border_width=1, border_color=BORDER_COLOR, wrap="word")
    status_box.grid(row=1, column=0, sticky="nsew")

    status_box._textbox.tag_config("normal", foreground=LOG_TEXT)
    status_box._textbox.tag_config("success", foreground=SUCCESS_COLOR)
    status_box._textbox.tag_config("warning", foreground=WARNING_COLOR)
    status_box._textbox.tag_config("error", foreground=ERROR_COLOR)
    status_box._textbox.tag_config("notice", foreground=WARNING_COLOR)
    status_box._textbox.tag_config("processing", foreground=INFO_COLOR)
    status_box._textbox.tag_config("fresh", background=LOG_BG)  # the newest line fades out of a soft highlight
    status_box._textbox.tag_config("stamp", foreground=LOG_STAMP)  # muted [time] prefix (created after the level tags so it wins)
    status_box._textbox.tag_config("flash", background="#3B2F0B")  # brief highlight when jumping to a warning
    status_box._textbox.tag_config("searchhit", background=LOG_HIT_BG)     # every search match
    status_box._textbox.tag_config("searchcur", background=LOG_CUR_BG, foreground=NAVY_BLUE)   # the current match
    status_box._textbox.tag_raise("searchhit")
    status_box._textbox.tag_raise("searchcur")

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
    _splash_step("Laying out…", 0.88)
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
            # "uniform" makes the two columns exactly 3 : 2 of the window, whatever their contents ask for. Without
            # it a column's width followed its content (a log scrollbar appearing, a long line), and the other
            # column, and everything in it, shifted to compensate.
            main_frame.grid_columnconfigure(0, weight=3, minsize=px(MIN_LEFT), uniform="panes")
            main_frame.grid_columnconfigure(1, weight=2, minsize=px(MIN_RIGHT), uniform="panes")
            main_frame.grid_rowconfigure(1, weight=1)
            main_frame.grid_rowconfigure(2, weight=0, minsize=0)
            left_col.grid(row=1, column=0, columnspan=1, sticky="nsew", padx=(0, PANE_GAP // 2), pady=(0, 0))
            log_col.grid(row=1, column=1, columnspan=1, sticky="nsew", padx=(PANE_GAP // 2, 0), pady=(0, 0))
        else:
            main_frame.grid_columnconfigure(0, weight=1, minsize=0, uniform="")
            main_frame.grid_columnconfigure(1, weight=0, minsize=0, uniform="")
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
        for _k in ("banner", "file"):          # re-cut the fixed-height texts for the new width / DPI
            if _k in _clamped:
                _render_clamped(_k)


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
    # Guaranteed fit on spawn
    # ======================================================================
    def fit_and_place():
        """One fixed window for every monitor: the same width, the same height and the same spacing everywhere.
        Nothing here depends on the screen except where the window is centred on a first launch."""
        app.update_idletasks()
        scale = widget_scale(app)
        _layout["dual_w"] = _layout["fixed_w"] = FIXED_WIDTH
        _layout["auto_compact"] = False
        set_density(True)
        apply_layout(width_u=FIXED_WIDTH)
        _layout["auto_compact"] = True
        height = int(_need_height_u() + 8)
        _layout["fixed_h"] = height
        x = int((app.winfo_screenwidth() - FIXED_WIDTH * scale) / 2)
        y = int((app.winfo_screenheight() - height * scale - 40) / 2)
        app.geometry(f"{FIXED_WIDTH}x{height}+{max(x, 0)}+{max(y, 0)}")
        try:
            app.resizable(False, False)
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
    app.bind("<Control-f>", focus_log_search)
    app.bind("<Control-F>", focus_log_search)
    app.bind("<Control-p>", open_preview)
    app.bind("<Control-P>", open_preview)

    def _cleaner_shortcut(index):
        def handler(_event=None):
            try:
                if _event is not None and _event.widget.winfo_class() == "Entry":
                    return          # typing a digit in the log search box is not a shortcut
            except Exception:
                pass
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
        if _layout.get("fixed_w"):
            # The size is fixed and was fitted to this screen, so only the position is restored (kept on screen).
            cur_w, cur_h = app.winfo_width(), app.winfo_height()
            if min(x + cur_w, vx + vw) - max(x, vx) < 160 or not (vy - 10 <= y <= vy + vh - 80):
                return
            y = max(vy, min(y, vy + vh - cur_h - 48))
            tk.Wm.wm_geometry(app, f"+{x}+{y}")
            app.update_idletasks()
            return
        pw = w * scale
        visible_w = min(x + pw, vx + vw) - max(x, vx)
        if visible_w < 160 or not (vy - 10 <= y <= vy + vh - 80):
            return
        if _layout.get("fixed_w"):       # fixed-size window: only the position is remembered
            tk.Wm.wm_geometry(app, f"+{x}+{y}")
            app.update_idletasks()
            return
        min_w, min_h = tk.Wm.wm_minsize(app)
        w = int(max(min(w, (vw - 40) / scale), min_w / scale))
        h = int(max(min(h, (vh - 80) / scale), min_h / scale))
        if _layout.get("fixed_w"):
            w = _layout["fixed_w"]            # an older saved width must not override the fixed one
        app.geometry(f"{w}x{h}+{x}+{y}")
        app.update_idletasks()

    def on_close():
        """Never leave a worker process running after the window closes."""
        _save_window_state()
        try:
            if _layout["scale_job"] is not None:
                app.after_cancel(_layout["scale_job"])
            if _anim["job"] is not None:
                app.after_cancel(_anim["job"])
            _stop_worker()
        finally:
            app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_close)

    # ======================================================================
    # Start
    # ======================================================================
    def _reveal_app():
        """Handover: the finished, fully laid-out window fades in while the loading screen fades out.
        Set FADE_IN_MS = 0 for an instant swap (a fade re-composites the window on every frame, which can
        stutter on slow graphics). Animations are switched on only after things settle."""
        try:
            app.lift()
        except Exception:
            pass

        def _finish():
            try:
                app.attributes("-alpha", 1.0)
            except Exception:
                pass
            if splash is not None:
                splash.close()
            app.after(400, lambda: _motion.update(on=True))   # state changes animate from here on

        if FADE_IN_MS <= 0:
            _finish()
            return
        t0 = time.perf_counter()

        def _step():
            t = min(1.0, (time.perf_counter() - t0) * 1000.0 / FADE_IN_MS)
            e = 1.0 - (1.0 - t) ** 3                          # ease-out
            try:
                app.attributes("-alpha", e)
                if splash is not None:
                    splash.set_alpha(1.0 - e)
            except Exception:
                pass
            if t < 1.0:
                app.after(12, _step)
            else:
                _finish()

        _step()

    def _settle_hidden(min_ms=400, quiet_ms=300, max_ms=5000):
        """Pumps the window's deferred work while it is still invisible behind the splash, and only returns
        once it has really stopped changing: CustomTkinter's DPI rescale (it polls every 100 ms), our own
        scale/monitor check (every 120 ms) and the refit passes that follow a monitor change. Without this
        those run after the reveal, and you watch the window re-fit itself to the monitor."""
        _layout["startup"] = True        # refit passes run back to back instead of 0.35 s / 1.25 s / 3 s apart
        t0 = time.perf_counter()
        quiet_since = None
        # Anything that moves, resizes, hides or re-shows the window (our own refits, CustomTkinter's titlebar
        # redraw, Windows' own DPI-change handling) lands here and restarts the quiet period.
        _last_event = {"t": time.perf_counter()}

        def _note_window_event(event):
            if event.widget is app:
                _last_event["t"] = time.perf_counter()

        _ids = [(seq, app.bind(seq, _note_window_event, add="+")) for seq in ("<Configure>", "<Map>", "<Unmap>")]
        try:
            while True:
                try:
                    app.update()
                except Exception:
                    break
                time.sleep(0.008)
                now = time.perf_counter()
                busy = (_layout.get("refit_job") is not None
                        or _resize_job is not None
                        or time.monotonic() < _layout["settle_until"]
                        or (now - _last_event["t"]) * 1000.0 < quiet_ms
                        or abs(widget_scale(app) - (_layout["scale"] or 0)) > 0.001)
                if busy:
                    quiet_since = None
                elif quiet_since is None:
                    quiet_since = now
                elapsed_ms = (now - t0) * 1000.0
                if elapsed_ms >= max_ms:
                    break
                if elapsed_ms >= min_ms and quiet_since is not None and (now - quiet_since) * 1000.0 >= quiet_ms:
                    break
        finally:
            _layout["startup"] = False
            for _seq, _fid in _ids:   # unbind only our handler, not the app's own bindings
                try:
                    app.unbind(_seq, _fid)
                except Exception:
                    pass

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
    if _layout.get("fixed_w"):
        try:
            app.resizable(False, False)   # size is fixed (also disables maximize); done while hidden
        except Exception:
            pass
    _layout["scale"] = widget_scale(app)
    schedule_layout()

    # Settle everything while the window is still invisible: realise the final geometry, run the
    # responsive layout pass and the clipping check now (instead of 40 ms / 350 ms after it is on screen),
    # then reveal. This is what removes the visible re-flow.
    try:
        app.update()
        apply_layout()
        app.update_idletasks()
    except Exception:
        pass
    _splash_step("Almost ready…", 0.99)
    _settle_hidden()
    _splash_step("Ready", 1.0)
    try:
        apply_layout()
        app.update_idletasks()
        app.update()          # paint every widget while still invisible, so the first visible frame is complete
        app.update_idletasks()
    except Exception:
        pass
    _reveal_app()

    # Optional diagnostics: set STRAT_PERF_LOG=1 to log every moment the UI loop was blocked for 80 ms or more.
    if os.environ.get("STRAT_PERF_LOG"):
        _rev_t = time.perf_counter()

        def _trace_window(event):
            if event.widget is app:
                log_message(f"Perf: +{(time.perf_counter() - _rev_t) * 1000:.0f} ms after reveal: "
                            f"{event.type.name if hasattr(event.type, 'name') else event.type} "
                            f"{event.width}x{event.height} scale={widget_scale(app):.2f}")

        app.bind("<Configure>", _trace_window, add="+")
        app.bind("<Map>", _trace_window, add="+")
        app.bind("<Unmap>", _trace_window, add="+")
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