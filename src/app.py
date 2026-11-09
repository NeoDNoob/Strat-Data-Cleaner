import threading
import customtkinter as ctk
from tkinter import filedialog
import datetime
import os
import ctypes
from pathlib import Path
import sys
import time
import traceback

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

SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

def log_message(message):
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
            elif any(w in message for w in ["Warning", "cancelled", "cleared"]):
                tag = "warning"
            elif any(w in message for w in ["Processing", "Consolidating", "Running"]):
                tag = "processing"

            status_box._textbox.insert("end", f"[{timestamp}] {message}\n", tag)
            status_box.see("end")
            status_box.configure(state="disabled")
            return
    except Exception:
        pass
    
    print(f"[Startup Log] {message}")

animation_job = None
timer_job = None
dot_count = 0
start_time = 0

def animate_process_state(button, base_text):
    """Cycles through spinner frames on the button and updates the live banner timer."""
    global animation_job, timer_job, dot_count
    if getattr(app, 'is_processing', False) and not getattr(app, 'cancel_requested', False):
        frame = SPINNER_FRAMES[dot_count % len(SPINNER_FRAMES)]
        elapsed = int(time.time() - start_time)
        m, s = divmod(elapsed, 60)
        time_str = f"{m} min {s} sec" if m > 0 else f"{s} seconds"
        
        button.configure(text=f"{frame} Processing...")
        live_status_banner.configure(text=f"RUNNING  •  {base_text}  •  Elapsed: {time_str}", text_color=INFO_COLOR)
        
        dot_count += 1
        animation_job = app.after(1000, lambda: animate_process_state(button, base_text))

def run_cleaner_threaded(cleaner_func, button, base_text, default_color, hover_color):
    """Runs the data cleaning task in a background thread with live banner updates and cancellation support."""
    if getattr(app, 'is_processing', False):
        return
        
    if not hasattr(app, 'selected_path') or not app.selected_path:
        log_message("Warning: Please upload a retail file before running the cleaner.")
        return
    
    file_name = Path(app.selected_path).name
    log_message(f"Processing file on '{file_name}' using {base_text}...")
    
    app.is_processing = True
    app.cancel_requested = False
    global start_time
    start_time = time.time()
    
    upload_btn.configure(state="disabled")
    clear_btn.configure(state="disabled")
    cancel_btn.configure(state="normal", fg_color=CANCEL_COLOR, hover_color=CANCEL_HOVER)
    live_status_banner.configure(text=f"Starting {base_text}...", text_color=INFO_COLOR)
    
    for other_btn in (cadence_btn, database_btn, brand_performance_btn, existing_db_btn):
        if other_btn != button:
            other_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    
    button.configure(fg_color=SURFACE_BG, hover_color=SURFACE_BG)
    
    global dot_count
    dot_count = 0
    animate_process_state(button, base_text)

    def background_execution():
        try:
            row_count, output_path = cleaner_func(app.selected_path)
            if getattr(app, 'cancel_requested', False):
                return
            elapsed = time.time() - start_time
            app.after(0, lambda: handle_success(row_count, output_path, elapsed, button, base_text, default_color, hover_color))
        except Exception as e:
            if getattr(app, 'cancel_requested', False):
                return
            error_trace = traceback.format_exc()
            app.after(0, lambda: handle_error(error_trace, button, base_text, default_color, hover_color))

    threading.Thread(target=background_execution, daemon=True).start()

def cancel_operation():
    """Immediately requests cancellation and restores UI state."""
    if getattr(app, 'is_processing', False):
        app.cancel_requested = True
        app.is_processing = False
        log_message("Warning: Operation cancelled by user.")
        
        global animation_job
        if animation_job:
            app.after_cancel(animation_job)
            
        live_status_banner.configure(text="Operation cancelled. Ready for the next task.", text_color=WARNING_COLOR)
            
        for btn, col, hov in [
            (cadence_btn, CADENCE_COLOR, CADENCE_HOVER),
            (database_btn, DB_COLOR, DB_HOVER),
            (brand_performance_btn, BRAND_COLOR, BRAND_HOVER),
            (existing_db_btn, EXISTING_DB_COLOR, EXISTING_DB_HOVER)
        ]:
            btn.configure(state="normal", fg_color=col, hover_color=hov)
            if btn == cadence_btn: btn.configure(text="Cadence Cleaner")
            elif btn == database_btn: btn.configure(text="Database Cleaner")
            elif btn == brand_performance_btn: btn.configure(text="Brand Performance Cleaner")
            elif btn == existing_db_btn: btn.configure(text="Existing DB Cleaner")

        cancel_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
        upload_btn.configure(state="normal")
        if hasattr(app, 'selected_path') and app.selected_path:
            clear_btn.configure(state="normal")

def format_duration(seconds):
    """Formats elapsed seconds into a clean human-readable string."""
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    m = int(seconds // 60)
    s = seconds % 60
    return f"{m} min {s:.1f} sec"

def handle_success(row_count, output_path, elapsed_seconds, button, base_text, default_color, hover_color):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
    
    live_status_banner.configure(text="Operation completed successfully.", text_color=SUCCESS_COLOR)
    button.configure(text=base_text, fg_color=default_color, hover_color=hover_color)
    
    for other_btn in (cadence_btn, database_btn, brand_performance_btn, existing_db_btn):
        if other_btn != button:
            color = (
                DB_COLOR if other_btn == database_btn
                else BRAND_COLOR if other_btn == brand_performance_btn
                else EXISTING_DB_COLOR if other_btn == existing_db_btn
                else CADENCE_COLOR
            )
            hover = (
                DB_HOVER if other_btn == database_btn
                else BRAND_HOVER if other_btn == brand_performance_btn
                else EXISTING_DB_HOVER if other_btn == existing_db_btn
                else CADENCE_HOVER
            )
            other_btn.configure(fg_color=color, hover_color=hover, state="normal")
            
    upload_btn.configure(state="normal")
    cancel_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    if hasattr(app, 'selected_path') and app.selected_path:
        clear_btn.configure(state="normal")

    duration_str = format_duration(elapsed_seconds)
    last_run_label.configure(text=f"Last run: {base_text}  •  {row_count:,} rows cleaned  •  {duration_str}", text_color=SUCCESS_COLOR)
    app.last_output_path = output_path
    log_message(f"Success! Cleaned {row_count} rows in {duration_str}.")
    log_message(f"Saved to: {output_path}")
    open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)

def handle_error(error_message, button, base_text, default_color, hover_color):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
        
    live_status_banner.configure(text="Operation failed. Review the activity log.", text_color=ERROR_COLOR)
    button.configure(text=base_text, fg_color=default_color, hover_color=hover_color)
    
    for other_btn in (cadence_btn, database_btn, brand_performance_btn, existing_db_btn):
        if other_btn != button:
            color = (
                DB_COLOR if other_btn == database_btn
                else BRAND_COLOR if other_btn == brand_performance_btn
                else EXISTING_DB_COLOR if other_btn == existing_db_btn
                else CADENCE_COLOR
            )
            hover = (
                DB_HOVER if other_btn == database_btn
                else BRAND_HOVER if other_btn == brand_performance_btn
                else EXISTING_DB_HOVER if other_btn == existing_db_btn
                else CADENCE_HOVER
            )
            other_btn.configure(fg_color=color, hover_color=hover, state="normal")
            
    upload_btn.configure(state="normal")
    cancel_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    if hasattr(app, 'selected_path') and app.selected_path:
        clear_btn.configure(state="normal")

    log_message("Error during cleaning process:")
    for line in error_message.strip().split("\n"):
        log_message(line)

# --- File Handlers ---
def process_selected_file(filename):
    if filename:
        short_name = Path(filename).name
        file_size = os.path.getsize(filename)
        size_text = f"{file_size / (1024 * 1024):.2f} MB" if file_size >= 1024 * 1024 else f"{file_size / 1024:.1f} KB"
        file_label.configure(text=f"{short_name}  •  {size_text}", text_color=SUCCESS_COLOR)
        upload_card.configure(border_color=BORDER_COLOR, border_width=1)
        app.selected_path = filename
        clear_btn.grid(row=0, column=2, padx=(5, 15))  
        
        cadence_btn.configure(state="normal", fg_color=CADENCE_COLOR, hover_color=CADENCE_HOVER, border_color=CADENCE_ACCENT)
        database_btn.configure(state="normal", fg_color=DB_COLOR, hover_color=DB_HOVER, border_color=DB_ACCENT)
        brand_performance_btn.configure(state="normal", fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, border_color=BRAND_ACCENT)
        existing_db_btn.configure(state="normal", fg_color=EXISTING_DB_COLOR, hover_color=EXISTING_DB_HOVER, border_color=EXISTING_DB_ACCENT)
        
        log_message(f"Retail file loaded: {short_name}")

def upload_file():
    filename = filedialog.askopenfilename(
        title="Select Retail File",
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if filename:
        process_selected_file(filename)
    else:
        log_message("Retail file selection cancelled.")

def clear_file():
    app.selected_path = None
    file_label.configure(text="No file selected. Browse or drop an Excel/CSV file.", text_color=MUTED_TEXT)
    upload_card.configure(border_color=BORDER_COLOR, border_width=1)
    clear_btn.grid_forget()  
    
    cadence_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    database_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    brand_performance_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    existing_db_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    
    if not hasattr(app, 'last_output_path') or not app.last_output_path:
        open_output_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    log_message("Retail file selection cleared.")

def open_cleaned_file():
    if hasattr(app, 'last_output_path') and app.last_output_path and os.path.exists(app.last_output_path):
        try:
            if sys.platform == "win32":
                os.startfile(app.last_output_path)
            elif sys.platform == "darwin":
                os.system(f"open '{app.last_output_path}'")
            else:
                os.system(f"xdg-open '{app.last_output_path}'")
            log_message(f"Opened file: {Path(app.last_output_path).name}")
        except Exception as e:
            log_message(f"Error opening file: {str(e)}")
    else:
        log_message("No generated file found to open.")

def open_destination_folder():
    output_dir = Path.home() / "Desktop" / "Cleaned Data"
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        if sys.platform == "win32":
            os.startfile(output_dir)
        elif sys.platform == "darwin":
            os.system(f"open '{output_dir}'")
        else:
            os.system(f"xdg-open '{output_dir}'")
        log_message("Opened destination folder: Cleaned Data")
    except Exception as e:
        log_message(f"Error opening folder: {str(e)}")

def clear_log():
    """Clears the activity log."""
    status_box.configure(state="normal")
    status_box.delete("1.0", "end")
    status_box.configure(state="disabled")
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


def custom_askyesno(title, message):
    dialog = ctk.CTkToplevel(app)
    dialog.title(title)
    dialog.geometry("340x160")
    dialog.resizable(False, False)
    dialog.configure(fg_color=NAVY_BLUE)
    
    try:
        icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
        dialog.iconbitmap(str(icon_path))
    except Exception:
        pass
    
    dialog.transient(app)
    dialog.grab_set()
    
    # Force geometry calculation for both app and dialog
    app.update_idletasks()
    dialog.update_idletasks()
    
    # Calculate center position relative to the main app window
    dialog_width = 340
    dialog_height = 160
    x = app.winfo_x() + (app.winfo_width() - dialog_width) // 2
    y = app.winfo_y() + (app.winfo_height() - dialog_height) // 2
    
    dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
    result = [False]
    
    def on_yes():
        result[0] = True
        dialog.destroy()
        
    def on_no():
        result[0] = False
        dialog.destroy()
        
    label = ctk.CTkLabel(dialog, text=message, font=("Arial", 12, "bold"), text_color="#FFFFFF", wraplength=300)
    label.pack(pady=(25, 20), padx=20)
    
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", padx=20, pady=(0, 20))
    btn_frame.grid_columnconfigure(0, weight=1)
    btn_frame.grid_columnconfigure(1, weight=1)
    
    yes_btn = ctk.CTkButton(btn_frame, text="Yes", font=("Arial", 12, "bold"), fg_color=ACCENT_ORANGE, hover_color=HOVER_ORANGE, text_color="#FFFFFF", corner_radius=8, height=36, command=on_yes)
    yes_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    
    no_btn = ctk.CTkButton(btn_frame, text="No", font=("Arial", 12, "bold"), fg_color=CANCEL_COLOR, hover_color=CANCEL_HOVER, text_color="#FFFFFF", corner_radius=8, height=36, command=on_no)
    no_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))
    
    dialog.wait_window()
    return result[0]

def run_cadence_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Cadence Cleaner on this file?"):
        from cadence_cleaner import process_cadence
        run_cleaner_threaded(process_cadence, cadence_btn, "Cadence Cleaner", CADENCE_COLOR, CADENCE_HOVER)

def run_database_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Database Cleaner on this file?"):
        from database_cleaner import process_database
        run_cleaner_threaded(process_database, database_btn, "Database Cleaner", DB_COLOR, DB_HOVER)

def run_brand_performance_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Brand Performance Cleaner on this file?"):
        from brand_performance_cleaner import process_brand_performance
        run_cleaner_threaded(process_brand_performance, brand_performance_btn, "Brand Performance Cleaner", BRAND_COLOR, BRAND_HOVER)

def run_existing_db_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Existing DB Cleaner on this file?"):
        from existing_db_cleaner import process_existing_database
        run_cleaner_threaded(process_existing_database, existing_db_btn, "Existing DB Cleaner", EXISTING_DB_COLOR, EXISTING_DB_HOVER)

# --- Main App Initialization ---
app = ctk.CTk()
app.title("Strat Data Cleaner v3.8.20")

icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
try:
    app.iconbitmap(str(icon_path))
except Exception:
    pass

window_width = 720
window_height = 820
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
center_x = int((screen_width - window_width) / 2)
center_y = int((screen_height - window_height) / 2)

app.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
app.resizable(True, True)
app.minsize(660, 760)

# --- Palette: Midnight Enterprise ---
NAVY_BLUE = "#080F1C"
CARD_BG = "#111D30"
SURFACE_BG = "#1E293B"
BORDER_COLOR = "#263B56"
WHITE = "#F1F5F9"

CLOSE_HOVER = "#7F1D1D"
ACCENT_ORANGE = "#3B82F6"
HOVER_ORANGE = "#2563EB"

# Each cleaner keeps its own identity through a restrained accent border.
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

CANCEL_COLOR = "#7F1D1D"
CANCEL_HOVER = "#991B1B"

OPEN_BTN_COLOR = "#0F766E"
OPEN_BTN_HOVER = "#0D9488"
OPEN_BTN_DISABLED = "#1E293B"

MUTED_TEXT = "#94A3B8"
LOG_BG = "#050B14"
LOG_TEXT = "#CBD5E1"
SUCCESS_COLOR = "#34D399"
WARNING_COLOR = "#FBBF24"
ERROR_COLOR = "#F87171"
INFO_COLOR = "#60A5FA"

app.configure(fg_color=NAVY_BLUE)

main_frame = ctk.CTkFrame(app, fg_color=NAVY_BLUE, corner_radius=0)
main_frame.pack(fill="both", expand=True, padx=25, pady=20)

# Header
header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
header_frame.pack(fill="x", pady=(0, 16))

title_label = ctk.CTkLabel(header_frame, text="Strat Data Cleaner", font=("Segoe UI", 25, "bold"), text_color=WHITE)
title_label.pack(side="left")

subtitle_label = ctk.CTkLabel(header_frame, text="STRATEGY DIVISION  /  DATA ANALYTICS", font=("Segoe UI", 10, "bold"), text_color=MUTED_TEXT)
subtitle_label.pack(side="right", anchor="s")

# Retail Section
retail_section_label = ctk.CTkLabel(main_frame, text="RETAIL DATA", font=("Segoe UI", 12, "bold"), text_color=WHITE)
retail_section_label.pack(anchor="w", pady=(0, 7))

upload_card = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
upload_card.pack(fill="x", pady=(0, 14))
upload_card.grid_columnconfigure(1, weight=1)

upload_btn = ctk.CTkButton(upload_card, text="Browse File", font=("Segoe UI", 12, "bold"), fg_color=ACCENT_ORANGE, hover_color=HOVER_ORANGE, text_color=WHITE, corner_radius=8, width=132, height=42, command=upload_file)
upload_btn.grid(row=0, column=0, padx=14, pady=14)

file_label = ctk.CTkLabel(upload_card, text="No file selected. Browse or drop an Excel/CSV file.", font=("Segoe UI", 11), text_color=MUTED_TEXT, anchor="w")
file_label.grid(row=0, column=1, sticky="ew", padx=10)

clear_btn = ctk.CTkButton(upload_card, text="×", font=("Segoe UI", 17, "bold"), fg_color=SURFACE_BG, hover_color=CLOSE_HOVER, text_color=WHITE, corner_radius=7, width=34, height=34, command=clear_file)

# Cleaner actions
cleaner_heading = ctk.CTkLabel(main_frame, text="CLEANING OPERATIONS", font=("Segoe UI", 12, "bold"), text_color=WHITE)
cleaner_heading.pack(anchor="w", pady=(0, 7))

btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
btn_frame.pack(fill="x", pady=(0, 10))
btn_frame.grid_columnconfigure((0, 1), weight=1, uniform="cleaners")

button_style = {
    "font": ("Segoe UI", 12, "bold"),
    "text_color": WHITE,
    "corner_radius": 10,
    "height": 48,
    "border_width": 1,
}
cadence_btn = ctk.CTkButton(btn_frame, text="Cadence Cleaner", fg_color=CADENCE_COLOR, hover_color=CADENCE_HOVER, border_color=CADENCE_ACCENT, state="disabled", command=run_cadence_cleaner, **button_style)
cadence_btn.grid(row=0, column=0, sticky="ew", padx=(0, 7), pady=(0, 8))

database_btn = ctk.CTkButton(btn_frame, text="Database Cleaner", fg_color=DB_COLOR, hover_color=DB_HOVER, border_color=DB_ACCENT, state="disabled", command=run_database_cleaner, **button_style)
database_btn.grid(row=0, column=1, sticky="ew", padx=(7, 0), pady=(0, 8))

brand_performance_btn = ctk.CTkButton(btn_frame, text="Brand Performance Cleaner", fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, border_color=BRAND_ACCENT, state="disabled", command=run_brand_performance_cleaner, **button_style)
brand_performance_btn.grid(row=1, column=0, sticky="ew", padx=(0, 7))

existing_db_btn = ctk.CTkButton(btn_frame, text="Existing DB Cleaner", fg_color=EXISTING_DB_COLOR, hover_color=EXISTING_DB_HOVER, border_color=EXISTING_DB_ACCENT, state="disabled", command=run_existing_db_cleaner, **button_style)
existing_db_btn.grid(row=1, column=1, sticky="ew", padx=(7, 0))

# Operation status and cancellation
operation_card = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
operation_card.pack(fill="x", pady=(0, 14))
operation_card.grid_columnconfigure(0, weight=1)

live_status_banner = ctk.CTkLabel(operation_card, text="Ready for a cleaning operation", font=("Segoe UI", 11), text_color=MUTED_TEXT, anchor="w")
live_status_banner.grid(row=0, column=0, sticky="ew", padx=14, pady=10)

cancel_btn = ctk.CTkButton(operation_card, text="Cancel", font=("Segoe UI", 11, "bold"), fg_color=CANCEL_COLOR, hover_color=CANCEL_HOVER, text_color=WHITE, corner_radius=7, height=32, width=92, state="disabled", command=cancel_operation)
cancel_btn.grid(row=0, column=1, padx=12, pady=7)

# Output Section
output_section_label = ctk.CTkLabel(main_frame, text="OUTPUT & DESTINATION", font=("Segoe UI", 12, "bold"), text_color=WHITE)
output_section_label.pack(anchor="w", pady=(0, 7))

output_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
output_frame.pack(fill="x", pady=(0, 14))
output_frame.grid_columnconfigure((0, 1), weight=1, uniform="outputs")

open_output_btn = ctk.CTkButton(output_frame, text="Open Cleaned File", font=("Segoe UI", 11, "bold"), fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER, text_color=WHITE, corner_radius=8, height=40, state="disabled", command=open_cleaned_file)
open_output_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

open_folder_btn = ctk.CTkButton(output_frame, text="Open Destination Folder", font=("Segoe UI", 11, "bold"), fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER, text_color=WHITE, corner_radius=8, height=40, command=open_destination_folder)
open_folder_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))

# Last operation summary
last_run_label = ctk.CTkLabel(main_frame, text="No completed operations in this session.", font=("Segoe UI", 10), text_color=MUTED_TEXT, anchor="w")
last_run_label.pack(fill="x", pady=(0, 10))

# Activity Log Section
log_header = ctk.CTkFrame(main_frame, fg_color="transparent")
log_header.pack(fill="x", pady=(0, 6))

status_label = ctk.CTkLabel(log_header, text="ACTIVITY LOG", font=("Segoe UI", 12, "bold"), text_color=WHITE)
status_label.pack(side="left")

log_actions = ctk.CTkFrame(log_header, fg_color="transparent")
log_actions.pack(side="right")

copy_log_btn = ctk.CTkButton(log_actions, text="Copy", font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG, hover_color="#334155", text_color=WHITE, corner_radius=6, height=27, width=58, command=copy_log)
copy_log_btn.pack(side="left", padx=(0, 6))

clear_log_btn = ctk.CTkButton(log_actions, text="Clear", font=("Segoe UI", 10, "bold"), fg_color=SURFACE_BG, hover_color="#334155", text_color=WHITE, corner_radius=6, height=27, width=58, command=clear_log)
clear_log_btn.pack(side="left")

status_box = ctk.CTkTextbox(main_frame, height=145, corner_radius=9, fg_color=LOG_BG, text_color=LOG_TEXT, font=("Consolas", 11), border_width=1, border_color=BORDER_COLOR)
status_box.pack(fill="both", expand=True, pady=(0, 10))

status_box._textbox.tag_config("normal", foreground=LOG_TEXT)
status_box._textbox.tag_config("success", foreground=SUCCESS_COLOR)
status_box._textbox.tag_config("warning", foreground=WARNING_COLOR)
status_box._textbox.tag_config("error", foreground=ERROR_COLOR)
status_box._textbox.tag_config("processing", foreground=INFO_COLOR)

status_box.insert("end", "[System] Application ready. Browse a retail file to begin.\\n", "normal")
status_box.configure(state="disabled")

footer_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
footer_frame.pack(side="bottom", fill="x", pady=(5, 0))

footer_label = ctk.CTkLabel(footer_frame, text="STRAT DATA CLEANER  •  v3.8.20", font=("Segoe UI", 9, "bold"), text_color=MUTED_TEXT)
footer_label.pack(side="left")

footer_credit = ctk.CTkLabel(footer_frame, text="Strategy Division · Data Analytics Team · 2026", font=("Segoe UI", 10), text_color=MUTED_TEXT)
footer_credit.pack(side="right")

app.mainloop()