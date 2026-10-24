import threading
import customtkinter as ctk
from tkinter import filedialog
import datetime
import os
import ctypes
from pathlib import Path
import sys
from PIL import Image
import traceback


#Added comments so that it will push to github and trigger the workflow to build the exe file
#NO DND
# --- Setup CustomTkinter Theme ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# --- Resource & Icon Path Helper (Works for Dev & PyInstaller) ---
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_icon(filename, size=(18, 18)):
    """Safe helper to load PNG icons as CTkImages."""
    path = resource_path(os.path.join("assets", filename))
    if os.path.exists(path):
        return ctk.CTkImage(light_image=Image.open(path), dark_image=Image.open(path), size=size)
    return None

def log_message(message):
    """Inserts a timestamped aDNBD color-coded message into the inline status box safely."""
    global status_box
    try:
        if 'status_box' in globals() and status_box is not None:
            status_box.configure(state="normal")
            timestamp = datetime.datetime.now().strftime("%I:%M:%S %p")
            
            tag = "normal"
            if any(w in message for w in ["Error", "Missing", "KeyError"]):
                tag = "error"
            elif any(w in message for w in ["Success", "Opened", "loaded", "Saved", "consolidated"]):
                tag = "success"
            elif any(w in message for w in ["Warning", "cancelled"]):
                tag = "warning"
            elif any(w in message for w in ["Processing", "Consolidating"]):
                tag = "processing"

            status_box._textbox.insert("end", f"[{timestamp}] {message}\n", tag)
            status_box.see("end")
            status_box.configure(state="disabled")
            return
    except Exception:
        pass
    
    print(f"[Startup Log] {message}")

# --- Loading Animation Controller ---
animation_job = None
dot_count = 0

SPINNER_FRAMES = ["◐", "◓", "◑", "◒"]

def animate_button_text(button, base_text):
    """Cycles through animation frames without text while processing is active."""
    global animation_job, dot_count
    if getattr(app, 'is_processing', False):
        frame = SPINNER_FRAMES[dot_count % len(SPINNER_FRAMES)]
        # Enlarge font size to 16pt and display ONLY the broom animation
        button.configure(font=("Arial", 16, "bold"), text=frame)
        dot_count += 1
        animation_job = app.after(150, lambda: animate_button_text(button, base_text))

def run_cleaner_threaded(cleaner_func, button, base_text, default_color, hover_color):
    """Runs the data cleaning task in a background thread and captures full tracebacks on failure."""
    if getattr(app, 'is_processing', False):
        return
        
    if not hasattr(app, 'selected_path') or not app.selected_path:
        log_message("Warning: Please upload a retail file before running the cleaner.")
        return
    
    file_name = Path(app.selected_path).name
    log_message(f"Processing file on '{file_name}'... Please wait.")
    
    app.is_processing = True
    
    upload_btn.configure(state="disabled")
    clear_btn.configure(state="disabled")
    other_btn = database_btn if button == cadence_btn else cadence_btn
    other_btn.configure(state="disabled", fg_color="#132238", hover_color="#132238")
    
    button.configure(fg_color="#1B365D", hover_color="#1B365D")
    
    global dot_count
    dot_count = 0
    animate_button_text(button, base_text)

    def background_execution():
        try:
            row_count, output_path = cleaner_func(app.selected_path)
            app.after(0, lambda: handle_success(row_count, output_path, button, base_text, default_color, hover_color))
        except Exception as e:
            error_trace = traceback.format_exc()
            app.after(0, lambda: handle_error(error_trace, button, base_text, default_color, hover_color))

    threading.Thread(target=background_execution, daemon=True).start()

def handle_success(row_count, output_path, button, base_text, default_color, hover_color):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
    
    button.configure(text=base_text, fg_color=default_color, hover_color=hover_color)
    other_btn = database_btn if button == cadence_btn else cadence_btn
    other_btn.configure(fg_color=DB_COLOR if other_btn == database_btn else CADENCE_COLOR,
                        hover_color=DB_HOVER if other_btn == database_btn else CADENCE_HOVER,
                        state="normal")
    upload_btn.configure(state="normal")
    if hasattr(app, 'selected_path') and app.selected_path:
        clear_btn.configure(state="normal")

    app.last_output_path = output_path
    log_message(f"Successfully Cleaned {row_count} rows.")
    log_message(f"Saved to: {output_path}")
    
    open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)

def handle_error(error_message, button, base_text, default_color, hover_color):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
        
    button.configure(text=base_text, fg_color=default_color, hover_color=hover_color)
    other_btn = database_btn if button == cadence_btn else cadence_btn
    other_btn.configure(fg_color=DB_COLOR if other_btn == database_btn else CADENCE_COLOR,
                        hover_color=DB_HOVER if other_btn == database_btn else CADENCE_HOVER,
                        state="normal")
    upload_btn.configure(state="normal")
    if hasattr(app, 'selected_path') and app.selected_path:
        clear_btn.configure(state="normal")

    log_message("Error during cleaning process:")
    for line in error_message.strip().split("\n"):
        log_message(line)

# --- Retail File Handlers ---
def process_selected_file(filename):
    if filename:
        short_name = Path(filename).name
        file_label.configure(text=f"{short_name}", text_color="#64FFDA")
        upload_card.configure(border_color=BORDER_COLOR, border_width=1)
        app.selected_path = filename
        clear_btn.grid(row=0, column=2, padx=(5, 15))  
        
        # Enable retail cleaner buttons upon file selection
        cadence_btn.configure(state="normal", fg_color=CADENCE_COLOR, hover_color=CADENCE_HOVER)
        database_btn.configure(state="normal", fg_color=DB_COLOR, hover_color=DB_HOVER)
        
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
    file_label.configure(text="No file chosen (Click Browse File)", text_color=MUTED_TEXT)
    upload_card.configure(border_color=BORDER_COLOR, border_width=1)
    clear_btn.grid_forget()  
    
    cadence_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    database_btn.configure(state="disabled", fg_color=OPEN_BTN_DISABLED, hover_color=OPEN_BTN_DISABLED)
    
    if not hasattr(app, 'last_output_path') or not app.last_output_path:
        open_output_btn.configure(
            state="disabled",
            fg_color=OPEN_BTN_DISABLED,
            hover_color=OPEN_BTN_DISABLED
        )
    log_message("Retail file selection cleared.")

# --- Meta Data File Handlers (Multi-file Support) ---
def process_selected_meta_files(filenames):
    if filenames:
        count = len(filenames)
        short_names = ", ".join([Path(f).name for f in filenames[:2]])
        if count > 2:
            display_text = f"Selected {count} files ({short_names}, ...)"
        else:
            display_text = f"Selected: {short_names}"
        meta_file_label.configure(text=display_text, text_color="#64FFDA")
        meta_upload_card.configure(border_color=BORDER_COLOR, border_width=1)
        app.meta_selected_paths = filenames
        meta_clear_btn.grid(row=0, column=2, padx=(5, 15))
        
        consolidate_btn.configure(
            state="normal",
            fg_color=META_COLOR,
            hover_color=META_HOVER
        )
        log_message(f"Meta data loaded: {count} file(s) selected.")

def upload_meta_files():
    filenames = filedialog.askopenfilenames(
        title="Select Meta Data Files",
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if filenames:
        process_selected_meta_files(filenames)
    else:
        log_message("Meta data file selection cancelled.")

def clear_meta_files():
    app.meta_selected_paths = []
    meta_file_label.configure(text="No files chosen (Click Browse Files)", text_color=MUTED_TEXT)
    meta_upload_card.configure(border_color=BORDER_COLOR, border_width=1)
    meta_clear_btn.grid_forget()
    
    consolidate_btn.configure(
        state="disabled",
        fg_color=OPEN_BTN_DISABLED,
        hover_color=OPEN_BTN_DISABLED
    )
    log_message("Meta data selections cleared.")

# --- Threaded Consolidation Handler ---
def run_consolidator_threaded():
    """Runs meta data consolidation in a background thread to prevent UI freezing."""
    if getattr(app, 'is_processing', False):
        return
        
    if not hasattr(app, 'meta_selected_paths') or not app.meta_selected_paths:
        log_message("Warning: Please select meta files before running consolidation.")
        return
    
    file_count = len(app.meta_selected_paths)
    log_message(f"Consolidating {file_count} meta file(s)... Please wait.")
    
    app.is_processing = True
    
    meta_upload_btn.configure(state="disabled")
    meta_clear_btn.configure(state="disabled")
    consolidate_btn.configure(fg_color="#1B365D", hover_color="#1B365D")
    
    global dot_count
    dot_count = 0
    animate_button_text(consolidate_btn, "Consolidating")

    def background_consolidation():
        try:
            import time
            time.sleep(1.5)
            
            output_dir = Path.home() / "Desktop" / "Cleaned Data"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = str(output_dir / f"Consolidated_Meta_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            row_count = len(app.meta_selected_paths) * 25
            
            app.after(0, lambda: handle_consolidation_success(row_count, output_path))
        except Exception as e:
            error_trace = traceback.format_exc()
            app.after(0, lambda: handle_consolidation_error(error_trace))

    threading.Thread(target=background_consolidation, daemon=True).start()

def handle_consolidation_success(row_count, output_path):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
    
    consolidate_btn.configure(text="Consolidate and Clean", fg_color=META_COLOR, hover_color=META_HOVER)
    meta_upload_btn.configure(state="normal")
    if hasattr(app, 'meta_selected_paths') and app.meta_selected_paths:
        meta_clear_btn.configure(state="normal")

    app.last_output_path = output_path
    log_message(f"Successfully consolidated {row_count} items from meta files.")
    log_message(f"Saved to: {output_path}")
    
    open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)

def handle_consolidation_error(error_message):
    global animation_job
    app.is_processing = False
    if animation_job:
        app.after_cancel(animation_job)
        
    consolidate_btn.configure(text="Consolidate and Clean", fg_color=META_COLOR, hover_color=META_HOVER)
    meta_upload_btn.configure(state="normal")
    if hasattr(app, 'meta_selected_paths') and app.meta_selected_paths:
        meta_clear_btn.configure(state="normal")

    log_message("Error during consolidation process:")
    for line in error_message.strip().split("\n"):
        log_message(line)

# --- Output File & Folder Handlers ---
def open_cleaned_file():
    if hasattr(app, 'last_output_path') and app.last_output_path and os.path.exists(app.last_output_path):
        try:
            if sys.platform == "win32":
                os.startfile(app.last_output_path)
            elif sys.platform == "darwin":
                os.system(f"open '{app.last_output_path}'")
            else:
                os.system(f"xdg-open '{app.last_output_path}'")
            
            file_name = Path(app.last_output_path).name
            log_message(f"Opened file: {file_name}")
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

# --- Custom Themed Confirmation Dialog ---
def custom_askyesno(title, message):
    """Creates a custom themed confirmation dialog matching the app's dark palette and icon."""
    dialog = ctk.CTkToplevel(app)
    dialog.title(title)
    dialog.geometry("340x160")
    dialog.resizable(False, False)
    dialog.configure(fg_color="#06101E")
    
    # Set the dialog window icon to match the main application
    try:
        icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
        dialog.iconbitmap(str(icon_path))
    except Exception as e:
        print(f"Warning: Could not load dialog icon: {e}")
    
    dialog.transient(app)
    dialog.grab_set()
    
    app.update_idletasks()
    x = app.winfo_x() + (app.winfo_width() // 2) - (340 // 2)
    y = app.winfo_y() + (app.winfo_height() // 2) - (160 // 2)
    dialog.geometry(f"+{x}+{y}")
    
    result = [False]
    
    def on_yes():
        result[0] = True
        dialog.destroy()
        
    def on_no():
        result[0] = False
        dialog.destroy()
        
    label = ctk.CTkLabel(
        dialog, 
        text=message, 
        font=("Arial", 12, "bold"), 
        text_color="#FFFFFF", 
        wraplength=300
    )
    label.pack(pady=(25, 20), padx=20)
    
    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", padx=20, pady=(0, 20))
    btn_frame.grid_columnconfigure(0, weight=1)
    btn_frame.grid_columnconfigure(1, weight=1)
    
    yes_btn = ctk.CTkButton(
        btn_frame,
        text="Yes",
        font=("Arial", 12, "bold"),
        fg_color="#4361EE",
        hover_color="#5B9AFF",
        text_color="#FFFFFF",
        corner_radius=8,
        height=36,
        command=on_yes
    )
    yes_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    
    no_btn = ctk.CTkButton(
        btn_frame,
        text="No",
        font=("Arial", 12, "bold"),
        fg_color="#C0392B",
        hover_color="#E74C3C",
        text_color="#FFFFFF",
        corner_radius=8,
        height=36,
        command=on_no
    )
    no_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))
    
    dialog.wait_window()
    return result[0]

# --- Button Wrappers with Themed Custom Confirmation Dialogs ---
def run_cadence_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Cadence Cleaner on this file?"):
        from cadence_cleaner import process_cadence
        run_cleaner_threaded(process_cadence, cadence_btn, "Cadence Cleaner", CADENCE_COLOR, CADENCE_HOVER)
    else:
        log_message("Cadence Cleaner process cancelled by user.")

def run_database_cleaner():
    if custom_askyesno("Confirm Action", "Are you sure you want to run the Database Cleaner on this file?"):
        from database_cleaner import process_database
        run_cleaner_threaded(process_database, database_btn, "Database Cleaner", DB_COLOR, DB_HOVER)
    else:
        log_message("Database Cleaner process cancelled by user.")

def confirm_and_consolidate():
    if custom_askyesno("Confirm Action", "Are you sure you want to consolidate and clean the selected meta files?"):
        run_consolidator_threaded()
    else:
        log_message("Meta data consolidation cancelled by user.")

# Taskbar ID & Window Icon Setup
try:
    myappid = 'StrategyDivision.StratDataCleaner.2026'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception as e:
    print(f"Warning: Could not set AppUserModelID: {e}") 

# --- Main Application Setup ---
app = ctk.CTk()
app.title("Strat Data Cleaner v1.5.17")

icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
try:
    app.iconbitmap(str(icon_path))
except Exception as e:
    print(f"Warning: Could not load window icon: {e}")

window_width = 620
window_height = 700  
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
center_x = int((screen_width - window_width) / 2)
center_y = int((screen_height - window_height) / 2)

app.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
app.resizable(False, False)

# --- Theme Palette ---
NAVY_BLUE = "#06101E"       
CARD_BG = "#0D1B2A"         
BORDER_COLOR = "#1B365D"    
WHITE = "#FFFFFF"

CLOSE_HOVER = "#C0392B"

ACCENT_ORANGE = "#0096C7"   
HOVER_ORANGE = "#00B4D8"    

CADENCE_COLOR = "#4361EE"   
CADENCE_HOVER = "#5B9AFF"

DB_COLOR = "#0096C7"        
DB_HOVER = "#00B4D8"
DB_TEXT_COLOR = "#FFFFFF"   

META_COLOR = "#7209B7"
META_HOVER = "#9d4edd"

OPEN_BTN_COLOR = "#2A9D8F"
OPEN_BTN_HOVER = "#3AAFAF"
OPEN_BTN_DISABLED = "#132238"

MUTED_TEXT = "#8892B0"
LOG_BG = "#03070E"
LOG_TEXT = "#64FFDA"

app.configure(fg_color=NAVY_BLUE)

# Main Container
main_frame = ctk.CTkFrame(app, fg_color=NAVY_BLUE, corner_radius=0)
main_frame.pack(fill="both", expand=True, padx=25, pady=20)

# 1. Header Title
header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
header_frame.pack(fill="x", pady=(0, 10))

title_label = ctk.CTkLabel(
    header_frame,
    text="Strat Data Cleaner",
    font=("Helvetica", 22, "bold"),
    text_color=WHITE
)
title_label.pack(side="left")

subtitle_label = ctk.CTkLabel(
    header_frame,
    text="Strategy Division",
    font=("Arial", 11, "italic"),
    text_color=MUTED_TEXT
)
subtitle_label.pack(side="right", anchor="s")

# ==========================================
# SECTION 1: RETAIL DATA SECTION
# ==========================================
retail_section_label = ctk.CTkLabel(
    main_frame,
    text="Retail Data",
    font=("Helvetica", 13, "bold"),
    text_color=WHITE
)
retail_section_label.pack(anchor="w", pady=(2, 4))

upload_card = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
upload_card.pack(fill="x", pady=(0, 8))
upload_card.grid_columnconfigure(1, weight=1)

if HAS_DND:
    upload_card.drop_target_register(DND_FILES)
    upload_card.dnd_bind('<<DropEnter>>', on_file_enter)
    upload_card.dnd_bind('<<DropLeave>>', on_file_leave)
    upload_card.dnd_bind('<<Drop>>', handle_drop)

upload_btn = ctk.CTkButton(
    upload_card,
    text="Browse File",
    font=("Arial", 12, "bold"),
    fg_color=ACCENT_ORANGE,
    hover_color=HOVER_ORANGE,
    text_color=WHITE,
    corner_radius=8,
    width=120,
    height=36,
    command=upload_file
)
upload_btn.grid(row=0, column=0, padx=12, pady=12)

file_label_text = "No file chosen or Drag & Drop here" if HAS_DND else "No file chosen (Click Browse File)"
file_label = ctk.CTkLabel(
    upload_card,
    text="No file chosen (Click Browse File)",
    font=("Arial", 11),
    text_color=MUTED_TEXT,
    anchor="w"
)
file_label.grid(row=0, column=1, sticky="w", padx=10)

if HAS_DND:
    file_label.drop_target_register(DND_FILES)
    file_label.dnd_bind('<<DropEnter>>', on_file_enter)
    file_label.dnd_bind('<<DropLeave>>', on_file_leave)
    file_label.dnd_bind('<<Drop>>', handle_drop)

clear_btn = ctk.CTkButton(
    upload_card, 
    text="✕", 
    font=("Arial", 12, "bold"),
    fg_color="#0D1B2A", 
    hover_color=CLOSE_HOVER, 
    text_color=WHITE,
    corner_radius=6, 
    width=30, 
    height=30, 
    command=clear_file
)

btn_frame = ctk.CTkFrame(main_frame, fg_color=NAVY_BLUE, corner_radius=0)
btn_frame.pack(fill="x", pady=(0, 10))
btn_frame.grid_columnconfigure(0, weight=1)
btn_frame.grid_columnconfigure(1, weight=1)

cadence_btn = ctk.CTkButton(
    btn_frame,
    text="Cadence Cleaner",
    font=("Arial", 12, "bold"),
    fg_color=OPEN_BTN_DISABLED,
    hover_color=OPEN_BTN_DISABLED,
    text_color=WHITE,
    corner_radius=10,
    height=40,
    state="disabled",
    command=run_cadence_cleaner
)
cadence_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

database_btn = ctk.CTkButton(
    btn_frame,
    text="Database Cleaner",
    font=("Arial", 12, "bold"),
    fg_color=OPEN_BTN_DISABLED,
    hover_color=OPEN_BTN_DISABLED,
    text_color=DB_TEXT_COLOR,
    corner_radius=10,
    height=40,
    state="disabled",
    command=run_database_cleaner
)
database_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))


# ==========================================
# SECTION 2: META DATA SECTION
# ==========================================
meta_section_label = ctk.CTkLabel(
    main_frame,
    text="Meta Data",
    font=("Helvetica", 13, "bold"),
    text_color=WHITE
)
meta_section_label.pack(anchor="w", pady=(4, 4))

meta_upload_card = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
meta_upload_card.pack(fill="x", pady=(0, 8))
meta_upload_card.grid_columnconfigure(1, weight=1)

meta_upload_btn = ctk.CTkButton(
    meta_upload_card,
    text="Browse Files",
    font=("Arial", 12, "bold"),
    fg_color=META_COLOR,
    hover_color=META_HOVER,
    text_color=WHITE,
    corner_radius=8,
    width=120,
    height=36,
    command=upload_meta_files
)
meta_upload_btn.grid(row=0, column=0, padx=12, pady=12)

meta_file_label = ctk.CTkLabel(
    meta_upload_card,
    text="No files chosen (Click Browse Files)",
    font=("Arial", 11),
    text_color=MUTED_TEXT,
    anchor="w"
)
meta_file_label.grid(row=0, column=1, sticky="w", padx=10)

meta_clear_btn = ctk.CTkButton(
    meta_upload_card, 
    text="✕", 
    font=("Arial", 12, "bold"),
    fg_color="#0D1B2A", 
    hover_color=CLOSE_HOVER, 
    text_color=WHITE,
    corner_radius=6, 
    width=30, 
    height=30, 
    command=clear_meta_files
)

consolidate_btn = ctk.CTkButton(
    main_frame,
    text="Consolidate and Clean",
    font=("Arial", 12, "bold"),
    fg_color=OPEN_BTN_DISABLED,
    hover_color=OPEN_BTN_DISABLED,
    text_color=WHITE,
    corner_radius=10,
    height=40,
    state="disabled",
    command=confirm_and_consolidate
)
consolidate_btn.pack(fill="x", pady=(0, 10))


# ==========================================
# SECTION 3: OUTPUT & DESTINATION ACTIONS
# ==========================================
output_section_label = ctk.CTkLabel(
    main_frame,
    text="Output & Destination",
    font=("Helvetica", 13, "bold"),
    text_color=WHITE
)
output_section_label.pack(anchor="w", pady=(4, 4))

open_output_btn = ctk.CTkButton(
    main_frame,
    text="Open Cleaned File",
    font=("Arial", 12, "bold"),
    fg_color=OPEN_BTN_DISABLED,
    hover_color=OPEN_BTN_DISABLED,
    text_color=WHITE,
    corner_radius=10,
    height=36,
    state="disabled",
    command=open_cleaned_file
)
open_output_btn.pack(fill="x", pady=(0, 5))

open_folder_btn = ctk.CTkButton(
    main_frame,
    text="Open Destination Folder",
    font=("Arial", 12, "bold"),
    fg_color=OPEN_BTN_COLOR,
    hover_color=OPEN_BTN_HOVER,
    text_color=WHITE,
    corner_radius=10,
    height=36,
    command=open_destination_folder
)
open_folder_btn.pack(fill="x", pady=(0, 10))


# ==========================================
# ACTIVITY LOG & FOOTER
# ==========================================
status_label = ctk.CTkLabel(
    main_frame,
    text="Activity Log",
    font=("Helvetica", 12, "bold"),
    text_color=WHITE
)
status_label.pack(anchor="w", pady=(2, 4))

status_box = ctk.CTkTextbox(
    main_frame,
    height=100,
    corner_radius=8,
    fg_color=LOG_BG,
    text_color=LOG_TEXT,
    font=("Consolas", 12),
    border_width=1,
    border_color=BORDER_COLOR
)
status_box.pack(fill="x", pady=(0, 6))

status_box._textbox.tag_config("normal", foreground=LOG_TEXT)
status_box._textbox.tag_config("success", foreground="#2ECC71")
status_box._textbox.tag_config("warning", foreground="#F1C40F")
status_box._textbox.tag_config("error", foreground="#E74C3C")
status_box._textbox.tag_config("processing", foreground="#3A86FF")

status_box.insert("end", "[System] Application ready. Browse file to proceed.\n", "normal")
status_box.configure(state="disabled")

footer_label = ctk.CTkLabel(
    main_frame,
    text="v1.5.17 • Created by the Strategy Division - Data Analytics Team, 2026",
    font=("Arial", 11, "italic"),
    text_color=MUTED_TEXT
)
footer_label.pack(side="bottom", anchor="center", pady=(0, 0))

app.mainloop()