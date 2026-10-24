import customtkinter as ctk
from tkinter import filedialog
import datetime
import os
import ctypes
from pathlib import Path
import sys
from PIL import Image
#Please work please work
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
    """Inserts a timestamped and color-coded message into the inline status box safely."""
    global status_box
    try:
        if 'status_box' in globals() and status_box is not None:
            status_box.configure(state="normal")
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            
            tag = "normal"
            if any(w in message for w in ["Error", "Missing", "KeyError"]):
                tag = "error"
            elif any(w in message for w in ["Success", "Opened", "loaded", "Saved"]):
                tag = "success"
            elif any(w in message for w in ["Warning", "cancelled"]):
                tag = "warning"
            elif any(w in message for w in ["Processing"]):
                tag = "processing"

            status_box._textbox.insert("end", f"[{timestamp}] {message}\n", tag)
            status_box.see("end")
            status_box.configure(state="disabled")
            return
    except Exception:
        pass
    
    print(f"[Startup Log] {message}")

def set_buttons_state(state):
    """Locks or unlocks action buttons to prevent double-clicking during processing."""
    cadence_btn.configure(state=state)
    database_btn.configure(state=state)
    upload_btn.configure(state=state)
    if hasattr(app, 'selected_path') and app.selected_path and state == "normal":
        clear_btn.configure(state="normal")
    else:
        clear_btn.configure(state=state)

def upload_file():
    filename = filedialog.askopenfilename(
        title="Select a File",
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if filename:
        short_name = filename.split('/')[-1]
        file_label.configure(text=f"Selected: {short_name}", text_color="#64FFDA")
        app.selected_path = filename
        clear_btn.grid(row=0, column=2, padx=(5, 15))  
        log_message(f"File loaded successfully: {short_name}")
    else:
        log_message("File selection cancelled.")

def clear_file():
    """Resets the file selection state and disables output references."""
    app.selected_path = None
    app.last_output_path = None  
    file_label.configure(text="No file chosen (XLSX, CSV, TXT)", text_color=MUTED_TEXT)
    clear_btn.grid_forget()  
    
    open_output_btn.configure(
        state="disabled",
        fg_color=OPEN_BTN_DISABLED,
        hover_color=OPEN_BTN_DISABLED
    )
    log_message("Selected file and output reference cleared.")

def open_cleaned_file():
    """Opens the most recently generated cleaned Excel file across operating systems."""
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

def run_cadence_cleaner():
    if not hasattr(app, 'selected_path') or not app.selected_path:
        log_message("Warning: Please upload a file before running Cadence Cleaner.")
        return
    
    file_name = app.selected_path.split('/')[-1]
    log_message(f"Processing Cadence File on '{file_name}'... Please wait.")
    set_buttons_state("disabled")
    app.update()  # Refresh GUI immediately so status and lock show up
    
    try:
        from cadence_cleaner import process_cadence
        row_count, output_path = process_cadence(app.selected_path)
        
        app.last_output_path = output_path
        log_message(f"Successfully Cleaned {row_count} rows.")
        log_message(f"Saved to: {output_path}")
        
        open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)
    except Exception as e:
        log_message(f"Error during cleaning: {str(e)}")
    finally:
        set_buttons_state("normal")

def run_database_cleaner():
    if not hasattr(app, 'selected_path') or not app.selected_path:
        log_message("Warning: Please upload a file before running Database Cleaner.")
        return
    
    file_name = app.selected_path.split('/')[-1]
    log_message(f"Processing Database File on '{file_name}'... Please wait.")
    set_buttons_state("disabled")
    app.update()  # Refresh GUI immediately
    
    try:
        from database_cleaner import process_database
        row_count, output_path = process_database(app.selected_path)
        
        app.last_output_path = output_path
        log_message(f"Successfully Cleaned {row_count} rows.")
        log_message(f"Saved to: {output_path}")
        
        open_output_btn.configure(state="normal", fg_color=OPEN_BTN_COLOR, hover_color=OPEN_BTN_HOVER)
    except Exception as e:
        log_message(f"Error during database cleaning: {str(e)}")
    finally:
        set_buttons_state("normal")
    
# --- Main Application Setup ---
app = ctk.CTk()
app.title("Strat Data Cleaner")

# Taskbar ID & Window Icon Setup
try:
    myappid = 'StrategyDivision.StratDataCleaner.2026'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception as e:
    print(f"Warning: Could not set AppUserModelID: {e}")

BASE_FOLDER = Path(__file__).resolve().parent.parent  
icon_path = BASE_FOLDER / "assets" / "strat_data_cleaner.ico"
try:
    app.iconbitmap(str(icon_path))
except Exception as e:
    print(f"Warning: Could not load window icon: {e}")

# Window Dimensions & Centering Logic
window_width = 620
window_height = 540
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
center_x = int((screen_width - window_width) / 2)
center_y = int((screen_height - window_height) / 2)

app.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
app.resizable(False, False)

# --- Theme Palette (Logo-Aligned & Distinct Action Colors) ---
NAVY_BLUE = "#06101E"       
CARD_BG = "#0D1B2A"         
BORDER_COLOR = "#1B365D"    
WHITE = "#FFFFFF"

# Refined Distinct Button Colors
ACCENT_ORANGE = "#0096C7"   # Browse File
HOVER_ORANGE = "#00B4D8"    # Hover for Browse File

CADENCE_COLOR = "#4361EE"   # Cadence Cleaner (Cobalt Blue)
CADENCE_HOVER = "#5B9AFF"

DB_COLOR = "#0096C7"        # Database Cleaner (Mint Teal)
DB_HOVER = "#00B4D8"
DB_TEXT_COLOR = "#FFFFFF"   # Dark high-contrast text for bright mint button

OPEN_BTN_COLOR = "#2A9D8F"
OPEN_BTN_HOVER = "#3AAFAF"
OPEN_BTN_DISABLED = "#132238"

MUTED_TEXT = "#8892B0"
LOG_BG = "#03070E"
LOG_TEXT = "#64FFDA"

app.configure(fg_color=NAVY_BLUE)

# Load Button Icons via CTkImage Helper
folder_icon = load_icon("folder.png", size=(16, 16))
cadence_icon = load_icon("cadence_cleaning.ico", size=(18, 18))
db_icon = load_icon("database_cleaning.ico", size=(18, 18))
open_icon = load_icon("open.png", size=(18, 18))
close_icon = load_icon("close.png", size=(14, 14))

# Main Container
main_frame = ctk.CTkFrame(app, fg_color=NAVY_BLUE, corner_radius=0)
main_frame.pack(fill="both", expand=True, padx=25, pady=20)

# 1. Header Title with Logo Brand Layout
header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
header_frame.pack(fill="x", pady=(0, 15))

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

# 2. Upload Section
upload_card = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
upload_card.pack(fill="x", pady=(0, 15))
upload_card.grid_columnconfigure(1, weight=1)

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
    image=folder_icon,
    compound="left",
    command=upload_file
)
upload_btn.grid(row=0, column=0, padx=15, pady=15)

file_label = ctk.CTkLabel(
    upload_card,
    text="No file chosen (XLSX, CSV, TXT)",
    font=("Arial", 11),
    text_color=MUTED_TEXT,
    anchor="w"
)
file_label.grid(row=0, column=1, sticky="w", padx=10)

clear_btn = ctk.CTkButton(
    upload_card,
    text="✕" if close_icon is None else "",
    font=("Arial", 12, "bold"),
    fg_color="#C0392B",
    hover_color="#E74C3C",
    text_color=WHITE,
    corner_radius=6,
    width=30,
    height=30,
    image=close_icon,
    command=clear_file
)

# 3. Action Section Title
action_label = ctk.CTkLabel(
    main_frame,
    text="Select Cleaning Operation",
    font=("Helvetica", 13, "bold"),
    text_color=WHITE
)
action_label.pack(anchor="w", pady=(5, 8))

# 4. Action Buttons Grid
btn_frame = ctk.CTkFrame(main_frame, fg_color=NAVY_BLUE, corner_radius=0)
btn_frame.pack(fill="x", pady=(0, 8))
btn_frame.grid_columnconfigure(0, weight=1)
btn_frame.grid_columnconfigure(1, weight=1)

cadence_btn = ctk.CTkButton(
    btn_frame,
    text="Cadence Cleaner",
    font=("Arial", 12, "bold"),
    fg_color=CADENCE_COLOR,
    hover_color=CADENCE_HOVER,
    text_color=WHITE,
    corner_radius=10,
    height=42,
    compound="left",
    command=run_cadence_cleaner
)
cadence_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

database_btn = ctk.CTkButton(
    btn_frame,
    text="Database Cleaner",
    font=("Arial", 12, "bold"),
    fg_color=DB_COLOR,
    hover_color=DB_HOVER,
    text_color=DB_TEXT_COLOR,  # High contrast text for mint teal
    corner_radius=10,
    height=42,
    compound="left",
    command=run_database_cleaner
)
database_btn.grid(row=0, column=1, sticky="ew", padx=(8, 0))

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
    image=open_icon,
    compound="left",
    command=open_cleaned_file
)
open_output_btn.pack(fill="x", pady=(0, 5))

# 5. Inline Status Log Box
status_label = ctk.CTkLabel(
    main_frame,
    text="Activity Log",
    font=("Helvetica", 12, "bold"),
    text_color=WHITE
)
status_label.pack(anchor="w", pady=(10, 5))

status_box = ctk.CTkTextbox(
    main_frame,
    height=90,
    corner_radius=8,
    fg_color=LOG_BG,
    text_color=LOG_TEXT,
    font=("Consolas", 12),
    border_width=1,
    border_color=BORDER_COLOR
)
status_box.pack(fill="x", pady=(0, 10))

status_box._textbox.tag_config("normal", foreground=LOG_TEXT)
status_box._textbox.tag_config("success", foreground="#2ECC71")
status_box._textbox.tag_config("warning", foreground="#F1C40F")
status_box._textbox.tag_config("error", foreground="#E74C3C")
status_box._textbox.tag_config("processing", foreground="#3A86FF") # Blue highlight for active work

status_box.insert("end", "[System] Application ready. Please upload a file.\n", "normal")
status_box.configure(state="disabled")

# 6. Footer Credit
footer_label = ctk.CTkLabel(
    main_frame,
    text="Created by the Strategy Division - Data Analytics Team, 2026",
    font=("Arial", 11, "italic"),
    text_color=MUTED_TEXT
)
footer_label.pack(side="bottom", anchor="center", pady=(2, 0))

app.mainloop()