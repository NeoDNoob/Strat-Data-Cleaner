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

# --- Safe Drag-and-Drop Import Fallback ---
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


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
            timestamp = datetime.datetime.now().strftime("%I:%M:%S %p")
            
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

# --- Loading Animation Controller ---
animation_job = None
dot_count = 0

SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

def animate_button_text(button, base_text):
    """Dynamically cycles through circular spinner frames while processing is active."""
    global animation_job, dot_count
    if getattr(app, 'is_processing', False):
        frame = SPINNER_FRAMES[dot_count % len(SPINNER_FRAMES)]
        button.configure(text=f"{frame}  {base_text}...")
        dot_count += 1
        animation_job = app.after(100, lambda: animate_button_text(button, base_text))

def run_cleaner_threaded(cleaner_func, button, base_text, default_color, hover_color):
    """Runs the data cleaning task in a background thread and captures full tracebacks on failure."""
    if getattr(app, 'is_processing', False):
        return
        
    if not hasattr(app, 'selected_path') or not app.selected_path:
        log_message("Warning: Please upload a file before running the cleaner.")
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

def process_selected_file(filename):
    if filename:
        short_name = Path(filename).name
        file_label.configure(text=f"Selected: {short_name}", text_color="#64FFDA")
        upload_card.configure(border_color=BORDER_COLOR, border_width=1)
        app.selected_path = filename
        clear_btn.grid(row=0, column=2, padx=(5, 15))  
        log_message(f"File loaded successfully: {short_name}")

def upload_file():
    filename = filedialog.askopenfilename(
        title="Select a File",
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*"))
    )
    if filename:
        process_selected_file(filename)
    else:
        log_message("File selection cancelled.")

# --- Drag & Drop Visual Animation Handlers ---
def on_file_enter(event):
    if not getattr(app, 'is_processing', False):
        upload_card.configure(border_color="#64FFDA", border_width=2)
        file_label.configure(text="Drop file to load...", text_color="#64FFDA")

def on_file_leave(event):
    if not getattr(app, 'is_processing', False) and (not hasattr(app, 'selected_path') or not app.selected_path):
        upload_card.configure(border_color=BORDER_COLOR, border_width=1)
        file_label.configure(text="No file chosen or Drag & Drop here", text_color=MUTED_TEXT)

def handle_drop(event):
    files = app.tk.splitlist(event.data)
    if files:
        filename = files[0].strip('{}')
        process_selected_file(filename)

def clear_file():
    app.selected_path = None
    app.last_output_path = None  
    file_label.configure(text="No file chosen or Drag & Drop here", text_color=MUTED_TEXT)
    upload_card.configure(border_color=BORDER_COLOR, border_width=1)
    clear_btn.grid_forget()  
    
    open_output_btn.configure(
        state="disabled",
        fg_color=OPEN_BTN_DISABLED,
        hover_color=OPEN_BTN_DISABLED
    )
    log_message("Selected file and output reference cleared.")

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
    """Opens the destination folder directory in file explorer across operating systems."""
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

def run_cadence_cleaner():
    from cadence_cleaner import process_cadence
    run_cleaner_threaded(process_cadence, cadence_btn, "Cadence Cleaner", CADENCE_COLOR, CADENCE_HOVER)

def run_database_cleaner():
    from database_cleaner import process_database
    run_cleaner_threaded(process_database, database_btn, "Database Cleaner", DB_COLOR, DB_HOVER)

# Taskbar ID & Window Icon Setup
try:
    myappid = 'StrategyDivision.StratDataCleaner.2026'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception as e:
    print(f"Warning: Could not set AppUserModelID: {e})") 

# --- Main Application Setup ---
app = ctk.CTk()
app.title("Strat Data Cleaner v1.0.11")

if HAS_DND:
    try:
        TkinterDnD.require(app)
    except Exception:
        HAS_DND = False

icon_path = resource_path(os.path.join("assets", "strat_data_cleaner.ico"))
try:
    app.iconbitmap(str(icon_path))
except Exception as e:
    print(f"Warning: Could not load window icon: {e}")

window_width = 620
window_height = 585
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

ACCENT_ORANGE = "#0096C7"   
HOVER_ORANGE = "#00B4D8"    

CADENCE_COLOR = "#4361EE"   
CADENCE_HOVER = "#5B9AFF"

DB_COLOR = "#0096C7"        
DB_HOVER = "#00B4D8"
DB_TEXT_COLOR = "#FFFFFF"   

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
upload_btn.grid(row=0, column=0, padx=15, pady=15)

file_label_text = "No file chosen or Drag & Drop here" if HAS_DND else "No file chosen (Click Browse File)"
file_label = ctk.CTkLabel(
    upload_card,
    text=file_label_text,
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
    text="✖", 
    font=("Arial", 12, "bold"),
    fg_color="#C0392B", 
    hover_color="#E74C3C", 
    text_color=WHITE,
    corner_radius=6, 
    width=30, 
    height=30, 
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
    command=run_cadence_cleaner
)
cadence_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

database_btn = ctk.CTkButton(
    btn_frame,
    text="Database Cleaner",
    font=("Arial", 12, "bold"),
    fg_color=DB_COLOR,
    hover_color=DB_HOVER,
    text_color=DB_TEXT_COLOR,
    corner_radius=10,
    height=42,
    command=run_database_cleaner
)
database_btn.grid(row=0, column=1, sticky="ew", padx=(8, 0))

# 5. Output Actions
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
open_folder_btn.pack(fill="x", pady=(0, 5))

# 6. Inline Status Log Box
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
status_box._textbox.tag_config("processing", foreground="#3A86FF")

status_box.insert("end", "[System] Application ready. Upload the file to begin.\n", "normal")
status_box.configure(state="disabled")

# 7. Footer Credit
footer_label = ctk.CTkLabel(
    main_frame,
    text="v1.0.11 • Created by the Strategy Division - Data Analytics Team, 2026",
    font=("Arial", 11, "italic"),
    text_color=MUTED_TEXT
)
footer_label.pack(side="bottom", anchor="center", pady=(2, 0))

app.mainloop()