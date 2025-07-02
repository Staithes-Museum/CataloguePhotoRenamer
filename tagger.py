import os
import json
import tkinter as tk
import shutil
from tkinter import messagebox
from PIL import Image, ImageTk
import pandas as pd
import tkinter.simpledialog as simpledialog
import tkinter.ttk as ttk # Import ttk for the modern progress bar

EXCEL_FILE = "Object numbers.xlsx"
IMAGE_FOLDER = "Images"
PROGRESS_FILE = "progress.json"

# ────────────────────────────────────────────────────────────
# Helpers for Excel + progress JSON
# ────────────────────────────────────────────────────────────

def load_descriptions_from_excel():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip() # Clean column names
    
    # Ensure all expected columns exist, fill missing with empty strings
    expected_cols = [
        "Description", "Object Number", "Sticker Number",
        "Imported Description", "Location"
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = '' # Add missing columns as empty strings/NaN

    # Filter out rows where 'Description' is empty or NaN
    df = df[df["Description"].notna() & (df["Description"].astype(str).str.strip() != '')]

    # Convert to list of dictionaries for easier lookup
    description_data = {}
    for _, row in df.iterrows():
        # Ensure the key itself is stripped when creating the dictionary
        stripped_desc_key = str(row["Description"]).strip()
        description_data[stripped_desc_key] = {
            "Object Number": str(row["Object Number"]).strip() if pd.notna(row["Object Number"]) else '',
            "Sticker Number": str(row["Sticker Number"]).strip() if pd.notna(row["Sticker Number"]) else '',
            "Imported Description": str(row["Imported Description"]).strip() if pd.notna(row["Imported Description"]) else '',
            "Location": str(row["Location"]).strip() if pd.notna(row["Location"]) else ''
        }
    
    # The descriptions_list will now contain the cleaned keys
    descriptions_list = list(description_data.keys())
    return descriptions_list, description_data

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            data = json.load(f)
            # Ensure new fields exist for backward compatibility
            data.setdefault("used_tags", [])
            data.setdefault("renamed", {})
            data.setdefault("index", 0)
            data.setdefault("image_files_order", [])
            data.setdefault("dont_know_files", []) # New field
            return data
    # Initial state if progress file doesn't exist
    return {"used_tags": [], "renamed": {}, "index": 0, "image_files_order": [], "dont_know_files": []}


# ────────────────────────────────────────────────────────────
# Main application class
# ────────────────────────────────────────────────────────────
class ImageTagger:
    ZOOM_MIN = 0.05    # 5% of original
    ZOOM_MAX = 10.0    # 1000%
    ZOOM_STEP = 1.2    # 20% per key-press / tick

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Museum Image Tagger")
        self.root.geometry("1100x800")

        # 🔹 data ------------------------------------------------------
        self.descriptions, self.description_data = load_descriptions_from_excel()

        p = load_progress()
        self.used_tags = set(p["used_tags"])
        self.renamed_files = p["renamed"]
        self.current_index = p["index"]
        self.dont_know_files = set(p["dont_know_files"]) # Initialize new set

        # Load initial image files, prioritizing saved order if available
        initial_disk_files = sorted(
            f for f in os.listdir(IMAGE_FOLDER)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        )

        if p.get("image_files_order") and len(p["image_files_order"]) > 0:
            # Reconstruct saved order, filtering out missing files and adding new ones
            saved_order = p["image_files_order"]
            # Keep files in saved_order that actually exist
            existing_saved_files = [f for f in saved_order if os.path.exists(os.path.join(IMAGE_FOLDER, f))]
            # Add any new files from disk that weren't in the saved order, sorted alphabetically at the end
            new_disk_files = sorted([f for f in initial_disk_files if f not in existing_saved_files])
            self.image_files = existing_saved_files + new_disk_files
            # Ensure current_index is valid for the reloaded list
            if self.current_index >= len(self.image_files):
                self.current_index = 0 if len(self.image_files) > 0 else 0
        else:
            # No saved order, use initial disk order
            self.image_files = initial_disk_files

        if not self.image_files:
            messagebox.showerror("No images found", "Put images in the Images/ folder first.")
            root.destroy()
            return

        # ── layout ---------------------------------------------------

        # Progress Bar and Filename at the top
        self.progress_frame = tk.Frame(root, bd=2, relief=tk.GROOVE)
        self.progress_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.progress_label = tk.Label(self.progress_frame, text="", anchor="w", font=("TkDefaultFont", 10, "bold"))
        self.progress_label.pack(side=tk.LEFT, padx=5)

        self.filename_label = tk.Label(self.progress_frame, text="", anchor="e", font=("TkDefaultFont", 9))
        self.filename_label.pack(side=tk.RIGHT, padx=5)

        # Custom progress bar using Canvas
        self.progress_canvas = tk.Canvas(self.progress_frame, bg="lightgray", height=20, highlightthickness=1, highlightbackground="gray")
        self.progress_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.progress_canvas.bind("<Button-1>", self._on_progressbar_click) # Bind click event
        self.progress_canvas.bind("<Configure>", self._on_progress_canvas_configure) # Bind configure for resizing


        self.canvas = tk.Canvas(root, bg="black", highlightthickness=0) # Default to black
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right = tk.Frame(root)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=8, pady=8)

        tk.Label(right, text="Select Description:").pack(anchor="w")
        yscroll = tk.Scrollbar(right)
        self.listbox = tk.Listbox(right, height=28, yscrollcommand=yscroll.set, exportselection=False)
        yscroll.config(command=self.listbox.yview)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack()
        for d in self.descriptions:
            self.listbox.insert(tk.END, d)

        btn = lambda t, c: tk.Button(right, text=t, command=c).pack(fill=tk.X, pady=1)
        btn("←  Previous",          self.prev_image)
        btn("→  Next",              self.next_image)
        btn("💾  Tag & Rename",      self.tag_and_rename)
        btn("❓  Don't know (send→)",self.send_to_end) # This button will now work as intended
        btn("🔁  Undo Rename",       self.undo_rename)
        btn("📥  Import object descriptions", self.import_object_descriptions)

        # Zoom controls
        zoom_frame = tk.Frame(right)
        zoom_frame.pack(pady=6, fill=tk.X)
        tk.Button(zoom_frame, text="＋", command=lambda: self.zoom_relative(self.ZOOM_STEP)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(zoom_frame, text="－", command=lambda: self.zoom_relative(1/self.ZOOM_STEP)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.zoom_slider = tk.Scale(right, from_=int(self.ZOOM_MAX*100), to=int(self.ZOOM_MIN*100),
                                            orient=tk.VERTICAL, command=self._slider_zoom, showvalue=False,
                                            length=160, resolution=10) # Added resolution=10 for 10% steps
        self.zoom_slider.pack(pady=4)

        # ── state vars ---------------------------------------------
        self.original_img: Image.Image | None = None
        self.zoom = 1.0    # effective scale (1 = original pixels)
        self.offset = [0.0, 0.0]     # top-left image corner in canvas coords
        self.drag_start = None       # (x,y) while dragging

        # ── event bindings -----------------------------------------
        for seq in ("<Control-plus>", "<Control-KP_Add>", "<Control-equal>"):
            root.bind_all(seq, lambda e: self.zoom_relative(self.ZOOM_STEP))
        for seq in ("<Control-minus>", "<Control-KP_Subtract>"):
            root.bind_all(seq, lambda e: self.zoom_relative(1/self.ZOOM_STEP))

        self.canvas.bind("<ButtonPress-1>", self._start_pan)
        self.canvas.bind("<B1-Motion>",       self._do_pan)
        self.canvas.bind("<MouseWheel>",      self._on_mousewheel)    # Windows/macOS
        self.canvas.bind("<Button-4>",        self._on_mousewheel)    # Linux scroll-up
        self.canvas.bind("<Button-5>",        self._on_mousewheel)    # Linux scroll-down
        
        # New: Right-click binding for listbox
        self.listbox.bind("<Button-3>", self._show_tag_info)


        # first image ----------------------------------------------
        self.display_image()

    # ────────────────────────────────────────────────────────────
    # Save progress method
    # ────────────────────────────────────────────────────────────
    def save_progress(self):
        """Saves the current state of the application to a JSON file."""
        data = {
            "used_tags": list(self.used_tags),
            "renamed": self.renamed_files,
            "index": self.current_index,
            "image_files_order": self.image_files, # Save the current order
            "dont_know_files": list(self.dont_know_files)
        }
        with open(PROGRESS_FILE, "w") as f:
            json.dump(data, f, indent=2)

    # ────────────────────────────────────────────────────────────
    # Import object descriptions
    # ────────────────────────────────────────────────────────────
    def import_object_descriptions(self):
        prompt = (
            "Paste rows copied from Google Sheets (columns C, D, P, Q, U). Each row should be tab or comma-separated.\n\n"
            "Required columns in order (case-insensitive column headers):\n"
            "C: Sticker Number\n"
            "D: Description\n"
            "P: Object Number\n"
            "Q: Imported Description\n"
            "U: Location\n\n"
            "Example row (tab-separated):\n"
            "12345\tCaptain Cook Bell\tSSESM.2019.338\tLarge brass bell with naval insignia\tFirstFloor/Room10/Cab1/Shelf05\n\n"
            "Press OK to continue or Cancel to abort."
        )
        raw_data = simpledialog.askstring("Import object descriptions", prompt)
        if raw_data is None:
            return  # User cancelled

        # Parse pasted data into list of dicts: [{'Description': ..., 'Object Number': ..., 'Location': ...}, ...]
        rows = self._parse_pasted_data(raw_data)
        if not rows:
            messagebox.showerror("Error", "No valid rows found in pasted data. Ensure all required columns (C, D, P, Q, U) are present and description (D) is not empty.")
            return

        # Show preview of first few rows for confirmation
        preview_lines = []
        for i, r in enumerate(rows[:5]):
            preview_lines.append(
                f"Desc: {r['Description']}\n"
                f"ObjNum: {r['Object Number']}\n"
                f"Sticker: {r['Sticker Number']}\n"
                f"ImpDesc: {r['Imported Description']}\n"
                f"Loc: {r['Location']}\n"
                "---"
            )
        preview_text = "\n".join(preview_lines)
        confirm = messagebox.askokcancel("Confirm import",
            f"Pasted data example (first {min(5, len(rows))} rows):\n\n{preview_text}\n\n"
            "Does this look right? Press OK to import or Cancel to abort."
        )

        if not confirm:
            return

        # Create a new DataFrame from the imported rows
        df_to_save = pd.DataFrame(rows)

        # Define the order of columns for saving to Excel
        output_columns = [
            "Description", "Object Number", "Sticker Number",
            "Imported Description", "Location"
        ]
        # Ensure all columns exist in the DataFrame before reordering,
        # fill with empty strings if a pasted column was missing for some reason.
        for col in output_columns:
            if col not in df_to_save.columns:
                df_to_save[col] = ''
            
        df_to_save = df_to_save[output_columns]


        # Save the new DataFrame to Excel, effectively overwriting the old content
        try:
            df_to_save.to_excel(EXCEL_FILE, index=False)
            messagebox.showinfo("Success", f"Imported {len(rows)} rows and updated {EXCEL_FILE}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save {EXCEL_FILE}:\n{e}")
            return # Exit if saving fails

        # Now, update the application's internal data
        # Reload descriptions from the now-updated Excel file
        self.descriptions, self.description_data = load_descriptions_from_excel()

        # Clear existing listbox items and insert new ones
        self.listbox.delete(0, tk.END)
        for d in self.descriptions:
            self.listbox.insert(tk.END, d)

        # Reset used tags and renamed files as the source data has changed
        self.used_tags.clear()
        self.renamed_files.clear()
        self.dont_know_files.clear() # Clear "don't know" flags

        # Re-scan image folder to set initial image_files order to current alphabetical disk order
        self.image_files = sorted(
            f for f in os.listdir(IMAGE_FOLDER)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        )
        self.current_index = 0 if self.image_files else 0 # Reset current image index

        self._update_list_colors()
        self._update_counter()
        self.display_image() # Re-display the current image
        self.save_progress() # Save the new progress state

    def _parse_pasted_data(self, raw_text):
        """
        Parse pasted text rows copied from Google Sheets entire rows.
        Extract columns:
          C = index 2 (Sticker Number)
          D = index 3 (Description)
          P = index 15 (Object Number)
          Q = index 16 (Imported Description)
          U = index 20 (Location)
        Returns list of dicts with keys for all relevant data.
        Skips rows with missing Description.
        """
        def sanitize_for_filename(text):
            """Sanitizes text to be suitable for a filename part."""
            if not isinstance(text, str):
                return ""
            # Replace common problematic characters with underscore
            # Keep alphanumeric, spaces, hyphens, and dots. Replace others with _
            sanitized = "".join(c if c.isalnum() or c in (' ', '-', '_', '.') else '_' for c in text).strip()
            # Replace multiple spaces/underscores with a single underscore
            sanitized = '_'.join(filter(None, sanitized.split())) # Handles spaces, and ensures single underscores
            sanitized = '_'.join(filter(None, sanitized.split('_'))) # Handles existing underscores
            return sanitized


        rows = []
        lines = [line.strip() for line in raw_text.strip().splitlines() if line.strip()]
        for line in lines:
            # Try tab split first, else comma
            if "\t" in line:
                parts = line.split("\t")
            else:
                parts = line.split(",")

            # Must have at least enough columns for all required indices (index 20 = Col U)
            if len(parts) <= 20:  
                continue

            # Extract data, handling potential empty strings for non-mandatory fields
            sticker_num = parts[2].strip() if len(parts) > 2 else ''
            desc = parts[3].strip() if len(parts) > 3 else ''
            obj_num = parts[15].strip() if len(parts) > 15 else ''
            imported_desc = parts[16].strip() if len(parts) > 16 else ''
            loc = parts[20].strip() if len(parts) > 20 else ''

            if not desc: # Description (D) is mandatory for a valid entry
                continue

            # Fallback logic for Object Number if empty
            if not obj_num:
                sanitized_fallback_desc = sanitize_for_filename(desc) # Sanitize description for filename
                sanitized_fallback_loc = sanitize_for_filename(loc)   # Sanitize location for filename

                if sanitized_fallback_loc and sanitized_fallback_desc:
                    obj_num = f"{sanitized_fallback_loc}_{sanitized_fallback_desc}"
                elif sanitized_fallback_loc:
                    obj_num = sanitized_fallback_loc
                elif sanitized_fallback_desc:
                    obj_num = sanitized_fallback_desc
                else:
                    obj_num = "Unknown_Object" # Generic fallback if both are empty

            rows.append({
                'Description': desc,
                'Object Number': obj_num, # This is now the sanitized fallback or original object number
                'Sticker Number': sticker_num,
                'Imported Description': imported_desc,
                'Location': loc
            })
        return rows
        
    # ────────────────────────────────────────────────────────────
    # Tag info popup
    # ────────────────────────────────────────────────────────────
    
    def _show_tag_info(self, event):
        """Displays a pop-up with detailed information for the right-clicked tag."""
        try:
            # Get the index of the clicked item
            index = self.listbox.nearest(event.y)
            # Ensure the index is valid
            if 0 <= index < len(self.descriptions):
                selected_desc_raw = self.listbox.get(index)
                selected_desc = selected_desc_raw.strip() # <--- Crucial: Strip whitespace from listbox selection

                details = self.description_data.get(selected_desc) # Use the stripped version for lookup

                if details:
                    # Create a new Toplevel window for the info
                    info_window = tk.Toplevel(self.root)
                    info_window.title(selected_desc) # Use the stripped version for title
                    info_window.transient(self.root) # Make it appear on top of main window
                    info_window.resizable(False, False)

                    def add_info_row(parent, title, value):
                        display_value = str(value) if value is not None else "N/A"
                        tk.Label(parent, text=title, font=("TkDefaultFont", 10, "bold"), anchor="w").pack(fill="x", padx=5, pady=(2,0))
                        tk.Label(parent, text=display_value, wraplength=300, anchor="w", justify="left").pack(fill="x", padx=5, pady=(0,2))

                    add_info_row(info_window, "Description:", selected_desc)
                    add_info_row(info_window, "Object Number:", details.get("Object Number", "N/A"))
                    add_info_row(info_window, "Sticker Number:", details.get("Sticker Number", "N/A"))
                    add_info_row(info_window, "Imported Description:", details.get("Imported Description", "N/A"))
                    add_info_row(info_window, "Location:", details.get("Location", "N/A"))

                    # Close button
                    tk.Button(info_window, text="Close", command=info_window.destroy).pack(pady=10)

                    # Update idletasks to calculate geometry after widgets are packed
                    info_window.update_idletasks()
                    
                    # Position the window near the mouse click after sizing is calculated
                    width = info_window.winfo_width()
                    height = info_window.winfo_height()
                    x_pos = event.x_root + 10
                    y_pos = event.y_root + 10

                    # Prevent window from going off-screen to the right or bottom
                    screen_width = self.root.winfo_screenwidth()
                    screen_height = self.root.winfo_screenheight()

                    if x_pos + width > screen_width:
                        x_pos = screen_width - width - 20 # 20px margin from right edge
                    if y_pos + height > screen_height:
                        y_pos = screen_height - height - 20 # 20px margin from bottom edge

                    info_window.geometry(f'+{x_pos}+{y_pos}')

                    # Call grab_set and focus_set AFTER the window is mapped
                    info_window.after(1, info_window.grab_set)
                    info_window.after(1, info_window.focus_set)
                    info_window.protocol("WM_DELETE_WINDOW", info_window.destroy) # Ensure proper close on X button
                else:
                    messagebox.showwarning("Info Not Found", 
                                           f"No detailed information found for '{selected_desc_raw}'.\n"
                                           "This usually means the description in your Excel file doesn't match exactly. "
                                           "Check for extra spaces or differing text.")
        except tk.TclError:
            pass # Or handle more gracefully if needed

    def _on_progressbar_click(self, event):
        """Jumps to an image in the list based on click position on the progress bar."""
        if not self.image_files:
            return

        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width == 0: # Avoid division by zero if widget not yet rendered
            return
        
        click_x = event.x
        total_images = len(self.image_files)
        
        # Calculate the new index based on the click's horizontal position
        # new_index will be 0-based
        new_index = int((click_x / canvas_width) * total_images)
        
        # Clamp the index to valid range
        new_index = max(0, min(total_images - 1, new_index))
        
        self.current_index = new_index
        self.display_image()
        self.save_progress() # Call the now-existing instance method

    def _on_progress_canvas_configure(self, event):
        """Redraws the progress blocks when the canvas is resized."""
        self._draw_progress_blocks()

    def _draw_progress_blocks(self):
        """Draws individual blocks on the progress canvas representing image statuses."""
        self.progress_canvas.delete("all") # Clear existing drawings

        canvas_width = self.progress_canvas.winfo_width()
        canvas_height = self.progress_canvas.winfo_height()
        total_images = len(self.image_files)

        if total_images == 0 or canvas_width == 0 or canvas_height == 0:
            return

        block_width = canvas_width / total_images

        for i, filename in enumerate(self.image_files):
            x1 = i * block_width
            x2 = (i + 1) * block_width
            
            fill_color = "gray" # Default for untagged
            outline_color = "" # No outline by default
            line_width = 1

            # Check if original filename is in renamed_files values, or if new filename is a key
            is_renamed = False
            for original_name, new_name in self.renamed_files.items():
                # For Windows case-insensitivity: compare lowercased versions for existing files
                if filename.lower() == new_name.lower(): # current file is a new name
                    is_renamed = True
                    break
                if filename.lower() == original_name.lower() and new_name.lower() != original_name.lower(): # current file is an original name that was renamed
                    is_renamed = True
                    break

            if is_renamed:
                fill_color = "blue"
            
            if filename in self.dont_know_files: # dont_know_files stores actual filenames, so exact match is fine
                fill_color = "red"
            
            if i == self.current_index:
                outline_color = "yellow"
                line_width = 2 # Thicker outline for current image

            self.progress_canvas.create_rectangle(x1, 0, x2, canvas_height, 
                                                    fill=fill_color, outline=outline_color, width=line_width)


    # ────────────────────────────────────────────────────────────
    # Image display helpers
    # ────────────────────────────────────────────────────────────
    def display_image(self):
        # Find the next valid image index if current one is missing
        if self.image_files:
            original_current_index = self.current_index
            found_valid_file = False
            for _ in range(len(self.image_files)): # Iterate up to length of list to find a valid file
                if os.path.exists(os.path.join(IMAGE_FOLDER, self.image_files[self.current_index])):
                    found_valid_file = True
                    break
                self.current_index = (self.current_index + 1) % len(self.image_files)
                if self.current_index == original_current_index: # Looped completely without finding one
                    break
            
            if not found_valid_file:
                messagebox.showerror("No images found", "All images in your list are missing from the folder or are inaccessible.")
                self.image_files = [] # Clear the list as it contains no valid files
                self.current_index = 0

        if not self.image_files:
            self.original_img = None
            self.canvas.delete("all")
            self.canvas.config(bg="black") # Ensure canvas is black if no image
            self._update_counter() # Will show "No images loaded."
            self._draw_progress_blocks() # Update empty progress bar
            return

        path = os.path.join(IMAGE_FOLDER, self.image_files[self.current_index])
        self.original_img = Image.open(path)

        # Determine canvas background color based on rename status
        current_filename = self.image_files[self.current_index]
        
        # Check against renamed_files values (new filenames) and dont_know_files
        is_renamed_visual = False
        for old_name, new_name in self.renamed_files.items():
            if current_filename.lower() == new_name.lower():
                is_renamed_visual = True
                break

        if current_filename in self.dont_know_files: # Check for "don't know" status (highest priority)
            self.canvas.config(bg="red")
        elif is_renamed_visual: 
            self.canvas.config(bg="blue")
        else: # Default for untagged
            self.canvas.config(bg="black")

        self._fit_to_window()
        self._update_counter()
        self._update_list_colors()
        self._draw_progress_blocks() # Redraw progress bar with current status


    def _fit_to_window(self):
        self.root.update_idletasks()
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 10 or ch < 10:
            self.root.after(50, self._fit_to_window)
            return
        if self.original_img: # Check if an image is loaded
            iw, ih = self.original_img.size
            self.zoom = min((cw-40)/iw, (ch-40)/ih, 1.0)  # 20-px margin
            self.offset = [(cw - iw*self.zoom)/2, (ch - ih*self.zoom)/2]
            self._apply_zoom_to_slider()
            self._render()
        else: # Clear canvas if no image
            self.canvas.delete("all")


    def _render(self):
        if not self.original_img: # Check if an image is loaded
            return

        iw, ih = self.original_img.size
        w, h = int(iw*self.zoom), int(ih*self.zoom)
        if w < 1 or h < 1: # Prevent errors with tiny sizes
            return
        img = self.original_img.resize((w, h), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(self.offset[0], self.offset[1], anchor=tk.NW, image=self.tk_img)

    # ─── Zoom logic ───────────────────────────────────────────────
    def zoom_relative(self, factor: float, centre: tuple | None = None):
        if not self.original_img: # Only zoom if an image is loaded
            return

        old_zoom = self.zoom
        new_zoom = max(self.ZOOM_MIN, min(self.ZOOM_MAX, old_zoom * factor))
        factor = new_zoom / old_zoom
        if abs(factor - 1.0) < 1e-3:
            return
        if centre is None:
            cw, ch = self.canvas.winfo_width()/2, self.canvas.winfo_height()/2
            centre = (cw, ch)
        cx, cy = centre
        # translate offset so that the point under 'centre' stays put
        ox, oy = self.offset
        self.offset[0] = cx - (cx - ox)*factor
        self.offset[1] = cy - (cy - oy)*factor
        self.zoom = new_zoom
        self._apply_zoom_to_slider()
        self._render()

    def _slider_zoom(self, val):
        target = float(val)/100.0 # val from slider is already stepped by resolution=10
        self.zoom_relative(target / self.zoom)

    def _apply_zoom_to_slider(self):
        self.zoom_slider.set(int(self.zoom*100))

    def _on_mousewheel(self, event):
        # Determine if scrolling up or down
        up = (event.delta > 0) if hasattr(event, "delta") else (event.num == 4)

        # Determine the zoom factor based on scroll direction
        zoom_factor = self.ZOOM_STEP if up else 1 / self.ZOOM_STEP

        # Get the current mouse position as the zoom center
        centre = (event.x, event.y)

        # Apply the zoom
        self.zoom_relative(zoom_factor, centre=centre)


    # ─── Panning ─────────────────────────────────────────────────
    def _start_pan(self, event):
        self.drag_start = (event.x, event.y)

    def _do_pan(self, event):
        if self.drag_start and self.original_img: # Only pan if an image is loaded
            dx = event.x - self.drag_start[0]
            dy = event.y - self.drag_start[1]
            self.offset[0] += dx
            self.offset[1] += dy
            self.drag_start = (event.x, event.y)
            self._render()
    
    # ─── Tag list / progress helpers ──
    def _update_list_colors(self):
        for i, desc in enumerate(self.descriptions):
            clr = "blue" if desc in self.used_tags else "black"
            self.listbox.itemconfig(i, fg=clr)

    def _update_counter(self):
        total_images = len(self.image_files)
        current_image_num = self.current_index + 1

        # Determine text color for progress and filename labels
        text_color = "black"
        if total_images > 0 and 0 <= self.current_index < total_images:
            current_filename = self.image_files[self.current_index]
            
            # Check against renamed_files values (new filenames) and dont_know_files for label color
            is_renamed_visual = False
            for old_name, new_name in self.renamed_files.items():
                if current_filename.lower() == new_name.lower():
                    is_renamed_visual = True
                    break

            # Prioritize "don't know" (red) over renamed (blue) for filename label
            if current_filename in self.dont_know_files:
                text_color = "red"
            elif is_renamed_visual:
                text_color = "blue"
            
        self.progress_label.config(fg=text_color)
        self.filename_label.config(fg=text_color)

        # Update progress label
        if total_images > 0 and 0 <= self.current_index < total_images:
            current_filename = self.image_files[self.current_index]
            self.progress_label.config(text=f"Image {current_image_num} / {total_images}")
            self.filename_label.config(text=f"Current: {current_filename}")
        else:
            self.progress_label.config(text="No images loaded.")
            self.filename_label.config(text="") # Clear filename if no images


    # ─── Navigation ------------------------------------------------
    def prev_image(self):
        if not self.image_files: return
        
        # Loop backward to find the previous valid image, wrapping around if necessary
        original_index = self.current_index
        attempts = 0
        while attempts < len(self.image_files):
            self.current_index = (self.current_index - 1 + len(self.image_files)) % len(self.image_files)
            if os.path.exists(os.path.join(IMAGE_FOLDER, self.image_files[self.current_index])):
                self.display_image()
                self.save_progress() # Call the now-existing instance method
                return # Found a valid file
            attempts += 1
        
        # If loop finishes, no valid image was found
        messagebox.showwarning("No Valid Images", "No previous valid image found. All images might be missing or inaccessible.")
        self.image_files = [] # Clear the list as it contains no valid files
        self.current_index = 0
        self.display_image() # This will call _update_counter and _draw_progress_blocks to reflect empty state

    def next_image(self):
        if not self.image_files: return
        
        # Loop forward to find the next valid image, wrapping around if necessary
        original_index = self.current_index
        attempts = 0
        while attempts < len(self.image_files):
            self.current_index = (self.current_index + 1) % len(self.image_files)
            if os.path.exists(os.path.join(IMAGE_FOLDER, self.image_files[self.current_index])):
                self.display_image()
                self.save_progress() # Save current state after changing image
                return # Found a valid file
            attempts += 1
        
        # If loop finishes, no valid image was found
        messagebox.showwarning("No Valid Images", "No next valid image found. All images might be missing or inaccessible.")
        self.image_files = [] # Clear the list as it contains no valid files
        self.current_index = 0
        self.display_image() # This will call _update_counter and _draw_progress_blocks to reflect empty state

    # ────────────────────────────────────────────────────────────
    # Tagging and Renaming
    # ────────────────────────────────────────────────────────────
    def tag_and_rename(self):
        if not self.image_files:
            messagebox.showwarning("No Image", "No image to tag.")
            return

        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select a description from the list.")
            return

        selected_description = self.descriptions[selected_indices[0]].strip()
        current_filename = self.image_files[self.current_index]
        base_name_current, ext_current = os.path.splitext(current_filename)

        # Get detailed info for the selected description
        details = self.description_data.get(selected_description)
        if not details:
            messagebox.showerror("Error", f"Could not find details for description: '{selected_description}'")
            return

        object_number = details.get("Object Number", "")
        # Ensure object_number is safe for filenames
        safe_object_number = "".join(c if c.isalnum() or c in (' ', '-', '_', '.') else '' for c in object_number).strip()
        safe_object_number = safe_object_number.replace(' ', '_')

        # Standardize the extension to lowercase for cross-platform consistency
        ext = ext_current.lower()

        # Initial new filename attempt
        new_filename_base = f"{safe_object_number}"
        new_filename = f"{new_filename_base}{ext}" # Combine base name with standardized extension
        
        old_path = os.path.join(IMAGE_FOLDER, current_filename)
        new_path = os.path.join(IMAGE_FOLDER, new_filename)

        # Check if the proposed new_filename (without suffix) is identical to the current filename
        # This handles cases where the file is already named correctly.
        if os.path.normcase(old_path) == os.path.normcase(new_path):
            messagebox.showinfo("No Change", f"File is already named '{new_filename}'.")
            self.used_tags.add(selected_description)
            self.dont_know_files.discard(current_filename) # Remove from don't know if tagged
            self.save_progress()
            self._update_list_colors()
            self._draw_progress_blocks() # Redraw for visual feedback even if no rename
            self.next_image() # Auto-advance even if no rename happened
            return

        # Handle duplicate filenames by adding a suffix (e.g., _1, _2)
        suffix = 1
        proposed_new_path = new_path # Start with the non-suffixed new_path

        # Loop to find a unique filename if a duplicate exists
        # Use os.path.normcase for case-insensitive comparison on Windows
        while os.path.exists(proposed_new_path) and os.path.normcase(proposed_new_path) != os.path.normcase(old_path):
            # The condition `os.path.normcase(proposed_new_path) != os.path.normcase(old_path)`
            # is crucial. It ensures we only add a suffix if the existing file is *not*
            # the current file itself (in a case-insensitive way on Windows).
            
            new_filename = f"{new_filename_base}_{suffix}{ext}"
            proposed_new_path = os.path.join(IMAGE_FOLDER, new_filename)
            suffix += 1

        # After the loop, `proposed_new_path` is the unique path we should use.
        # Update `new_filename` and `new_path` to reflect the final chosen name (with suffix if any).
        new_path = proposed_new_path
        # Extract the filename part from the final `new_path`
        new_filename = os.path.basename(new_path)

        try:
            # Rename the file
            shutil.move(old_path, new_path)
    #        messagebox.showinfo("Success", f"Renamed '{current_filename}' to '{new_filename}'")

            # Update internal state
            self.used_tags.add(selected_description)
            self.renamed_files[current_filename] = new_filename # Store old -> new mapping
            self.dont_know_files.discard(current_filename) # Remove from don't know if tagged

            # Update the image_files list with the new filename
            self.image_files[self.current_index] = new_filename

            self.save_progress() # Save after updating internal state
            
            # Refresh UI
            self._update_list_colors()
            self.next_image() # Auto-advance to the next image
            
        except FileNotFoundError:
            messagebox.showerror("Error", f"Original file not found: {old_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {e}")

    def undo_rename(self):
        if not self.image_files:
            messagebox.showwarning("No Image", "No image to undo rename for.")
            return

        current_filename = self.image_files[self.current_index]
        
        original_filename = None
        # Iterate through renamed_files to find the original name, comparing case-insensitively
        for old_name, new_name in self.renamed_files.items():
            if current_filename.lower() == new_name.lower(): # Match current file (new name) to a stored new_name
                original_filename = old_name
                break
        
        if not original_filename:
            # Check if the current file *was* the original name of a renamed file that now exists under a new name
            # This handles cases where you've moved past a renamed file and come back to its new name
            # This check is a bit complex due to potential case differences, so relying primarily on the above.
            # However, if it's not a value (new name), it means it's either untouched or an original name.
            is_original_name_in_map = False
            for old_n, new_n in self.renamed_files.items():
                if current_filename.lower() == old_n.lower() and current_filename.lower() != new_n.lower():
                    is_original_name_in_map = True
                    break
            
            if is_original_name_in_map:
                messagebox.showinfo("Not Renamed (Currently)", f"'{current_filename}' was renamed. To undo, you need to be on the renamed file itself (e.g., 'OBJECT_XYZ.jpg'), not its original name.")
                return

            messagebox.showinfo("Not Renamed", f"'{current_filename}' has not been renamed by this application.")
            return

        old_path = os.path.join(IMAGE_FOLDER, current_filename)
        new_path = os.path.join(IMAGE_FOLDER, original_filename)

        if not os.path.exists(old_path):
            messagebox.showerror("File Missing", f"The current file '{current_filename}' is missing from the '{IMAGE_FOLDER}' folder.")
            # Attempt to clean up renamed_files if the new file is gone
            if original_filename in self.renamed_files and self.renamed_files[original_filename].lower() == current_filename.lower():
                del self.renamed_files[original_filename]
                self.save_progress()
                self._draw_progress_blocks()
            return
        
        # If the target original filename exists, confirm overwrite, considering case-insensitivity
        if os.path.exists(new_path) and os.path.normcase(new_path) != os.path.normcase(old_path):
            overwrite = messagebox.askyesno(
                "File Exists",
                f"A file named '{original_filename}' already exists.\nOverwrite it?"
            )
            if not overwrite:
                return

        try:
            shutil.move(old_path, new_path)
            messagebox.showinfo("Success", f"Undid rename: '{current_filename}' reverted to '{original_filename}'")

            # Update internal state
            if original_filename in self.renamed_files:
                del self.renamed_files[original_filename] # Remove the mapping

            # Remove from used_tags - this logic is fine as used_tags stores descriptions, not filenames
            # self.used_tags = {tag for tag in self.used_tags if tag not in self.descriptions} # This line seems incorrect

            self.dont_know_files.discard(current_filename) # If it was marked don't know, remove it
            self.dont_know_files.discard(original_filename) # Ensure original is also not marked as don't know

            # Update the image_files list with the original filename
            self.image_files[self.current_index] = original_filename
            

            self.save_progress()
            self.display_image() # Re-display the current image, which is now original_filename
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to undo rename: {e}")

    def send_to_end(self):
        if not self.image_files:
            messagebox.showwarning("No Image", "No image to move.")
            return

        current_filename = self.image_files[self.current_index]

        if current_filename in self.dont_know_files:
            messagebox.showinfo("Already Marked", f"'{current_filename}' is already marked as 'Don't know'.")
            self.next_image() # Still advance
            return

        # Add to the "don't know" set
        self.dont_know_files.add(current_filename)
        
        # Remove from renamed_files if it was previously renamed, as "don't know" takes precedence
        # Find if this file was a *new name* from a previous rename operation, considering case
        original_name_of_current = None
        for old_n, new_n in list(self.renamed_files.items()): # Use list() to allow modification during iteration
            if current_filename.lower() == new_n.lower():
                original_name_of_current = old_n
                break
        
        if original_name_of_current:
            del self.renamed_files[original_name_of_current] # Remove the entry from renamed_files
            
        # Re-order image_files: move current image to the end
        self.image_files.pop(self.current_index)
        self.image_files.append(current_filename)

        # The current_index stays the same logically, as the image that was *at* this index
        # is now gone, and the next image in the list naturally moves into its place.
        # If the list is empty after moving, set index to 0.
        if not self.image_files:
            self.current_index = 0
        elif self.current_index >= len(self.image_files):
            self.current_index = 0 # If it was the last item, loop back to start

        messagebox.showinfo("Moved", f"'{current_filename}' marked as 'Don't know' and moved to end of list.")
        
        self.save_progress() # Save after moving and updating state
        self.display_image() # Refresh display (will naturally show the new image at self.current_index)


if __name__ == "__main__":
    # Create the Images folder if it doesn't exist
    if not os.path.exists(IMAGE_FOLDER):
        os.makedirs(IMAGE_FOLDER)

    # Create a dummy Excel file if it doesn't exist for first run
    if not os.path.exists(EXCEL_FILE):
        df_empty = pd.DataFrame(columns=["Description", "Object Number", "Sticker Number", "Imported Description", "Location"])
        df_empty.to_excel(EXCEL_FILE, index=False)
        print(f"Created empty {EXCEL_FILE}. Please populate it with descriptions.")

    root = tk.Tk()
    app = ImageTagger(root)
    root.mainloop()
