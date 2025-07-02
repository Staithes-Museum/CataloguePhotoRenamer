import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
from pillow_heif import register_heif_opener # Import the HEIF opener
import subprocess # New: For running external scripts
import sys # New: For correctly calling bundled Python scripts

# Register the HEIF opener so Pillow can read HEIC files
register_heif_opener()

# --- Configuration for Tagger Script (ensure these match tagger.py) ---
TAGGER_IMAGE_FOLDER_NAME = "Images" # The folder tagger.py looks for
TAGGER_SCRIPT_NAME = "tagger.py"    # The name of your tagger script

class ImageProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Image Processing Tool")

        self.input_dir = ""
        self.output_dir = "ProcessedImages" # Default output directory

        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Labels and Entry for Input Directory
        self.label_input = tk.Label(master, text="Select Input Image Folder:")
        self.label_input.grid(row=1, column=0, padx=5, pady=5, sticky="w") # Shifted row for new button

        self.entry_input = tk.Entry(master, width=50)
        self.entry_input.grid(row=1, column=1, padx=5, pady=5)

        self.btn_browse = tk.Button(master, text="Browse", command=self.browse_input_directory)
        self.btn_browse.grid(row=1, column=2, padx=5, pady=5)

        # New: Button to run tagger.py first
        self.btn_run_tagger = tk.Button(master, text="Rename object images first!", command=self.run_tagger_script)
        self.btn_run_tagger.grid(row=0, column=0, columnspan=3, pady=10, ipadx=10, ipady=5) # Placed at the top


        # Buttons for Image Operations (shifted down)
        self.btn_create_tiny = tk.Button(master, text="Create Tiny Thumbnails (150px wide)", command=self.create_tiny_thumbnails)
        self.btn_create_tiny.grid(row=2, column=0, columnspan=3, pady=5)

        self.btn_create_small = tk.Button(master, text="Create Small Thumbnails (1024px wide)", command=self.create_small_thumbnails)
        self.btn_create_small.grid(row=3, column=0, columnspan=3, pady=5)

        self.btn_convert_tif_to_jpg = tk.Button(master, text="Convert .tif to .jpg", command=self.convert_tif_to_jpg)
        self.btn_convert_tif_to_jpg.grid(row=4, column=0, columnspan=3, pady=5)

        # New: HEIC to JPEG conversion button
        self.btn_convert_heic_to_jpg = tk.Button(master, text="Convert .heic to .jpg", command=self.convert_heic_to_jpg)
        self.btn_convert_heic_to_jpg.grid(row=5, column=0, columnspan=3, pady=5)

        # New: Rename TEMP_123 to TEMP.123 button
        self.btn_rename_temp_files = tk.Button(master, text="Rename TEMP_XXX.ext to TEMP.XXX.ext", command=self.rename_temp_files)
        self.btn_rename_temp_files.grid(row=6, column=0, columnspan=3, pady=5)


        # Status Label
        self.status_label = tk.Label(master, text="Status: Ready", fg="blue")
        self.status_label.grid(row=7, column=0, columnspan=3, pady=5)


    def browse_input_directory(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_dir = folder_selected
            self.entry_input.delete(0, tk.END)
            self.entry_input.insert(0, self.input_dir)
            self.status_label.config(text=f"Status: Input directory set to {self.input_dir}", fg="blue")

    def _get_cleaned_base_name(self, original_filepath):
        """
        Extracts the base name from a filepath and removes common suffixes
        (-small, -tiny) and certain format extensions (like .heic, .tif)
        if they are embedded in the base name itself.
        """
        base_name = os.path.basename(original_filepath)
        name_without_ext, _ = os.path.splitext(base_name)

        # Remove any existing -small or -tiny suffixes before applying new ones
        if name_without_ext.endswith("-small"):
            name_without_ext = name_without_ext[:-len("-small")]
        if name_without_ext.endswith("-tiny"):
            name_without_ext = name_without_ext[:-len("-tiny")]
        
        # Remove common image extensions if they somehow got into name_without_ext
        current_name_lower = name_without_ext.lower()
        for ext_to_remove in ('.heic', '.heif', '.tif', '.tiff', '.jpg', '.jpeg', '.png', '.bmp'):
            if current_name_lower.endswith(ext_to_remove):
                name_without_ext = name_without_ext[:-len(ext_to_remove)]
                break
                
        return name_without_ext

    def process_images(self, suffix, target_width=None, convert_to_jpg=False, specific_exts=None):
        if not self.input_dir:
            messagebox.showwarning("Warning", "Please select an input directory first.")
            return

        supported_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.heic', '.heif')
        processed_count = 0
        skipped_count = 0

        self.status_label.config(text="Status: Processing...", fg="orange")
        self.master.update_idletasks() # Update GUI immediately

        for root, _, files in os.walk(self.input_dir):
            for filename in files:
                file_path = os.path.join(root, filename)
                original_file_ext = os.path.splitext(filename)[1].lower()

                if specific_exts and original_file_ext not in specific_exts:
                    continue # Skip if specific extensions are required and doesn't match

                if not specific_exts and original_file_ext not in supported_extensions:
                    continue # Skip if not a supported image file (for general ops)

                try:
                    # 1. Get the cleaned base name (e.g., "TEMP.123" from "TEMP_123.jpg")
                    cleaned_base_name = self._get_cleaned_base_name(file_path)

                    # 2. Determine the final output extension
                    final_output_ext = ".jpg" if convert_to_jpg else original_file_ext

                    # 3. Construct the full desired output filename including suffix
                    output_filename = f"{cleaned_base_name}{suffix}{final_output_ext}"
                    output_filepath = os.path.join(self.output_dir, output_filename)

                    img = Image.open(file_path)

                    # Resize if a target width is specified
                    if target_width:
                        original_width, original_height = img.size
                        if original_width > target_width: # Only resize if larger than target
                            new_height = int((target_width / original_width) * original_height)
                            img = img.resize((target_width, new_height), Image.LANCZOS)

                    # Save the image
                    if final_output_ext == ".jpg": # If the target output format is JPG
                        img = img.convert("RGB") # Ensure it's in a format compatible with JPG
                        img.save(output_filepath, "jpeg")
                    else: # If keeping original format (e.g., if a future operation saves PNG as PNG)
                        img.save(output_filepath)


                    self.status_label.config(text=f"Processed: {os.path.basename(file_path)} -> {os.path.basename(output_filepath)}")
                    self.master.update_idletasks()
                    processed_count += 1

                except Exception as e:
                    self.status_label.config(text=f"Error processing {filename}: {e}", fg="red")
                    self.master.update_idletasks()
                    skipped_count += 1
                    # print(f"Error processing {file_path}: {e}") # For debugging

        messagebox.showinfo("Processing Complete",
                            f"Finished processing images.\nProcessed: {processed_count}\nSkipped (errors/unsupported): {skipped_count}")
        self.status_label.config(text="Status: Ready", fg="blue")


    def create_tiny_thumbnails(self):
        # Always convert to JPG for thumbnails, and apply the -tiny suffix
        self.process_images(suffix="-tiny", target_width=150, convert_to_jpg=True)

    def create_small_thumbnails(self):
        # Always convert to JPG for thumbnails, and apply the -small suffix
        self.process_images(suffix="-small", target_width=1024, convert_to_jpg=True)

    def convert_tif_to_jpg(self):
        # Convert TIFs to JPGs, no special suffix for the conversion itself
        self.process_images(suffix="", convert_to_jpg=True, specific_exts=['.tif', '.tiff'])

    def convert_heic_to_jpg(self):
        # Convert HEICs to JPGs, no special suffix for the conversion itself
        self.process_images(suffix="", convert_to_jpg=True, specific_exts=['.heic', '.heif'])


    def rename_temp_files(self):
        if not self.input_dir:
            messagebox.showwarning("Warning", "Please select an input directory first.")
            return

        renamed_count = 0
        skipped_count = 0
        self.status_label.config(text="Status: Renaming TEMP_ files...", fg="orange")
        self.master.update_idletasks()

        for root, _, files in os.walk(self.input_dir):
            for filename in files:
                name_without_ext, ext = os.path.splitext(filename)

                # Check if the base name starts with "TEMP_" (case-insensitive)
                # and if the part after "TEMP_" is purely digits.
                if name_without_ext.lower().startswith("temp_"):
                    potential_number_part = name_without_ext[len("temp_"):]
                    
                    if potential_number_part.isdigit():
                        original_filepath = os.path.join(root, filename)
                        
                        # Replace the first underscore in the base name with a dot
                        new_name_without_ext = name_without_ext.replace('_', '.', 1) # Replace only the first occurrence
                        new_filename = new_name_without_ext + ext
                        new_filepath = os.path.join(root, new_filename)

                        try:
                            # Avoid renaming if the file already exists with the new name
                            if os.path.exists(new_filepath):
                                print(f"Skipping rename for '{filename}': '{new_filename}' already exists.")
                                skipped_count += 1
                                continue

                            os.rename(original_filepath, new_filepath)
                            self.status_label.config(text=f"Renamed: {filename} -> {new_filename}")
                            self.master.update_idletasks()
                            renamed_count += 1
                        except OSError as e:
                            self.status_label.config(text=f"Error renaming {filename}: {e}", fg="red")
                            self.master.update_idletasks()
                            skipped_count += 1
                    else:
                        skipped_count += 1
                else:
                    skipped_count += 1

        messagebox.showinfo("Renaming Complete",
                            f"Finished renaming files.\nRenamed: {renamed_count}\nSkipped: {skipped_count}")
        self.status_label.config(text="Status: Ready", fg="blue")

    def run_tagger_script(self):
        """
        Creates the IMAGE_FOLDER if it doesn't exist and then runs the tagger.py script.
        """
        # 1. Create the IMAGE_FOLDER if it doesn't exist
        if not os.path.exists(TAGGER_IMAGE_FOLDER_NAME):
            try:
                os.makedirs(TAGGER_IMAGE_FOLDER_NAME)
                messagebox.showinfo("Folder Created", f"Created folder: '{TAGGER_IMAGE_FOLDER_NAME}'")
            except Exception as e:
                messagebox.showerror("Folder Creation Error", f"Could not create folder '{TAGGER_IMAGE_FOLDER_NAME}': {e}")
                self.status_label.config(text=f"Error creating folder: {e}", fg="red")
                return

        # 2. Try to locate tagger.py
        # When bundled with PyInstaller, sys.executable points to the bundled python interpreter
        # and scripts are often extracted to a temporary directory accessible relative to that.
        # This assumes tagger.py is in the same directory as this main script or its bundle entry.
        tagger_script_path = os.path.join(os.path.dirname(sys.argv[0]), TAGGER_SCRIPT_NAME)

        if not os.path.exists(tagger_script_path):
            # Fallback for development if tagger.py is in current working directory
            if os.path.exists(TAGGER_SCRIPT_NAME):
                tagger_script_path = TAGGER_SCRIPT_NAME
            else:
                messagebox.showerror("Script Not Found", f"'{TAGGER_SCRIPT_NAME}' not found. Please ensure it's in the same directory as this executable/script.")
                self.status_label.config(text=f"Error: {TAGGER_SCRIPT_NAME} not found.", fg="red")
                return

        self.status_label.config(text=f"Status: Running '{TAGGER_SCRIPT_NAME}'...", fg="orange")
        self.master.update_idletasks()

        try:
            # Run the tagger script using the Python interpreter that's running this script
            # This is robust for both development and PyInstaller bundled executables.
            result = subprocess.run([sys.executable, tagger_script_path], check=True, capture_output=True, text=True)
            
            # Display output/errors from tagger.py if any
            if result.stdout:
                print(f"Output from {TAGGER_SCRIPT_NAME}:\n{result.stdout}")
            if result.stderr:
                messagebox.showerror(f"{TAGGER_SCRIPT_NAME} Errors", f"Errors occurred in {TAGGER_SCRIPT_NAME}:\n{result.stderr}")
                self.status_label.config(text=f"'{TAGGER_SCRIPT_NAME}' finished with errors.", fg="red")
            else:
                messagebox.showinfo("Tagger Script Complete", f"'{TAGGER_SCRIPT_NAME}' executed successfully.")
                self.status_label.config(text=f"Status: '{TAGGER_SCRIPT_NAME}' completed.", fg="blue")

        except subprocess.CalledProcessError as e:
            # This catches errors if the tagger script itself exits with a non-zero status
            messagebox.showerror("Tagger Script Error", f"'{TAGGER_SCRIPT_NAME}' failed with exit code {e.returncode}.\nOutput:\n{e.stdout}\nErrors:\n{e.stderr}")
            self.status_label.config(text=f"'{TAGGER_SCRIPT_NAME}' failed.", fg="red")
        except FileNotFoundError:
            # This should ideally be caught by the os.path.exists check, but as a safeguard
            messagebox.showerror("Script Not Found", f"The Python interpreter could not find '{TAGGER_SCRIPT_NAME}'.")
            self.status_label.config(text=f"Error: Python interpreter issue with {TAGGER_SCRIPT_NAME}.", fg="red")
        except Exception as e:
            messagebox.showerror("An Unexpected Error Occurred", f"An unexpected error occurred while trying to run '{TAGGER_SCRIPT_NAME}': {e}")
            self.status_label.config(text=f"Error running {TAGGER_SCRIPT_NAME}: {e}", fg="red")

# Main part of the script
if __name__ == "__main__":
    root = tk.Tk()
    app = ImageProcessorApp(root)
    root.mainloop()
