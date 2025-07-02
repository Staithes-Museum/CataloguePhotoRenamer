# 📸 CataloguePhotoRenamer

Used at **Staithes Museum** to rename photographs of museum objects for our digital catalogue.

---

## 🧾 About this Programme

As part of our cataloguing process, we take a lot of photos of our museum objects. But keeping those photos organised, searchable, and easy to use? That’s a whole other challenge.

After a lot of head-scratching, we settled on a simple but powerful system: every photo file gets renamed using the object’s catalogue number (like `SSESM.2019.45`) followed by a brief description of the object. It’s a tidy solution—but renaming hundreds of photos by hand? NVery repetitive, and prone to errors.  So… we built our own software to do it for us!

This tool:

- Imports object numbers and descriptions from a spreadsheet
- Automatically renames image files to match
- Saves hours of time
- Reduces human error
- Keeps everything beautifully consistent

---

## 👩‍💻 Want to Try It?

Download it for free from the **Staithes Museum GitHub**.  
_Note: This application is designed and tested for **Linux** only. It won’t work on Windows._

---

## 🗂️ Project Structure

```
CataloguePhotoRenamer/
├── tagger.py               # Main application script
├── processimages.py        # Optional: Image utility functions (resize, format conversion)
├── Object numbers.xlsx     # Your spreadsheet of object metadata
├── Images/                 # Folder containing your photos (.jpg, .jpeg, .png)
└── progress.json           # Automatically created to save progress
```

---

## ⚙️ Dependencies

Install required libraries using pip:

```bash
pip install Pillow pandas openpyxl
```

- `Pillow`: Image handling
- `pandas`: Spreadsheet processing
- `openpyxl`: Enables reading `.xlsx` files via `pandas`

---

## 📄 Preparing Your Excel File

The application reads metadata from `Object numbers.xlsx`.

Required columns (case-insensitive, no leading/trailing spaces):

| Column Name          | Purpose                                                                 |
|----------------------|-------------------------------------------------------------------------|
| **Description**      | Required. Appears in app listbox. Must be unique & concise.             |
| **Object Number**    | Primary identifier for filenames and folder structure                   |
| **Sticker Number**   | Additional ID, included in filenames                                    |
| **Imported Description** | Optional extra detail (limit: 50 characters)                        |
| **Location**         | Used for filename/folder if Object Number is missing                    |

💡 If you're copy/pasting rows from another spreadsheet, the program expects the following columns:

- Column C → Sticker Number  
- Column D → Description  
- Column P → Object Number  
- Column Q → Imported Description  
- Column U → Location

An example Excel file is provided: `Sample object details.xlsx`.

---

## 🖼️ Adding Images

Place your image files (`.jpg`, `.jpeg`, `.png`) into the `Images/` folder.

---

## ▶️ Running the Application

In the terminal, navigate to the project folder and run:

```bash
python3 tagger.py
```

---

## 🖥️ Application Interface

![App Screenshot](https://github.com/user-attachments/assets/a0b752c9-cdb2-4bc9-af6d-97f1b9faff83)

### Main Layout:

- **Left Panel**: Displays the image being tagged
- **Right Panel**: Contains navigation, description list, and controls

### Progress Bar (Top):

- **Grey** = Untagged  
- **Blue** = Tagged  
- **Red** = “Don’t know”  
- **Yellow border** = Current image

Click anywhere on the bar to jump to that part of your collection.

---

## 🧭 Basic Controls

### Navigation

- `← Previous` / `→ Next`: Browse through images
- Click the **progress bar** to jump

### Zoom & Pan

- `+` / `-`: Zoom in and out
- `Ctrl` + `+` or `-`: Zoom with keyboard
- **Mouse Wheel**: Zoom while hovering
- **Drag**: Pan the image
- **Double Click**: Reset zoom
- **Slider**: Vertical zoom control

### Tagging an Image

1. Select a description from the right-hand list  
2. Click `💾 Tag & Rename` or double-click the description  
3. The image will be renamed and moved automatically  
4. The app moves to the next image

**Blue text** = Already-used description  
**Right-click** = View full description and object details

---

## 📁 Renaming & Folder Logic

### Folder Destination Rules

- If **Object Number** exists → use it to rename and sort  
- If only **Location** exists → use Location in filename/folder (slashes converted to `.`)
- If both are missing → image stays in Images/ folder with `untagged_` prefix

### Filename Conflicts

- If filename exists → app appends `_1`, `_2`, etc. to avoid clashes  
- If filename already matches target → no action is taken

---

## ✨ Extra Features

### ❓ Don't Know (Send →)

- Marks image with red background
- Skips it for now and returns to it later
- Click again to unmark

### 🔁 Undo Rename

- Reverts last renaming of current image  
- Restores original name/location  
- _Note: This only works if files weren’t moved or deleted manually_

### Auto Save

Progress (via `progress.json`) is saved after:
- Renaming
- Marking “Don’t know”
- Navigating between images

---

## 🛠️ Troubleshooting

| Issue | Solution |
|-------|----------|
| **No images found** | Make sure your image files are in `Images/` and are `.jpg`, `.jpeg`, or `.png` |
| **Excel file not found / Excel error** | Ensure `Object numbers.xlsx` is in the project folder and not open elsewhere |
| **Missing Description** | Check your Excel file; rows without a description are skipped |
| **Image won’t rename/move** | Check for duplicates or file permission issues |
| **Black image or failed load** | Try reinstalling Pillow: `pip install --upgrade Pillow` |

---

## 📣 Credits

Built by the team at **Staithes Museum**  
We hope this saves you time and headache — it certainly did for us!
