# 📁 FileWhip Excel Tool — Recursive File Audit & Categorization

## Overview
**FileWhip** is a VBA-powered Excel tool designed to:
- Recursively scan folders and subfolders
- List all files with metadata (name, path, size, type, last modified)
- Flag duplicates, clutter, and old files
- Categorize files by type (e.g., Document, Image, Archive)
- Generate a summary by category
- Offer flexible file relocation options
- Log all moves and support undo operations

---

## 🔧 Setup Instructions

### 1. Open the Excel Workbook
- Launch Excel and open the workbook containing the FileWhip VBA macros.

### 2. Enable Macros
- When prompted, click **Enable Content** to allow macros to run.
- If macros are disabled, go to:
  `File > Options > Trust Center > Trust Center Settings > Macro Settings`  
  and enable **"Enable all macros"** (recommended only for trusted files).

---

## 🚀 How to Use

*** You can use the UI in the "Instructions" tab or see below: ***

### Step 1: Run the File Listing
- Press `Alt + F8` to open the **Macro dialog**.
- Select `ListFilesRecursively_Tabbed` and click **Run**.
- Choose the folder you want to scan.
- The tool will:
  - Create one or more sheets named `FileList_1`, `FileList_2`, etc.
  - List all files with metadata
  - Flag files for review
  - Categorize by type
  - Generate a `CategorySummary` sheet

### Step 2: Review Recommendations
- In each `FileList_*` sheet:
  - Column A: Action recommendations (e.g., "Review: Duplicate", "Old file")
  - Column B: Duplicate reference path
  - Highlight colors:
    - 🟧 Orange: Duplicate
    - 🩶 Gray: Clutter type
    - 🟥 Red: Old file
    - 🟨 Yellow: Marked for move
    - 🟩 Green: Successfully moved

### Step 3: Move Files
- This action will **only** move 🟨 Yellow highlighted filenames in Column C
- Press `Alt + F8`, run `MoveMarkedFilesWithOptions`.
- Choose a destination folder.
- Confirm whether to move:
  - Individual files
  - File’s containing folder
  - Parent folder of the file’s folder
- A `MoveLog` sheet will be created to track all changes.

### Step 4: Undo Moves (Optional)
- To reverse any moves:
  - Highlight entries in `MoveLog` (Column A) with **yellow fill**
  - Run `UndoMovesFromLog` from the Macro dialog
  - Successfully restored items will turn **green**

---

## 📊 Sheets Overview

| Sheet Name         | Purpose                                      |
|--------------------|----------------------------------------------|
| `FileList_*`       | Lists scanned files with metadata and flags  |
| `CategorySummary`  | Aggregated file counts and sizes by category |
| `MoveLog`          | Tracks all file/folder moves with timestamps |

---

## 🧠 Notes & Tips

- The tool supports up to **50,000 rows per sheet** before creating a new one.
- File types are categorized using a comprehensive extension dictionary.
- You can customize:
  - `cutoffDate` in `FlagForCleanup` for aging logic
  - `clutterTypes` array for unwanted extensions
- All moves are logged and reversible if needed.

---

## 🔒 Security Reminder
Always run macros from **trusted sources only**. This tool uses file system access via `Scripting.FileSystemObject`, which can modify files and folders.

---

## 📞 Support
For questions or enhancements, contact the tool author or submit feedback via your preferred channel.