# FileWhip v2.0 - Python

A powerful file management tool converted from VBA/Excel to Python with a modern customtkinter GUI. FileWhip helps you scan, categorize, analyze, and clean up files on your system.

## Features

### Core Functionality
- **Recursive File Scanning** - Scan folders and list all files with metadata (name, path, size, type, modified date)
- **File Categorization** - Automatically categorize files by extension (Music, Video, Documents, Images, Programming, etc.)
- **Category Summary** - View file counts and total sizes by category
- **Cleanup Flagging** - Flag files for cleanup based on:
  - Duplicate detection (name + size or hash-based)
  - Clutter file types (tmp, log, bak, old, dmp - customizable)
  - Old files (configurable cutoff date)
- **File Moving** - Move marked files with options:
  - Move individual files
  - Move file's containing folder
  - Move parent folder
- **Move Logging** - All moves are logged to JSON file and displayed in Move Log tab
- **Undo Functionality** - Undo all successful moves from the log

### New Optimizations & Upgrades (v2.0)

#### Performance Optimizations
- **Single-Pass File Scanning** - Optimized from two-pass to single-pass scanning with dynamic progress estimation
- **Thread-Safe GUI Updates** - Implemented queue-based GUI updates to prevent crashes during long operations
- **Cancel Operations** - Added cancel button to stop long-running operations (scan, move)
- **Hash-Based Duplicate Detection** - Optional MD5 hash-based detection for more accurate duplicate finding (slower but more precise)

#### User Interface Improvements
- **Search & Filter** - Real-time search by filename/path and filter by category
- **Marked Files Filter** - Quick filter to show only marked files
- **Theme Toggle** - Switch between dark and light mode
- **Settings Dialog** - Customize cutoff date, clutter types, and duplication detection method
- **Keyboard Shortcuts**:
  - `Ctrl+S` - Save scan results
  - `Ctrl+O` - Load scan results
  - `Ctrl+E` - Export to CSV
  - `Escape` - Cancel current operation

#### Data Management
- **Export to CSV** - Export file list with all metadata to CSV format
- **Save/Load Scan Results** - Save scan results to JSON and load them later for analysis
- **Configuration Persistence** - Settings saved to config.json file

#### Visualization
- **Charts Tab** - Visualize file distribution with:
  - Pie chart showing file count by category (top 10)
  - Bar chart showing storage usage by category (top 10)
  - Dark-themed charts matching the application theme

## Installation

1. Install Python 3.8 or higher
2. Install dependencies:
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install customtkinter matplotlib
```

## Usage

Run the application:
```bash
python filewhip.py
```

### Basic Workflow
1. **Select a folder** to scan using the Browse button
2. **Click "Scan Files"** to recursively scan the folder
3. **Click "Categorize Files"** to categorize files by type
4. **Click "Flag for Cleanup"** to identify potential cleanup candidates
5. **Review marked files** in the file list (yellow-highlighted)
6. **Select destination folder** for moving files
7. **Choose move options** (file, folder, or parent folder)
8. **Click "Move Marked Files"** to move flagged files
9. **Use "Undo Moves"** if needed to restore files

### Advanced Features

#### Settings
Click the ⚙️ Settings button to configure:
- **Cutoff Date** - Date threshold for flagging old files
- **Clutter Types** - File extensions to flag as clutter (comma-separated)
- **Hash-Based Duplication** - Enable MD5 hash detection for more accurate duplicate finding (slower)

#### Search & Filter
- Use the search box to filter by filename or path
- Use the category dropdown to filter by file category
- Check "Marked Only" to show only flagged files

#### Export & Save
- **Export to CSV** - Export current file list with all metadata
- **Save Scan** - Save scan results to JSON for later analysis
- **Load Scan** - Load previously saved scan results

#### Visualization
- Go to the "Visualization" tab
- Click "Generate Charts" to see file distribution
- Requires matplotlib to be installed

## Configuration

Configuration is stored in `config.json` in the same directory as the application. Default settings:

```json
{
  "cutoff_date": "2023-01-01",
  "clutter_types": ["tmp", "log", "bak", "old", "dmp"],
  "use_hash_duplication": false,
  "appearance_mode": "System",
  "theme": "blue"
}
```

## File Categories

FileWhip recognizes 200+ file extensions across categories:
- Music (mp3, wav, flac, etc.)
- Video (mp4, avi, mkv, etc.)
- Documents (pdf, docx, txt, etc.)
- Images (jpg, png, svg, etc.)
- Programming (py, js, java, cpp, etc.)
- Spreadsheets (xlsx, csv, ods, etc.)
- Archives (zip, rar, 7z, etc.)
- System files (dll, sys, tmp, etc.)
- And many more...

## Future Upgrade Suggestions

### High Priority
1. **Virtual Treeview** - Implement virtual treeview for handling 100k+ files without performance issues
2. **File Preview** - Add quick preview panel for images, text files, and documents
3. **Batch Operations** - Select multiple files manually for batch operations
4. **Regex Search** - Add regex support for advanced file filtering

### Medium Priority
5. **Size-Based Filters** - Filter files by size ranges (e.g., >100MB, <1KB)
6. **Date-Based Filters** - Filter files by date ranges (created, modified, accessed)
7. **Duplicate Groups** - Group duplicates together for easier review
8. **Integration with Recycle Bin** - Move to system recycle bin instead of direct deletion
9. **File Content Search** - Search within text files for specific content
10. **Custom Categories** - Allow users to define custom file type categories

### Low Priority
11. **Cloud Storage Support** - Scan and manage files in cloud storage (Google Drive, Dropbox)
12. **Network Scanning** - Scan network shares and mapped drives
13. **Scheduling** - Schedule automatic scans and cleanups
14. **Reports** - Generate detailed HTML/PDF reports
15. **Multi-Language Support** - Add internationalization support
16. **Plugin System** - Allow custom plugins for additional functionality
17. **Dark Mode Auto-Detect** - Automatically detect system theme preference
18. **File Comparison** - Side-by-side file comparison for duplicates
19. **Archive Preview** - Preview contents of zip/rar archives
20. **File Hash Verification** - Verify file integrity against known hashes

## Performance Considerations

- **Large Scans**: For folders with 50,000+ files, consider using the "Save Scan" feature to save results and analyze later
- **Hash Detection**: Hash-based duplicate detection is slower but more accurate. Enable only for smaller folders or specific use cases
- **Memory Usage**: File lists are stored in memory. Very large scans (>100k files) may use significant RAM

## Troubleshooting

**Application crashes during scan**: Use the cancel button to stop long operations. The application now uses thread-safe GUI updates.

**Charts not showing**: Install matplotlib with `pip install matplotlib`

**Settings not saving**: Ensure write permissions in the application directory

**Move operations failing**: Check that:
- Destination folder exists and is writable
- Files are not in use by other applications
- You have sufficient permissions

## License

This is a converted version of the original FileWhip VBA application.

## Credits

Original VBA version: FileWhip v1.1
Python conversion: FileWhip v2.0
GUI Framework: customtkinter
