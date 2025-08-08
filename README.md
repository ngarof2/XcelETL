# XcelETL
Custom Toolkits and Automations

## FileWhip 
(Created with MicroSoft Excel VBA)
### 🧹 Whip Your Files Into Shape

### **Purpose**
A robust framework for deep file audits, extension-based classification, intelligent cleanup, and user-directed relocation.  
Ideal for managing sprawling repositories with transparency and precision.

---

## 🔄 Core Workflow Modules

### 📁 Recursive Scanning (`ListFilesRecursively_Tabbed`)
- User selects root folder  
- Files listed across tabbed `FileList_*` sheets (~50K rows each)  
- Includes full path, metadata, and split directories

### 🧠 Categorization & Cleanup
- Categorizes via 200+ extension dictionary  
- Flags clutter by date, duplication, or extension type  
- Smart reclassification of “Uncategorized” entries  
- Summarized in `CategorySummary`

### 🚚 Flexible File Movement (`MoveMarkedFilesWithOptions`)
- Moves color-flagged (yellow) entries based on user-selected granularity:
  - 🟩 **File Only** (Green flag)  
  - 🟧 **File Folder** (Orange flag)  
  - 🟥 **Parent Folder** (Red flag)
- Destination selected via folder picker  
- Generates detailed audit trail in `MoveLog`:
  - Timestamps  
  - Action type  
  - Excel row locator

---

## 🧩 Technical Highlights

| **Feature**                  | **Details**                                                                 |
|-----------------------------|------------------------------------------------------------------------------|
| Modular recursion           | Automatically creates new sheets when row threshold exceeded                |
| Contextual path splitting   | Breaks full path into granular folder columns for sorting/filtering         |
| Color-coded action markers  | Yellow = flagged for move; red/orange/green reflect move type               |
| Granular user control       | Prompts user for move scope (file, folder, parent) at runtime               |
| Robust logging              | Tracks success/failure with comments and Immediate Window debug output      |
| Error resilience            | Uses `On Error Resume Next` and custom error messages for graceful failure  |

---

## 💡 Next-Level Extensions

- Build a UI dashboard to configure filters and triggers visually  
- Add simulation mode with non-destructive logging  
- Detect and handle filename collisions at target location  
- Integrate archive or backup logic for moved folders  
- Export summary of missed or skipped moves with resolution hints
