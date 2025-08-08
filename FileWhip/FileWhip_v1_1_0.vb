Sub ListFilesRecursively_Tabbed()
    Dim ws As Worksheet
    Dim r As Long
    Dim fDialog As FileDialog
    Dim startFolder As String
    Dim sheetCount As Integer

    ' File dialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder"
        If .Show <> -1 Then Exit Sub
        startFolder = .SelectedItems(1)
    End With

    ' Create first sheet
    sheetCount = 1
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "FileList_" & sheetCount
    Call WriteHeaders(ws)
    r = 2

    ' Start recursive listing
    Call RecursiveFileScan_Tabbed(startFolder, ws, r, sheetCount)
    Call FlagForCleanup
    Call CategorizeFilesByType
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "FileList_" Then
            With ws
                .Activate
                .Columns("I:I").Cut
                .Columns("H:H").Insert Shift:=xlToRight
            End With
        End If
    Next ws

    Sheets("CategorySummary").Activate

        Cells.Select
    Cells.EntireColumn.AutoFit
    MsgBox "Recursive listing completed across " & sheetCount & " sheet(s).", vbInformation
End Sub

Sub RecursiveFileScan_Tabbed(ByVal folderPath As String, ByRef ws As Worksheet, ByRef r As Long, ByRef sheetCount As Integer)
    Dim fso As Object, folder As Object, subfolder As Object, file As Object
    Dim pathParts() As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' List files in this folder
    For Each file In folder.Files
        If r > 50000 Then
            ' Start new sheet
            sheetCount = sheetCount + 1
            Set ws = ThisWorkbook.Sheets.Add(After:=ws)
            ws.Name = "FileList_" & sheetCount
            Call WriteHeaders(ws)
            r = 2
        End If

        ws.Cells(r, 1).Value = file.Name
        ws.Cells(r, 2).Value = file.Path
        ws.Cells(r, 3).Value = Format(file.Size / 1024, "0.00")
        ws.Cells(r, 4).Value = fso.GetExtensionName(file.Path)
        ws.Cells(r, 5).Value = Format(file.DateLastModified, "yyyy-mm-dd hh:nn:ss")


        pathParts = Split(file.Path, "\")
        For i = 0 To UBound(pathParts)
            ws.Cells(r, 6 + i).Value = pathParts(i)
        Next i
        r = r + 1
    Next file

    ' Dive into subfolders
    For Each subfolder In folder.SubFolders
        Call RecursiveFileScan_Tabbed(subfolder.Path, ws, r, sheetCount)
    Next subfolder
    

End Sub

Sub WriteHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "Filename"
    ws.Cells(1, 2).Value = "Full Path"
    ws.Cells(1, 3).Value = "Size (KB)"
    ws.Cells(1, 4).Value = "File Type"
    ws.Cells(1, 5).Value = "Last Modified"
    ws.Cells(1, 6).Value = "Path Element 1"
End Sub

Sub FlagForCleanup()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim cutoffDate As Date, fileDate As Date
    Dim fileName As String, fileSizeKB As Double, fileKey As String, ext As String
    Dim reason As String, duplicatePath As String
    Dim nameSizeDict As Object, pathDict As Object
    Dim clutterTypes As Variant, clutterExt As Variant
    Dim actionCol As Long, dupCol As Long
    Dim nameCol As Long, pathCol As Long, sizeCol As Long, extCol As Long, dateCol As Long

    cutoffDate = DateSerial(2023, 1, 1) ' ? Customize as needed
    Set nameSizeDict = CreateObject("Scripting.Dictionary")
    Set pathDict = CreateObject("Scripting.Dictionary")
    clutterTypes = Array("tmp", "log", "bak", "old", "dmp")

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "FileList_" Then
            ' Insert columns A and B
            ws.Columns(1).Insert Shift:=xlToRight
            ws.Columns(1).Insert Shift:=xlToRight
            ws.Cells(1, 1).Value = "Action Recommendation"
            ws.Cells(1, 2).Value = "Duplicate Reference"

            lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
            actionCol = 1: dupCol = 2
            nameCol = 3: pathCol = 4: sizeCol = 5: extCol = 6: dateCol = 7

            For r = 2 To lastRow
                fileName = Trim(ws.Cells(r, nameCol).Value)
                fileSizeKB = ws.Cells(r, sizeCol).Value
                ext = LCase(Trim(ws.Cells(r, extCol).Value))
                fileKey = fileName & "|" & Format(fileSizeKB, "0.00")
                reason = ""

                ' ?? Duplicate check
                If nameSizeDict.exists(fileKey) Then
                    ws.Cells(r, nameCol).Interior.Color = RGB(255, 165, 0)
                    reason = reason & "Duplicate (name + size); "
                    ws.Cells(r, dupCol).Value = pathDict(fileKey)
                Else
                    nameSizeDict.Add fileKey, True
                    pathDict.Add fileKey, ws.Cells(r, pathCol).Value
                End If

                ' ?? Clutter extension
                For Each clutterExt In clutterTypes
                    If ext = clutterExt Then
                        ws.Cells(r, nameCol).Interior.Color = RGB(200, 200, 200)
                        reason = reason & "Clutter file type; "
                        Exit For
                    End If
                Next clutterExt

                ' ??? Old file
                If IsDate(ws.Cells(r, dateCol).Value) Then
                    fileDate = ws.Cells(r, dateCol).Value
                    If fileDate < cutoffDate Then
                        ws.Cells(r, nameCol).Interior.Color = RGB(255, 200, 200)
                        reason = reason & "Old file (" & Format(fileDate, "yyyy-mm-dd") & "); "
                    End If
                End If

                ' ?? Write recommendation
                If reason <> "" Then
                    ws.Cells(r, actionCol).Value = "Review: " & reason
                End If
            Next r
        End If
    Next ws


End Sub


Sub CategorizeFilesByType()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim ext As String, fileName As String
    Dim categoryCol As Long
    Dim category As String
    Dim knownTypes As Object

    Set knownTypes = CreateObject("Scripting.Dictionary")

  ' ?? Music
knownTypes.Add "aac", "Music"
knownTypes.Add "aiff", "Music"
knownTypes.Add "alac", "Music"
knownTypes.Add "flac", "Music"
knownTypes.Add "mod", "Music"
knownTypes.Add "mp3", "Music"
knownTypes.Add "m4a", "Music"
knownTypes.Add "ogg", "Music"
knownTypes.Add "snd", "Music"
knownTypes.Add "wav", "Music"
knownTypes.Add "xm", "Music"

' ?? Video
knownTypes.Add "3gp", "Video"
knownTypes.Add "avi", "Video"
knownTypes.Add "mkv", "Video"
knownTypes.Add "mov", "Video"
knownTypes.Add "mp4", "Video"
knownTypes.Add "mpeg", "Video"
knownTypes.Add "mpg", "Video"
knownTypes.Add "mts", "Video"
knownTypes.Add "vob", "Video"
knownTypes.Add "webm", "Video"
knownTypes.Add "wmv", "Video"

' ?? Spreadsheets
knownTypes.Add "csv", "Spreadsheet"
knownTypes.Add "ods", "Spreadsheet"
knownTypes.Add "xls", "Spreadsheet"
knownTypes.Add "xlsb", "Spreadsheet"
knownTypes.Add "xlsx", "Spreadsheet"
knownTypes.Add "xlsm", "Macro Spreadsheet"
knownTypes.Add "numbers", "Spreadsheet"

' ?? Documents
knownTypes.Add "doc", "Document"
knownTypes.Add "docx", "Document"
knownTypes.Add "djvu", "Document"
knownTypes.Add "epub", "Document"
knownTypes.Add "md", "Document"
knownTypes.Add "mobi", "Document"
knownTypes.Add "odt", "Document"
knownTypes.Add "pdf", "Document"
knownTypes.Add "rst", "Document"
knownTypes.Add "rtf", "Document"
knownTypes.Add "tex", "Document"
knownTypes.Add "txt", "Document"

' ??? Images
knownTypes.Add "ai", "Image"
knownTypes.Add "apng", "Image"
knownTypes.Add "bmp", "Image"
knownTypes.Add "exr", "Image"
knownTypes.Add "gif", "Image"
knownTypes.Add "heic", "Image"
knownTypes.Add "ico", "Image"
knownTypes.Add "indd", "Image"
knownTypes.Add "jpg", "Image"
knownTypes.Add "jpeg", "Image"
knownTypes.Add "png", "Image"
knownTypes.Add "psd", "Image"
knownTypes.Add "raw", "Image"
knownTypes.Add "svg", "Image"
knownTypes.Add "tiff", "Image"
knownTypes.Add "webp", "Image"
knownTypes.Add "xcf", "Image"

' ??? Presentations
knownTypes.Add "key", "Presentation"
knownTypes.Add "odp", "Presentation"
knownTypes.Add "ppt", "Presentation"
knownTypes.Add "pptx", "Presentation"
knownTypes.Add "pptm", "Presentation"
knownTypes.Add "sxi", "Presentation"

' ?? Text-based Data & Web
knownTypes.Add "css", "Web Style"
knownTypes.Add "html", "Web Page"
knownTypes.Add "json", "Data File"
knownTypes.Add "ndjson", "Data File"
knownTypes.Add "tsv", "Data File"
knownTypes.Add "xml", "Data File"
knownTypes.Add "yaml", "Config File"
knownTypes.Add "yml", "Config File"
knownTypes.Add "hjson", "Data File"
knownTypes.Add "toml", "Config File"
knownTypes.Add "ini", "Config File"

' ?? Config and System
knownTypes.Add "au", "Adobe Audition Sound File"
knownTypes.Add "bak", "Backup"
knownTypes.Add "cfg", "Config File"
knownTypes.Add "desktop", "Config File"
knownTypes.Add "dll", "System Component"
knownTypes.Add "env", "Environment Settings"
knownTypes.Add "lock", "System Metadata"
knownTypes.Add "log", "Log File"
knownTypes.Add "ocx", "System Component"
knownTypes.Add "plist", "Config File"
knownTypes.Add "resx", "Resource File"
knownTypes.Add "service", "System File"
knownTypes.Add "swp", "Swap File"
knownTypes.Add "sys", "System File"
knownTypes.Add "tmp", "Temporary File"
knownTypes.Add "xaml", "App Markup"

' ?? Connection & Credentials
knownTypes.Add "cer", "Certificate"
knownTypes.Add "crt", "Certificate"
knownTypes.Add "jks", "Certificate"
knownTypes.Add "kdbx", "Password Database"
knownTypes.Add "p12", "Certificate"
knownTypes.Add "pem", "Private Key"
knownTypes.Add "pfx", "Certificate"
knownTypes.Add "ppk", "PuTTY Key"
knownTypes.Add "ovpn", "VPN Config"
knownTypes.Add "ssh", "SSH Config"

' ?? Archives & Installers
knownTypes.Add "7z", "Archive"
knownTypes.Add "apk", "Installer"
knownTypes.Add "appx", "Installer"
knownTypes.Add "bz2", "Archive"
knownTypes.Add "cab", "Archive"
knownTypes.Add "deb", "Installer"
knownTypes.Add "exe", "Executable"
knownTypes.Add "gz", "Archive"
knownTypes.Add "msi", "Installer"
knownTypes.Add "rar", "Archive"
knownTypes.Add "rpm", "Installer"
knownTypes.Add "tar", "Archive"
knownTypes.Add "xz", "Archive"
knownTypes.Add "zip", "Archive"

' ?? Disk Images
knownTypes.Add "bin", "Disk Image"
knownTypes.Add "cue", "Disk Image"
knownTypes.Add "dmg", "Disk Image"
knownTypes.Add "img", "Disk Image"
knownTypes.Add "iso", "Disk Image"
knownTypes.Add "vhd", "Disk Image"
knownTypes.Add "vmdk", "Disk Image"

' ????? Programming & Scripting
knownTypes.Add "bat", "Batch Script"
knownTypes.Add "c", "C/C++ Code"
knownTypes.Add "clj", "Clojure Code"
knownTypes.Add "cpp", "C/C++ Code"
knownTypes.Add "cs", "C# Code"
knownTypes.Add "dart", "Dart Code"
knownTypes.Add "erl", "Erlang Code"
knownTypes.Add "fs", "F# Code"
knownTypes.Add "go", "Go Code"
knownTypes.Add "groovy", "Groovy Code"
knownTypes.Add "ipynb", "Python Notebook"
knownTypes.Add "java", "Java Code"
knownTypes.Add "jl", "Julia Code"
knownTypes.Add "js", "JavaScript"
knownTypes.Add "jsx", "JavaScript"
knownTypes.Add "lua", "Lua Code"
knownTypes.Add "ml", "OCaml Code"
knownTypes.Add "php", "PHP Code"
knownTypes.Add "pl", "Perl Script"
knownTypes.Add "ps1", "PowerShell Script"
knownTypes.Add "py", "Python Code"
knownTypes.Add "r", "R Script"
knownTypes.Add "rb", "Ruby Code"
knownTypes.Add "rdl", "Report Definition"
knownTypes.Add "rs", "Rust Code"
knownTypes.Add "scala", "Scala Code"
knownTypes.Add "sh", "Shell Script"
knownTypes.Add "sql", "SQL Script"
knownTypes.Add "swift", "Swift Code"
knownTypes.Add "ts", "TypeScript"
knownTypes.Add "tsx", "TypeScript"
knownTypes.Add "vb", "VB Script"
knownTypes.Add "vbs", "VB Script"

' ?? BI / Data Science
knownTypes.Add "arrow", "Data File"
knownTypes.Add "feather", "Data File"
knownTypes.Add "mat", "MATLAB Data"
knownTypes.Add "parquet", "Data File"
knownTypes.Add "pbix", "Power BI File"
knownTypes.Add "rds", "R Data File"
knownTypes.Add "sav", "SPSS Data"
knownTypes.Add "sas", "SAS Script"
knownTypes.Add "dta", "Stata Data"
knownTypes.Add "orc", "Data File"

' ?? CAD / 3D
knownTypes.Add "3ds", "3D Model"
knownTypes.Add "blend", "3D Model"
knownTypes.Add "dae", "3D Model"
knownTypes.Add "dwg", "CAD Drawing"
knownTypes.Add "dxf", "CAD Drawing"
knownTypes.Add "fbx", "3D Model"
knownTypes.Add "max", "3D Model"
knownTypes.Add "obj", "3D Model"
knownTypes.Add "skp", "3D Model"
knownTypes.Add "stl", "3D Model"

' ?? Game Assets
knownTypes.Add "dat", "Game Data"
knownTypes.Add "gba", "ROM Image"
knownTypes.Add "lvl", "Game Level"
knownTypes.Add "map", "Game Map"
knownTypes.Add "nes", "ROM Image"
knownTypes.Add "pak", "Game Archive"
knownTypes.Add "rom", "ROM Image"
knownTypes.Add "wad", "Game Package"

' ?? Miscellaneous / System
knownTypes.Add "accdt", "Access Template"
knownTypes.Add "ad", "Advertisement File"
knownTypes.Add "adp", "Access Project"
knownTypes.Add "bak1", "Backup"
knownTypes.Add "bfc", "System File"
knownTypes.Add "btr", "Database File"
knownTypes.Add "cpl", "Control Panel Item"
knownTypes.Add "cur", "Cursor Image"
knownTypes.Add "dbb", "Database File"
knownTypes.Add "dmp", "Memory Dump"
knownTypes.Add "ds_store", "System Metadata"
knownTypes.Add "dtd", "Document Type Definition"
knownTypes.Add "dylib", "macOS Dynamic Library"
knownTypes.Add "exd", "Office Cached Control"
knownTypes.Add "fon", "Font File"
knownTypes.Add "gitignore", "Git Config"
knownTypes.Add "hlp", "Help File"
knownTypes.Add "idx", "Index File"
knownTypes.Add "inf", "Setup Information"
knownTypes.Add "inuse", "System Metadata"
knownTypes.Add "jar", "Java Archive"
knownTypes.Add "jnilib", "Java Native Library"
knownTypes.Add "ldf", "SQL Server Log"
knownTypes.Add "lnk", "Shortcut"
knownTypes.Add "mds", "Media Descriptor"
knownTypes.Add "mui", "Multilingual User Interface"
knownTypes.Add "nfo", "Info File"
knownTypes.Add "old", "Legacy File"
knownTypes.Add "pdb", "Program Database"
knownTypes.Add "pf", "Prefetch File"
knownTypes.Add "pkg", "Package"
knownTypes.Add "plb", "Oracle Library"
knownTypes.Add "pref", "Preferences File"
knownTypes.Add "props", "Properties File"
knownTypes.Add "pyc", "Compiled Python"
knownTypes.Add "rc", "Resource Script"
knownTypes.Add "rc2", "Resource Script"
knownTypes.Add "reg", "Registry File"
knownTypes.Add "rsp", "Response File"
knownTypes.Add "sdb", "System Database"
knownTypes.Add "suo", "Solution User Options"
knownTypes.Add "swf", "Flash Movie"
knownTypes.Add "sym", "Symbol File"
knownTypes.Add "template", "Template File"
knownTypes.Add "tlb", "Type Library"
knownTypes.Add "ttc", "TrueType Collection"
knownTypes.Add "ttf", "TrueType Font"
knownTypes.Add "url", "Internet Shortcut"
knownTypes.Add "vsix", "Visual Studio Extension"
knownTypes.Add "xconf", "XML Config"
knownTypes.Add "xsd", "XML Schema"
knownTypes.Add "xslt", "XSL Transform"
knownTypes.Add "zst", "Zstandard Archive"

' ?? Game Assets (Additional)
knownTypes.Add "smc", "ROM Image"
knownTypes.Add "srm", "Game Save"
knownTypes.Add "qpx", "Game Data"
knownTypes.Add "pck", "Game Package"
knownTypes.Add "part", "Game Asset Fragment"
knownTypes.Add "graph", "Game Graph"
knownTypes.Add "trg", "Game Trigger"
knownTypes.Add "fpk", "Game Package"
knownTypes.Add "pk", "Game Package"
knownTypes.Add "dancing", "Game Animation"
knownTypes.Add "ifo", "DVD Info File"
knownTypes.Add "bup", "DVD Backup File"
knownTypes.Add "m3u", "Playlist"
knownTypes.Add "m3u8", "Playlist"
knownTypes.Add "pls", "Playlist"
knownTypes.Add "sfv", "Checksum File"

' ????? Programming & Scripting (Additional)
knownTypes.Add "sln", "Solution File"
knownTypes.Add "csproj", "C# Project"
knownTypes.Add "vcproj", "Visual C++ Project"
knownTypes.Add "vssettings", "Visual Studio Settings"
knownTypes.Add "dsp", "Developer Studio Project"
knownTypes.Add "dsw", "Developer Studio Workspace"
knownTypes.Add "ncb", "Intellisense Database"
knownTypes.Add "opt", "Developer Options"
knownTypes.Add "def", "Module Definition"
knownTypes.Add "clw", "ClassWizard File"
knownTypes.Add "aps", "Application Support"
knownTypes.Add "bpr", "Borland Project"
knownTypes.Add "dfm", "Delphi Form"
knownTypes.Add "dpr", "Delphi Project"
knownTypes.Add "pas", "Pascal Source"
knownTypes.Add "h", "C/C++ Header"
knownTypes.Add "hpp", "C++ Header"


' ?? Documents (Additional)
knownTypes.Add "dot", "Word Template"
knownTypes.Add "wps", "Works Document"
knownTypes.Add "wri", "Write Document"
knownTypes.Add "pub", "Publisher Document"
knownTypes.Add "pot", "PowerPoint Template"
knownTypes.Add "mix", "PhotoDraw Document"
knownTypes.Add "xps", "XPS Document"
knownTypes.Add "svgz", "Compressed SVG"
knownTypes.Add "jfif", "JPEG Format"
knownTypes.Add "odc", "Office Data Connection"
knownTypes.Add "accdb", "Access Database"
knownTypes.Add "database", "Database File"
knownTypes.Add "dwproj", "Data Warehouse Project"
knownTypes.Add "user", "User Profile"
knownTypes.Add "dim", "Dimension File"
knownTypes.Add "role", "Role Definition"
knownTypes.Add "cube", "Cube Definition"
knownTypes.Add "partitions", "Partition File"
knownTypes.Add "kwf", "Workflow File"
knownTypes.Add "lib", "Library File"

' ?? Archives (Additional)
knownTypes.Add "tgz", "Archive"

' ?? Disk Images (Additional)
knownTypes.Add "mdf", "Disk Image"
knownTypes.Add "toast", "Disk Image"

    ' ?? Process FileList sheets
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "FileList_" Then
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            categoryCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
            ws.Cells(1, categoryCol).Value = "FileCategory"

            For r = 2 To lastRow
                fileName = ws.Cells(r, 3).Value
                If InStrRev(fileName, ".") > 0 Then
                    ext = LCase(Mid(fileName, InStrRev(fileName, ".") + 1))
                    If knownTypes.exists(ext) Then
                        category = knownTypes(ext)
                    Else
                        category = GuessCategory(ext)
                    End If
                    ws.Cells(r, categoryCol).Value = category
                Else
                    ws.Cells(r, categoryCol).Value = "Unknown"
                End If
            Next r
        End If
    Next ws
    
    Call GenerateCategorySummary

End Sub

Function GuessCategory(ext As String) As String
    Select Case ext
        Case "lst", "toc", "manifest"
            GuessCategory = "Index/List"
        Case "sqlite", "db", "mdb"
            GuessCategory = "Database"
        Case "lock", "bak", "swp", "temp"
            GuessCategory = "Backup/Temp"
        Case "config", "settings", "preferences"
            GuessCategory = "Settings"
        Case "lic", "license"
            GuessCategory = "License File"
        Case "scr", "theme", "wallpaper"
            GuessCategory = "Display Asset"
        Case Else
            GuessCategory = "Unclassified"
    End Select
End Function

Sub ListFilesRecursively_Tabbed()
    Dim ws As Worksheet
    Dim r As Long
    Dim fDialog As FileDialog
    Dim startFolder As String
    Dim sheetCount As Integer

    ' File dialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder"
        If .Show <> -1 Then Exit Sub
        startFolder = .SelectedItems(1)
    End With

    ' Create first sheet
    sheetCount = 1
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "FileList_" & sheetCount
    Call WriteHeaders(ws)
    r = 2

    ' Start recursive listing
    Call RecursiveFileScan_Tabbed(startFolder, ws, r, sheetCount)
    Call FlagForCleanup
    Call CategorizeFilesByType
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "FileList_" Then
            With ws
                .Activate
                .Columns("I:I").Cut
                .Columns("H:H").Insert Shift:=xlToRight
            End With
        End If
    Next ws

    Sheets("CategorySummary").Activate

        Cells.Select
    Cells.EntireColumn.AutoFit
    MsgBox "Recursive listing completed across " & sheetCount & " sheet(s).", vbInformation
End Sub

Sub RecursiveFileScan_Tabbed(ByVal folderPath As String, ByRef ws As Worksheet, ByRef r As Long, ByRef sheetCount As Integer)
    Dim fso As Object, folder As Object, subfolder As Object, file As Object
    Dim pathParts() As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' List files in this folder
    For Each file In folder.Files
        If r > 50000 Then
            ' Start new sheet
            sheetCount = sheetCount + 1
            Set ws = ThisWorkbook.Sheets.Add(After:=ws)
            ws.Name = "FileList_" & sheetCount
            Call WriteHeaders(ws)
            r = 2
        End If

        ws.Cells(r, 1).Value = file.Name
        ws.Cells(r, 2).Value = file.Path
        ws.Cells(r, 3).Value = Format(file.Size / 1024, "0.00")
        ws.Cells(r, 4).Value = fso.GetExtensionName(file.Path)
        ws.Cells(r, 5).Value = Format(file.DateLastModified, "yyyy-mm-dd hh:nn:ss")


        pathParts = Split(file.Path, "\")
        For i = 0 To UBound(pathParts)
            ws.Cells(r, 6 + i).Value = pathParts(i)
        Next i
        r = r + 1
    Next file

    ' Dive into subfolders
    For Each subfolder In folder.SubFolders
        Call RecursiveFileScan_Tabbed(subfolder.Path, ws, r, sheetCount)
    Next subfolder
    

End Sub

Sub WriteHeaders(ws As Worksheet)
    ws.Cells(1, 1).Value = "Filename"
    ws.Cells(1, 2).Value = "Full Path"
    ws.Cells(1, 3).Value = "Size (KB)"
    ws.Cells(1, 4).Value = "File Type"
    ws.Cells(1, 5).Value = "Last Modified"
    ws.Cells(1, 6).Value = "Path Element 1"
End Sub

Sub MoveMarkedFilesWithOptions()
    Dim fso          As Object
    Dim ws           As Worksheet, logWS As Worksheet
    Dim r            As Long, logRow As Long
    Dim lastRow      As Long
    Dim filePath     As String, fileName As String
    Dim sourceFolder As String, parentFolder As String
    Dim targetRoot   As String
    Dim response     As VbMsgBoxResult
    Dim moveFile     As Boolean, moveFolder As Boolean, moveParent As Boolean
    Dim movedCount   As Long
    Dim colorYellow  As Long

    ' ?? Ask user for destination root folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Root Folder"
        If .Show <> -1 Then Exit Sub
        targetRoot = .SelectedItems(1)
    End With
    If Right(targetRoot, 1) <> "\" Then targetRoot = targetRoot & "\"

    ' ? Ask which move types to apply
    response = MsgBox("Move individual files?", vbYesNo + vbQuestion, "Move Options")
    moveFile = (response = vbYes)

    response = MsgBox("Move file's containing folder?", vbYesNo + vbQuestion, "Move Options")
    moveFolder = (response = vbYes)

    response = MsgBox("Move parent folder of file's folder?", vbYesNo + vbQuestion, "Move Options")
    moveParent = (response = vbYes)

    Set fso = CreateObject("Scripting.FileSystemObject")
    colorYellow = vbYellow
    movedCount = 0

    ' ?? Load or create MoveLog sheet
    On Error Resume Next
    Set logWS = ThisWorkbook.Sheets("MoveLog")
    If logWS Is Nothing Then
        Set logWS = ThisWorkbook.Sheets.Add
        logWS.Name = "MoveLog"
        With logWS
            .Cells(1, 1).Value = "Filename"
            .Cells(1, 2).Value = "Original Path"
            .Cells(1, 3).Value = "New Path"
            .Cells(1, 4).Value = "Change Comment"
            .Cells(1, 5).Value = "Timestamp"
            .Cells(1, 6).Value = "Tab Locater"
        End With
    End If
    On Error GoTo 0
    logRow = logWS.Cells(logWS.Rows.Count, 1).End(xlUp).Row + 1

    ' ?? Loop through FileList sheets
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 9) = "FileList_" Then
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For r = 2 To lastRow
                If ws.Cells(r, 1).Interior.Color = colorYellow Then
                    filePath = ws.Cells(r, 2).Value
                    fileName = ws.Cells(r, 1).Value

                    If fso.FileExists(filePath) Then
                        sourceFolder = fso.GetParentFolderName(filePath)
                        parentFolder = fso.GetParentFolderName(sourceFolder)

                        ' ?? Move parent folder
                        If moveParent And fso.FolderExists(parentFolder) Then
                            On Error Resume Next
                            fso.moveFolder parentFolder, targetRoot & fso.GetFileName(parentFolder)
                            If Err.Number = 0 Then
                                ws.Cells(r, 1).Interior.Color = RGB(255, 0, 0)
                                movedCount = movedCount + 1
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = parentFolder
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(parentFolder)
                                    .Cells(logRow, 4).Value = "Parent Folder"
                                    .Cells(logRow, 5).Value = Now
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                            Else
                                Debug.Print "[Parent Folder Fail] " & Err.Description & " :: " & parentFolder
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = parentFolder
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(parentFolder)
                                    .Cells(logRow, 4).Value = "Parent Folder ?"
                                    .Cells(logRow, 5).Value = "ERROR: " & Err.Description
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                                Err.Clear
                            End If
                            logRow = logRow + 1
                            On Error GoTo 0
                        End If

                        ' ?? Move file folder
                        If moveFolder And fso.FolderExists(sourceFolder) Then
                            On Error Resume Next
                            fso.moveFolder sourceFolder, targetRoot & fso.GetFileName(sourceFolder)
                            If Err.Number = 0 Then
                                ws.Cells(r, 1).Interior.Color = RGB(255, 165, 0)
                                movedCount = movedCount + 1
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = sourceFolder
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(sourceFolder)
                                    .Cells(logRow, 4).Value = "File Folder"
                                    .Cells(logRow, 5).Value = Now
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                            Else
                                Debug.Print "[File Folder Fail] " & Err.Description & " :: " & sourceFolder
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = sourceFolder
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(sourceFolder)
                                    .Cells(logRow, 4).Value = "File Folder ?"
                                    .Cells(logRow, 5).Value = "ERROR: " & Err.Description
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                                Err.Clear
                            End If
                            logRow = logRow + 1
                            On Error GoTo 0
                        End If

                        ' ?? Move individual file
                        If moveFile Then
                            On Error Resume Next
                            fso.moveFile filePath, targetRoot & fso.GetFileName(filePath)
                            If Err.Number = 0 Then
                                ws.Cells(r, 1).Interior.Color = RGB(0, 255, 0)
                                movedCount = movedCount + 1
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = filePath
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(filePath)
                                    .Cells(logRow, 4).Value = "File Only"
                                    .Cells(logRow, 5).Value = Now
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                            Else
                                Debug.Print "[File Move Fail] " & Err.Description & " :: " & filePath
                                With logWS
                                    .Cells(logRow, 1).Value = fileName
                                    .Cells(logRow, 2).Value = filePath
                                    .Cells(logRow, 3).Value = targetRoot & fso.GetFileName(filePath)
                                    .Cells(logRow, 4).Value = "File Only ?"
                                    .Cells(logRow, 5).Value = "ERROR: " & Err.Description
                                    .Cells(logRow, 6).Value = ws.Name & ":" & r
                                End With
                                Err.Clear
                            End If
                            logRow = logRow + 1
                            On Error GoTo 0
                        End If

                    Else
                        Debug.Print "[Missing File] " & filePath
                        With logWS
                            .Cells(logRow, 1).Value = fileName
                            .Cells(logRow, 2).Value = filePath
                            .Cells(logRow, 3).Value = "N/A"
                            .Cells(logRow, 4).Value = "File Missing ?"
                            .Cells(logRow, 5).Value = "File does not exist"
                            .Cells(logRow, 6).Value = ws.Name & ":" & r
                        End With
                        logRow = logRow + 1
                    End If
                End If
            Next r
        End If
    Next ws

    MsgBox "Completed. Total moved: " & movedCount & vbCrLf & "Check 'MoveLog' and Immediate Window for details.", vbInformation
End Sub

Sub UndoMovesFromLog()
    Dim fso As Object
    Dim logWS As Worksheet
    Dim r As Long, lastRow As Long
    Dim sourcePath As String, targetPath As String, itemType As String
    Dim colorYellow As Long
    Dim undoCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    colorYellow = vbYellow
    undoCount = 0

    ' Check for MoveLog sheet
    On Error Resume Next
    Set logWS = ThisWorkbook.Sheets("MoveLog")
    On Error GoTo 0
    If logWS Is Nothing Then
        MsgBox "MoveLog sheet not found.", vbExclamation
        Exit Sub
    End If

    lastRow = logWS.Cells(logWS.Rows.Count, 1).End(xlUp).Row

    ' Loop through MoveLog entries
    For r = 2 To lastRow
        If logWS.Cells(r, 1).Interior.Color = colorYellow Then
            sourcePath = logWS.Cells(r, 2).Value
            targetPath = logWS.Cells(r, 3).Value
            itemType = logWS.Cells(r, 4).Value

            On Error Resume Next
            Select Case itemType
                Case "File Only"
                    If fso.FileExists(targetPath) Then
                        fso.moveFile targetPath, sourcePath
                        If Err.Number = 0 Then
                            logWS.Cells(r, 1).Interior.Color = RGB(0, 255, 0) ' Green for success
                            undoCount = undoCount + 1
                        Else
                            logWS.Cells(r, 1).Interior.Color = RGB(255, 0, 0) ' Red for failure
                            logWS.Cells(r, 5).Value = "UNDO ERROR: " & Err.Description
                            Err.Clear
                        End If
                    End If

                Case "File Folder", "Parent Folder"
                    If fso.FolderExists(targetPath) Then
                        fso.moveFolder targetPath, sourcePath
                        If Err.Number = 0 Then
                            logWS.Cells(r, 1).Interior.Color = RGB(0, 255, 0)
                            undoCount = undoCount + 1
                        Else
                            logWS.Cells(r, 1).Interior.Color = RGB(255, 0, 0)
                            logWS.Cells(r, 5).Value = "UNDO ERROR: " & Err.Description
                            Err.Clear
                        End If
                    End If

                Case Else
                    logWS.Cells(r, 1).Interior.Color = RGB(255, 0, 0)
                    logWS.Cells(r, 5).Value = "UNDO ERROR: Unknown type"
            End Select
            On Error GoTo 0
        End If
    Next r

    MsgBox "Undo completed. Total restored: " & undoCount, vbInformation
End Sub
