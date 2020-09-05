Attribute VB_Name = "FileSystem"
Option Explicit
'Function to get local path of workbook. Requried, since if our workbook is on OneDrive or otherwise shared a URL may be returned through ThisWorkbook.Name.
'Based on https://social.msdn.microsoft.com/Forums/office/en-US/1331519b-1dd1-4aa0-8f4f-0453e1647f57/how-to-get-physical-path-instead-of-url-onedrive?forum=officegeneral
Public Function GetWokrbookPath(ByVal Path As String) As String
    Dim FinalPath As String, Slashes As Byte
    FinalPath = Path
    If InStr(1, FinalPath, OneDriveCommercial) <> 0 Then
        'We are using OneDrive Commercial
        'Find "/Documents" position in file URL
        Slashes = InStr(1, FinalPath, "/Documents") + Len("/Documents")
        'Get the ending file path without pointer in OneDrive
        FinalPath = Mid(FinalPath, Slashes, Len(FinalPath) - Slashes + 1)
        'Add OneDrive
        FinalPath = FindOneDrive(FinalPath)
    ElseIf InStr(1, FinalPath, OneDrive) <> 0 Then
        'We are using regular OneDrive
        'Locate OneDrive portion and remove it
        Slashes = InStr(Len(OneDrive) + 1, FinalPath, "/")
        FinalPath = Mid(ThisWorkbook.Path, Slashes)
        'Add OneDrive
        FinalPath = FindOneDrive(FinalPath)
    End If
    'Correcting slashes and adding ending one
    FinalPath = Replace(FinalPath, "/", "\") & "\"
    'Correct double slashes at the end, in case there are such. Need to do that only in the end of the string, because we may be working with network paths.
    'Should not normally happen, more of a precaution.
    FinalPath = Replace(FinalPath, "\\", "\")
    If Right(FinalPath, 2) = "\\" Then
        FinalPath = Left(FinalPath, Len(FinalPath) - 1)
    End If
    'Replacing html entoty for space with actual space
    FinalPath = Replace(FinalPath, "%20", " ")
    GetWokrbookPath = FinalPath
End Function
'Locate path in OneDrive directories
Private Function FindOneDrive(ByVal Path As String) As String
    Dim Shell As Object, OneDrivePath As String
    Set Shell = CreateObject("WScript.Shell")
    'Get local path from registry
    On Error Resume Next
    OneDrivePath = Shell.RegRead("HKEY_CURRENT_USER\Environment\OneDrive")
    If OneDrivePath = vbNullString Then
        OneDrivePath = Shell.RegRead("HKEY_CURRENT_USER\Environment\OneDriveConsumer")
    Else
         If Dir(OneDrivePath & Path, vbNormal) <> vbNullString Or Dir(OneDrivePath & Path, vbDirectory) <> vbNullString Then
            FindOneDrive = OneDrivePath & Path
         End If
    End If
    If OneDrivePath = vbNullString Then
        OneDrivePath = Shell.RegRead("HKEY_CURRENT_USER\Environment\OneDriveCommercial")
    Else
        If Dir(OneDrivePath & Path, vbNormal) <> vbNullString Or Dir(OneDrivePath & Path, vbDirectory) <> vbNullString Then
            FindOneDrive = OneDrivePath & Path
         End If
    End If
    On Error GoTo 0
    If OneDrivePath = vbNullString Then
        OneDrivePath = Path
    Else
        If Dir(OneDrivePath & Path, vbNormal) <> vbNullString Or Dir(OneDrivePath & Path, vbDirectory) <> vbNullString Then
            FindOneDrive = OneDrivePath & Path
         End If
    End If
End Function
'Fucntion to check if directory exists and optionally create it
Public Function CheckDir(ByVal Directory As String, Optional Create = False, Optional DateSelector As Integer = 1) As Boolean
    On Error GoTo Catcher
    CheckDir = False
    'Update value, in case it's a setting ID, not an actual value and replace YYYYMMDD
    Directory = DateAnchorsReplace(StringOrSetting(Directory), DateSelector)
    If Dir(Directory, vbDirectory) <> vbNullString Then
        CheckDir = True
    Else
        If Create = False Then
            CheckDir = False
        Else
            MkDir Directory
            If Dir(Directory, vbDirectory) <> vbNullString Then
                CheckDir = True
            Else
                CheckDir = False
            End If
        End If
    End If
Catcher:
'Graceful exit in case of errors on MkDir or Dir
End Function
'Check if file exists
Public Function CheckFile(ByVal FileMask As String) As Boolean
    CheckFile = False
    If Trim(FileMask) = "" Then
        CheckFile = False
        Exit Function
    End If
    On Error Resume Next
    If Dir(FileMask, vbNormal) <> vbNullString Then
        CheckFile = True
    End If
    If Err.Number <> 0 Then
        CheckFile = False
    End If
    On Error GoTo 0
End Function
Private Function CheckFileList(ByVal FileListInit As String, Optional ByVal Minutes As Long = 0, Optional ByVal Reverse = False, Optional DateSelector As Integer = 1) As Boolean
    Dim filelist() As String
    Dim FileElement As Variant
    Dim FileCount As Integer, MsgChoice As Integer
    Dim Lost As String, Found As String
    Dim CurTime As Date, FileTime As Date
    Dim LostCount As Integer: LostCount = 0
    'Set current time
    CurTime = Now()
    'Set Lost and Found to empty string as precaution
    Lost = ""
    Found = ""
    'Update value, in case it's a setting ID, not an actual value
    FileListInit = StringOrSetting(FileListInit)
    'Fill array from the string that was passed
    filelist = Split(FileListInit, GetSetting("ArrayDelimiter"), -1, 1)
    For Each FileElement In filelist
        If IsEmpty(FileElement) = False And FileElement <> "" Then
            'Trim whitespace. Have to use Replace first, since VBA's Trim does not trim new lines
            FileElement = Trim(Replace(Replace(Replace(Replace(FileElement, vbCrLf, ""), vbNewLine, ""), vbCr, ""), vbLf, ""))
            'Replace YYYYMMDD
            FileElement = DateAnchorsReplace(FileElement, DateSelector)
            'Count files
            FileCount = CountFile(FileElement)
            If FileCount > 0 Then
                FileTime = NewestFileTime(FileElement)
                If Reverse = False Then
                    If Minutes > 0 Then
                        'Get latest timestamp
                        If Abs(DateDiff("n", FileTime, CurTime)) <= Minutes Then
                            'Add to Found files list
                            Found = Found & vbCrLf & FileElement & " (count: " & FileCount & ", time: " & FileTime & ")"
                        Else
                            'Add to Lost files list
                            If LostCount < 10 Then
                                Lost = Lost & vbCrLf & FileElement & " (count: " & FileCount & ", time: " & FileTime & ")"
                                LostCount = LostCount + 1
                            Else
                                'No need to scan further, since we will not save the info, so optimizing a bit
                                Exit For
                            End If
                        End If
                    End If
                Else
                    Found = Found & vbCrLf & FileElement & " (count: " & FileCount & ", time: " & FileTime & ")"
                End If
            Else
                'Add to Lost files list
                If LostCount < 10 Then
                    Lost = Lost & vbNewLine & FileElement
                    LostCount = LostCount + 1
                Else
                    'No need to scan further, since we will not save the info, so optimizing a bit
                    Exit For
                End If
            End If
        End If
    Next FileElement
    If Lost <> "" Then
        'Files are missing
        If Reverse = False Then
            If LostCount > 10 Then
                MsgBox "Some files are missing or outdated. Showing first 10 of them:" & vbCrLf & Lost, vbOKOnly + vbExclamation + vbApplicationModal, "Files are missing or outdated!"
            Else
                MsgBox LostCount & " files are missing or outdated:" & vbCrLf & Lost, vbOKOnly + vbExclamation + vbApplicationModal, "Files are missing or outdated!"
            End If
            CheckFileList = False
        Else
            If Found = "" Then
                MsgBox "No files found.", vbOKOnly + vbInformation + vbApplicationModal, "No files found"
                CheckFileList = True
            Else
                MsgBox "Some files found: " & vbCrLf & Found, vbOKOnly + vbExclamation + vbApplicationModal, "Some files found"
                CheckFileList = False
            End If
        End If
    Else
        'Files found
        If Reverse = False Then
            If Minutes > 0 Then
                MsgBox "All files found:" & vbCrLf & Found, vbOKOnly + vbInformation + vbApplicationModal, "All files found"
            Else
                MsgBox "All files found.", vbOKOnly + vbInformation + vbApplicationModal, "All files found"
            End If
            CheckFileList = True
        Else
            MsgBox "Some files found: " & vbCrLf & Found, vbOKOnly + vbExclamation + vbApplicationModal, "Some files found"
            CheckFileList = False
        End If
    End If
End Function
'Count files
Public Function CountFile(ByVal FileMask As String) As Integer
    CountFile = 0
    Dim File As Variant
    On Error Resume Next
    File = Dir(FileMask, vbNormal)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    If File <> vbNullString Then
        While (File <> vbNullString)
            CountFile = CountFile + 1
            File = Dir
        Wend
    End If
    On Error GoTo 0
End Function
Private Function NewestFileTime(ByVal FileMask As String) As Date
    Dim File As Variant
    Dim TimeStamp As Date, TempTime As Date
    Dim Folder As String
    TimeStamp = 0
    TempTime = 0
    'Get folder string to utilize further on, since DIR returns only filename, but FileDateTime requires full path
    Folder = Left(FileMask, InStrRev(FileMask, "\"))
    File = Dir(FileMask, vbNormal)
    If File <> vbNullString Then
        TimeStamp = FileDateTime(Folder & File)
        While (File <> vbNullString)
            TempTime = FileDateTime(Folder & File)
            If TempTime > TimeStamp Then
                TimeStamp = TempTime
            End If
            File = Dir
        Wend
    End If
    NewestFileTime = TimeStamp
End Function
Private Function OpenFile(ByVal ToOpen As String, Optional DateSelector As Integer = 1) As Boolean
    'Check if Empty
    If IsEmpty(ToOpen) = True Or ToOpen = "" Then
        Call WriteLog("Failed to open 'empty path'")
        OpenFile = False
    End If
    'Update value, in case it's a setting ID, not an actual value
    ToOpen = DateAnchorsReplace(StringOrSetting(ToOpen), DateSelector)
    'Check if an URI
    If InStr(1, ToOpen, "://") <= 0 And InStr(1, ToOpen, ":\\") <= 0 Then
        'Not an URI
        'Check if file exists
        If CheckFile(ToOpen) = False Then
            'Setting found, but not the file it refers to - exit the function
            Call WriteLog("Failed to open '" & ToOpen & "': path not found")
            OpenFile = False
            Exit Function
        End If
    End If
    On Error Resume Next
    Call CreateObject("Shell.Application").ShellExecute(ToOpen)
    If Err.Number <> 0 Then
        Call WriteLog("Failed to open '" & ToOpen & "' path with error #" & Err.Number)
        OpenFile = False
    Else
        Call WriteLog("Opened '" & ToOpen & "' path")
        OpenFile = True
    End If
    On Error GoTo 0
End Function
Public Function ListFiles(ByVal ToList As String, Optional FullPath As Boolean = False, Optional Final As Boolean = True, Optional DateSelector As Integer = 1) As Variant
    'Update value, in case it's a setting ID, not an actual value
    ToList = DateAnchorsReplace(StringOrSetting(ToList), DateSelector)
    Dim filelist() As String, Files() As Variant
    Dim FileElement As Variant, File As Variant, Folder As String
    'Defaults as precaution
    Folder = ""
    'ReDim Files(0)
    'Fill array from the string that was passed
    filelist = Split(ToList, GetSetting("ArrayDelimiter"), -1, 1)
    For Each FileElement In filelist
        Folder = Left(FileElement, InStrRev(FileElement, "\"))
        On Error Resume Next
        File = Dir(FileElement, vbNormal)
        'Skip this loop, if failed to open directory for any reason
        If Err.Number <> 0 Then
            GoTo NextIteration
        End If
        On Error GoTo 0
        While (File <> vbNullString)
            If FullPath = True Then
                File = Folder & File
            End If
            'Increase array to include new element
            'Catching error, in case this is our first loop and array is empty Doing this to avoid empty element in the array
            On Error Resume Next
            ReDim Preserve Files(UBound(Files) + 1)
            If Err.Number <> 0 Then
                ReDim Files(0)
            End If
            On Error GoTo 0
            Files(UBound(Files)) = File
            File = Dir
        Wend
NextIteration:
        On Error GoTo 0
    Next FileElement
    'Checking if array was properly initiated. If it was not - it's empty
    On Error Resume Next
    ReDim Preserve Files(UBound(Files))
    If Err.Number <> 0 Then
        If Final = True Then
            Call WriteLog("No files found matching the criteria given")
            MsgBox "No files found matching the criteria given.", vbOKOnly + vbInformation + vbApplicationModal, "No files found"
        End If
        ListFiles = False
    Else
        If Final = True Then
            Call WriteLog("Following files found:" & vbCrLf & Join(Files, vbCrLf))
            MsgBox "Following files found:" & vbCrLf & Join(Files, vbCrLf), vbOKOnly + vbInformation + vbApplicationModal, "Files found"
            ListFiles = True
        Else
            ListFiles = Files
        End If
    End If
    On Error GoTo 0
End Function
Public Function FindInFile(ByVal Haystack As String, ByVal Needle As String, Optional ByVal ShowMatches As Boolean = False, Optional ByVal Charset As String = "UTF-8", Optional DateSelector As Integer = 1) As Boolean
    'Set default result
    FindInFile = False
    'Update value, in case it's a setting ID, not an actual value
    Haystack = DateAnchorsReplace(StringOrSetting(Haystack), DateSelector)
    If CheckFile(Haystack) = True Then
        On Error Resume Next
        'Create stream to read from file
        Dim adoStream As Object
        Dim Hay As String
        Set adoStream = CreateObject("ADODB.Stream")
        If Err.Number <> 0 Then
            Call WriteLog("Failed to create stream for '" & Haystack & "' with error #" & Err.Number)
            MsgBox "Failed to create stream for '" & Haystack & "' with error #" & Err.Number, vbOKOnly + vbExclamation + vbApplicationModal, "Failed to stream"
            On Error GoTo 0
            Exit Function
        End If
        'Set character set
        adoStream.Charset = Charset
        If Err.Number <> 0 Then
            Call WriteLog("Failed to set charset for '" & Haystack & "' with error #" & Err.Number)
            MsgBox "Failed to set charset for '" & Haystack & "' with error #" & Err.Number, vbOKOnly + vbExclamation + vbApplicationModal, "Failed to set charset"
            On Error GoTo 0
            Exit Function
        End If
        'Open file
        adoStream.Open
        If Err.Number <> 0 Then
            Call WriteLog("Failed to open stream for '" & Haystack & "' with error #" & Err.Number)
            MsgBox "Failed to open stream for '" & Haystack & "' with error #" & Err.Number, vbOKOnly + vbExclamation + vbApplicationModal, "Failed to stream"
            On Error GoTo 0
            Exit Function
        End If
        adoStream.LoadFromFile Haystack
        If Err.Number <> 0 Then
            Call WriteLog("Failed to open file '" & Haystack & "' with error #" & Err.Number)
            MsgBox "Failed to open file '" & Haystack & "' with error #" & Err.Number, vbOKOnly + vbExclamation + vbApplicationModal, "Haystack is not from hay"
            On Error GoTo 0
            Exit Function
        End If
        'Read file
        Hay = adoStream.ReadText
        'Create regexp
        Dim Regexp As Object
        Set Regexp = CreateObject("vbscript.regexp")
        If Err.Number <> 0 Then
            Call WriteLog("Failed to create regexp object for '" & Haystack & "' with error #" & Err.Number)
            MsgBox "Failed to create regexp object for '" & Haystack & "' with error #" & Err.Number, vbOKOnly + vbExclamation + vbApplicationModal, "Failed to regexp"
            On Error GoTo 0
            Exit Function
        End If
        'Setup regexp
        If ShowMatches = True Then
            'Forcing the needle to be a pattern. If we don't, then the list will show only the needle itself, not the full string, where the match was found
            Regexp.Pattern = ".*" & Needle & ".*"
        Else
            Regexp.Pattern = Needle
        End If
        Regexp.Global = True
        Regexp.IgnoreCase = True
        'Get matches (if any)
        Dim Matches As Object, Match As Integer, MatchesList As String
        MatchesList = ""
        Set Matches = Regexp.Execute(Hay)
        If Matches.Count > 0 Then
            Call WriteLog(Matches.Count & " matches found in '" & Haystack & "'")
            If ShowMatches = True Then
                For Match = Matches.Count To 1 Step -1
                    MatchesList = MatchesList & Matches.Item(Match - 1).Value
                    If (Matches.Count - Match + 1) = 5 Then
                        Exit For
                    End If
                Next Match
                MsgBox "Last matches (up to 5) from '" & Haystack & "':" & vbCrLf & MatchesList, vbOKOnly + vbInformation + vbApplicationModal, "Matches found"
            End If
            FindInFile = True
        Else
            Call WriteLog("No matches found in '" & Haystack & "'")
            If MsgBox("No matches found in '" & Haystack & "'" & vbCrLf & vbCrLf & "Do you want to open the file to manually check it?", vbYesNo + vbExclamation + vbApplicationModal, "No matches") = vbYes Then
                Call OpenFile(Haystack)
            End If
        End If
        On Error GoTo 0
    Else
        Call WriteLog("'" & Haystack & "' not found")
        MsgBox "'" & Haystack & "' not found", vbOKOnly + vbExclamation + vbApplicationModal, "Haystack not found"
    End If
End Function
'Function to attach screenshots as objects into cells based on their names
Private Function ScreenshotsToObjects(ByVal ScreenSheet As Worksheet, ByVal CurDate As String) As Boolean
    Dim ScreensPath As String
    Dim File As Variant, CellAddress As String, FinalObject As OLEObject, CellToPaste As Range
    'Define path to Screenshots Excel file
    ScreensPath = GetSetting("ScreenshotsPath")
    If IsEmpty(ScreensPath) = True Or ScreensPath = "" Then
        'Set default path
        ScreensPath = GetWokrbookPath(ThisWorkbook.Path) & "Screenshots"
    End If
    If CheckDir(ScreensPath, False) = False Then
        ScreenshotsToObjects = False
        Exit Function
    End If
    On Error Resume Next
    File = Dir(ScreensPath & "\" & CurDate & "_*.jpg", vbNormal)
    If Err.Number <> 0 Then
        On Error GoTo 0
        ScreenshotsToObjects = False
        Exit Function
    End If
    On Error GoTo 0
    If File <> vbNullString Then
        While (File <> vbNullString)
            CellAddress = Replace(Replace(File, CurDate & "_", ""), ".jpg", "")
            'Checking if valid address
            On Error Resume Next
            Set CellToPaste = ScreenSheet.Range(CellAddress)
            If Err.Number = 0 Then
                Set FinalObject = ScreenSheet.OLEObjects.Add(Filename:=ScreensPath & "\" & File, Link:=False, DisplayAsIcon:=True, IconLabel:=CurDate & "_" & CellAddress)
                If Err.Number <> 0 Then
                    Call WriteLog("Failed to attach screenshot '" & ScreensPath & "\" & File & "' with error #" & Err.Number)
                    MsgBox "Failed to attach screenshot with error #" & Err.Number & "!" & vbCrLf & "It will be retained as '" & ScreensPath & "\" & File & "'.", vbOKOnly + vbExclamation + vbApplicationModal, "Failed to attach"
                Else
                    Kill ScreensPath & "\" & File
                    With FinalObject
                        .ShapeRange.AlternativeText = CurDate & "_" & CellAddress
                        .ShapeRange.LockAspectRatio = msoFalse
                        .Top = CellToPaste.Top
                        .Left = CellToPaste.Left
                        .Width = CellToPaste.Width
                        .Height = CellToPaste.Height
                    End With
                End If
            End If
            On Error GoTo 0
            File = Dir
        Wend
    End If
    ScreenshotsToObjects = True
End Function
