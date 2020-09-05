Attribute VB_Name = "WindowsOS"
Option Explicit
'Function to connect to a PC using RDP
Public Function RDP(ByVal HostName As String, Optional ByVal SizePercentage As Integer = 100, Optional ByVal CurrentUser As Boolean = True, Optional ByVal Cache As Boolean = False, Optional ByVal AsAdmin As Boolean = False) As Boolean
    Dim MSTSC As Long
    Dim ConnectionString As String
    'Update HostName in case it's a setting ID
    HostName = StringOrSetting(HostName)
    ConnectionString = "/v:" & HostName
    'Sanitze SizePercentage. Not using values <10, since unlikely it will be usable.
    If SizePercentage < 10 Then
        SizePercentage = 10
    End If
    If SizePercentage > 100 Then
        SizePercentage = 100
    End If
    If SizePercentage = 100 Then
        'Open in FullScreen mode and use current screen resolution
        ConnectionString = ConnectionString & " /f /w:" & GetSystemMetrics32(ScreenWidth) & " /h:" & GetSystemMetrics32(ScreenHeight)
    Else
        'Open in window mode with resolution based on percentage of current resolution
        ConnectionString = ConnectionString & " /w:" & GetSystemMetrics32(ScreenWidth) / 100 * SizePercentage & " /h:" & GetSystemMetrics32(ScreenHeight) / 100 * SizePercentage
    End If
    If CurrentUser = False Then
        'Force login credentials request
        ConnectionString = ConnectionString & " /prompt"
    End If
    If AsAdmin = True Then
        'Elevate connection permissions
        ConnectionString = ConnectionString & " /admin"
    End If
    If Cache = False Then
        'Prevent storage of credentials
        ConnectionString = ConnectionString & " /public"
    End If
    On Error Resume Next
    MSTSC = Shell("mstsc " & ConnectionString, vbNormalFocus)
    If Err.Number <> 0 Then
        Call WriteLog("Failed to connect to '" & HostName & "' workstation with error #" & Err.Number)
        RDP = False
    Else
        If MSTSC = 0 Then
            Call WriteLog("Failed to start MSTSC to connect to '" & HostName & "' workstation")
            RDP = False
        Else
            RDP = True
        End If
    End If
    On Error GoTo 0
End Function
'Function to check if service is started
Public Function ServiceCheck(ByVal ServiceName As String, Optional ByVal WorkStation As String = "", Optional ByVal AutoStart As Boolean = False, Optional ByVal ReStart As Boolean = False, Optional ByVal SecondsWait As Integer = 10) As Boolean
    Dim WMI As Object, Services As Object, Service As Object, RequestResult As Integer
    'Update ServiceName in case it's a setting ID
    ServiceName = StringOrSetting(ServiceName)
    'Set WMI
    If WorkStation = "" Then
        WorkStation = VBA.Environ$("COMPUTERNAME")
    Else
        'Update WorkStation in case it's a setting ID
        WorkStation = StringOrSetting(WorkStation)
    End If
    On Error Resume Next
    Set WMI = GetObject("WinMgmts:{impersonationLevel=impersonate}//" & WorkStation)
    'Failed to connect to WMI
    If Err.Number <> 0 Then
        Call WriteLog("Failed to connect to WMI of '" & WorkStation & "' with error #" & Err.Number)
        'Show same (almost) message to user
        MsgBox "Failed to connect to WMI of '" & WorkStation & "' with error #" & Err.Number & "!", vbCritical + vbOKOnly + vbApplicationModal, "Connection to WMI failed!"
        ServiceCheck = False
        Exit Function
    End If
    On Error GoTo 0
    'Get list of services with appropriate name or display name
    Set Services = WMI.ExecQuery("Select * from Win32_Service Where Name='" & ServiceName & "' OR DisplayName='" & ServiceName & "'")
    If Services.Count = 1 Then
        'Still need to use loop to go through the "list"
        For Each Service In Services
            If Service.State = "Running" Then
                If ReStart = True Then
                    RequestResult = Service.StopService
                    If RequestResult <> 0 Then
                        'Something went wrong with service start/resume
                        Call WriteLog("Failed to stop service '" & ServiceName & "' with error:" & ServiceControlResponse(RequestResult))
                        MsgBox "Failed to stop service '" & ServiceName & "' with error:" & vbCrLf & ServiceControlResponse(RequestResult) & ".", vbCritical + vbOKOnly + vbApplicationModal, "Service stop failed!"
                        ServiceCheck = False
                        Exit Function
                    End If
                    'Wait defined amounts of seconds, in case service may need some times to stop
                    Sleep (SecondsWait * 1000)
                    RequestResult = Service.StartService
                    If RequestResult <> 0 Then
                        'Something went wrong with service start/resume
                        Call WriteLog("Failed to restart service '" & ServiceName & "' with error:" & ServiceControlResponse(RequestResult))
                        MsgBox "Failed to restart service '" & ServiceName & "' with error:" & vbCrLf & ServiceControlResponse(RequestResult) & ".", vbCritical + vbOKOnly + vbApplicationModal, "Service restart failed!"
                        ServiceCheck = False
                        Exit Function
                    End If
                    'Wait defined amounts of seconds, in case service may need some times to start
                    Sleep (SecondsWait * 1000)
                    'Update service state
                    Set Services = WMI.ExecQuery("Select * from Win32_Service Where Name='" & ServiceName & "' OR DisplayName='" & ServiceName & "'")
                    For Each Service In Services
                        If Service.State = "Running" Then
                            'Return
                            ServiceCheck = True
                            Exit Function
                        Else
                            Call WriteLog("Failed to restart service '" & ServiceName & "' with no trackable error. Current state: '" & Service.State & "'")
                            MsgBox "Failed to restart service '" & ServiceName & "' with no trackable error." & vbCrLf & "Current state: '" & Service.State & "'.", vbCritical + vbOKOnly + vbApplicationModal, "Service restart failed!"
                            ServiceCheck = False
                            Exit Function
                        End If
                    Next Service
                Else
                    'Return
                    ServiceCheck = True
                    Exit Function
                End If
            Else
                If AutoStart = True Then
                    'Start if stopped, continue if paused. If anything else - exit, since it's most likely an intermitent state and better wait
                    If Service.State = "Stopped" Then
                        RequestResult = Service.StartService
                    ElseIf Service.State = "Paused" Then
                        RequestResult = Service.ResumeService
                    Else
                        Call WriteLog("Service '" & ServiceName & "' is in intermitent state ('" & Service.State & "')")
                        MsgBox "Service '" & ServiceName & "' is in intermitent state ('" & Service.State & "')." & vbCrLf & "Please, trye again later.", vbExclamation + vbOKOnly + vbApplicationModal, "Service in intermitent state!"
                        ServiceCheck = False
                        Exit Function
                    End If
                    If RequestResult <> 0 Then
                        'Something went wrong with service start/resume
                        Call WriteLog("Failed to start service '" & ServiceName & "' with error:" & ServiceControlResponse(RequestResult))
                        MsgBox "Failed to start service '" & ServiceName & "' with error:" & vbCrLf & ServiceControlResponse(RequestResult) & ".", vbCritical + vbOKOnly + vbApplicationModal, "Service start failed!"
                        ServiceCheck = False
                        Exit Function
                    End If
                Else
                    ServiceCheck = False
                    Exit Function
                End If
            End If
        Next Service
        'If we we starting the service, we need to get its new state, since reciving "0" from control command does not mean it actually started
        If AutoStart = True Then
            'Wait defined amounts of seconds, in case service may need some times to start
            Sleep (SecondsWait * 1000)
            'Update service state
            Set Services = WMI.ExecQuery("Select * from Win32_Service Where Name='" & ServiceName & "' OR DisplayName='" & ServiceName & "'")
            For Each Service In Services
                If Service.State = "Running" Then
                    'Return
                    ServiceCheck = True
                Else
                    Call WriteLog("Failed to start service '" & ServiceName & "' with no trackable error. Current state: '" & Service.State & "'")
                    MsgBox "Failed to start service '" & ServiceName & "' with no trackable error." & vbCrLf & "Current state: '" & Service.State & "'.", vbCritical + vbOKOnly + vbApplicationModal, "Service start failed!"
                    ServiceCheck = False
                End If
            Next Service
        End If
    ElseIf Service.Count = 0 Then
        Call WriteLog("Service '" & ServiceName & "' is missing or user has no access to it")
        'Show same (almost) message to user
        MsgBox "Service '" & ServiceName & "' is missing or you lack access to it!", vbCritical + vbOKOnly + vbApplicationModal, "Service not found!"
        ServiceCheck = False
    ElseIf Service.Count > 1 Then
        Call WriteLog("Multiple services with names '" & ServiceName & "' were found")
        'Show same (almost) message to user
        MsgBox "Multiple services with names '" & ServiceName & "' were found!" & vbCrLf & "Exiting function due to ambiguity.", vbCritical + vbOKOnly + vbApplicationModal, "Multiple services found!"
        ServiceCheck = False
    End If
End Function
'Convert codes of service control errors to human readable text
Private Function ServiceControlResponse(ByVal Code As Integer) As String
    'This is the list of only those codes, that may be valid in case of starting/resuming
    Select Case Code
        Case 2
            ServiceControlResponse = "no access"
        Case 4
            ServiceControlResponse = "the requested control code is not valid, or it is unacceptable to the service"
        Case 7
            ServiceControlResponse = "the service did not respond to the start request in a timely fashion"
        Case 8
            ServiceControlResponse = "unknown failure when starting the service"
        Case 9
            ServiceControlResponse = "the directory path to the service executable file was not found"
        Case 10
            ServiceControlResponse = "the service is already running"
        Case 12
            ServiceControlResponse = "a dependency this service relies on has been removed from the system"
        Case 13
            ServiceControlResponse = "the service failed to find the service needed from a dependent service"
        Case 14
            ServiceControlResponse = "the service is disabled"
        Case 15
            ServiceControlResponse = "the service does not have the correct authentication to run on the system"
        Case 16
            ServiceControlResponse = "this service is being removed from the system"
        Case 17
            ServiceControlResponse = "the service has no execution thread"
        Case 18
            ServiceControlResponse = "the service has circular dependencies when it starts"
        Case 19
            ServiceControlResponse = "a service is running under the same name"
        Case 20
            ServiceControlResponse = "the service name has invalid characters"
        Case 21
            ServiceControlResponse = "invalid parameters have been passed to the service"
        Case 22
            ServiceControlResponse = "the account under which this service runs is either invalid or lacks the permissions to run the service"
        Case Else
            ServiceControlResponse = "code " & CStr(Code)
    End Select
End Function
Public Function ToClipBoard(ByVal Text As String) As Boolean
    'Based on https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
    'But with some modifications, most notable one is usage of LongPtr instead of Long, because otherwise it fails to work on 64-bit system
    If IsEmpty(Text) = True Or Trim(Text) = "" Then
        ToClipBoard = False
    Else
        On Error Resume Next
        Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr, hClipMemory As LongPtr, X As LongPtr
        'Allocate moveable global memory
        hGlobalMemory = GlobalAlloc(&H42, Len(Text) + 1)
        'Lock the block to get a far pointer to this memory.
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        'Copy the string to this global memory.
        lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)
        'Unlock the memory.
        If GlobalUnlock(hGlobalMemory) <> 0 Then
            ToClipBoard = False
        Else
            'Open the Clipboard to copy data to.
            If OpenClipboard(0&) = 0 Then
                ToClipBoard = False
            Else
                'Clear the Clipboard.
                X = EmptyClipboard()
                'Copy the data to the Clipboard.
                hClipMemory = SetClipboardData(1, hGlobalMemory)
                If Err.Number <> 0 Then
                    ToClipBoard = False
                Else
                    ToClipBoard = True
                End If
            End If
        End If
        Call CloseClipboard
        On Error GoTo 0
    End If
End Function
Public Function VNC(ByVal PathToVNC As String) As Boolean
    Dim VNCexe As String
    VNCexe = GetSetting("VNCPath")
    'Check if executable exists
    If CheckFile(VNCexe) = True Then
        'Update value, in case it's a setting ID, not an actual value
        PathToVNC = StringOrSetting(PathToVNC)
        'Check if .vnc configuration file exists
        If CheckFile(PathToVNC) = True Then
            Dim ShellObject As Object
            ShellObject = Shell(VNCexe + " /config " + PathToVNC, vbNormalFocus)
            VNC = True
        Else
            Call WriteLog("'" & PathToVNC & "' was not found")
            MsgBox "'" & PathToVNC & "' was not found!", vbOKOnly + vbExclamation + vbApplicationModal, "No VNC configuration found"
            VNC = False
        End If
    Else
        Call WriteLog("'" & VNCexe & "' was not found")
        MsgBox "'" & VNCexe & "' was not found!", vbOKOnly + vbExclamation + vbApplicationModal, "No VNC executable"
        VNC = False
    End If
End Function
Option Explicit
Private Function AttachScreenshot(CellToPaste As Range) As Boolean
    Dim ScreenshotName As String, ScreensPath As String, ScreensExcel As Workbook, ScreenSheet As Worksheet
    Dim ChartTemp As ChartObject, ChartAreaTemp As Chart
    Dim Screenshot As Shape, TempFilePath As String
    'Define name for the shape
    ScreenshotName = "ScreenShot" & CellToPaste.Address
    'Define path to Screenshots Excel file
    ScreensPath = GetSetting("ScreenshotsPath")
    If IsEmpty(ScreensPath) = True Or ScreensPath = "" Then
        'Set default path
        ScreensPath = GetWokrbookPath(ThisWorkbook.Path) & "Screenshots"
    End If
    If CheckDir(ScreensPath, True) = False Then
        Call WriteLog("Object pasted is not an image")
        MsgBox "Folder '" & ScreensPath & "' does not exist and failed to be created!", vbOKOnly + vbCritical + vbApplicationModal, "No folder!"
        AttachScreenshot = False
        GoTo SelectPasteCell
    End If
    'Set path where screenshot will be saved before creating an object from it
    TempFilePath = ScreensPath & "\" & Format(ReturnRange("CurrentDayCell").Value2, "YYYYMMDD") & "_" & AlphaNumeric(CellToPaste.Address) & ".jpg"
    'Check that we have a picture in clipboard
    If ScreenShotInClipBoard = True Then
        'Create workbook
        Set ScreensExcel = Workbooks.Add()
        'Add sheet to store screenshots in
        ScreensExcel.Worksheets.Add.Name = "Screenshots"
        'Set sheet to object for ease of access
        Set ScreenSheet = ScreensExcel.Worksheets("Screenshots")
        'Paste screenshot
        ScreenSheet.Paste
        Set Screenshot = ScreenSheet.Shapes(Selection.Name)
        'Set name of the shape for consistency
        Screenshot.Name = ScreenshotName
        'Check if we actually have pasted an image and not something else
        If Screenshot.Type <> msoPicture Then
            'Delete the object, if it's not an image and exit function
            ScreensExcel.Close
            Call WriteLog("Object pasted is not an image")
            MsgBox "Object pasted is not an image!", vbOKOnly + vbCritical + vbApplicationModal, "Not an image"
            AttachScreenshot = False
            GoTo SelectPasteCell
        End If
        'Create chart to export image through it
        On Error Resume Next
        Set ChartTemp = ScreenSheet.ChartObjects.Add(0, 0, Screenshot.Width, Screenshot.Height)
        If Err.Number <> 0 Then
            ScreensExcel.Close
            Call WriteLog("Failed to create Chart with error #" & Err.Number)
            MsgBox "Failed to create Chart with error #" & Err.Number & "!", vbOKOnly + vbCritical + vbApplicationModal, "Failed to paste"
            AttachScreenshot = False
            GoTo SelectPasteCell
        End If
        On Error GoTo 0
        Set ChartAreaTemp = ChartTemp.Chart
        'Export to image
        With ChartAreaTemp
            .ChartArea.Select
            .Paste
            On Error Resume Next
            .Export Filename:=TempFilePath, FilterName:="JPEG", Interactive:=False
            If Err.Number <> 0 Then
                ScreensExcel.Close
                Call WriteLog("Failed to export Chart with error #" & Err.Number)
                MsgBox "Failed to export Chart with error #" & Err.Number & "!", vbOKOnly + vbCritical + vbApplicationModal, "Failed to export"
                AttachScreenshot = False
                GoTo SelectPasteCell
            End If
            On Error GoTo 0
        End With
        ScreensExcel.Close
        AttachScreenshot = True
    Else
        Call WriteLog("No image found in clipboard")
        MsgBox "No image found in clipboard!", vbOKOnly + vbExclamation + vbApplicationModal, "No image"
        AttachScreenshot = False
    End If
SelectPasteCell:
    ThisWorkbook.Worksheets("RunSheet").Range(CellToPaste.Address).Select
End Function
'Function to check if we have an image in clipboard
'Taken from https://www.mrexcel.com/board/threads/vba-paste-screenshot-from-clipboard-onto-cell-with-screenshot-automatically-sizing-to-cells-dimensions.962029/
Private Function ScreenShotInClipBoard() As Boolean
    Dim sClipboardFormatName As String, sBuffer As String
    Dim CF_Format As LongPtr, i As Long
    Dim bDtataInClipBoard As Boolean
    If OpenClipboard(0) Then
        CF_Format = EnumClipboardFormats(0&)
        Do While CF_Format <> 0
            sClipboardFormatName = String(255, vbNullChar)
            i = GetClipboardFormatName(CF_Format, sClipboardFormatName, 255)
            sBuffer = sBuffer & Left(sClipboardFormatName, i)
            bDtataInClipBoard = True
            CF_Format = EnumClipboardFormats(CF_Format)
        Loop
        CloseClipboard
    End If
    ScreenShotInClipBoard = bDtataInClipBoard And Len(sBuffer) = 0
End Function
'Function to launch a URL in Private/Incognito mode
Public Function InPrivate(ByVal URL As String, Optional Browser As String = "Edge", Optional BrowserPath As String = "") As Boolean
    Dim PrivateBrowser As Long, Shortcut As String
    Shortcut = ""
    'Ensuring we use only 1st argument and ignore the rest, if anything was set
    URL = StringOrSetting(Trim(URL))
    'Check if argument is empty
    If IsEmpty(URL) Or URL = "" Then
        Call WriteLog("Empty URL provided")
        MsgBox "Empty URL provided!", vbOKOnly + vbCritical + vbApplicationModal, "No URL!"
        InPrivate = False
        Exit Function
    End If
    'Check path if it was provided
    If BrowserPath <> "" Then
        If LCase(Right(BrowserPath, 4)) <> ".exe" Then
            Call WriteLog("Path '" & BrowserPath & "' does not seem to be an .exe")
            MsgBox "Path '" & BrowserPath & "' does not seem to be an .exe", vbOKOnly + vbCritical + vbApplicationModal, "Not an exe!"
            InPrivate = False
            Exit Function
        End If
        If CheckFile(BrowserPath) = False Then
            Call WriteLog("Path '" & BrowserPath & "' was not found")
            MsgBox "Path '" & BrowserPath & "' was not found", vbOKOnly + vbCritical + vbApplicationModal, "Browser not found!"
            InPrivate = False
            Exit Function
        End If
    End If
    
    Select Case LCase(Browser)
        Case "internetexplorer", "internet explorer", "explorer", "ie"
            If BrowserPath <> "" Then
                Shortcut = BrowserPath & " -private " & URL
            ElseIf CheckFile("C:\Program Files\Internet Explorer\iexplore.exe") = True Then
                Shortcut = "C:\Program Files\Internet Explorer\iexplore.exe -private " & URL
            ElseIf CheckFile("C:\Program Files (x86)\Internet Explorer\iexplore.exe") = True Then
                Shortcut = "C:\Program Files (x86)\Internet Explorer\iexplore.exe -private " & URL
            End If
        Case "edge", "msedge", "microsoft edge"
            If BrowserPath <> "" Then
                Shortcut = BrowserPath & " -inprivate " & URL
                'Edge is stored in x86 folder on x64 systems for some reason
            ElseIf CheckFile("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") = True Then
                Shortcut = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"" -inprivate " & URL
            ElseIf CheckFile("C:\Program Files\Microsoft\Edge\Application\msedge.exe") = True Then
                Shortcut = "C:\Program Files\Microsoft\Edge\Application\msedge.exe -inprivate " & URL
            End If
        Case "chrome", "google chrome"
            If BrowserPath <> "" Then
                Shortcut = BrowserPath & " -incognito " & URL
            ElseIf CheckFile("C:\Program Files\Google\Chrome\Application\chrome.exe") = True Then
                Shortcut = "C:\Program Files\Google\Chrome\Application\chrome.exe -incognito " & URL
            ElseIf CheckFile("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") = True Then
                Shortcut = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -incognito " & URL
            End If
        Case "opera"
            If BrowserPath <> "" Then
                Shortcut = BrowserPath & " --private " & URL
            ElseIf CheckFile("C:\Program Files\Opera\launcher.exe") = True Then
                Shortcut = "C:\Program Files\Opera\launcher.exe --private " & URL
            ElseIf CheckFile("C:\Program Files (x86)\Opera\launcher.exe") = True Then
                Shortcut = "C:\Program Files (x86)\Opera\launcher.exe --private " & URL
            End If
        Case "firefox"
            If BrowserPath <> "" Then
                Shortcut = BrowserPath & " -private-window " & URL
            ElseIf CheckFile("C:\Program Files\Mozilla Firefox\firefox.exe") = True Then
                Shortcut = "C:\Program Files\Mozilla Firefox\firefox.exe -private-window " & URL
            ElseIf CheckFile("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") = True Then
                Shortcut = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe -private-window " & URL
            End If
        Case Else
            Call WriteLog("Unsupported browser '" & Browser & "' specified")
            MsgBox "Unsupported browser '" & Browser & "' specified", vbOKOnly + vbCritical + vbApplicationModal, "Unsupported browser!"
            InPrivate = False
            Exit Function
    End Select
    
    If Shortcut = "" Then
        Call WriteLog("Failed to determine shortcut")
        MsgBox "Failed to determine shortcut", vbOKOnly + vbCritical + vbApplicationModal, "Empty shortcut!"
        InPrivate = False
        Exit Function
    End If
    
    On Error Resume Next
    PrivateBrowser = Shell(Shortcut, vbNormalFocus)
    If Err.Number <> 0 Then
        Call WriteLog("Failed to open '" & URL & "' with error #" & Err.Number)
        MsgBox "Failed to open '" & URL & "' with error #" & Err.Number, vbOKOnly + vbCritical + vbApplicationModal, "Failed to open URL!"
        InPrivate = False
    Else
        If PrivateBrowser = 0 Then
            Call WriteLog("Failed to start '" & Browser)
            MsgBox "Failed to start '" & Browser, vbOKOnly + vbCritical + vbApplicationModal, "Failed to open URL!"
            InPrivate = False
        Else
            Call WriteLog("Started Internet Explorer to open '" & URL & "'")
            InPrivate = True
        End If
    End If
    On Error GoTo 0
End Function
'Function to export all the code of the workbook. Requries programmatic access being allowed in MS Office Trust Center.
Private Sub VBAExport()
    Dim VBModule As Object
    Dim VBAPath As String
    VBAPath = GetSetting("ExportCodePath")
    If IsEmpty(VBAPath) = True Or VBAPath = "" Then
        'Set default path
        VBAPath = GetWokrbookPath(ThisWorkbook.Path) & "SourceCode\"
    End If
    If CheckDir(VBAPath, True) = True Then
        For Each VBModule In ThisWorkbook.VBProject.VBComponents
            'Excluding any modules designed specifically for 3rd parties
            If InStr(1, VBModule.Name, "custom", vbTextCompare) = 0 Then
                VBModule.Export (VBAPath & VBModule.Name & ".bas")
            End If
        Next VBModule
        MsgBox "Modules exported!", vbOKOnly + vbInformation + vbApplicationModal, "Modules exported!"
    End If
End Sub
