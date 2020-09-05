VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Application.EnableEvents = False
    If ThisWorkbook.MultiUserEditing = True Then
        Application.Run "ClearSessions"
    End If
    If Success = True Then
        Application.Run "Welcomen"
        Application.StatusBar = "Workbook saved successfully"
    Else
        Application.StatusBar = "Failed to save workbook"
        Call WriteLog("Failed to save workbook with error #" & Err.Number)
        MsgBox "Failed to save workbook with error #" & Err.Number, vbCritical + vbOKOnly + vbApplicationModal, "Failed to save"
    End If
    Application.EnableEvents = True
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim BoxResult As Integer
    If EditMode() = True Then
        BoxResult = MsgBox("Workbook is in editor mode!" & vbCrLf & "It will not be accessible by regular users!" & vbCrLf & "Are you sure you want to exit?", vbYesNo + vbExclamation + vbApplicationModal + vbDefaultButton2 + vbMsgBoxSetForeground, "Leave Editor mode?")
        If BoxResult = vbNo Then
            Cancel = True
        End If
    End If
    'Save before actual exit
    'Disabling events to avoid BeforeSave loop
    Application.EnableEvents = False
    On Error Resume Next
    'Remove any custom views
    Dim ViewToRemove As CustomView
    For Each ViewToRemove In ThisWorkbook.CustomViews
        ViewToRemove.Delete
    Next ViewToRemove
    'Actually save
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        Application.StatusBar = "Failed to save workbook"
        Call WriteLog("Failed to save workbook on closure with error #" & Err.Number)
        BoxResult = MsgBox("Failed to save workbook with error #" & Err.Number & vbCrLf & "Do you still want to exit?", vbCritical + vbYesNo + vbApplicationModal, "Failed to save")
        If BoxResult = vbNo Then
            Cancel = True
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    'Log exit
    Call WriteLog("Logged out")
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Application.StatusBar = "Optimizing Excel environment..."
    Call Optimize(True)
    Application.Run "UnWelcomen"
    Application.Run "Init"
    Application.Run "AllowUI"
    Application.StatusBar = "Beautifying Excel environment..."
    Call Optimize(False)
End Sub
Private Sub Workbook_Open()
    'Ensure Status bar is enabled
    Application.DisplayStatusBar = True
    Application.StatusBar = "Optimizing Excel environment..."
    
    'Disable page breaks for optimization
    Application.StatusBar = "Hiding page breaks..."
    ThisWorkbook.Worksheets("RunSheet").DisplayPageBreaks = False
    ThisWorkbook.Worksheets("Settings").DisplayPageBreaks = False
    ThisWorkbook.Worksheets("EditorManual").DisplayPageBreaks = False
    
    'Attempt to reduce issues with data validation (.Validate.Add) by disabling OneDrive's AutoSave
    Application.StatusBar = "Disabling OneDrive/SharePoint auto-saving..."
    If ThisWorkbook.AutoSaveOn = True Then
        ThisWorkbook.AutoSaveOn = False
    End If
    
    'Ensure autocalculation is disabled for some extra performance
    Application.StatusBar = "Disabling automatic calculation..."
    Application.Calculation = xlCalculationManual
    
    Application.StatusBar = "Optimizing Excel environment..."
    Call Optimize(True)

    'Reapply protection, if enabled
    Application.StatusBar = "Checking protection..."
    Application.Run "AllowUI", True
    'Initialize ranged names, styles, formating
    Application.StatusBar = "RunSheet initialization..."
    Application.Run "Init"
    
    'Freeze panes on RunSheet
    Application.StatusBar = "Freezing panes..."
    ThisWorkbook.Worksheets("RunSheet").Activate
    ThisWorkbook.Worksheets("RunSheet").Calculate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 4
        .FreezePanes = True
    End With
    
    'Show optional discalimer on start-up, if set
    If IsEmpty(GetSetting("Disclaimer")) = False And GetSetting("Disclaimer") <> "" Then
        Application.StatusBar = "Showing disclaimer..."
        MsgBox GetSetting("Disclaimer"), vbOKOnly + vbApplicationModal + vbInformation + vbDefaultButton1 + vbMsgBoxSetForeground, "Disclaimer"
    End If
    
    Application.StatusBar = "Welcoming..."
    Application.Run "Welcomen"
    
    'Log success
    Call WriteLog("Logged in")
    
    Application.StatusBar = "Beautifying Excel environment..."
    Call Optimize(False)
    Application.StatusBar = "Ready for work"
End Sub
