Attribute VB_Name = "Security"
Option Explicit
'Obviously change this on production and be careful whom you give access to the code and to the workbook in general
Private Const WorkbookPassword As String = "1234567890"
'Function to get username with domain
Public Function FullName() As String
    FullName = VBA.Environ$("USERDOMAIN") + "\" + VBA.Environ$("USERNAME")
End Function
'Updates name of current user in UI
Private Sub Welcomen()
    ReturnRange("WelcomeCell").Value2 = "Welcome, " + FullName() + "!"
    Call ApplyStyle(ReturnRange("WelcomeCell"), "WelcomeCell")
End Sub
'If we will be saving the welcome cell, it will be updated everytime for all users, this sub is to rever that on save
Private Sub UnWelcomen()
    ReturnRange("WelcomeCell").Value2 = "Welcome, $UserName!"
End Sub
'Function to check if current user is an editor
Public Function IsEditor() As Boolean
    If IsEmpty(GetSetting("Editors")) Or GetSetting("Editors") = "" Then
        'If no users are defined consider that this first launch before actual setup and allow editing
        IsEditor = True
    Else
        If IsInArray(FullName(), Split(GetSetting("Editors"), GetSetting("ArrayDelimiter"), -1, 1)) = True Then
            IsEditor = True
        Else
            IsEditor = False
        End If
    End If
End Function
'Using this to prevent accidental selection of a hidden row, when using arrows
Private Sub NextVisible(ByVal CurrentCell As Range, Optional ByVal SkipCompleted As Boolean = True)
    'Do this only if not in Editor mode to allow editor selecting rows or groups of cells
    If EditMode() = False Then
findnextcell:
        If ActiveCell.EntireRow.Hidden = True Then
            'Skip hidden rows
            ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible).Activate
            GoTo findnextcell
        Else
            If CurrentCell.Worksheet.Name = "RunSheet" Then
                If ActiveCell.Value2 = "Completed" And SkipCompleted = True Then
                    'Skip cells marked as completed
                    ActiveCell.Offset(1, 0).Activate
                Else
                    'Skip delimiters
                    If RangeExists("RunSheetDelimiters") = True Then
                        If Not Intersect(ActiveCell, ReturnRange("RunSheetDelimiters")) Is Nothing Then
                            ActiveCell.Offset(1, 0).Activate
                        End If
                    End If
                End If
            End If
        End If
        ActiveCell.Cells(1, 1).Select
        'Change selection in case the selected one is locked or hidden still
        Call SafeCell
    End If
End Sub
'Function to (re)apply protection with UserInterfaceOnly
Private Sub AllowUI(Optional ByVal ProtectCheck As Boolean = True)
    Dim Sheet As Worksheet
    For Each Sheet In ThisWorkbook.Worksheets
        'If we are just starting workbook or updating this on save, we need to check, whether a sheet is even protected and if we need to protect it (that is not editor mode)
        'We also need to check that workbook is not shared, because otherwise it will fail
        'UserInterfaceOnly is set to False for consistent behaviour
        'Allowing FormattingColumns and FormattingRows to allow hiding of elements
        If (ProtectCheck = False Or (ProtectCheck = True And Sheet.ProtectContents = True)) And ThisWorkbook.MultiUserEditing = False Then
            Sheet.Protect WorkbookPassword, _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, _
            AllowInsertingHyperlinks:=False, _
            AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, _
            AllowSorting:=False, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=False, _
            UserInterfaceOnly:=False
        End If
    Next Sheet
End Sub
Private Sub CellLock()
    Dim Sheet As Worksheet
    'Lock all cells first
    For Each Sheet In ThisWorkbook.Worksheets
        Sheet.Cells.Locked = True
    Next Sheet
    'Unlock specific cells, which we can interract with
    ReturnRange("RunSheetTimeColumnData").Locked = False
    ReturnRange("RunSheetUserColumnData").Locked = False
    ReturnRange("RunSheetStatusColumnData").Locked = False
    ReturnRange("SettingsEditable").Locked = False
    ReturnRange("CurrentDayCell").Locked = False
    ReturnRange("PreviousDayCell").Locked = False
    ReturnRange("NextDayCell").Locked = False
    'Cells, that should not be used by users, but we can't lock them because of shared workbook
    ReturnRange("WelcomeCell").Locked = False
    ReturnRange("EditorSwitchCell").Locked = False
    ReturnRange("LateSwitchCell").Locked = False
    ReturnRange("SettingsEditModeValue").Locked = False
    ReturnRange("SettingsLateFlagValue").Locked = False
    If RangeExists("RunSheetButtons", "RunSheet") = True Then
        ReturnRange("RunSheetButtons").Locked = False
    End If
End Sub
Private Sub UnprotectSheet(ByVal SheetName As Worksheet)
    SheetName.Unprotect WorkbookPassword
End Sub
Private Sub ProtectMe(ByVal Safe As Boolean)
    Dim Sheet As Worksheet
    If Safe = True Then
        'Hide EditorManual sheet
        If ThisWorkbook.Worksheets("EditorManual").Visible <> xlSheetVeryHidden Then
            ThisWorkbook.Worksheets("EditorManual").Visible = xlSheetVeryHidden
        End If
        'Locking cells
        Application.StatusBar = "Locking cells..."
        Application.Run "CellLock"
        'Protect worksheets
        Application.StatusBar = "Protecting sheets..."
        Call AllowUI(False)
        'Protect workbook
        Application.StatusBar = "Protecting workbook..."
        ThisWorkbook.Protect WorkbookPassword, True, True
        'Share workbook
        Application.StatusBar = "Sharing workbook..."
'        If ThisWorkbook.MultiUserEditing = False Then
'            On Error Resume Next
'            ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, AccessMode:=xlShared
'            If Err.Number <> 0 Then
'                Call WriteLog("Failed to save workbook on sharing with error #" & Err.Number)
'                MsgBox "Failed to save workbook with error #" & Err.Number, vbCritical + vbOKOnly + vbApplicationModal, "Failed to save"
'            End If
'            On Error GoTo 0
'            ThisWorkbook.AutoUpdateFrequency = 5
'            ThisWorkbook.AutoUpdateSaveChanges = True
'            ThisWorkbook.ConflictResolution = xlLocalSessionChanges
'            ThisWorkbook.KeepChangeHistory = False
'            ThisWorkbook.PersonalViewPrintSettings = False
'            ThisWorkbook.PersonalViewListSettings = False
'        End If
    Else
        'The order should be reversed here
        If ThisWorkbook.MultiUserEditing = True Then
            Application.StatusBar = "Unsharing workbook..."
            ThisWorkbook.ExclusiveAccess
        End If
        Application.StatusBar = "Unprotecting workbook..."
        ThisWorkbook.Unprotect WorkbookPassword
        Application.StatusBar = "Unprotecting sheets..."
        For Each Sheet In ThisWorkbook.Worksheets
            If Sheet.ProtectContents = True Then
                Sheet.Unprotect WorkbookPassword
            End If
        Next Sheet
        If ThisWorkbook.Worksheets("EditorManual").Visible <> xlSheetVisible Then
            ThisWorkbook.Worksheets("EditorManual").Visible = xlSheetVisible
        End If
        'Save workbook again, but only if it's not exclusive already (this will result in error)
        If ThisWorkbook.MultiUserEditing = True Then
            On Error Resume Next
            ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, AccessMode:=xlExclusive
            If Err.Number <> 0 Then
                Call WriteLog("Failed to save workbook on unsharing with error #" & Err.Number)
                MsgBox "Failed to save workbook with error #" & Err.Number, vbCritical + vbOKOnly + vbApplicationModal, "Failed to save"
            End If
            On Error GoTo 0
        End If
    End If
End Sub
'Function to clear dead sessions based on https://superuser.com/questions/961918/how-do-you-prevent-corruption-of-shared-excel-files but without potential issues with date formating
Private Sub ClearSessions()
    Dim TimeLimit As Integer
    Dim Users As Variant
    Dim User As Integer
    
    Application.StatusBar = "Clearing old sessions..."
    'Get time
    TimeLimit = CInt(Format(GetSetting("SessionTimeOut"), "nn")) + CInt(Format(GetSetting("SessionTimeOut"), "hh")) * 60
    'Get users
    Users = ThisWorkbook.UserStatus
    
    For User = UBound(Users) To 1 Step -1
        If DateDiff("n", Users(User, 2), Now()) > TimeLimit Then
            ThisWorkbook.RemoveUser (User)
        End If
    Next
    Application.StatusBar = "Cleared old sessions"
End Sub
'Select a safe cell to avoid potential issues with objects or locked cells
Private Sub SafeCell()
    If TypeName(Selection) <> "Range" Then
        If ThisWorkbook.ActiveSheet.Name = "RunSheet" Then
            ReturnRange("RunSheetStatusColumnData").Cells(1, 1).Select
        ElseIf ThisWorkbook.ActiveSheet.Name = "Settings" Then
            ReturnRange("SettingsEditable").Cells(1, 1).Select
        End If
    Else
        If Selection.Locked = True Or Selection.EntireRow.Hidden = True Or Selection.EntireColumn.Hidden = True Then
            If ThisWorkbook.ActiveSheet.Name = "RunSheet" Then
                If ReturnRange("RunSheetStatusColumnData").Cells(Selection.Row - 4, 1).Locked = False And ReturnRange("RunSheetStatusColumnData").Cells(Selection.Row - 4, 1).EntireRow.Hidden = False And ReturnRange("RunSheetStatusColumnData").Cells(Selection.Row - 4, 1).EntireColumn.Hidden = False Then
                    ReturnRange("RunSheetStatusColumnData").Cells(Selection.Row - 4, 1).Select
                Else
                    ReturnRange("RunSheetStatusColumnData").Cells(1, 1).Select
                End If
            ElseIf ThisWorkbook.ActiveSheet.Name = "Settings" Then
                If ReturnRange("SettingsEditable").Cells(Selection.Row - 1, 1).Locked = False And ReturnRange("SettingsEditable").Cells(Selection.Row, 1).EntireRow.Hidden = False And ReturnRange("SettingsEditable").Cells(Selection.Row, 1).EntireColumn.Hidden = False Then
                    ReturnRange("SettingsEditable").Cells(Selection.Row - 1, 1).Select
                Else
                    ReturnRange("SettingsEditable").Cells(1, 1).Select
                End If
            End If
        End If
    End If
End Sub
