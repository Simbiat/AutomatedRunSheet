Attribute VB_Name = "Filtering"
Option Explicit
Private Sub LateRows()
    Dim LateCell As Range, LateCellsExist As Boolean
    LateCellsExist = LateExists()
    Set LateCell = ReturnRange("LateSwitchCell")
    If LateMode() = True Then
        Application.StatusBar = "Enabling late mode..."
        Call ApplyFilter("RunSheetIsLateSpecialColumnData", False)
        If LateCellsExist = True Or (LateCellsExist = False And EditMode() = True) Then
            LateCell.Value2 = "Late Mode is ON"
        Else
            LateCell.ClearContents
            LateCell.ClearFormats
        End If
    Else
        Application.StatusBar = "Disabling late mode..."
        Call ApplyFilter("RunSheetIsLateSpecialColumnData", True, True)
        If LateCellsExist = True Or (LateCellsExist = False And EditMode() = True) Then
            LateCell.Value2 = "Late Mode is OFF"
        Else
            LateCell.ClearContents
            LateCell.ClearFormats
        End If
    End If
End Sub
Private Sub ApplyFilter(ByVal RangeName As String, ByVal Hide As Boolean, Optional ByVal WhatToHide As Boolean = False)
    Dim SubRange As Range
    If EditMode() = True Then
        'Ensure that all rows are always shown in Editor Mode
        ReturnRange(RangeName).EntireRow.Hidden = False
    Else
        For Each SubRange In ReturnRange(RangeName)
            If Hide = True And SubRange.Value2 = WhatToHide Then
                SubRange.EntireRow.Hidden = True
            Else
                SubRange.EntireRow.Hidden = False
            End If
        Next SubRange
    End If
End Sub
Private Sub EditorColumns()
    Dim IsEditorValue As Boolean, EditModeValue As Boolean, EditorButton As Range
    IsEditorValue = IsEditor()
    EditModeValue = EditMode()
    Set EditorButton = ReturnRange("EditorSwitchCell")
    If IsEditorValue = False Then
        Call SetSetting("EditMode", False)
    End If
    If EditModeValue = True And IsEditorValue = True Then
        Application.StatusBar = "Unhiding columns..."
        ReturnRange("RunSheetEditorColumns").EntireColumn.Hidden = False
        ReturnRange("SettingsIsEditableColumnData").EntireColumn.Hidden = False
        ReturnRange("SettingsIsCustomColumnData").EntireColumn.Hidden = False
        Call ApplyFilter("SettingsIsEditableColumnData", False)
        ReturnRange("RunSheetTimeStartColumnData").EntireColumn.Hidden = False
        ReturnRange("RunSheetTimeEndColumnData").EntireColumn.Hidden = False
        ReturnRange("RunSheetProcessingBlockColumnData").EntireColumn.Hidden = False
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
        EditorButton.Value2 = "Editor Mode is ON"
    Else
        Application.StatusBar = "Hiding columns..."
        ReturnRange("RunSheetEditorColumns").EntireColumn.Hidden = True
        Call ApplyFilter("SettingsIsEditableColumnData", True, False)
        ReturnRange("SettingsIsEditableColumnData").EntireColumn.Hidden = True
        ReturnRange("SettingsIsCustomColumnData").EntireColumn.Hidden = True
        'Hiding TimeStart and TimeEnd columns if they are empty to make up for slightly cleaner look
        If IsTimeColumnEmpty(ReturnRange("RunSheetTimeStartColumnData")) = True Then
            ReturnRange("RunSheetTimeStartColumnData").EntireColumn.Hidden = True
        End If
        If IsTimeColumnEmpty(ReturnRange("RunSheetTimeEndColumnData")) = True Then
            ReturnRange("RunSheetTimeEndColumnData").EntireColumn.Hidden = True
        End If
        'Hiding processing blocks if there are none set
        If BlocksExist() = False Then
            ReturnRange("RunSheetProcessingBlockColumnData").EntireColumn.Hidden = True
        End If
        Application.DisplayCommentIndicator = xlNoIndicator
        EditorButton.Value2 = "Editor Mode is OFF"
    End If
    Application.Run "ButtonToggleFormat", EditorButton
    If IsEditorValue = False Then
        'Hide button if user is not an editor
        EditorButton.ClearContents
        EditorButton.ClearFormats
    End If
End Sub
Private Function IsTimeColumnEmpty(ByVal CellRange As Range) As Boolean
    Dim cell As Range
    Dim ZeroTime As Date
    'Setting up our ZeroTime for optimization of the loop
    'Using Format to ensure consistency
    ZeroTime = TimeValue(Format("00:00", GetSetting("TimeFormat")))
    IsTimeColumnEmpty = True
    For Each cell In CellRange
        'For some reason if we do both IFs in 1 line IsTime always validates as True
        If IsTime(cell.Text) = True Then
            If TimeValue(Format(cell.Text, GetSetting("TimeFormat"))) <> ZeroTime Then
                IsTimeColumnEmpty = False
                Exit Function
            End If
        End If
    Next cell
End Function
Private Function BlocksExist() As Boolean
    Dim cell As Range
    BlocksExist = False
    For Each cell In ReturnRange("RunSheetProcessingBlockColumnData")
        If IsEmpty(cell.Value2) = False And Trim(cell.Value2) <> "" Then
            BlocksExist = True
            Exit Function
        End If
    Next cell
End Function
Public Function EditMode() As Boolean
    EditMode = CBool(GetSetting("EditMode"))
End Function
Public Function LateMode() As Boolean
    LateMode = CBool(GetSetting("LateFlag"))
End Function
Private Sub SpecialWorkDays()
    Dim CurDate As String, ReverseLogic As Boolean
    'Get reverse logic flag
    ReverseLogic = GetSetting("SpecialDaysReverse")
    'Get current date
    CurDate = Format(ReturnRange("CurrentDayCell").Value2, GetSetting("DateFormat"))
    'Check if date is present in lists of special dates and hide appropriate steps
    If ReverseLogic = False Then
        If IsInArray(CurDate, DatesArray(GetSetting("FirstSpecialDays"))) = True Then
            Call ApplyFilter("RunSheetIsFirstSpecialColumnData", True, False)
        Else
            Call ApplyFilter("RunSheetIsFirstSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("RegularSpecialDays"))) = True Then
            Call ApplyFilter("RunSheetIsRegularSpecialColumnData", True, False)
        Else
            Call ApplyFilter("RunSheetIsRegularSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("LastSpecialDays"))) = True Then
            Call ApplyFilter("RunSheetIsLastSpecialColumnData", True, False)
        Else
            Call ApplyFilter("RunSheetIsLastSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("FirstSpecialWorkDays"))) = True Then
            Call ApplyFilter("RunSheetIsFirstSpecialWorkColumnData", True, False)
        Else
            Call ApplyFilter("RunSheetIsFirstSpecialWorkColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("LastSpecialWorkDays"))) = True Then
            Call ApplyFilter("RunSheetIsLastSpecialWorkColumnData", True, False)
        Else
            Call ApplyFilter("RunSheetIsLastSpecialWorkColumnData", False)
        End If
    Else
        If IsInArray(CurDate, DatesArray(GetSetting("FirstSpecialDays"))) = False Then
            Call ApplyFilter("RunSheetIsFirstSpecialColumnData", True, True)
        Else
            Call ApplyFilter("RunSheetIsFirstSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("RegularSpecialDays"))) = False Then
            Call ApplyFilter("RunSheetIsRegularSpecialColumnData", True, True)
        Else
            Call ApplyFilter("RunSheetIsRegularSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("LastSpecialDays"))) = False Then
            Call ApplyFilter("RunSheetIsLastSpecialColumnData", True, True)
        Else
            Call ApplyFilter("RunSheetIsLastSpecialColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("FirstSpecialWorkDays"))) = False Then
            Call ApplyFilter("RunSheetIsFirstSpecialWorkColumnData", True, True)
        Else
            Call ApplyFilter("RunSheetIsFirstSpecialWorkColumnData", False)
        End If
        If IsInArray(CurDate, DatesArray(GetSetting("LastSpecialWorkDays"))) = False Then
            Call ApplyFilter("RunSheetIsLastSpecialWorkColumnData", True, True)
        Else
            Call ApplyFilter("RunSheetIsLastSpecialWorkColumnData", False)
        End If
    End If
End Sub

