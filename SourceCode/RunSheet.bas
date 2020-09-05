VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RunSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim OldValue As String
Dim OldText As String
Dim UsedRows As Integer
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Prevent autocomplete for Status column
    Application.EnableAutoComplete = Intersect(Target, ReturnRange("RunSheetStatusColumnData")) Is Nothing
    'Save old value (for tracking)
    If Target.Cells.CountLarge = 1 Then
        OldValue = CStr(Target.Value2)
        OldText = CStr(Target.Text)
    Else
        If Target.Address = Target.EntireRow.Address Then
            'Counting used rows to track removed rows
            UsedRows = ThisWorkbook.Worksheets("RunSheet").UsedRange.Rows.Count
        End If
    End If
    
    'Need to send False for SkipCompleted to allow manual selection of cells marked as Completed, for example if a cell was marked by mistake
    Application.Run "NextVisible", Target, False
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.CountLarge = 1 Then
        Call Optimize(True)
        Application.Run "TrackChange", Target, OldValue, OldText
        Call Optimize(False)
    Else
        'Prevent removal of columns
        If Target.Address = Target.EntireColumn.Address Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
        End If
        
        If Target.Address = Target.EntireRow.Address Then
            'Check if row was inserted or removed within our main grid
            If Not Intersect(Target.EntireRow.Cells(1, 7), ReturnRange("RunSheetTypeColumnData").Resize(ReturnRange("RunSheetTypeColumnData").Rows.Count + Target.Rows.Count, ReturnRange("RunSheetTypeColumnData").Columns.Count)) Is Nothing Then
                Call Optimize(True)
                Dim RowRange As Range
                Dim RowsCount As Integer
                'Will need rows count to do offset when handling rows removal
                RowsCount = Target.Rows.Count
                If ThisWorkbook.Worksheets("RunSheet").UsedRange.Rows.Count < UsedRows Then
                    'Undo rows removal
                    Application.Undo
                    For Each RowRange In Target.Rows
                        'Check if step was a button, in order to remove its settings as well
                        If LCase(RowRange.Cells(1, 1).Offset(-RowsCount, 6).Value2) = "button" Then
                            Application.Run "ButtonUpdate", RowRange.Cells(1, 1).Offset(-RowsCount, 8).Value2, True
                        End If
                        'Properly removing the row now
                        RowRange.Cells(1, 1).Offset(-RowsCount, 0).EntireRow.Delete
                        Call WriteLog("Removed step '" & RowRange.Cells(1, 1).Offset(-RowsCount, 9).Value2 & "'")
                    Next RowRange
                    Application.Run "UpdateRanges"
                ElseIf ThisWorkbook.Worksheets("RunSheet").UsedRange.Rows.Count > UsedRows Then
                    'Row(s) was added (probably)
                    'Style the row as a regular step, if it was inserted inside our main grid
                    For Each RowRange In Target.Rows
                        RowRange.EntireRow.Cells(1, 7).Value2 = "Regular"
                        RowRange.EntireRow.Cells(1, 9).Value2 = "New Step"
                        Application.Run "StepStyle", RowRange.EntireRow.Cells(1, 7), True
                        Call WriteLog("Inserted new row in 'RunSheet' at '" & RowRange.Address & "'")
                    Next RowRange
                    Application.Run "UpdateRanges"
                Else
                    'Most likely rows were moved
                    For Each RowRange In Target.Rows
                        Call WriteLog("Step '" & RowRange.EntireRow.Cells(1, 9).Value2 & "' moved to '" & RowRange.Address & "'")
                    Next RowRange
                End If
                Call Optimize(False)
            End If
        End If
    End If
    
    Application.Run "NextVisible", Target
    Application.StatusBar = "Ready for work"
End Sub
Private Sub Worksheet_Activate()
    ThisWorkbook.Worksheets("RunSheet").DisplayPageBreaks = False
    ThisWorkbook.Worksheets("RunSheet").Calculate
    'Counting used rows to track removed rows
    UsedRows = ThisWorkbook.Worksheets("RunSheet").UsedRange.Rows.Count
    'Skipping error, since on start this will fail (before Init())
    On Error GoTo skiperror
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 4
        .FreezePanes = True
    End With
skiperror:
End Sub
'Need to use double-click for buttons,because I was not able to make several solutions, I've found, work
Private Sub WorkSheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not ReturnRange("RunSheetButtons") Is Nothing Then
        Application.EnableEvents = False
        If Not Intersect(Target, ReturnRange("RunSheetButtons")) Is Nothing Then
            'Cancel standard behaviour
            Cancel = True
            'Do nothing, if it's a disabled Editor/Late button, essentially selecting the cell as a regular one
            If Not Intersect(Target, ReturnRange("EditorSwitchCell")) Is Nothing And IsEditor() = False Then
                Exit Sub
            End If
            If Not Intersect(Target, ReturnRange("LateSwitchCell")) Is Nothing And LateExists() = False Then
                Exit Sub
            End If
            'Select next cell down to make the animation visible
            Target.Offset(1, 0).Select
            Dim FontColor As Long, BackColor As Long
            FontColor = Target.Font.Color
            BackColor = Target.Interior.Color
            'Animate button down
            Application.Run "AnimateButtonClick", Target, ToneRGB(RangeColorToString(Target.Font.Color), True, 1), ToneRGB(RangeColorToString(Target.Interior.Color), True, 1)
            'Using Sleep to allow visual queue for button click
            Sleep 30
            Application.ScreenUpdating = False
            'Actual processing
            If Not Intersect(Target, ReturnRange("LateSwitchCell")) Is Nothing Or Not Intersect(Target, ReturnRange("EditorSwitchCell")) Is Nothing Then
                If Not Intersect(Target, ReturnRange("LateSwitchCell")) Is Nothing Then
                    Application.Run "LateButton", FontColor, BackColor
                    Dim LateCellsExist
                    LateCellsExist = LateExists()
                    If LateCellsExist = False Or (LateCellsExist = False And EditMode() = False) Then
                        Application.Run "AnimateButtonClick", Target, Target.Font.Color, Target.Interior.Color
                        Application.Run "ButtonToggleFormat", Target
                    End If
                ElseIf Not Intersect(Target, ReturnRange("EditorSwitchCell")) Is Nothing Then
                    Application.Run "EditModeButton", FontColor, BackColor
                    Application.Run "AnimateButtonClick", Target, Target.Font.Color, Target.Interior.Color
                    Application.Run "ButtonToggleFormat", Target
                End If
            Else
                If Not Intersect(Target, ReturnRange("EndOfDayCell")) Is Nothing Then
                    Application.DisplayAlerts = False
                    Application.Run "EndOfDayButton"
                    Application.DisplayAlerts = True
                Else
                    Application.Run "CustomButtonClick", Target
                End If
                Application.Run "AnimateButtonClick", Target, FontColor, BackColor
            End If
            'Animate button up
            Sleep 30
            Application.ScreenUpdating = True
        End If
        Application.Run "NextVisible", Target
        Application.EnableEvents = True
    End If
End Sub

