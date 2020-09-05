Option Explicit
Dim OldValue As String
Dim OldText As String
Dim UsedRows As Integer
Private Sub Worksheet_Activate()
    ThisWorkbook.Worksheets("Settings").DisplayPageBreaks = False
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    ThisWorkbook.Worksheets("Settings").Calculate
    'Counting used rows to track removed rows
    UsedRows = ThisWorkbook.Worksheets("Settings").UsedRange.Rows.Count
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.CountLarge = 1 Then
        Call Optimize(True)
        Application.Run "TrackChange", Target, OldValue, OldText
        Call Optimize(False)
    End If
    
    'Prevent removal of columns
    If Target.Address = Target.EntireColumn.Address Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
    
    If Target.Address = Target.EntireRow.Address Then
        'Check if row was inserted or removed within our main grid
        If Not Intersect(Target.EntireRow.Cells(1, 1), ReturnRange("SettingsIDColumnData").Resize(ReturnRange("SettingsIDColumnData").Rows.Count + Target.Rows.Count, ReturnRange("SettingsIDColumnData").Columns.Count)) Is Nothing Then
            Call Optimize(True)
            Dim RowRange As Range
            Dim RowsCount As Integer
            'Will need rows count to do offset when handling rows removal
            RowsCount = Target.Rows.Count
            If ThisWorkbook.Worksheets("Settings").UsedRange.Rows.Count < UsedRows Then
                'Undo rows removal
                Application.Undo
                For Each RowRange In Target.Rows
                    'Check if this is a custom setting, since we would like to protect system settings from removal
                    If RowRange.Cells(1, 1).Offset(-RowsCount, 2).Value2 = True Then
                        'Properly removing the row now
                        Call WriteLog("Removed setting '" & RowRange.Cells(1, 1).Offset(-RowsCount, 0).Value2 & "'")
                        RowRange.Cells(1, 1).Offset(-RowsCount, 0).EntireRow.Delete
                    End If
                Next RowRange
                Application.Run "UpdateRanges"
            Else
                'Row(s) was added (probably)
                For Each RowRange In Target.Rows
                    Call WriteLog("Inserted new row in 'Settings' at '" & RowRange.Address & "'")
                Next RowRange
                Application.Run "UpdateRanges"
            End If
            Call Optimize(False)
        End If
    End If
    
    Application.Run "NextVisible", Target
    Application.StatusBar = "Ready for work"
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Save old value (for tracking)
    If Target.Cells.CountLarge = 1 Then
        OldValue = CStr(Target.Value2)
        OldText = CStr(Target.Text)
    Else
        If Target.Address = Target.EntireRow.Address Then
            'Counting used rows to track removed rows
            UsedRows = ThisWorkbook.Worksheets("Settings").UsedRange.Rows.Count
        End If
    End If
    
    Application.Run "NextVisible", Target
End Sub
