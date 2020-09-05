Attribute VB_Name = "Ranges"
Option Explicit
'When we have several workbooks opened some of the functions with named ranges may fail, so we attempt to return ranges, that are linked to this workbook
Public Function ReturnRange(ByVal NamedRange As String, Optional ByVal SheetName As String = "") As Range
    If SheetName = "" Then
        Set ReturnRange = Range(ThisWorkbook.Name & "!" & NamedRange)
    Else
        Set ReturnRange = Range("[" & ThisWorkbook.Name & "]" & SheetName & "!" & NamedRange)
    End If
End Function
'Function to check if range exists (or not empty) based on https://stackoverflow.com/questions/12611900/test-if-range-exists-in-vba
Public Function RangeExists(ByVal RangeToCheck As String, Optional ByVal SheetName As String = "") As Boolean
    RangeExists = False
    Dim TempRange
    On Error Resume Next
    Set TempRange = ReturnRange(RangeToCheck, SheetName)
    If Err = 1004 Then
        RangeExists = False
    Else
        RangeExists = True
    End If
    Err.Clear
    On Error GoTo 0
End Function
Public Sub UpdateRanges()
    Dim EndRow As Integer
    Dim RangeTitle As Range, CellType As Range
    Dim StepMaxLength As Integer
    StepMaxLength = CInt(GetSetting("StepMaxLength"))
    
    'Get last row for RunSheet
    If RangeExists("NewStepMenu") = True Then
        EndRow = CInt(ReturnRange("NewStepMenu").Row - 1)
    Else
        EndRow = CInt(ThisWorkbook.Worksheets("RunSheet").UsedRange.Rows.Count)
    End If
    
    'Clear named ranges
    Dim RangeName As Name
    Application.StatusBar = "Clearing named ranges..."
    'Have to skip errors, becasue in some cases get name "=#NAME?" out of nowhere (probably realted to Excel's hidden ranges it creates when using some functions)
    On Error Resume Next
    For Each RangeName In ThisWorkbook.Names
        RangeName.Delete
    Next RangeName
    On Error GoTo 0
    
    Application.StatusBar = "Creating named ranges..."
    
    'Special menu for adding a new step at the end of the current list
    ReturnRange(ThisWorkbook.Worksheets("RunSheet").Cells(EndRow + 1, 7).Address, "RunSheet").Name = "NewStepMenu"
    
    'Settings sheet
    'Titles row
    ReturnRange("A" & CStr(1) & ":" & ThisWorkbook.Worksheets("Settings").Cells(1, Columns.Count).End(xlToLeft).Address, "Settings").Name = "SettingsTitleRow"
    'Columns
    For Each RangeTitle In ReturnRange("SettingsTitleRow").Columns
        'Name column
        ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(ThisWorkbook.Worksheets("Settings").UsedRange.Rows.Count - 1, 0).Address, "Settings").Name = "Settings" & RangeTitle.Value2 & "ColumnData"
        'Special ranges dependant on setting type
        Select Case LCase(RangeTitle.Value2)
            Case "id"
                For Each CellType In ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")
                    If IsEmpty(CellType.Value2) = True Or CellType.Value2 = "" Then
                        'MAX() works only on numeric values, which allows us not to touch non-numeric ones, if we ever choose to use such
                        'Essentially below we are replicating a auto-increment, but not limited to integer
                        CellType.Value2 = Application.WorksheetFunction.Max(CellType.EntireColumn) + 1
                    Else
                        CellType.Value2 = AlphaNumeric(CellType.Value2)
                    End If
                    If LCase(CellType.Value2) = "editmode" Then
                        CellType.Offset(0, 4).Name = "SettingsEditModeValue"
                    End If
                    If LCase(CellType.Value2) = "lateflag" Then
                        CellType.Offset(0, 4).Name = "SettingsLateFlagValue"
                    End If
                Next CellType
            Case "type"
                For Each CellType In ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")
                    'Set default if an unsupported value
                    If IsError(Application.Match(LCase(CellType.Value2), Array("string", "array", "boolean", "integer", "range", "time", "color"), False)) Then
                        CellType.Value2 = "String"
                    End If
                    Select Case LCase(CellType)
                        'Turn cell types to named ranges
                        Case "range"
                            If IsEmpty(CellType.Offset(0, 3).Value2) = False And CellType.Offset(0, 3).Value2 <> "" Then
                                'Suppress error in case some bad value was entered into the setting
                                On Error Resume Next
                                ReturnRange(CellType.Offset(0, 3).Value2, "RunSheet").Name = CellType.Offset(0, -1).Value2
                                On Error GoTo 0
                            End If
                        'Settings for colors to apply coloring logic
                        Case "color"
                            'Checking if range is currently empty, to avoid error
                            If RangeExists("SettingsColors") = True Then
                                Union(ReturnRange("SettingsColors"), CellType.Offset(0, 3)).Name = "SettingsColors"
                            Else
                                CellType.Offset(0, 3).Name = "SettingsColors"
                            End If
                        'Add to time cells range for styling
                        Case "time"
                            'Checking if range is currently empty, to avoid error
                            If RangeExists("SettingsTimeCells") = True Then
                                Union(ReturnRange("SettingsTimeCells"), CellType.Offset(0, 3)).Name = "SettingsTimeCells"
                            Else
                                CellType.Offset(0, 3).Name = "SettingsTimeCells"
                            End If
                        'Add to boolean cells range for styling
                        Case "boolean"
                            'Checking if range is currently empty, to avoid error
                            If RangeExists("SettingsBooleanCells") = True Then
                                Union(ReturnRange("SettingsBooleanCells"), CellType.Offset(0, 3)).Name = "SettingsBooleanCells"
                            Else
                                CellType.Offset(0, 3).Name = "SettingsBooleanCells"
                            End If
                    End Select
                Next CellType
            Case "iseditable"
                'Add to boolean cells range for styling
                If RangeExists("SettingsBooleanCells") = True Then
                    Union(ReturnRange("SettingsBooleanCells"), ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")).Name = "SettingsBooleanCells"
                Else
                    ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData").Name = "SettingsBooleanCells"
                End If
                'Add cells, that are editable in non-editor mode to unlock them later
                For Each CellType In ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")
                    If CellType.Value2 = True Then
                        If RangeExists("SettingsEditable") = True Then
                            Union(ReturnRange("SettingsEditable"), CellType.Offset(0, 1)).Name = "SettingsEditable"
                        Else
                            CellType.Offset(0, 1).Name = "SettingsEditable"
                        End If
                    End If
                Next CellType
            Case "iscustom"
                'Set defaul value if not already set
                For Each CellType In ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")
                    If IsEmpty(CellType.Value2) = True Or CellType.Value2 = "" Then
                        CellType.Value2 = True
                    End If
                Next CellType
                'Add to boolean cells range for styling
                If RangeExists("SettingsBooleanCells") = True Then
                    Union(ReturnRange("SettingsBooleanCells"), ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData")).Name = "SettingsBooleanCells"
                Else
                    ReturnRange("Settings" & RangeTitle.Value2 & "ColumnData").Name = "SettingsBooleanCells"
                End If
        End Select
    Next RangeTitle
    
    'Set range for buttons (using main ones)
    Union(ReturnRange("LateSwitchCell"), ReturnRange("EditorSwitchCell"), ReturnRange("EndOfDayCell")).Name = "RunSheetButtons"
    
    'RunSheet sheet
    'Titles row
    ReturnRange("A" & 4 & ":" & ThisWorkbook.Worksheets("RunSheet").Cells(4, Columns.Count).End(xlToLeft).Address, "RunSheet").Name = "RunSheetTitleRow"
    'Columns
    For Each RangeTitle In ReturnRange("RunSheetTitleRow").Columns
        'Size the comment box (in some cases it seems to span to the bottom of the sheet)
        If Not RangeTitle.Comment Is Nothing Then
            'Heisenbug that seems to fail to properly initialize comment's shape sometimes, so ignoring it
            On Error Resume Next
            RangeTitle.Comment.Shape.TextFrame2.AutoSize = True
            On Error GoTo 0
        End If
        Select Case LCase(RangeTitle.Value2)
            Case "islatespecial", "isfirstspecial", "isregularspecial", "islastspecial", "isfirstspecialwork", "islastspecialwork"
                'Create range
                ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet").Name = "RunSheet" & RangeTitle.Value2 & "ColumnData"
                'Add to boolean cells range for styling
                Call AddToRange(ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData"), "RunSheetBooleanCells", False)
                'Add to editor-only columns
                Call AddToRange(ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData"), "RunSheetEditorColumns", False)
                'Set default value for all empty cells
                'Use Resume Next to skip error, if there are no empty cell
                On Error Resume Next
                If LCase(RangeTitle.Value2) = "islatespecial" Then
                    ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData").SpecialCells(xlCellTypeBlanks).Value2 = False
                Else
                    ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData").SpecialCells(xlCellTypeBlanks).Value2 = True
                End If
                On Error GoTo 0
            Case "type"
                'Create range
                ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet").Name = "RunSheet" & RangeTitle.Value2 & "ColumnData"
                'Add to editor-only columns
                Call AddToRange(ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData"), "RunSheetEditorColumns", False)
                For Each CellType In ReturnRange("RunSheet" & RangeTitle.Value2 & "ColumnData")
                    'Set default if an unsupported value
                    If IsError(Application.Match(LCase(CellType.Value2), Array("regular", "button", "delimiter"), False)) Then
                        CellType.Value2 = "Regular"
                    End If
                    'Style cells based on type
                    Application.Run "StepStyle", CellType
                    Select Case LCase(CellType.Value2)
                        'Set range for delimiters
                        Case "delimiter"
                            Call AddToRange(Union(CellType.Offset(0, 1), CellType.Offset(0, 2), CellType.Offset(0, 3), CellType.Offset(0, 4), CellType.Offset(0, 5), CellType.Offset(0, 6), CellType.Offset(0, 7), CellType.Offset(0, 8)), "RunSheetDelimiters", False)
                        Case "button"
                            'Set range for buttons
                            Call AddToRange(CellType.Offset(0, 2), "RunSheetButtons", False)
                    End Select
                Next CellType
            Case "stepname"
                For Each CellType In ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet")
                    'Sanitize step name
                    If IsEmpty(CellType.Value2) = True Or CellType.Value2 = "" Then
                        CellType.Value2 = "New Step"
                    Else
                        'Ensure length complies with global setting
                        CellType.Value2 = Left(CellType.Value2, StepMaxLength)
                    End If
                    'Add to range
                    Call AddToRange(CellType, "RunSheet" & RangeTitle.Value2 & "ColumnData")
                Next CellType
            Case "timestart", "timeend"
                For Each CellType In ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet")
                    'Sanitize time value
                    If IsTime(CellType.Text) = False Then
                        If RangeExists("RunSheetDelimiters") = True Then
                            If Intersect(CellType, ReturnRange("RunSheetDelimiters")) Is Nothing Then
                                CellType.Value2 = "00:00"
                            End If
                        Else
                            CellType.Value2 = "00:00"
                        End If
                    End If
                    'Add to range
                    Call AddToRange(CellType, "RunSheet" & RangeTitle.Value2 & "ColumnData")
                Next CellType
            Case "description", "time", "user", "status"
                For Each CellType In ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet")
                    'Add to range
                    Call AddToRange(CellType, "RunSheet" & RangeTitle.Value2 & "ColumnData")
                Next CellType
            Case "processingblock"
                For Each CellType In ReturnRange(RangeTitle.Offset(1, 0).Address & ":" & RangeTitle.Offset(EndRow - 4, 0).Address, "RunSheet")
                    'Ensure length complies with global setting
                    CellType.Value2 = Left(CellType.Value2, StepMaxLength)
                    'Add to range
                    Call AddToRange(CellType, "RunSheet" & RangeTitle.Value2 & "ColumnData")
                Next CellType
        End Select
    Next RangeTitle
    
    'Set range of UI elements on RunSheet, values of which we do not allow changing
    Union(ReturnRange("LateSwitchCell"), ReturnRange("EditorSwitchCell"), ReturnRange("WelcomeCell"), ReturnRange("CurrentDayTextCell"), _
                            ReturnRange("PreviousDayTextCell"), ReturnRange("NextDayTextCell"), ReturnRange("EndOfDayCell")).Name = "RunSheetUI"
End Sub
Private Sub AddToRange(ByVal CellType As Range, ByVal RangeName As String, Optional ByVal DelimiterCheck As Boolean = True)
    'Check if Delimiters range if we want to (not requried for some ranges)
    If DelimiterCheck = True Then
        If RangeExists("RunSheetDelimiters") = True Then
            'Check if cell is among the delimiters
            If Not Intersect(CellType, ReturnRange("RunSheetDelimiters")) Is Nothing Then
                'Cancel adding cell to range
                Exit Sub
            End If
        End If
    End If
    'Check if range exists
    If RangeExists(RangeName) = True Then
        'Update range
        Union(ReturnRange(RangeName), CellType).Name = RangeName
    Else
        'Create range
        CellType.Name = RangeName
    End If
End Sub
