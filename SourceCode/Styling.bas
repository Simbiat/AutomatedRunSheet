Attribute VB_Name = "Styling"
Option Explicit
'Add update custom styles used for cells
Private Sub UpdateStyles()
    'Create collection of styles specific to RunSheet
    'Using collection instead of array, because at the time of writing, I do not have a final number in mind and feel this is more appropriate
    Application.StatusBar = "Updating styles..."
    Dim RunSheetStyles As New Collection, RunSheetStyle As Variant, StyleToSet As Style
    RunSheetStyles.Add "Delimiter"
    RunSheetStyles.Add "StepName"
    RunSheetStyles.Add "TimeCell"
    RunSheetStyles.Add "DateCell"
    RunSheetStyles.Add "MissedTime"
    RunSheetStyles.Add "Description"
    RunSheetStyles.Add "StepType"
    RunSheetStyles.Add "SettingType"
    RunSheetStyles.Add "Button"
    RunSheetStyles.Add "ButtonTrue"
    RunSheetStyles.Add "ButtonFalse"
    RunSheetStyles.Add "BooleanTrue"
    RunSheetStyles.Add "BooleanFalse"
    RunSheetStyles.Add "DateTimeCell"
    RunSheetStyles.Add "StatusCell"
    RunSheetStyles.Add "DateTimeCellCompleted"
    RunSheetStyles.Add "StatusCellCompleted"
    RunSheetStyles.Add "DateTimeCellInProgress"
    RunSheetStyles.Add "StatusCellInProgress"
    RunSheetStyles.Add "DateTimeCellFailed"
    RunSheetStyles.Add "StatusCellFailed"
    RunSheetStyles.Add "DateTimeCellSkipped"
    RunSheetStyles.Add "StatusCellSkipped"
    RunSheetStyles.Add "WelcomeCell"
    For Each RunSheetStyle In RunSheetStyles
        'Check if style exists and add it, if it does not, otherwise update it accordingly
        If StyleExists("RunSheet" & RunSheetStyle) = True Then
            Set StyleToSet = ThisWorkbook.Styles("RunSheet" & RunSheetStyle)
        Else
            Set StyleToSet = ThisWorkbook.Styles.Add("RunSheet" & RunSheetStyle)
        End If
        With StyleToSet
            .IncludeFont = True
            .IncludeNumber = True
            .IncludeAlignment = True
            .IncludeBorder = True
            .IncludePatterns = True
            .IncludeProtection = True
            .AddIndent = False
            .IndentLevel = 0
            .FormulaHidden = True
            'MSDN says, that MergeCells is RW, but it appears it's not so
            '.MergeCells = False
            .Orientation = xlHorizontal
            .ReadingOrder = xlContext
            .ShrinkToFit = False
            'Set font styling to that from "Normal" builtin style by default
            'Skip errors here, in case it's missing for some reason
            'Sadly cant just do .Font = .Font - results in error
            On Error Resume Next
            .Font.Bold = ThisWorkbook.Styles("Normal").Font.Bold
            .Font.Color = ThisWorkbook.Styles("Normal").Font.Color
            .Font.ColorIndex = ThisWorkbook.Styles("Normal").Font.ColorIndex
            .Font.FontStyle = ThisWorkbook.Styles("Normal").Font.FontStyle
            .Font.Italic = ThisWorkbook.Styles("Normal").Font.Italic
            .Font.Size = ThisWorkbook.Styles("Normal").Font.Size
            .Font.Strikethrough = ThisWorkbook.Styles("Normal").Font.Strikethrough
            .Font.Subscript = ThisWorkbook.Styles("Normal").Font.Subscript
            .Font.Superscript = ThisWorkbook.Styles("Normal").Font.Superscript
            .Font.ThemeColor = ThisWorkbook.Styles("Normal").Font.ThemeColor
            .Font.ThemeFont = ThisWorkbook.Styles("Normal").Font.ThemeFont
            .Font.TintAndShade = ThisWorkbook.Styles("Normal").Font.TintAndShade
            .Font.Underline = ThisWorkbook.Styles("Normal").Font.Underline
            .Interior.Color = ThisWorkbook.Styles("Normal").Interior.Color
            .Interior.Pattern = ThisWorkbook.Styles("Normal").Interior.Pattern
            .Borders.LineStyle = ThisWorkbook.Styles("Normal").Borders.LineStyle
            .Borders.Color = ThisWorkbook.Styles("Normal").Borders.Color
            .Borders.Weight = ThisWorkbook.Styles("Normal").Borders.Weight
            .HorizontalAlignment = ThisWorkbook.Styles("Normal").HorizontalAlignment
            .VerticalAlignment = ThisWorkbook.Styles("Normal").VerticalAlignment
            .NumberFormat = ThisWorkbook.Styles("Normal").NumberFormat
            .NumberFormatLocal = ThisWorkbook.Styles("Normal").NumberFormatLocal
            .WrapText = ThisWorkbook.Styles("Normal").WrapText
            .Locked = ThisWorkbook.Styles("Normal").Locked
            On Error GoTo 0
            'Global but potentially not for "Normal" style
            .VerticalAlignment = xlVAlignCenter
            .WrapText = False
            Select Case RunSheetStyle
                Case "WelcomeCell"
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlVAlignCenter
                    .Font.Bold = True
                    .Borders.LineStyle = xlLineStyleNone
                    .NumberFormat = "General"
                    .Locked = True
                Case "Delimiter"
                    .Font.Color = StringToRGB(GetSetting("ColorDelimiter"))
                    .Interior.Color = .Font.Color
                    .Interior.Pattern = xlPatternSolid
                    .HorizontalAlignment = xlCenter
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    'Using "General" instead of "@" (text) since we may want to use formulas in description
                    .NumberFormat = "General"
                    .Locked = True
                Case "StepName"
                    .Font.Bold = True
                    .HorizontalAlignment = xlLeft
                Case "MissedTime"
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                    .Borders.Weight = xlThick
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                Case "TimeCell"
                    .Font.Color = RGB(0, 0, 0)
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                Case "DateCell"
                    .NumberFormat = GetSetting("DateFormat")
                Case "DateTimeCell"
                    .NumberFormat = GetSetting("DateTimeFormat")
                Case "Description"
                    .WrapText = True
                    .Font.Color = RGB(0, 0, 0)
                    .Interior.Color = xlNone
                    .HorizontalAlignment = xlLeft
                Case "Button"
                    Dim ButtonColor As String
                    ButtonColor = GetSetting("ColorBackButton")
                    .Font.Color = StringToRGB(GetSetting("ColorFontButton"))
                    .Interior.Color = StringToRGB(ButtonColor)
                    .Borders(xlLeft).Color = ToneRGB(ButtonColor, False)
                    .Borders(xlTop).Color = .Borders(xlLeft).Color
                    .Borders(xlRight).Color = ToneRGB(ButtonColor, True)
                    .Borders(xlBottom).Color = .Borders(xlRight).Color
                Case "BooleanTrue"
                    .Font.Color = RGB(0, 75, 0)
                    .Interior.Color = RGB(144, 238, 144)
                Case "BooleanFalse"
                    .Font.Color = RGB(139, 0, 0)
                    .Interior.Color = RGB(255, 204, 203)
                Case "ButtonTrue"
                    .Font.Color = RGB(0, 75, 0)
                    .Interior.Color = RGB(144, 238, 144)
                    .Borders(xlLeft).Color = RGB(196, 246, 196)
                    .Borders(xlTop).Color = .Borders(xlLeft).Color
                    .Borders(xlRight).Color = RGB(77, 127, 77)
                    .Borders(xlBottom).Color = .Borders(xlRight).Color
                    .Locked = False
                Case "ButtonFalse"
                    .Font.Color = RGB(139, 0, 0)
                    .Interior.Color = RGB(255, 204, 203)
                    .Borders(xlLeft).Color = RGB(255, 228, 227)
                    .Borders(xlTop).Color = .Borders(xlLeft).Color
                    .Borders(xlRight).Color = RGB(136, 109, 108)
                    .Borders(xlBottom).Color = .Borders(xlRight).Color
                    .Locked = False
            End Select
            'Some styling, that applies to multiple styles
            Select Case RunSheetStyle
                Case "Delimiter", "StepName", "StepType", "SettingType"
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    .NumberFormat = "@"
                    .Locked = True
                Case "TimeCell", "MissedTime"
                    .HorizontalAlignment = xlLeft
                    .Interior.Color = xlNone
                    .NumberFormat = GetSetting("TimeFormat")
                    .Locked = True
                Case "DateCell", "DateTimeCell", "StatusCell"
                    .HorizontalAlignment = xlCenter
                    .Interior.Color = xlNone
                    .Borders.LineStyle = xlLineStyleNone
                    .Locked = False
                Case "DateTimeCellCompleted", "StatusCellCompleted"
                    .Font.Color = StringToRGB(GetSetting("ColorFontCompleted"))
                    .Interior.Color = StringToRGB(GetSetting("ColorBackCompleted"))
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .Locked = False
                Case "DateTimeCellInProgress", "StatusCellInProgress"
                    .Font.Color = StringToRGB(GetSetting("ColorFontInProgress"))
                    .Interior.Color = StringToRGB(GetSetting("ColorBackInProgress"))
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .Locked = False
                Case "DateTimeCellFailed", "StatusCellFailed"
                    .Font.Color = StringToRGB(GetSetting("ColorFontFailed"))
                    .Interior.Color = StringToRGB(GetSetting("ColorBackFailed"))
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .Locked = False
                Case "DateTimeCellSkipped", "StatusCellSkipped"
                    .Font.Color = StringToRGB(GetSetting("ColorFontSkipped"))
                    .Interior.Color = StringToRGB(GetSetting("ColorBackSkipped"))
                    .Borders.Weight = xlThin
                    .Borders.Color = .Font.Color
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .Locked = False
                Case "BooleanTrue", "BooleanFalse"
                    .Borders.Color = .Font.Color
                    .Borders.Weight = xlThin
                    .Borders.LineStyle = xlContinuous
                    .NumberFormat = "General"
                    .HorizontalAlignment = xlCenter
                Case "Button", "ButtonTrue", "ButtonFalse"
                    .Font.Bold = True
                    .Borders.Weight = xlThick
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlVAlignCenter
                    .Interior.Pattern = xlPatternSolid
            End Select
            Select Case RunSheetStyle
                Case "StatusCell", "StatusCellCompleted", "StatusCellInProgress", "StatusCellFailed", "StatusCellSkipped"
                    .NumberFormat = "@"
                Case "DateTimeCellCompleted", "DateTimeCellInProgress", "DateTimeCellFailed", "DateTimeCellSkipped"
                    .NumberFormat = GetSetting("DateTimeFormat")
            End Select
            .Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
            .Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
        End With
    Next RunSheetStyle
End Sub
Private Sub SetFormats()
    Application.StatusBar = "Styling cells..."
    
    'Set formats to cells
    If EditMode() = True Then
        Call UpdateStyles
        'Static texts
        ReturnRange("CurrentDayTextCell").Value2 = "Current Date:"
        ReturnRange("PreviousDayTextCell").Value2 = "Previous Date:"
        ReturnRange("NextDayTextCell").Value2 = "Next Date:"
        ReturnRange("EndOfDayCell").Value2 = "End Of Day"
        
        'Settings sheet styling
        Call ApplyStyle(ReturnRange("SettingsIDColumnData"), "StepName")
        Call ApplyStyle(ReturnRange("SettingsDescriptionColumnData"), "Description")
        Call ApplyStyle(ReturnRange("SettingsValueColumnData"), "Description")
        Call ApplyStyle(ReturnRange("SettingsTypeColumnData"), "SettingType")
        Call ApplyStyle(ReturnRange("SettingsBooleanCells"), "Boolean")
        'Restore time formating for time values in Settings
        Call ApplyStyle(ReturnRange("SettingsTimeCells"), "TimeCell")
        'Color settings with type "Color"
        Call ColorsFormat(ReturnRange("SettingsColors"))
        
        'Runsheet styling
        Call ApplyStyle(ReturnRange("RunSheetTimeColumnData"), "DateTimeCell")
        Call ApplyStyle(ReturnRange("RunSheetTimeStartColumnData"), "TimeCell")
        Call ApplyStyle(ReturnRange("RunSheetTimeEndColumnData"), "TimeCell")
        Call ApplyStyle(ReturnRange("RunSheetStepNameColumnData"), "StepName")
        Call ApplyStyle(ReturnRange("RunSheetProcessingBlockColumnData"), "StepName")
        Call ApplyStyle(ReturnRange("RunSheetDescriptionColumnData"), "Description")
        'Styping steps based on type
        Call ApplyStyle(ReturnRange("RunSheetTypeColumnData"), "StepType")
        'Styling boolean cells
        Call ApplyStyle(ReturnRange("RunSheetBooleanCells"), "Boolean")
        'Styling delimiters
        If RangeExists("RunSheetDelimiters", "RunSheet") = True Then
            Call ApplyStyle(ReturnRange("RunSheetDelimiters"), "Delimiter")
        End If
        
        Application.StatusBar = "Styling 'New step' button..."
        Call NewStepButton
    End If
    
    'Styling buttons
    If RangeExists("RunSheetButtons", "RunSheet") = True Then
        Call ApplyStyle(ReturnRange("RunSheetButtons"), "Button")
    End If
    'Styling specific UI elements and status cells, since user may modify them
    Call ApplyStyle(ReturnRange("CurrentDayCell"), "DateCell")
    Call ApplyStyle(ReturnRange("PreviousDayCell"), "DateCell")
    Call ApplyStyle(ReturnRange("NextDayCell"), "DateCell")
        
    Call ButtonToggleFormat(ReturnRange("LateSwitchCell"))
    Call ButtonToggleFormat(ReturnRange("EditorSwitchCell"))
    
    'Style status cells along with user and time
    Call ApplyStyle(ReturnRange("RunSheetStatusColumnData"), "Status")
End Sub
'Apply the custom style to range
Public Sub ApplyStyle(ByVal RangeToStyle As Range, ByVal StyleName As String)
    Dim SubRange As Range, TimeCell As Range, UserCell As Range
    
    'Reset to "Normal"
    On Error Resume Next
    Call StyleFallback(RangeToStyle, "Normal")
    On Error GoTo 0
    'Apply style if not Boolean or Status (otherwise apply it during sanitization)
    If StyleName <> "Boolean" And StyleName <> "Status" Then
        Call StyleFallback(RangeToStyle, "RunSheet" & StyleName)
    End If
    
    'Autofit step names
    If StyleName = "StepName" Then
        RangeToStyle.EntireColumn.AutoFit
    End If
    
    'Extra styling of descriptions (this can't be added to style)
    If StyleName = "Description" Then
        RangeToStyle.ColumnWidth = GetSetting("DescriptionWidth")
        RangeToStyle.Hyperlinks.Delete
        RangeToStyle.EntireRow.AutoFit
    End If
    
    If StyleName <> "Boolean" And StyleName <> "Status" Then
        Call Validation(RangeToStyle, StyleName)
    End If
    
    'Using loop for two types, since we need to apply some extra logic in some cases (like sanitization)
    If StyleName = "Boolean" Then
        For Each SubRange In RangeToStyle
            'Applying validation
            Call Validation(SubRange, StyleName)
            'Applying style based on value
            If SubRange.Value2 = True Then
                Call StyleFallback(SubRange, "RunSheet" & StyleName & "True")
            Else
                Call StyleFallback(SubRange, "RunSheet" & StyleName & "False")
            End If
        Next SubRange
    End If
    
    If StyleName = "Status" Then
        For Each SubRange In RangeToStyle
            Set TimeCell = SubRange.Offset(0, -2)
            Set UserCell = SubRange.Offset(0, -1)
            'Applying validation
            Call Validation(SubRange, StyleName)
            'Styling based on value
            If IsEmpty(SubRange.Value2) Or SubRange.Value2 = "" Then
                'Remove any text
                Union(SubRange, TimeCell, UserCell).ClearContents
                'Apply style
                Call StyleFallback(SubRange, "RunSheetStatusCell")
                Call StyleFallback(UserCell, "RunSheetStatusCell")
                Call StyleFallback(TimeCell, "RunSheetDateTimeCell")
            Else
                Select Case SubRange.Value2
                    Case "Completed"
                        Call StyleFallback(SubRange, "RunSheetStatusCellCompleted")
                        Call StyleFallback(UserCell, "RunSheetStatusCellCompleted")
                        Call StyleFallback(TimeCell, "RunSheetDateTimeCellCompleted")
                    Case "In Progress"
                        Call StyleFallback(SubRange, "RunSheetStatusCellInProgress")
                        Call StyleFallback(UserCell, "RunSheetStatusCellInProgress")
                        Call StyleFallback(TimeCell, "RunSheetDateTimeCellInProgress")
                    Case "Failed"
                        Call StyleFallback(SubRange, "RunSheetStatusCellFailed")
                        Call StyleFallback(UserCell, "RunSheetStatusCellFailed")
                        Call StyleFallback(TimeCell, "RunSheetDateTimeCellFailed")
                    Case "Skipped"
                        Call StyleFallback(SubRange, "RunSheetStatusCellSkipped")
                        Call StyleFallback(UserCell, "RunSheetStatusCellSkipped")
                        Call StyleFallback(TimeCell, "RunSheetDateTimeCellSkipped")
                End Select
            End If
        Next SubRange
    End If
End Sub
Private Function Validation(ByVal RangeToStyle As Range, ByVal StyleName As String) As Boolean
    'Doing this only in case of Editor Mode
    If EditMode() = True Then
        On Error GoTo failedvalidation
        With RangeToStyle.Validation
            'Delete any validation logic, if it was applied to the range
            .Delete
            'Apply the validation logic
            Select Case StyleName
                Case "StepName"
                    .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:=GetSetting("StepMaxLength")
                    .IgnoreBlank = False
                    .ShowError = True
                    .InCellDropdown = False
                    .ShowInput = False
                Case "StepType"
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Regular,Button,Delimiter", Formula2:=""
                    .IgnoreBlank = False
                    .ShowError = True
                    .InCellDropdown = True
                    .ShowInput = False
                Case "SettingType"
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Array,Boolean,Color,Integer,Range,String,Time", Formula2:=""
                    .IgnoreBlank = False
                    .ShowError = True
                    .InCellDropdown = True
                    .ShowInput = False
                Case "Boolean"
                    'Sanitizing value, in case it was changed before validation got applied to cell
                    RangeToStyle.Value2 = ToBoolean(RangeToStyle.Value2)
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=LocalisedBoolean(True) & "," & LocalisedBoolean(False), Formula2:=""
                    .IgnoreBlank = True
                    .ShowError = False
                    .InCellDropdown = True
                    .ShowInput = False
                Case "Status"
                    'Sanitize value
                    If IsInArray(LCase(RangeToStyle.Value2), Array("completed", "in progress", "failed", "skipped")) = False Then
                        RangeToStyle.ClearContents
                    End If
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Completed,In Progress,Failed,Skipped", Formula2:=""
                    .IgnoreBlank = True
                    .ShowError = False
                    .InCellDropdown = True
                    .ShowInput = False
                'Not actually a style, but still
                Case "NewStepMenu"
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Regular,Button,Delimiter", Formula2:=""
                    .IgnoreBlank = False
                    .ShowError = True
                    .InCellDropdown = True
                    .ShowInput = False
            End Select
        End With
        On Error GoTo 0
        Validation = True
    Else
        Validation = True
    End If
    Exit Function
failedvalidation:
    MsgBox "Failed to apply validation to " & RangeToStyle.Address & " on " & RangeToStyle.Worksheet.Name & "!" & "Common cause is issues with OneDrive or SharePoint connectivity." & vbCrLf & "All macros have been cancelled!" & vbCrLf & "Please, close workbook without saving and open it again!", vbOKOnly + vbApplicationModal + vbCritical, "Failed to apply validation"
    On Error GoTo 0
    Validation = False
    End
End Function
Private Sub StyleFallback(ByVal CellRange As Range, StyleName As String)
    'Use common function to apply the style
    'Tracking errors to allow fallback in case of failure
    On Error Resume Next
    CellRange.Style = StyleName
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    'In case we failed to apply style
    With CellRange
        'Have to skipp errors, because a style may be missing certain type of setting in it, which is by design
        On Error Resume Next
        'Content placement
        .IndentLevel = ThisWorkbook.Styles(StyleName).IndentLevel
        .Orientation = ThisWorkbook.Styles(StyleName).Orientation
        .ReadingOrder = ThisWorkbook.Styles(StyleName).ReadingOrder
        .ShrinkToFit = ThisWorkbook.Styles(StyleName).ShrinkToFit
        .HorizontalAlignment = ThisWorkbook.Styles(StyleName).HorizontalAlignment
        .VerticalAlignment = ThisWorkbook.Styles(StyleName).VerticalAlignment
        .WrapText = ThisWorkbook.Styles(StyleName).WrapText
        'Font
        With .Font
            .Bold = ThisWorkbook.Styles(StyleName).Font.Bold
            .Color = ThisWorkbook.Styles(StyleName).Font.Color
            .ColorIndex = ThisWorkbook.Styles(StyleName).Font.ColorIndex
            .FontStyle = ThisWorkbook.Styles(StyleName).Font.FontStyle
            .Italic = ThisWorkbook.Styles(StyleName).Font.Italic
            .Size = ThisWorkbook.Styles(StyleName).Font.Size
            .Strikethrough = ThisWorkbook.Styles(StyleName).Font.Strikethrough
            .Subscript = ThisWorkbook.Styles(StyleName).Font.Subscript
            .Superscript = ThisWorkbook.Styles(StyleName).Font.Superscript
            .ThemeColor = ThisWorkbook.Styles(StyleName).Font.ThemeColor
            .ThemeFont = ThisWorkbook.Styles(StyleName).Font.ThemeFont
            .TintAndShade = ThisWorkbook.Styles(StyleName).Font.TintAndShade
            .Underline = ThisWorkbook.Styles(StyleName).Font.Underline
        End With
        'Background style
        .Interior.Color = ThisWorkbook.Styles(StyleName).Interior.Color
        .Interior.Pattern = ThisWorkbook.Styles(StyleName).Interior.Pattern
        'Doing Borders without specification may seem to result in wrong application of the style in some cases, thus doing it for each border separately
        'Also doing them in a loop with some checks for a small optimization
        Dim ArrayElement As Variant
        For Each ArrayElement In Array(xlLeft, xlTop, xlRight, xlBottom, xlDiagonalDown, xlDiagonalUp, xlInsideHorizontal, xlInsideVertical)
            'Line style first
            .Borders(ArrayElement).LineStyle = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).LineStyle
            'Check if it's not "no border", because if it is and we apply color or weight, it will change to Thin
            If .Borders(ArrayElement).LineStyle <> xlLineStyleNone Then
                'Color
                .Borders(ArrayElement).Color = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).Color
                'Weight
                .Borders(ArrayElement).Weight = ThisWorkbook.Styles(StyleName).Borders(ArrayElement).Weight
            End If
        Next ArrayElement
        'Formatting
        .NumberFormat = ThisWorkbook.Styles(StyleName).NumberFormat
        .NumberFormatLocal = ThisWorkbook.Styles(StyleName).NumberFormatLocal
        .FormulaHidden = ThisWorkbook.Styles(StyleName).FormulaHidden
        'Protection
        .Locked = ThisWorkbook.Styles(StyleName).Locked
        On Error GoTo 0
    End With
End Sub
'Function to check of style exists based on https://stackoverflow.com/questions/17209989/excel-2010-vba-find-out-if-formatting-style-exists
Private Function StyleExists(ByVal StyleName As String) As Boolean
    On Error Resume Next
    StyleExists = Len(ThisWorkbook.Styles(StyleName).Name) > 0
    On Error GoTo 0
End Function
Private Sub ButtonToggleFormat(ByVal Button As Range)
    Dim truetoggle As Boolean, LateCellsExist As Boolean, LateCell As Boolean, EditCell As Boolean, EditModeValue As Boolean
    truetoggle = False
    LateCellsExist = LateExists()
    EditModeValue = EditMode()
    If Not Intersect(Button, ReturnRange("LateSwitchCell")) Is Nothing Then
        LateCell = True
    Else
        LateCell = False
    End If
    If Not Intersect(Button, ReturnRange("EditorSwitchCell")) Is Nothing Then
        EditCell = True
    Else
        EditCell = False
    End If
    Dim BackColor As String
    If LateCell = True Then
        truetoggle = LateMode()
        'Disable late mode if we do not have late special steps
        If LateCellsExist = False And truetoggle = True Then
            If EditModeValue = False Then
                Call SetSetting("LateFlag", False)
                truetoggle = False
            End If
        End If
    ElseIf EditCell = True Then
        truetoggle = EditModeValue
    End If
    With Button
        'Button should be 'red' if it's enabled, to indicate caution, since it's not expected to be used on daily basis
        If truetoggle = True Then
            Call ApplyStyle(Button, "ButtonFalse")
        Else
            Call ApplyStyle(Button, "ButtonTrue")
        End If
        If EditCell = True And IsEditor() = False Then
            .Interior.Pattern = xlPatternGray50
        End If
        If LateCell = True And LateCellsExist = False Then
            .Interior.Pattern = xlPatternGray50
        End If
    End With
End Sub
Private Sub ColorsFormat(ByVal ColorRange As Range)
    Dim SubCell As Range
    For Each SubCell In ColorRange
        If IsEmpty(SubCell.Value2) = False And SubCell.Value2 <> "" Then
            SubCell.Interior.Color = StringToRGB(SubCell.Value2)
            SubCell.Font.Color = SubCell.Interior.Color
        Else
            SubCell.Value2 = RangeColorToString(SubCell.Interior.Color)
        End If
    Next SubCell
End Sub
Private Sub StepStyle(ByVal StepRange As Range, Optional ByVal Reset As Boolean = False)
    If Reset = True Then
        'Reset flags
        StepRange.Offset(0, -6).Value2 = False
        StepRange.Offset(0, -5).Value2 = True
        StepRange.Offset(0, -4).Value2 = True
        StepRange.Offset(0, -3).Value2 = True
        StepRange.Offset(0, -2).Value2 = True
        StepRange.Offset(0, -1).Value2 = True
        Call ApplyStyle(StepRange.Offset(0, -6), "Boolean")
        Call ApplyStyle(StepRange.Offset(0, -5), "Boolean")
        Call ApplyStyle(StepRange.Offset(0, -4), "Boolean")
        Call ApplyStyle(StepRange.Offset(0, -3), "Boolean")
        Call ApplyStyle(StepRange.Offset(0, -2), "Boolean")
        Call ApplyStyle(StepRange.Offset(0, -1), "Boolean")
        'Clear contents
        StepRange.Offset(0, 6).ClearContents
        StepRange.Offset(0, 7).ClearContents
        StepRange.Offset(0, 8).ClearContents
    End If
    If LCase(StepRange.Value2) = "delimiter" Then
        'Clear values
        StepRange.Offset(0, 1).ClearContents
        StepRange.Offset(0, 2).Value2 = "Delimiter"
        StepRange.Offset(0, 3).ClearContents
        StepRange.Offset(0, 4).ClearContents
        StepRange.Offset(0, 5).ClearContents
        'Style Cells
        Call ApplyStyle(StepRange.Offset(0, 1), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 2), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 3), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 4), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 5), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 6), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 7), "Delimiter")
        Call ApplyStyle(StepRange.Offset(0, 8), "Delimiter")
    Else
        'Update styles
        Call ApplyStyle(StepRange.Offset(0, 1), "StepName")
        If LCase(StepRange.Value2) = "regular" Then
            Call ApplyStyle(StepRange.Offset(0, 2), "StepName")
        Else
            Call ApplyStyle(StepRange.Offset(0, 2), "Button")
        End If
        Call ApplyStyle(StepRange.Offset(0, 3), "Description")
        'Clear values and style cells
        If IsTime(StepRange.Offset(0, 4).Text) = False Then
            StepRange.Offset(0, 4).Value2 = "00:00"
        End If
        Call ApplyStyle(StepRange.Offset(0, 4), "TimeCell")
        If IsTime(StepRange.Offset(0, 5).Text) = False Then
            StepRange.Offset(0, 5).Value2 = "00:00"
        End If
        Call ApplyStyle(StepRange.Offset(0, 5), "TimeCell")
        Call ApplyStyle(StepRange.Offset(0, 8), "Status")
    End If
End Sub
Private Sub NewStepButton()
    Dim Menu As Range
    Set Menu = ReturnRange("NewStepMenu")
    If EditMode() = True Then
        Call ApplyStyle(Menu, "Button")
        Menu.Value2 = "Add Step"
        Call Validation(Menu, "NewStepMenu")
    Else
        'This may fail in some wierd conditions, but stince this is a button meant for editors only - ignore the error
        On Error Resume Next
        Menu.Validation.Delete
        On Error GoTo 0
        Menu.ClearContents
        Menu.ClearFormats
    End If
End Sub
