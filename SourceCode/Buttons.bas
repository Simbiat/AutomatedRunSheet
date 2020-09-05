Attribute VB_Name = "Buttons"
Option Explicit
Private Sub CustomButtonClick(ByVal Button As Range)
    Dim ButtonID As String, FunctionToRun As String, ArgumentsString As String, Arguments() As String, FunctionResult As Boolean
    Dim ErrorMsg As String, CurrentAlerts As Boolean
    Call WriteLog("Clicked '" & Button.Value2 & "' button")
    ButtonID = AlphaNumeric(Button.Value2)
    FunctionToRun = GetSetting(ButtonID & "Function")
    ArgumentsString = GetSetting(ButtonID & "Arguments")
    If LCase(FunctionToRun) <> LCase(ButtonID & "Function" & "=N/A") Then
        If LCase(ArgumentsString) <> LCase(ButtonID & "Arguments" & "=N/A") Then
            'Process arguments
            Arguments = TrimArray(Split(ArgumentsString, GetSetting("ArrayDelimiter"), -1, 1))
            'Call custom functions
            On Error Resume Next
            Select Case LCase(FunctionToRun)
                Case "attachscreenshot"
                    'This one needs alerts disabled, so ensuring that
                    CurrentAlerts = Application.DisplayAlerts
                    Application.DisplayAlerts = False
                    FunctionResult = Application.Run(FunctionToRun, Button.Offset(0, 6))
                    Application.DisplayAlerts = CurrentAlerts
                Case "vnc"
                    'Redim array, to ensure we have 2 element (1 reserved and 1 for actual function)
                    ReDim Preserve Arguments(1)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("No path to .vnc provided!")
                        Exit Sub
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1))
                Case "findinfile"
                    'Redim array, to ensure we have 6 element (1 reserved and 5 for actual function)
                    ReDim Preserve Arguments(5)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Path to haystack not provided!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Call ButtonError("Needle not provided!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "0"
                    End If
                    If IsEmpty(Arguments(4)) = True Or Arguments(4) = "" Then
                        Arguments(4) = "UTF-8"
                    End If
                    If IsEmpty(Arguments(5)) = True Or Arguments(5) = "" Then
                        Arguments(5) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), Arguments(2), ToBoolean(Arguments(3)), Arguments(4), CInt(Arguments(5)))
                Case "toclipboard"
                    'Redim array, to ensure we have 2 element (1 reserved and 1 for actual function)
                    ReDim Preserve Arguments(1)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Clipboard text is not set!")
                        Exit Sub
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1))
                Case "listfiles"
                    'Redim array, to ensure we have 4 elements (1 reserved and 3 for actual function). 1 argument is static for buttons, so we sent static value
                    ReDim Preserve Arguments(3)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Mask to check is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = "0"
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), ToBoolean(Arguments(2)), True, CInt(Arguments(3)))
                Case "checkdir"
                    'Redim array, to ensure we have 4 elements (1 reserved and 3 for actual function)
                    ReDim Preserve Arguments(3)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Directory to check is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = "0"
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), ToBoolean(Arguments(2)), CInt(Arguments(3)))
                Case "servicecheck"
                    'Redim array, to ensure we have 6 elements (1 reserved and 5 for actual function)
                    ReDim Preserve Arguments(5)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Service name is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = ""
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "0"
                    End If
                    If IsEmpty(Arguments(4)) = True Or Arguments(4) = "" Then
                        Arguments(4) = "0"
                    End If
                    If IsEmpty(Arguments(5)) = True Or Arguments(5) = "" Then
                        Arguments(5) = "10"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), Arguments(2), ToBoolean(Arguments(3)), ToBoolean(Arguments(4)), CInt(Arguments(5)))
                Case "rdp"
                    'Redim array, to ensure we have 6 elements (1 reserved and 5 for actual function)
                    ReDim Preserve Arguments(5)
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Host name is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = "100"
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "1"
                    End If
                    If IsEmpty(Arguments(4)) = True Or Arguments(4) = "" Then
                        Arguments(4) = "0"
                    End If
                    If IsEmpty(Arguments(5)) = True Or Arguments(5) = "" Then
                        Arguments(5) = "0"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), CInt(Arguments(2)), ToBoolean(Arguments(3)), ToBoolean(Arguments(4)), ToBoolean(Arguments(5)))
                Case "openfile"
                    'Redim array, to ensure we have 3 elements (1 reserved and 2 for actual function)
                    ReDim Preserve Arguments(2)
                    'Sanitize values, as precaution and also exit, if mandatory arguments are not set
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("File to open is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), CInt(Arguments(2)))
                Case "checkfilelist"
                    'Redim array, to ensure we have 4 elements (1 reserved and 3 for actual function)
                    ReDim Preserve Arguments(3)
                    'Sanitize values, as precaution and also exit, if mandatory arguments are not set
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("File list is not specified!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Arguments(2) = "0"
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Arguments(3) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), CLng(Arguments(2)), CInt(Arguments(3)))
                Case "sendmail"
                    'Redim array, to ensure we have 9 elements (1 reserved and 8 for actual function)
                    ReDim Preserve Arguments(8)
                    'Sanitize values, as precaution and also exit, if mandatory arguments are not set
                    If IsEmpty(Arguments(1)) = True Or Arguments(1) = "" Then
                        Call ButtonError("Subject for SendMail is not definied!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(2)) = True Or Arguments(2) = "" Then
                        Call ButtonError("Text for SendMail is not definied!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(3)) = True Or Arguments(3) = "" Then
                        Call ButtonError("Recipient for SendMail is not definied!")
                        Exit Sub
                    End If
                    If IsEmpty(Arguments(4)) = True Or Arguments(4) = "" Then
                        Arguments(4) = ""
                    End If
                    If IsEmpty(Arguments(5)) = True Or Arguments(5) = "" Then
                        Arguments(5) = ""
                    End If
                    If IsEmpty(Arguments(6)) = True Or Arguments(6) = "" Then
                        Arguments(6) = ""
                    End If
                    If IsEmpty(Arguments(7)) = True Or Arguments(7) = "" Then
                        Arguments(7) = "0"
                    End If
                    If IsEmpty(Arguments(8)) = True Or Arguments(8) = "" Then
                        Arguments(8) = "1"
                    End If
                    FunctionResult = Application.Run(FunctionToRun, Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), CInt(Arguments(7)), CInt(Arguments(7)))
                Case Else
                    'When calling the function, we remove first argument, which was reserved, since it should not be required by the actual function
                    'We send the whole array, though, so you will need to handle parsing it in your function
                    FunctionResult = Application.Run(FunctionToRun, RemoveFirstElement(Arguments))
            End Select
            If Err.Number = 1004 Then
                Call ButtonError("Function '" & FunctionToRun & "' assigned to '" & Button.Text & "' button does not exist!")
                Exit Sub
            ElseIf Err.Number <> 0 Then
                Call ButtonError("Function '" & FunctionToRun & "' assigned to '" & Button.Text & "' failed with error #" & Err.Number & "!")
                Exit Sub
            End If
            On Error GoTo 0
            'Logic for cell marking
            'Attaching screenshot does not really need arguments, so covering it separately
            If LCase(FunctionToRun) = "attachscreenshot" Then
                If ToBoolean(FunctionResult) = False Then
                    Application.Run "MarkCell", Button.Offset(0, 6), "Failed"
                Else
                    Application.Run "MarkCell", Button.Offset(0, 6), "Completed"
                End If
            Else
                If Arguments(0) <> "" Then
                    'Using Case to make it somewhat more readable and avoid action in case of unsupported value
                    Select Case LCase(Arguments(0))
                        Case "in progress", "inprogress"
                            Application.Run "MarkCell", Button.Offset(0, 6), "In Progress"
                        Case "completed", "failed", "skipped"
                            Application.Run "MarkCell", Button.Offset(0, 6), Arguments(0)
                        Case "autocompleted", "autoinprogress", "autoskipped"
                            If ToBoolean(FunctionResult) = False Then
                                'Mark as failed
                                Application.Run "MarkCell", Button.Offset(0, 6), "Failed"
                            Else
                                Select Case LCase(Arguments(0))
                                    Case "autocompleted"
                                        Application.Run "MarkCell", Button.Offset(0, 6), "Completed"
                                    Case "autoinprogress"
                                        Application.Run "MarkCell", Button.Offset(0, 6), "In Progress"
                                    Case "autoskipped"
                                        Application.Run "MarkCell", Button.Offset(0, 6), "Skipped"
                                End Select
                            End If
                    End Select
                End If
            End If
        Else
            Call ButtonError("No arguments are assigned to '" & Button.Text & "' button!")
        End If
    Else
        Call ButtonError("No function is assigned to '" & Button.Text & "' button!")
    End If
End Sub
Private Sub ButtonError(ByVal ErrorMsg As String)
    Call WriteLog(ErrorMsg)
    MsgBox ErrorMsg, vbOKOnly + vbExclamation + vbApplicationModal, "Dummy button"
End Sub
'Borders are simply switched, but since we are calling this from outside, need to pass colors for background and font
Private Sub AnimateButtonClick(ByVal Button As Range, ByVal Font As Long, ByVal Back As Long)
    With Button
        'We need to check if above cell is also a button, in order to retain consistency of colors with it
        If .Borders(xlEdgeTop).Color = .Offset(-1, 0).Borders(xlEdgeRight).Color Then
            .Borders(xlEdgeRight).Color = .Borders(xlEdgeLeft).Color
            .Borders(xlEdgeLeft).Color = .Borders(xlEdgeBottom).Color
            .Borders(xlEdgeBottom).Color = .Borders(xlEdgeRight).Color
        Else
            .Borders(xlEdgeLeft).Color = .Borders(xlEdgeRight).Color
            .Borders(xlEdgeRight).Color = .Borders(xlEdgeTop).Color
            .Borders(xlEdgeTop).Color = .Borders(xlEdgeLeft).Color
            .Borders(xlEdgeBottom).Color = .Borders(xlEdgeRight).Color
        End If

        If .Interior.Pattern = xlPatternSolid Then
            .Interior.Pattern = xlPatternLightUp
            .Interior.PatternColor = .Borders(xlEdgeRight).Color
        Else
            .Interior.Pattern = xlPatternSolid
        End If
        .Font.Color = Font
        .Interior.Color = Back
    End With
End Sub
Private Sub LateButton(ByVal FontColor As Long, ByVal BackColor As Long)
    Dim Button As Range, LateCellsExist As Boolean, EditModeValue As Boolean
    Set Button = ReturnRange("LateSwitchCell")
    LateCellsExist = LateExists()
    EditModeValue = EditMode()
    'Toggle value
    If LateCellsExist = False Or EditModeValue = True Then
        'If no late cells or we are in Edit Mode, force True to ensure nothing is hidden
        Call SetSetting("LateFlag", True)
    Else
        Call SetSetting("LateFlag", Toggle(LateMode()))
    End If
    'Change color for button
    If LateCellsExist = False Or EditModeValue = False Then
        Application.Run "ButtonToggleFormat", Button
        FontColor = Button.Font.Color
        BackColor = Button.Interior.Color
        Application.Run "AnimateButtonClick", Button, ToneRGB(RangeColorToString(FontColor), True, 1), ToneRGB(RangeColorToString(BackColor), True, 1)
    End If
    Application.Run "LateRows"
    If LateCellsExist = True Or EditModeValue = True Then
        'Since we have already updated color - update color and recreate click animation
        Button.Interior.Pattern = xlPatternLightUp
        Application.Run "AnimateButtonClick", Button, FontColor, BackColor
        'Button.Interior.Pattern = xlPatternSolid
        Call WriteLog("Switched late mode to " & LateMode())
        Application.StatusBar = "Switched late mode to " & LateMode()
    Else
        Call WriteLog("Switched late mode to " & LateMode())
        Application.StatusBar = "Switched late mode to " & LateMode()
    End If
End Sub
Private Sub EditModeButton(ByVal FontColor As Long, ByVal BackColor As Long)
    Dim Button As Range
    Set Button = ReturnRange("EditorSwitchCell")
    If Toggle(EditMode()) = True Then
        Application.Run "ProtectMe", False
        Call SetSetting("EditMode", Toggle(EditMode()))
        Call SetSetting("LateFlag", True)
        Application.Run "Init"
    Else
        Call SetSetting("EditMode", Toggle(EditMode()))
        Call SetSetting("LateFlag", False)
        Application.Run "Init"
        Application.Run "ProtectMe", True
    End If
    Application.Run "ButtonToggleFormat", Button
    FontColor = Button.Font.Color
    BackColor = Button.Interior.Color
    Application.Run "AnimateButtonClick", Button, ToneRGB(RangeColorToString(FontColor), True, 1), ToneRGB(RangeColorToString(BackColor), True, 1)
    'Activate RunSheet, because aftet unhiding EditorManual it gets activated for some reason
    ThisWorkbook.Worksheets("RunSheet").Activate
    'Since we have already updated color - update pattern and recreate click animation
    Button.Interior.Pattern = xlPatternLightUp
    Application.Run "AnimateButtonClick", Button, FontColor, BackColor
    ReturnRange("RunSheetStatusColumnData").Cells(1, 1).Select
    Call WriteLog("Switched editor mode to " & EditMode())
    Application.StatusBar = "Switched editor mode to " & EditMode()
End Sub
Private Sub EndOfDayButton()
    Call WriteLog("Initiated End Of Day")
    Application.Run "EndOfDay"
    Application.StatusBar = "End Of Day Completed"
End Sub
Public Sub ButtonUpdate(ByVal ButtonName As String, Optional ByVal Remove As Boolean = False)
    Dim ButtonID As String
    ButtonID = AlphaNumeric(ButtonName)
    'Process only if there no duplicates. Mostly important for removal, but may improve performance for adding as well
    'Comparing against 2, since if removing we still have a button, and if adding it we already have a button, thus we will always have 1
    If ButtonCount(ButtonID) < 2 Then
        If Remove = True Then
            'Remove settings
            Application.Run "RemoveSetting", ButtonID & "Function"
            Application.Run "RemoveSetting", ButtonID & "Arguments"
        Else
            'Add settings
            Application.Run "AddSetting", ButtonID & "Function", "String"
            Application.Run "AddSetting", ButtonID & "Arguments", "Array"
        End If
    Else
        If Remove = False Then
            'If we are adding a button and there is already a button with same name, warn that we have buttons with same name already
            MsgBox "Button '" & ButtonName & "' already exists!" & vbCrLf & "Settings from previously registered button will be used.", vbOKOnly + vbInformation + vbApplicationModal, "Button already exists!"
        End If
    End If
    'Update and style ranges
    Application.Run "Init"
End Sub
'Count occurances of same button
Private Function ButtonCount(ByVal ButtonID As String) As Integer
    Dim Button As Range
    'We need to count and rely on "Remove" flag, because if we are not removing we expect to find, at least, 1 button, and we are - no
    ButtonCount = 0
    For Each Button In ReturnRange("RunSheetButtons")
        If LCase(AlphaNumeric(Button.Value2)) = LCase(ButtonID) Then
            ButtonCount = ButtonCount + 1
        End If
    Next Button
End Function
Private Function ButtonSettingsRenaming(ByVal OldValue As String, ByVal NewValue As String) As Boolean
    Dim OldButtonID As String, NewButtonID As String
    Dim OldButtonCount As Integer, NewButtonCount As Integer
    Dim DuplicateAlert As Integer
    'Setting values for slight optimization
    OldButtonID = AlphaNumeric(OldValue)
    NewButtonID = AlphaNumeric(NewValue)
    OldButtonCount = ButtonCount(OldButtonID)
    NewButtonCount = ButtonCount(NewButtonID)
    If NewButtonCount < 2 Then
        'NewButtonCount < 2 means that we only have 1 button with this new ID - the one we just created. Using <2 instead of =1 to cover unlikely situation we get 0 or negative values
        If OldButtonCount < 1 Then
            'No buttons with old ID, thus simply renaming the settings
            Application.Run "RenameSetting", OldButtonID & "Function", NewButtonID & "Function"
            Application.Run "RenameSetting", OldButtonID & "Arguments", NewButtonID & "Arguments"
            ButtonSettingsRenaming = True
        Else
            'Means we need to save original values for buttons with old IDs, but we also need new ones, so we create new settings,...
            Application.Run "AddSetting", NewButtonID & "Function", "String", GetSettingEditable(OldButtonID & "Function")
            Application.Run "AddSetting", NewButtonID & "Arguments", "Array", GetSettingEditable(OldButtonID & "Arguments")
            '...update ranges to be able to locate them,...
            Call UpdateRanges
            '...and then set their values to those of the old settings
            Call SetSetting(NewButtonID & "Function", GetSetting(OldButtonID & "Function"))
            Call SetSetting(NewButtonID & "Arguments", GetSetting(OldButtonID & "Arguments"))
            ButtonSettingsRenaming = True
        End If
    Else
        'Let user choose whether to copy old settings to new ones or not. Allow cancelling renaming as well
        DuplicateAlert = MsgBox("There are already " & NewButtonCount & " '" & NewValue & "' buttons! Do you want to update their settings?" & vbCrLf & _
                                "Press 'Yes' to copy the settings from '" & OldValue & "' button to '" & NewValue & "'." & vbCrLf & _
                                "Press 'No' to continue using settings from '" & NewValue & "' button." & vbCrLf & "Press 'Cancel' to cancel renaming of the button." _
                                , vbYesNoCancel + vbExclamation + vbApplicationModal, "Duplicate buttons found!")
        If DuplicateAlert = vbCancel Then
            'Cancel renaming
            ButtonSettingsRenaming = False
            Exit Function
        ElseIf DuplicateAlert = vbYes Then
            'Update settings for buttons
            Call SetSetting(NewButtonID & "Function", GetSetting(OldButtonID & "Function"))
            Call SetSetting(NewButtonID & "Arguments", GetSetting(OldButtonID & "Arguments"))
        End If
        If OldButtonCount < 1 Then
            'If no duplicates remain for old ID - remove its settings
            Application.Run "RemoveSetting", OldButtonID & "Function"
            Application.Run "RemoveSetting", OldButtonID & "Arguments"
        End If
        ButtonSettingsRenaming = True
    End If
    'need to apply the same logic when changing type of the button?
End Function
