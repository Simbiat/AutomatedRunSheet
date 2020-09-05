Attribute VB_Name = "Tracking"
Option Explicit
Private Sub MarkCell(ByVal Target As Range, Optional ByVal Status = "", Optional ByVal SelectAfter As Range = Nothing)
    Dim FontColor As Long, BackColor As Long
    Dim CellText As String, TimeStamp As String, Username As String, Computer As String, LogLine As String, BlockText As String
    Dim SubTarget As Range, TimeCell As Range, UserCell As Range, StepCell As Range
    Dim SetTime As Boolean: SetTime = False
    Dim SymbolsComplete As String, SymbolsInProgress As String, SymbolsFailed As String, SymbolsSkipped As String, SymbolsClear As String
    Dim DateFormat As String
    
    'Get settings to optimize in case of a loop
    DateFormat = GetSetting("DateTimeFormat")
    SymbolsComplete = GetSetting("MarkSymbolsCompleted")
    SymbolsInProgress = GetSetting("MarkSymbolsInProgress")
    SymbolsFailed = GetSetting("MarkSymbolsFailed")
    SymbolsSkipped = GetSetting("MarkSymbolsSkipped")
    SymbolsClear = GetSetting("MarkSymbolsClear")
    
    'Using Environ$ instead of Environ to ensure we get a string
    Username = FullName()
    Computer = VBA.Environ$("COMPUTERNAME")
    TimeStamp = Format(Now(), GetSetting("DateTimeFormat"))
    
    For Each SubTarget In Target
        'Process only if the target's row is visible
        If SubTarget.EntireRow.Hidden = False Then
            Set TimeCell = SubTarget.Offset(0, -2)
            Set UserCell = SubTarget.Offset(0, -1)
            Set StepCell = SubTarget.Offset(0, -6)
            BlockText = " in '" & ReturnRange("RunSheetProcessingBlockColumnData").Cells(SubTarget.Row - 4, 1).Value2 & "' block"
            If BlockText = " in '' block" Then
                BlockText = ""
            End If
            'Update cell text
            'Using vbTextCompare to do case-insensetive comparisson
            If Status <> "" Then
                Select Case LCase(Status)
                    Case "purge"
                        SubTarget.Value2 = ""
                    Case "completed"
                        SubTarget.Value2 = "Completed"
                        SetTime = True
                    Case "failed"
                        SubTarget.Value2 = "Failed"
                        SetTime = True
                    Case "skipped"
                        SubTarget.Value2 = "Skipped"
                        SetTime = True
                    Case "in progress", "inprogress"
                        SubTarget.Value2 = "In Progress"
                        SetTime = True
                End Select
            Else
                If SubTarget.Value2 = "Completed" Or SymbolCheck(SubTarget.Value2, SymbolsComplete) = True Then
                    SubTarget.Value2 = "Completed"
                    SetTime = True
                ElseIf SubTarget.Value2 = "Failed" Or SymbolCheck(SubTarget.Value2, SymbolsFailed) = True Then
                    SubTarget.Value2 = "Failed"
                    SetTime = True
                ElseIf SubTarget.Value2 = "Skipped" Or SymbolCheck(SubTarget.Value2, SymbolsSkipped) = True Then
                    SubTarget.Value2 = "Skipped"
                    SetTime = True
                ElseIf SubTarget.Value2 = "In Progress" Or SymbolCheck(SubTarget.Value2, SymbolsInProgress) = True Then
                    SubTarget.Value2 = "In Progress"
                    SetTime = True
                ElseIf SymbolCheck(SubTarget.Value2, SymbolsClear) = True Then
                    SubTarget.Value2 = ""
                End If
            End If
            'Apply style
            Call ApplyStyle(SubTarget, "Status")
            'Apply the colors and log details if we are not clearing the cell
            If SetTime = True Then
                TimeCell.Value2 = Format(Now(), DateFormat)
                UserCell.Value2 = Username
                LogLine = "Step '" & StepCell.Value2 & "'" & BlockText & " (" & Target.Address & ") marked as '" & SubTarget.Value2 & "'"
            Else
                LogLine = "Step '" & StepCell.Value2 & "'" & BlockText & " (" & Target.Address & ") cleared"
            End If
            'Update time frame cells if they are used for the step
            Application.Run "TimeFrameCheck", Target
        End If
    Next SubTarget
    
    'Select next visible cells
    If SelectAfter Is Nothing Then
        Application.Run "NextVisible", Target
    Else
        SelectAfter.Activate
    End If
    
    Call WriteLog(LogLine, TimeStamp, Computer, Username)
End Sub
'Using below function to check fo symbol/key instead of originally planned InStr, because it returns 1 for empty strings (which are possible for automations)
Private Function SymbolCheck(ByVal Symbol As String, ByVal SymbolsList As String) As Boolean
    If IsEmpty(Symbol) = True Or Symbol = "" Then
        'If string is empty - do not do search
        SymbolCheck = False
    Else
        'If not do array search
        SymbolCheck = IsInArray(Symbol, Split(SymbolsList, GetSetting("ArrayDelimiter"), -1, 1))
    End If
End Function
Private Sub TrackChange(ByVal Target As Range, ByRef OldValue As String, ByRef OldText As String)
    Dim LogText As String
    Dim BoolFlag As Boolean
    Dim StepName As String
    If Target.Worksheet.Name = "RunSheet" Then
        StepName = ThisWorkbook.Worksheets("RunSheet").Cells(Target.Row, 9).Value2
    Else
        StepName = ThisWorkbook.Worksheets("Settings").Cells(Target.Row, 1).Value2
    End If
    LogText = "Changed value of " & Target.Address & " on '" & Target.Worksheet.Name & "''"
    BoolFlag = False
    If Target.Worksheet.Name = "RunSheet" Then
        If Not Intersect(Target, ReturnRange("RunSheetStatusColumnData")) Is Nothing Then
            Call MarkCell(Target)
            'Exiting the sub, because MarkCell did all that we needed
            Exit Sub
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsLateSpecialColumnData")) Is Nothing Then
            LogText = "Switched IsLateSpecial for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsFirstSpecialColumnData")) Is Nothing Then
            LogText = "Switched IsFirstSpecial for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsRegularSpecialColumnData")) Is Nothing Then
            LogText = "Switched IsRegularSpecial for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsLastSpecialColumnData")) Is Nothing Then
            LogText = "Switched IsLastSpecial for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsFirstSpecialWorkColumnData")) Is Nothing Then
            LogText = "Switched IsFirstSpecialWork for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetIsLastSpecialWorkColumnData")) Is Nothing Then
            LogText = "Switched IsLastSpecialWork for step '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("RunSheetTypeColumnData")) Is Nothing Then
            Application.Run "StepStyle", Target, True
            If LCase(Target.Value2) = "button" Then
                'Add button specific settings
                Application.Run "ButtonUpdate", StepName, False
            ElseIf LCase(OldValue) = "button" Then
                'Remove button specific settings
                Application.Run "ButtonUpdate", StepName, True
            End If
            Application.Run "UpdateRanges"
            LogText = "Switched Type for step '" & StepName & "'"
        ElseIf Not Intersect(Target, ReturnRange("NewStepMenu")) Is Nothing Then
            If EditMode() = True And Target.Value2 <> OldValue Then
                Application.Run "StepStyle", Target, True
                Application.Run "ApplyStyle", Target, "StepType"
                Target.Offset(1, 0).Name = "NewStepMenu"
                Application.Run "UpdateRanges"
                Application.Run "NewStepButton"
                If LCase(Target.Value2) = "button" Then
                    'Add button specific settings
                    Application.Run "ButtonUpdate", "New Step", False
                End If
                LogText = "Switched Type for step 'New Step'"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("RunSheetProcessingBlockColumnData")) Is Nothing Then
            'Ensure length complies with global setting
            Target.Value2 = Left(Target.Value2, GetSetting("StepMaxLength"))
            LogText = "Changed Processing Block for step '" & StepName & "'"
        ElseIf Not Intersect(Target, ReturnRange("RunSheetStepNameColumnData")) Is Nothing Then
            If EditMode() = True And Target.Value2 <> OldValue Then
                If LCase(Target.Offset(0, -2).Value2) = "button" Then
                    'Ensure length complies with global setting
                    Target.Value2 = Left(Target.Value2, GetSetting("StepMaxLength"))
                    'Renaming of settings if we have a button, but if it returns False - revert the button name
                    If (Application.Run("ButtonSettingsRenaming", OldValue, Target.Value2)) = False Then
                        Target.Value2 = OldValue
                    End If
                End If
                LogText = "Changed Step Name"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("RunSheetDescriptionColumnData")) Is Nothing Then
            LogText = "Changed Description for step '" & StepName & "'"
        ElseIf Not Intersect(Target, ReturnRange("RunSheetTimeStartColumnData")) Is Nothing Then
            'Prevent changing into non-time format
            If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                LogText = "Changed Time Start flag for step '" & StepName & "'"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("RunSheetTimeEndColumnData")) Is Nothing Then
            'Prevent changing into non-time format
            If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                LogText = "Changed End Start flag for step '" & StepName & "'"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("RunSheetTimeColumnData")) Is Nothing Then
            Target.Value2 = OldValue
        ElseIf Not Intersect(Target, ReturnRange("RunSheetUserColumnData")) Is Nothing Then
            Target.Value2 = OldValue
        ElseIf Not Intersect(Target, ReturnRange("RunSheetTitleRow")) Is Nothing Then
            Target.Value2 = OldValue
        ElseIf Not Intersect(Target, ReturnRange("CurrentDayCell")) Is Nothing Then
            If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                Application.Run "DateChange"
                Application.Run "SpecialWorkDays"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("PreviousDayCell")) Is Nothing Then
            If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                LogText = "Changed previous day value"
            Else
                Target.Value2 = OldValue
            End If
        ElseIf Not Intersect(Target, ReturnRange("NextDayCell")) Is Nothing Then
            If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                LogText = "Changed next day value"
            Else
                Target.Value2 = OldValue
            End If
        End If
    ElseIf Target.Worksheet.Name = "Settings" Then
        If Not Intersect(Target, ReturnRange("SettingsIDColumnData")) Is Nothing Then
            LogText = "Changed ID for setting"
        ElseIf Not Intersect(Target, ReturnRange("SettingsTypeColumnData")) Is Nothing Then
            LogText = "Changed Type for setting '" & StepName & "'"
            If LCase(Target.Value2) = "boolean" Then
                BoolFlag = True
            ElseIf LCase(Target.Value2) = "time" Then
                Call ApplyStyle(Target.Offset(0, 3), "TimeCell")
            ElseIf LCase(Target.Value2) = "color" Then
                Application.Run "ColorsFormat", Target.Offset(0, 3)
            End If
        ElseIf Not Intersect(Target, ReturnRange("SettingsIsCustomColumnData")) Is Nothing Then
            'We do not want this value to be easily changeable to avoid confusion and decrease chances of accidental removal of system settings
            Target.Value2 = OldValue
        ElseIf Not Intersect(Target, ReturnRange("SettingsIsEditableColumnData")) Is Nothing Then
            LogText = "Switched IsEditable for setting '" & StepName & "'"
            BoolFlag = True
        ElseIf Not Intersect(Target, ReturnRange("SettingsValueColumnData")) Is Nothing Then
            If LCase(Target.Offset(0, -3).Value2) = "time" Then
                'Prevent changing into non-time format
                If IsTime(Target.Text) = True And Target.Value2 <> OldValue Then
                    LogText = "Changed Value for setting '" & StepName & "'"
                    Call ApplyStyle(Target, "TimeCell")
                Else
                    Target.Value2 = OldValue
                End If
            ElseIf LCase(Target.Offset(0, -3).Value2) = "color" Then
                Application.Run "ColorsFormat", Target
                LogText = "Changed Value for setting '" & StepName & "'"
            Else
                LogText = "Changed Value for setting '" & StepName & "'"
                If LCase(Target.Offset(0, -3).Value2) = "boolean" Then
                    BoolFlag = True
                End If
            End If
            If Target.Value2 <> OldValue Then
                'Check if one of the styles related settings was changed and update styles accordingly
                If Not IsError(Application.Match(LCase(Target.Offset(0, -4).Value2), Array("dateformat", "timeformat", _
                                                                                "datetimeformat", "colorbackcompleted", _
                                                                                "colorfontcompleted", "colorbackfailed", _
                                                                                "colorfontfailed", "colorbackskipped", _
                                                                                "colorfontskipped", "colorbackinprogress", _
                                                                                "colorfontinprogress", "colorfontbutton", _
                                                                                "colorbackbutton", "colordelimiter", _
                                                                                "stepmaxlength" _
                                                                                ), False)) Then
                    Application.ScreenUpdating = False
                    Application.Run "UpdateStyles"
                    Application.ScreenUpdating = True
                End If
                'Check if delimiter was changed
                If LCase(Target.Offset(0, -4).Value2) = "arraydelimiter" Then
                    Dim BoxResult As Integer
                    BoxResult = MsgBox("Changing delimiter will update all settings of 'Array' type." & vbCrLf & "Be sure that your current settings do not use the new value of delimiter or it may break some functions." & vbCrLf & "Are you sure you want to change delimiter from '" & OldValue & "' to '" & Target.Value2 & "'?", vbYesNo + vbExclamation + vbApplicationModal + vbDefaultButton2 + vbMsgBoxSetForeground, "Really change delimiter?")
                    If BoxResult = vbNo Then
                        Target.Value2 = OldValue
                    Else
                        Application.Run "DelimiterUpdate", OldValue, Target.Value2
                    End If
                End If
            End If
        ElseIf Not Intersect(Target, ReturnRange("SettingsDescriptionColumnData")) Is Nothing Then
            LogText = "Changed Description for setting '" & StepName & "'"
        ElseIf Not Intersect(Target, ReturnRange("SettingsTitleRow")) Is Nothing Then
            Target.Value2 = OldValue
        End If
    End If
    
    If EditMode() = True And IsEditor() = True Then
        If BoolFlag = True Then
            If LCase(Target.Value2) = "boolean" Then
                Call ApplyStyle(Target.Offset(0, 3), "Boolean")
            Else
                Call ApplyStyle(Target, "Boolean")
            End If
        End If
        If OldValue <> Target.Value2 Then
            'Times and dates will return Integers, which are not human readable, so using their Text values instead
            If IsTime(Target.Text) = True Then
                Call WriteLog(LogText & " from '" & OldText & "' to '" & Target.Text & "'" & " (" & Target.Address & ")")
            Else
                Call WriteLog(LogText & " from '" & OldValue & "' to '" & Target.Value2 & "'" & " (" & Target.Address & ")")
            End If
        End If
    Else
        Target.Value2 = OldValue
    End If
    
    'Selecting next cell down
    Target.Offset(1, 0).Select
    Application.Run "NextVisible", Target
    'Updating old values
    'Needed, since we can use the drop-down menu, which will not change the selection after value change
    'This results in not logging further changes for cells: change True to False, it's logged, but if you change False to True in the same cell without deselecting it first - it's not logged
    OldValue = Selection.Value2
    OldText = Selection.Text
End Sub
Public Function WriteLog(ByVal LogLine As String, Optional ByVal TimeStamp As String = "", Optional ByVal Computer As String = "", Optional ByVal Username As String = "") As String
    If GetSetting("Logging") = True Then
        Dim LogPath As String, LogName As String, FileNumber As Integer, FullLogLine As String
        FullLogLine = ""
        LogPath = GetSetting("LogPath")
        LogName = GetSetting("LogName")
        If GetSetting("LogRotate") = True Then
            LogName = Format(Now(), "yyyymmdd") & ".log"
            If IsEmpty(LogPath) = True Or LogPath = "" Then
                'Set default log path
                LogPath = GetWokrbookPath(ThisWorkbook.Path) & "logs\"
            End If
        Else
            If IsEmpty(LogName) = True Or LogName = "" Then
                'Set default log name
                LogName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & ".log"
            End If
            If IsEmpty(LogPath) = True Or LogPath = "" Then
                'Set default log path
                LogPath = GetWokrbookPath(ThisWorkbook.Path)
            End If
        End If
        'Add slash to path, it's not there
        If Right(LogPath, 1) <> "\" Then
            LogPath = LogPath & "\"
        End If
        'Check if directory exists and create it
        If CheckDir(LogPath, True) = True Then
            'Get FileNumber as per recommendations from Microsoft manual
            FileNumber = FreeFile
            'Open file in shared mode for appending the line
            Open LogPath & LogName For Append Access Write Shared As #FileNumber
            If TimeStamp = "" Then
                TimeStamp = Format(Now(), GetSetting("DateTimeFormat"))
            End If
            If Username = "" Then
                Username = FullName()
            End If
            If Computer = "" Then
                Computer = VBA.Environ$("COMPUTERNAME")
            End If
            'Write to file
            FullLogLine = TimeStamp & vbTab & Computer & vbTab & Username & vbTab & LogLine
            Print #FileNumber, FullLogLine
            'Close file
            Close #FileNumber
            'StatusBar has limit of 255 symbols, so need to trim it
            Application.StatusBar = Left(FullLogLine, 255)
        End If
        WriteLog = FullLogLine
    Else
        WriteLog = ""
    End If
End Function
Private Sub EndOfDay()
    Dim BackupAuto As Boolean, BackupPath As String, BackupName As String
    Dim BackupBook As Workbook, CurDate As String, ColumnRow As Integer, NewSheet As Worksheet
    BackupAuto = GetSetting("BackupAuto")
    BackupPath = GetSetting("LogPath")
    'Define current date as string, to search for perop screenshots
    CurDate = Format(ReturnRange("CurrentDayCell").Value2, "YYYYMMDD")
    
    'Creating backup of the file, if it's enabled
    If BackupAuto = True Then
        'Default the name, if it's empty
        If IsEmpty(BackupPath) = True Or BackupPath = "" Then
            BackupPath = GetWokrbookPath(ThisWorkbook.Path) & "backup\"
        End If
        'Add slash to path, it's not there
        If Right(BackupPath, 1) <> "\" Then
            BackupPath = BackupPath & "\"
        End If
        'Define preliminary file name
        BackupName = Format(Now(), "yyyymmdd") & ".xlsx"
        'Update file name if file already exists
        If CheckFile(BackupPath & BackupName) = True Then
            BackupName = Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
        End If
        
        'Check if directory exists and create it
        If CheckDir(BackupPath, True) = True Then
            'Create new workbook
            Set BackupBook = Workbooks.Add
            'Copy sheet to new book
            ThisWorkbook.Worksheets("RunSheet").Copy After:=BackupBook.Worksheets(1)
            Set NewSheet = BackupBook.Worksheets("RunSheet")
            'Unprotect the copied sheet
            Application.Run "UnprotectSheet", NewSheet
            'Remove default sheet
            BackupBook.Worksheets(1).Delete
            'Attach screenshots
            Application.Run "ScreenshotsToObjects", NewSheet, CurDate
            'Save new book
            'Despite MSDN telling it has no limit on password - it does. Maximum length is 15 symbols
            BackupBook.SaveAs Filename:=BackupPath & BackupName, FileFormat:=xlWorkbookDefault, AccessMode:=xlExclusive, AddToMru:=False, ReadOnlyRecommended:=True, WriteResPassword:=RandomString(15)
            
            BackupBook.Close
            ThisWorkbook.Activate
        End If
    End If

    'Clear cells
    Call MarkCell(ReturnRange("RunSheetStatusColumnData"), "Purge", ReturnRange("RunSheetStatusColumnData").Cells(1, 1))
    'Update dates
    Application.Run "DateChange"
    'Filter steps
    Application.Run "SpecialWorkDays"
    
    'Remove previously saved screenshots
    Dim Screenshot As Shape
    For Each Screenshot In ThisWorkbook.Worksheets("RunSheet").Shapes
        If LCase(Left(Screenshot.Name, 10)) = "screenshot" Then
            Screenshot.Delete
        End If
    Next Screenshot
    
    'Save updated file
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        Call WriteLog("Failed to save workbook on end of day with error #" & Err.Number)
        MsgBox "Failed to save workbook with error #" & Err.Number, vbCritical + vbOKOnly + vbApplicationModal, "Failed to save"
    End If
    On Error GoTo 0
End Sub
Private Function SendMail(ByVal Subject As String, ByVal MailText As String, ByVal SendTo As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "", Optional ByVal Attachment As String = "", Optional ByVal Importance As Integer = 0, Optional DateSelector As Integer = 1) As Boolean
    Dim OutlookInstance As Object, MailToSend As Object
    'Attempt to get Outlook object
    Application.StatusBar = "Creating mail..."
    On Error Resume Next
    Set OutlookInstance = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        Call WriteLog("Failed to initiate Outlook instance with error #" & Err.Number)
        SendMail = False
        On Error GoTo 0
        Exit Function
    End If
    Set MailToSend = OutlookInstance.CreateItem(0)
    If Err.Number <> 0 Then
        Call WriteLog("Failed to create Outlook mail item with error #" & Err.Number)
        SendMail = False
        On Error GoTo 0
        Exit Function
    End If
    'Update arguments, in case they are actually settings' IDs
    SendTo = StringOrSetting(SendTo)
    CC = StringOrSetting(CC)
    BCC = StringOrSetting(BCC)
    'Also replace new lines with <br>
    MailText = nl2br(DateAnchorsReplace(StringOrSetting(MailText), DateSelector))
    Subject = DateAnchorsReplace(StringOrSetting(Subject), DateSelector)
    Attachment = DateAnchorsReplace(StringOrSetting(Attachment), DateSelector)
    With MailToSend
        .Subject = Subject
        .To = SendTo
        .CC = CC
        .BCC = BCC
        'Define importance
        Select Case Importance
            Case 0, 1, 2
                .Importance = Importance
            Case Else
                .Importance = 0
        End Select
        'Forcing Unicode: is there even a reason not to use it nowadays?
        .InternetCodepage = 65001
        'Doing display first to attach default signature if present
        .Display
        If GetSetting("MailSignatureConcat") = True Then
            .HTMLBODY = MailText & .HTMLBODY
        Else
            .HTMLBODY = Left(.HTMLBODY, InStr(InStr(1, .HTMLBODY, "<body", vbTextCompare), .HTMLBODY, ">", vbTextCompare)) & MailText & Mid(.HTMLBODY, InStr(InStr(1, .HTMLBODY, "<body", vbTextCompare), .HTMLBODY, ">", vbTextCompare) + 1)
        End If
        If IsEmpty(Attachment) = False And Attachment <> "" Then
            Application.StatusBar = "Attaching file..."
            'Add attachment only if file actually exists
            If CheckFile(Attachment) = True Then
                .Attachments.Add Attachment
                If Err.Number <> 0 Then
                    Call WriteLog("Failed attached '" & Attachment & "' with error #" & Err.Number)
                    SendMail = False
                End If
            Else
                Call WriteLog("'" & Attachment & "' was not found: skipping attachement")
                SendMail = False
            End If
        End If
        If GetSetting("AutoMailSend") = True Then
            .Send
            If Err.Number <> 0 Then
                Call WriteLog("Failed to send mail item with error #" & Err.Number)
                SendMail = False
            Else
                Call WriteLog("Mail item sent")
                SendMail = True
            End If
        Else
           Call WriteLog("Mail item prepared")
           SendMail = True
        End If
    End With
    On Error GoTo 0
End Function
