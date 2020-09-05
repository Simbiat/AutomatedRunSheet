Attribute VB_Name = "SettingsControl"
Option Explicit
'Function to get setting value
Public Function GetSetting(ByVal SetName As String) As String
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            If LCase(Setting.Offset(0, 1).Value2) = "color" Then
                GetSetting = RangeColorToString(Setting.Offset(0, 4).Interior.Color)
            Else
                GetSetting = Setting.Offset(0, 4).Value2
            End If
            Exit Function
        End If
    Next Setting
    'Return a string identifying, that setting was not found, in case nothing is found
    'It's unlikely a setting will have a value like this
    GetSetting = SetName & "=N/A"
End Function
'Function to get setting's IsEditable flag. Requried only for copying settings for buttons.
Public Function GetSettingEditable(ByVal SetName As String) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            GetSettingEditable = ToBoolean(Setting.Offset(0, 3).Value2)
            Exit Function
        End If
    Next Setting
    'Return False by default
    GetSettingEditable = False
End Function
'Function to determine if we have a setting or a string and update the value, if it's a setting
Public Function StringOrSetting(ByVal SetName As String) As String
    Dim PotentialSetting As String
    'Do not search for a setting if value is empty (slight optimization)
    If IsEmpty(SetName) = True Or SetName = "" Then
        StringOrSetting = ""
        Exit Function
    End If
    PotentialSetting = GetSetting(SetName)
    If LCase(PotentialSetting) = LCase(SetName & "=N/A") Then
        'Return original string
        StringOrSetting = SetName
    Else
        'Return value of the setting
        StringOrSetting = PotentialSetting
    End If
End Function
Public Function SetSetting(ByVal SetName As String, ByVal SetValue As Variant) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            On Error Resume Next
            Setting.Offset(0, 4).Value2 = SetValue
            If Err.Number <> 0 Then
                SetSetting = False
                Exit Function
            End If
            On Error GoTo 0
            If LCase(Setting.Offset(0, 1).Value2) = "boolean" Then
                'Style the boolean value
                Call ApplyStyle(Setting.Offset(0, 4), "Boolean")
            End If
            SetSetting = True
            Exit Function
        End If
    Next Setting
    SetSetting = False
End Function
Private Function RemoveSetting(ByVal SetName As String) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            On Error Resume Next
            'Remove setting
            Setting.EntireRow.Delete
            If Err.Number <> 0 Then
                Call WriteLog("Setting '" & SetName & "' failed to be removed")
                RemoveSetting = False
            Else
                Call WriteLog("Setting '" & SetName & "' is removed")
                RemoveSetting = True
            End If
            On Error GoTo 0
            Exit Function
        End If
    Next Setting
    Call WriteLog("Setting '" & SetName & "' is already removed")
    RemoveSetting = True
End Function
Private Function AddSetting(ByVal SetName As String, ByVal SetType As String, Optional SetEditable As Boolean = False) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            'Setting already exists
            Call WriteLog("Setting '" & SetName & "' is already present")
            AddSetting = False
            Exit Function
        End If
    Next Setting
    'This is, indeed, a new setting add it
    With ThisWorkbook.Worksheets("Settings").Cells(ThisWorkbook.Worksheets("Settings").UsedRange.Rows.Count + 1, 1)
        .Value2 = SetName
        .Offset(0, 1).Value2 = SetType
        .Offset(0, 2).Value2 = True
        .Offset(0, 3).Value2 = SetEditable
    End With
    Call WriteLog("Setting '" & SetName & "' added")
    AddSetting = False
End Function
Private Function RenameSetting(ByVal SetName As String, ByVal NewName As String) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Value2) = LCase(SetName) Then
            Setting.Value2 = NewName
            Call WriteLog("Setting '" & SetName & "' renamed to '" & NewName & "'")
            RenameSetting = True
            Exit Function
        End If
    Next Setting
    'Setting was not found
    Call WriteLog("Setting '" & SetName & "' was not found")
    RenameSetting = False
End Function
Private Function DelimiterUpdate(ByVal OldDelimiter As String, ByVal NewDelimiter As String) As Boolean
    Dim Setting As Range
    'Using For Each to prevent Find function in UI to get filled in by value we are searching for
    On Error Resume Next
    For Each Setting In ReturnRange("SettingsIDColumnData")
        If LCase(Setting.Offset(0, 1).Value2) = "array" Then
            Setting.Offset(0, 4).Value2 = Replace(Setting.Offset(0, 4).Value2, OldDelimiter, NewDelimiter, 1, -1, vbTextCompare)
            If Err.Number <> 0 Then
                DelimiterUpdate = False
            End If
        End If
    Next Setting
    On Error GoTo 0
    DelimiterUpdate = True
End Function
