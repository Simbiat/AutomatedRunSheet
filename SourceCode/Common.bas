Attribute VB_Name = "Common"
Option Explicit
Public Sub Optimize(Optional ByVal Off As Boolean = True)
    Dim Action As String
    If Off = True Then
        Action = "Disabling "
    Else
        Action = "Enabling "
    End If
    With Application
        .StatusBar = Action & "events..."
        .EnableEvents = Not (Off)
        .StatusBar = Action & "screen updating..."
        .ScreenUpdating = Not (Off)
        .StatusBar = Action & "printer communications..."
        .PrintCommunication = Not (Off)
        .StatusBar = Action & "alerts..."
        .DisplayAlerts = Not (Off)
        .StatusBar = Action & "animations..."
        .EnableAnimations = Not (Off)
        .StatusBar = Action & "macro animations..."
        .EnableMacroAnimations = Not (Off)
    End With
End Sub
Private Sub Init()
    'Using Application.Run since most of Subs and Functions are Private for access control
    
    'In case a locked cell or an object is selected - select first step. Doing this to avoid potential issues when shared
    Application.Run "SafeCell"
    
    If EditMode() = True Then
        'Check if a non-editor is trying to open the book while it's in editor mode
        If IsEditor() = False Then
            'For some reason at this point msgbox returns location of the file instead of username if I call FullName() function
            MsgBox "Dear " + VBA.Environ$("USERDOMAIN") + "\" + VBA.Environ$("USERNAME") + ", " + vbCrLf + "the RunSheet is currently being edited and you do not appear to be registered editor, thus it will be closed." + vbCrLf + "Please, try again later.", vbOKOnly + vbApplicationModal + vbInformation + vbDefaultButton1 + vbMsgBoxSetForeground, "Ongoing maintenance detected!"
            ThisWorkbook.Close Savechanges:=False
        Else
            'In case we had some issue with the book and it got partially unprotected, we need to properly unprotect to avoid some possible errors in other functions (like styles' update)
            Application.Run "ProtectMe", False
            Application.Run "UpdateRanges"
        End If
    End If
    
    Application.Run "SetFormats"
    Application.Run "TimeFrameCheck"
    
    'Hide/show editor columns
    Application.Run "EditorColumns"
    'Hide rows, that do not correspond to special days (if we have a special day)
    Application.Run "SpecialWorkDays"
    'Hide/show late rows
    Application.Run "LateRows"
End Sub
'Based on https://stackoverflow.com/questions/15723672/how-to-remove-all-non-alphanumeric-characters-from-a-string-except-period-and-sp
'Using it instead of more progressive Regex to avoid unnecessary non-obvious dependencies. Don't think it will affect performance much
Public Function AlphaNumeric(ByVal StringToChange As String) As String
    Dim Symbol As Integer
    Dim FinalString As String
    For Symbol = 1 To Len(StringToChange)
        Select Case Asc(Mid(StringToChange, Symbol, 1))
            Case 48 To 57, 65 To 90, 97 To 122:
                FinalString = FinalString & Mid(StringToChange, Symbol, 1)
        End Select
    Next
    'Limit the length of IDs to something more sensible
    AlphaNumeric = Left(FinalString, GetSetting("StepMaxLength"))
End Function
'Function to toggle boolean
Public Function Toggle(ByVal ToSwitch As Boolean) As Boolean
    Toggle = Not ToSwitch
End Function
'Function to check if time based on https://excel.tips.net/T003292_Checking_for_Time_Input.html
'Works for dates, too
Function IsTime(ByVal TimeText As String) As Boolean
    On Error Resume Next
    IsTime = IsDate(TimeValue(Format(TimeText, GetSetting("TimeFormat"))))
    On Error GoTo 0
End Function
'Random string generator
Public Function RandomString(Optional ByVal Length As Long = 32) As String
    Dim CharacterBank() As String
    Dim X As Long
    'Test Length Input and set it's minimum to 1
    If Length < 1 Then
        Length = 1
    End If
    CharacterBank = Split("a b c d e f g h i j k l m n o p q r s t u v w x y z A B C D E F G H I J K L M N O P Q R S T U V W X Y Z 0 1 2 3 4 5 6 7 8 9", " ")
    'Randomly Select Characters One-by-One
    For X = 1 To Length
        Randomize
        RandomString = RandomString & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next X
End Function
'Function to get localized text boolean value
'Sadly have to use a predefined cell with appropriate value
Public Function LocalisedBoolean(ByVal BoolVal As Boolean) As String
    If BoolVal = True Then
        LocalisedBoolean = ReturnRange("C11", "EditorManual").Text
    Else
        LocalisedBoolean = ReturnRange("C12", "EditorManual").Text
    End If
End Function
'Check if late special steps exist
Public Function LateExists() As Boolean
    LateExists = IsInArray(True, Application.Transpose(ReturnRange("RunSheetIsLateSpecialColumnData").Value2))
End Function
'Convert all newline symbols to <br> tag
Public Function nl2br(ByVal Text As String) As String
    nl2br = Replace(Replace(Replace(Replace(Text, vbCrLf, "<br>"), vbNewLine, "<br>"), vbCr, "<br>"), vbLf, "<br>")
End Function
'Function to enhance regular CBool
Public Function ToBoolean(ByVal ToConvert As Variant) As Boolean
    'Check type of the variable
    Select Case VarType(ToConvert)
        'Semantically 'empty' types
        Case 0, 1, 8192, 8193
            ToBoolean = False
        'Numbers
        Case 2, 3, 4, 5, 6, 7, 14, 17, 20
            'Convert to an Integer first
            If CInt(ToConvert) > 0 Then
                ToBoolean = True
            Else
                ToBoolean = False
            End If
        'Strings or types that can be directly converted to string
        Case 8, 10, 11
            On Error Resume Next
            ToConvert = CStr(ToConvert)
            If Err.Number <> 0 Then
                'Means we might have an onject with default property
                If ToConvert Is Nothing Then
                    ToBoolean = False
                Else
                    ToBoolean = True
                End If
            Else
                Select Case LCase(ToConvert)
                    Case "true", "истина", "yes", "да", "ya"
                        ToBoolean = True
                    Case "false", "ложь", "no", "нет", "nein", "", "n/a", "null", "nan"
                        ToBoolean = False
                    Case Else
                        'Doing this to avoid type mismatch error
                        If IsNumeric(ToConvert) Then
                            If CInt(ToConvert) > 0 Then
                                ToBoolean = True
                            Else
                                ToBoolean = False
                            End If
                        Else
                            'As some other programming languages treat a non empty string as True
                            ToBoolean = True
                        End If
                End Select
            End If
            On Error GoTo 0
        'Objects
        Case 9, 13
            If ToConvert Is Nothing Then
                ToBoolean = False
            Else
                ToBoolean = True
            End If
        'Arrays
        Case Else
            '8192 is added to any regular value identifying array of respective types. 8192 and 8193 would mean arrays of 'empty' elements, so we exclude them
            If VarType(ToConvert) = 12 Or VarType(ToConvert) = 36 Or VarType(ToConvert) >= 8194 Then
                If UBound(ToConvert) - LBound(ToConvert) + 1 < 1 Then
                    ToBoolean = False
                Else
                    ToBoolean = True
                End If
            Else
                ToBoolean = False
            End If
    End Select
End Function
