Attribute VB_Name = "DateTime"
Option Explicit
'Highlighting steps that have not bee completed on time
Private Sub TimeFrameCheck(Optional ByVal CellRange As Range = Nothing)
    Application.StatusBar = "Checking time cells for expiration..."
    Dim CurTime As Date
    Dim TimeFormat As String, ZeroTime As String, TimeStartMargin As Date, TimeEndMargin As Date
    Dim StatusCell As Range, TimeStartCell As Range, TimeEndCell As Range
    Dim GoodTimeCells As Range, BadTimeCells As Range
    TimeFormat = GetSetting("TimeFormat")
    ZeroTime = Format("00:00", TimeFormat)
    TimeStartMargin = TimeValue(Format(GetSetting("TimeStartMargin"), TimeFormat))
    TimeEndMargin = TimeValue(Format(GetSetting("TimeEndMargin"), TimeFormat))
    CurTime = Time()
    If CellRange Is Nothing Then
        Set CellRange = ReturnRange("RunSheetStatusColumnData")
    End If
    For Each StatusCell In CellRange
        Set TimeStartCell = StatusCell.Offset(0, -4)
        Set TimeEndCell = StatusCell.Offset(0, -3)
        If StatusCell.Value2 <> "Completed" And StatusCell.Value2 <> "Skipped" Then
            If TimeEndCell.Text <> ZeroTime Then
                If (CurTime - TimeEndCell.Value2 - TimeEndMargin) > 0 Then
                    If BadTimeCells Is Nothing Then
                        Set BadTimeCells = TimeEndCell
                    Else
                        Set BadTimeCells = Union(BadTimeCells, TimeEndCell)
                    End If
                Else
                    If GoodTimeCells Is Nothing Then
                        Set GoodTimeCells = TimeEndCell
                    Else
                        Set GoodTimeCells = Union(GoodTimeCells, TimeEndCell)
                    End If
                End If
            ElseIf TimeStartCell.Text <> ZeroTime Then
                If (CurTime - TimeStartCell.Value2 - TimeStartMargin) > 0 Then
                    If BadTimeCells Is Nothing Then
                        Set BadTimeCells = TimeStartCell
                    Else
                        Set BadTimeCells = Union(BadTimeCells, TimeStartCell)
                    End If
                Else
                    If GoodTimeCells Is Nothing Then
                        Set GoodTimeCells = TimeStartCell
                    Else
                        Set GoodTimeCells = Union(GoodTimeCells, TimeStartCell)
                    End If
                End If
            End If
        Else
            If GoodTimeCells Is Nothing Then
                Set GoodTimeCells = Union(TimeStartCell, TimeEndCell)
            Else
                Set GoodTimeCells = Union(GoodTimeCells, TimeStartCell, TimeEndCell)
            End If
        End If
    Next StatusCell
    'Apply styles
    If Not GoodTimeCells Is Nothing Then
        Call ApplyStyle(GoodTimeCells, "TimeCell")
    End If
    If Not BadTimeCells Is Nothing Then
        Call ApplyStyle(BadTimeCells, "MissedTime")
    End If
End Sub
'Get next (or previous) working day, which is not a weekend according to our setup (not relying on system)
Private Function NextWorkDay(ByVal DateToChange As Date, Optional ByVal Previous As Boolean = False) As Date
    If Previous = True Then
        NextWorkDay = DateAdd("d", -1, DateToChange)
    Else
        NextWorkDay = DateAdd("d", 1, DateToChange)
    End If
    While IsWorkDay(NextWorkDay) = True
        If Previous = True Then
            NextWorkDay = DateAdd("d", -1, NextWorkDay)
        Else
            NextWorkDay = DateAdd("d", 1, NextWorkDay)
        End If
    Wend
End Function
Private Function IsWorkDay(ByVal DateToCheck As Date) As Boolean
    IsWorkDay = True
    Dim Weekends() As String
    'Get days considered as weekends
    Weekends() = Split(GetSetting("WeekEnds"), GetSetting("ArrayDelimiter"))
    If IsInArray(CStr(Application.Weekday(DateToCheck)), Weekends) = True Then
        Dim DateFormated As String
        DateFormated = Format(DateToCheck, GetSetting("DateFormat"))
        If IsInArray(DateFormated, DatesArray(GetSetting("FirstSpecialDays"))) = True Then
            IsWorkDay = False
        End If
        If IsInArray(DateFormated, DatesArray(GetSetting("RegularSpecialDays"))) = True Then
            IsWorkDay = False
        End If
        If IsInArray(DateFormated, DatesArray(GetSetting("LastSpecialDays"))) = True Then
            IsWorkDay = False
        End If
    Else
        IsWorkDay = False
    End If
End Function
Public Function DatesArray(ByVal Dates As String) As String()
    Dim DatesList() As String
    Dim SubDate As Integer
    Dim DateFormat As String
    DateFormat = GetSetting("DateFormat")
    'Fill array from the string that was passed
    DatesList = Split(Dates, GetSetting("ArrayDelimiter"), -1, 1)
    For SubDate = LBound(DatesList) To UBound(DatesList)
        'Format the dates according to common setup
        DatesList(SubDate) = Format(DatesList(SubDate), DateFormat)
    Next SubDate
    DatesArray = DatesList
End Function
Private Sub DateChange()
    Dim CurCurDate As Date, CurPrevDate As Date, CurNextDate As Date
    Dim NewCurDate As Date, NewPrevDate As Date, NewNextDate As Date
    
    'Get current values
    CurCurDate = ReturnRange("CurrentDayCell").Value2
    CurPrevDate = ReturnRange("PreviousDayCell").Value2
    CurNextDate = ReturnRange("NextDayCell").Value2
    
    'New values
    If CurCurDate > CurPrevDate And CurNextDate > CurCurDate Then
        NewCurDate = CurNextDate
        NewPrevDate = CurCurDate
    Else
        NewCurDate = CurCurDate
        NewPrevDate = NextWorkDay(CurCurDate, True)
    End If
    NewNextDate = NextWorkDay(NewCurDate)
    
    ReturnRange("CurrentDayCell").Value2 = NewCurDate
    Call WriteLog("Changed Current date from " & CurCurDate & " to " & NewCurDate)
    ReturnRange("PreviousDayCell").Value2 = NewPrevDate
    Call WriteLog("Changed Current date from " & CurPrevDate & " to " & NewPrevDate)
    ReturnRange("NextDayCell").Value2 = NewNextDate
    Call WriteLog("Changed Current date from " & CurNextDate & " to " & NewNextDate)
End Sub
'Replaces YYYY, MM and DD with a date
Public Function DateAnchorsReplace(ByVal Text As String, Optional DateSelector As Integer = 1) As String
    Dim YYYY As String, YY As String, MM As String, DD As String
    Dim DateToUse As Date
    'Choose date from those setup in the RunSheet
    Select Case DateSelector
        Case 0
            DateToUse = ReturnRange("PreviousDayCell").Value2
        Case 2
            DateToUse = ReturnRange("NextDayCell").Value2
        Case Else
            DateToUse = ReturnRange("CurrentDayCell").Value2
    End Select
    'Get components of the date
    YYYY = Format(DateToUse, "YYYY")
    YY = Format(DateToUse, "YY")
    MM = Format(DateToUse, "MM")
    DD = Format(DateToUse, "DD")
    'Replace anchors
    DateAnchorsReplace = Replace(Replace(Replace(Replace(Text, "%YYYY%", YYYY), "%MM%", MM), "%DD%", DD), "%YY%", YY)
End Function
