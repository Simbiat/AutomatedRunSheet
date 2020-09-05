Attribute VB_Name = "Colors"
Option Explicit
'Functions dealing with colors
Public Function StringToRGB(ByVal ColorString As String) As Long
    Dim ColorArray() As String
    ColorArray = ColorStringToArray(ColorString)
    StringToRGB = RGB(CInt(ColorArray(0)), CInt(ColorArray(1)), CInt(ColorArray(2)))
End Function
'Helper function to convert our RGB strings to array, that we will work with in other functions
Private Function ColorStringToArray(ByVal ColorString As String) As String()
    Dim ColorArray() As String
    Dim ArraySize As Integer: ArraySize = 0
    'Fill array from the string that was passed
    ColorArray = Split(ColorString, GetSetting("ArrayDelimiter"), -1, 1)
    'Get  size of the array
    ArraySize = UBound(ColorArray) - LBound(ColorArray) + 1
    'Fill missing entries if any
    If ArraySize = 0 Then
        ReDim Preserve ColorArray(0 To 2) As String
        ColorArray(0) = "0"
        ColorArray(1) = "0"
        ColorArray(2) = "0"
    ElseIf ArraySize = 1 Then
        ReDim Preserve ColorArray(0 To 2) As String
        ColorArray(1) = "0"
        ColorArray(2) = "0"
    ElseIf ArraySize = 2 Then
        ReDim Preserve ColorArray(0 To 2) As String
        ColorArray(2) = "0"
    ElseIf ArraySize > 3 Then
        ReDim Preserve ColorArray(0 To 2) As String
    End If
    'Check that strings in the array are actually valid RGB values
    ColorArray(0) = ColorCheck(ColorArray(0))
    ColorArray(1) = ColorCheck(ColorArray(1))
    ColorArray(2) = ColorCheck(ColorArray(2))
    ColorStringToArray = ColorArray
End Function
Private Function ColorCheck(ByVal ColorValue As String) As String
    ColorCheck = ColorValue
    If CInt(Val(ColorCheck)) < 0 Then
        ColorCheck = "0"
    End If
    If CInt(Val(ColorCheck)) > 255 Then
        ColorCheck = "255"
    End If
    'Ensure, that we have an number. If string has no numbers, below will return 0
    If CInt(Val(ColorCheck)) = 0 Then
        ColorCheck = "0"
    End If
End Function
'Based on https://www.thespreadsheetguru.com/blog/vba-macro-code-lighten-darken-fill-colors-excel (1st approach) and comments there
Public Function ToneRGB(ByVal ColorString As String, Optional Darken As Boolean = False, Optional ToneLevel As Integer = 7) As Long
    Dim ColorArray() As String
    ColorArray = ColorStringToArray(ColorString)
    If Darken = True Then
        ToneRGB = RGB(Round(CInt(ColorArray(0)) - (CInt(ColorArray(0)) * (ToneLevel / 15)), 0), Round(CInt(ColorArray(1)) - (CInt(ColorArray(1)) * (ToneLevel / 15)), 0), Round(CInt(ColorArray(2)) - (CInt(ColorArray(2)) * (ToneLevel / 15)), 0))
    Else
        ToneRGB = RGB(Round(CInt(ColorArray(0)) + ((255 - CInt(ColorArray(0))) * (ToneLevel / 15)), 0), Round(CInt(ColorArray(1)) + ((255 - CInt(ColorArray(1))) * (ToneLevel / 15)), 0), Round(CInt(ColorArray(2)) + ((255 - CInt(ColorArray(2))) * (ToneLevel / 15)), 0))
    End If
End Function
Public Function RangeColorToString(RangeColor As Long) As String
    RangeColorToString = (RangeColor Mod 256) & GetSetting("ArrayDelimiter") & ((RangeColor \ 256) Mod 256) & GetSetting("ArrayDelimiter") & (RangeColor \ 65536)
End Function

