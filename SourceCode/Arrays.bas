Attribute VB_Name = "Arrays"
Option Explicit
'Function to check if a value is in array based on https://stackoverflow.com/questions/38267950/check-if-a-Value2-is-in-an-array-or-not-with-excel-vba
Public Function IsInArray(ByVal Needle As Variant, ByVal Haystack As Variant, Optional ByVal IgnoreCase = True) As Boolean
    Dim Straw As Integer
    For Straw = LBound(Haystack) To UBound(Haystack)
        'Unlike in original code, we simulate strict type comparisson
        If VarType(Haystack(Straw)) = VarType(Needle) Then
            'If we are dealing with string - ignore case by default
            If VarType(Needle) = 8 And IgnoreCase = True Then
                If LCase(Haystack(Straw)) = LCase(Needle) Then
                    IsInArray = True
                    Exit Function
                End If
            Else
                If Haystack(Straw) = Needle Then
                    IsInArray = True
                    Exit Function
                End If
            End If
        End If
    Next Straw
    IsInArray = False
End Function
'Function to remove first element of an array
Public Function RemoveFirstElement(ByRef OriginalArray() As String) As String()
    Dim ArrayElement As Integer, NewArray() As String
    ReDim NewArray(UBound(OriginalArray) - 1)
    For ArrayElement = LBound(OriginalArray) To UBound(OriginalArray)
        If ArrayElement > 0 Then
            NewArray(ArrayElement - 1) = OriginalArray(ArrayElement)
        End If
    Next ArrayElement
    RemoveFirstElement = NewArray
End Function
'Function to trim all strings in an array
Public Function TrimArray(ByVal ToTrim As Variant) As String()
    Dim ArrayElement As Integer
    For ArrayElement = LBound(ToTrim) To UBound(ToTrim)
        ToTrim(ArrayElement) = Trim(ToTrim(ArrayElement))
    Next ArrayElement
    TrimArray = ToTrim
End Function
