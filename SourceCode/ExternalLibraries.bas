Attribute VB_Name = "ExternalLibraries"
Option Explicit
#If VBA7 Then
    'Required for Sleep
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    'Required to get current resolution
    'Not sure what Alias is for (it does not work for me at all), but retaining just in case
    Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As LongPtr) As LongPtr
    'Required for using clipboard
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As LongPtr, ByVal lpString As String, ByVal nMaxCount As LongPtr) As Long
#Else
    'Required for Sleep
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    'Required to get current resolution
    Public Declare Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    'Required for using clipboard
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#End If

