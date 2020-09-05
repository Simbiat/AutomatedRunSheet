VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditorManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    Call Optimize(True)
    ThisWorkbook.Worksheets("EditorManual").DisplayPageBreaks = False
    
    'Styling elements in manual as visual cue
    Call ApplyStyle(ReturnRange("A3", "EditorManual"), "Button")

    With ReturnRange("A6", "EditorManual")
        .Interior.Color = StringToRGB(GetSetting("ColorBackCompleted"))
        .Font.Color = StringToRGB(GetSetting("ColorFontCompleted"))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = .Font.Color
        .Borders.Weight = xlThin
    End With
    With ReturnRange("A7", "EditorManual")
        .Interior.Color = StringToRGB(GetSetting("ColorBackInProgress"))
        .Font.Color = StringToRGB(GetSetting("ColorFontInProgress"))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = .Font.Color
        .Borders.Weight = xlThin
    End With
    With ReturnRange("A8", "EditorManual")
        .Interior.Color = StringToRGB(GetSetting("ColorBackFailed"))
        .Font.Color = StringToRGB(GetSetting("ColorFontFailed"))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = .Font.Color
        .Borders.Weight = xlThin
    End With
    With ReturnRange("A9", "EditorManual")
        .Interior.Color = StringToRGB(GetSetting("ColorBackSkipped"))
        .Font.Color = StringToRGB(GetSetting("ColorFontSkipped"))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = .Font.Color
        .Borders.Weight = xlThin
    End With
    Call ApplyStyle(ReturnRange("A10", "EditorManual"), "MissedTime")
    Call ApplyStyle(ReturnRange("C11:C12", "EditorManual"), "Boolean")
    'For delimiters do not change font color, so it's possible to read it
    Call ApplyStyle(ReturnRange("A4,A11", "EditorManual"), "Delimiter")
    With ReturnRange("A4,A11", "EditorManual")
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Format description cells
    Call ApplyStyle(ReturnRange("B2:B4,B6:B11,B13:B19,D2:D20,F2:F13", "EditorManual"), "Description")
    
    ThisWorkbook.Worksheets("EditorManual").Calculate
    Call Optimize(False)
End Sub

