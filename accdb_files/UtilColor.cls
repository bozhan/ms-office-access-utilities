VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ColorDec
  White = 16777215              'RGB(255, 255, 255)
  Black = 0                     'RGB(0, 0, 0)
  Blue = 14390640               'RGB(112, 149, 219)
  Blue_Highlight = -2147483635
  Blue_Accent1 = 12419407
  Blue_80_Percent = 15849926    'RGB(198,217,241)
  Purple_60_Percent = 14336460  'RGB(204,193,218)
  Gray = 11842740               'RGB(180, 180, 180)
  Gray_Light = 15790320         'RGB(240, 240, 240)
  Gray_Alternate = 15921906
  Red = 6118894                 'RGB(238, 93, 93)
  Red_Warning = 6843350
  Red_Warning_Light = 9737945
  Green = 9164625
  Green_Light = 14745056        'RGB(224, 253, 224)
End Enum

Public Enum ColorIndex
  None = -4142
  Automatic = -4105
  Black = 1
  White = 2
  Grey = 48
  Lightgrey = 15
  Darkgrey = 16
  Red = 3
  Green = 43
  LightGreen = 35
  Yellow = 6
  Lightyellow = 36
  Blue = 41
  Lightblue = 37
  Purple = 17
  Lightpurple = 39
  Darkpurple = 47
  Orange = 45
  Lightorange = 44
  Darkorange = 46
End Enum

Public Function GetRed(c As ColorDec) As Long
  GetRed = c Mod 256
End Function

Public Function GetGreen(c As ColorDec) As Long
  GetGreen = (c \ 256) Mod 256
End Function

Public Function GetBlue(c As ColorDec) As Long
  GetBlue = (c \ 256 \ 256) Mod 256
End Function


'TODO Revise
Private Sub testMatchColorWithTolerance()
  Dim TargetColor, Tolerance, TmpColor, TargetColorRed
  Dim TargetColorGreen, TargetColorBlue, Cell, MyColor
  Dim Red, Green, Blue, Distance
  
  TargetColor = RGB(0, 255, 0)
  Tolerance = 200
  'Extract the color components. The extraction is unneeded, but if the target
  'color came from the color of a selected cell, it would be needed.
  TmpColor = TargetColor
  TargetColorRed = TmpColor Mod 256
  TmpColor = TmpColor \ 256
  TargetColorGreen = TmpColor Mod 256
  TmpColor = TmpColor \ 256
  TargetColorBlue = TmpColor Mod 256

  For Each Cell In ActiveSheet.UsedRange.Cells
    MyColor = Cell.Characters.Font.Color 'Color is a number
    'Extract the RGB components of the color
    Red = MyColor Mod 256
    MyColor = MyColor \ 256
    Green = MyColor Mod 256
    MyColor = MyColor \ 256
    Blue = MyColor Mod 256
    'Find the distance from the target color
    Distance = ((Red - TargetColorRed) ^ 2 + _
                (Green - TargetColorGreen) ^ 2 + _
                (Blue - TargetColorBlue) ^ 2) ^ 0.5
    If Distance < Tolerance Then
      Cell.Interior.Color = RGB(230, 230, 230) 'Mark the cell using its background color
    End If
  Next
End Sub


