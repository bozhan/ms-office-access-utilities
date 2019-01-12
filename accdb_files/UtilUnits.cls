VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MagnitudeOrder
  tera = 12
  giga = 9
  mega = 6
  kilo = 3
  hecto = 2
  deca = 1
  one = 0
  deci = -1
  centi = -2
  milli = -3
  micro = -6
  nano = -9
  pico = -12
End Enum

Private Const TPI As Integer = 1440 'twips per inch
Private Const PPI As Integer = 72 'points per inch
Private Const IPC As Single = 0.3937 'Inch per cm
Private Const TPC As Integer = 567 'twips per cm

Property Get TwipsPerInch() As Integer
  TwipsPerInch = TPI
End Property

Property Get PixelsPerInch() As Integer
  PixelsPerInch = PPI
End Property

Property Get InchesPerCm() As Single
  InchesPerCm = IPC
End Property

Property Get TwipsPerCm() As Integer
  TwipsPerCm = TPC
End Property

Public Function TwipsToInch(Value As Single) As Variant
    TwipsToInch = Value / TwipsPerInch
End Function

Public Function InchToTwips(Value As Single) As Variant
    InchToTwips = Value * TwipsPerInch
End Function

Public Function PointsToInch(Value As Single) As Variant
    PointsToInch = Value / PixelsPerInch
End Function

Public Function InchToPoints(Value As Single) As Variant
    InchToPoints = Value * PixelsPerInch
End Function

Public Function TwipsToCm(Value As Single) As Variant
    TwipsToCm = Value * TwipsPerCm
End Function

Public Function CmToTwips(Value As Single) As Variant
    CmToTwips = Value / TwipsPerCm
End Function

Public Function PointsToCm(Value As Single) As Variant
    PointsToCm = Value / (PixelsPerInch / InchesPerCm)
End Function

Public Function CmToPoints(Value As Single) As Variant
  If IsNumeric(Value) Then
    CmToPoints = (Value * InchesPerCm) * PixelsPerInch
  End If
End Function

Public Function SingleToString(ByVal Value As Single) As String
  SingleToString = Replace(CStr(Value), ",", ".")
End Function

Public Function StringToSingle(ByVal Value As String) As Single
  StringToSingle = CSng(Replace(Value, ".", ","))
End Function

