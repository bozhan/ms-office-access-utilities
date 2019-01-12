VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : PerformanceTimer
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
  lpPerformanceCount As LARGE_INTEGER _
) As Long

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
  lpFrequency As LARGE_INTEGER _
) As Long

Private start As LARGE_INTEGER
Private finish As LARGE_INTEGER
Private freq As Double
Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double
Dim Low As Double
  Low = LI.LowPart
  If Low < 0 Then Low = Low + TWO_32
  LI2Double = LI.HighPart * TWO_32 + Low
End Function

Public Sub StartCounter()
Dim perfFreq As LARGE_INTEGER
  QueryPerformanceFrequency perfFreq
  freq = LI2Double(perfFreq)
  QueryPerformanceCounter start
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Descr.    : returns elapsed time in miliseconds
'---------------------------------------------------------------------------------------
Property Get ElapsedTime( _
  Optional ByVal timeMagnitude As MagnitudeOrder = MagnitudeOrder.milli, _
  Optional ByVal roundTo As Long = 10 _
) As Double
Dim crStart As Double
Dim crStop As Double
  If freq = 0 Then
    ElapsedTime = -1
    Exit Property
  End If
  QueryPerformanceCounter finish
  crStart = LI2Double(start)
  crStop = LI2Double(finish)
  
  ElapsedTime = Round(((crStop - crStart) / (freq * (10 ^ timeMagnitude))), roundTo)
End Property

Property Get StartTime() As Double
  StartTime = LI2Double(start)
End Property

Public Function StopWatch() As Double
  StopWatch = Me.ElapsedTime
  StartCounter
End Function


