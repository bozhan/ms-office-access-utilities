VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilRegExPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get DatePattern() As String
Dim rxDay As String, rxMonth As String, rxYear As String, rxDateSeperator As String
  
  rxDay = "(0[1-9]{1}|[1-9]{1}|[12][0-9]{1}|3[01]{1})"
  rxMonth = "(0[1-9]{1}|[1-9]{1}|1[012]{1})"
  rxYear = "((19|20)?[0-9]{2})"
  rxDateSeperator = "[\.\/\-]{1}"
  
  DatePattern = _
  "(" & "(" & rxDay & rxDateSeperator & rxMonth & rxDateSeperator & rxYear & ")" & _
  "|" & "(" & rxMonth & rxDateSeperator & rxDay & rxDateSeperator & rxYear & ")" & _
  "|" & "(" & rxYear & rxDateSeperator & rxDay & rxDateSeperator & rxMonth & ")" & _
  "|" & "(" & rxDay & rxDateSeperator & rxYear & rxDateSeperator & rxMonth & ")" & _
  "|" & "(" & rxMonth & rxDateSeperator & rxYear & rxDateSeperator & rxDay & ")" & _
  "|" & "(" & rxYear & rxDateSeperator & rxMonth & rxDateSeperator & rxDay & ")" & ")"
End Property

Property Get NotDatePattern() As String
Dim rxNotDay As String, rxNotMonth As String, rxNotYear As String, rxNotDateSeperator As String
  
  rxNotDay = "([^\d]+)"
  rxNotMonth = "([^\d]+)"
  rxNotYear = "([^\d]+)"
  rxNotDateSeperator = "[^\.\/\-]?"

  NotDatePattern = _
  "^(" & "(" & rxNotDay & rxNotDateSeperator & rxNotMonth & rxNotDateSeperator & rxNotYear & ")" & _
  "|" & "(" & rxNotMonth & rxNotDateSeperator & rxNotDay & rxNotDateSeperator & rxNotYear & ")" & _
  "|" & "(" & rxNotYear & rxNotDateSeperator & rxNotDay & rxNotDateSeperator & rxNotMonth & ")" & _
  "|" & "(" & rxNotDay & rxNotDateSeperator & rxNotYear & rxNotDateSeperator & rxNotMonth & ")" & _
  "|" & "(" & rxNotMonth & rxNotDateSeperator & rxNotYear & rxNotDateSeperator & rxNotDay & ")" & _
  "|" & "(" & rxNotYear & rxNotDateSeperator & rxNotMonth & rxNotDateSeperator & rxNotDay & ")" & ")$"
End Property

Property Get ValueInQuotesPattern() As String
  
  ValueInQuotesPattern = "([" & """" & "'])(?:(?=(\\?))\2.)*?\1"

End Property
'---------------------------------------------------------------------------------------
' Procedure : PCMNumberPattern
' Descr.    : Provides the patter for a PCm number consisting of
'   [1 Digit]-[2 Digits]-[6 Digits] or
'   [1 Digit]-[2 Digits]-[6 Digits]-[2 Digits]-[3 Digits]
'---------------------------------------------------------------------------------------
Property Get PCMNumberPattern() As String
Dim pcmShort As String
Dim pcmLong As String

  pcmShort = "([0-9]{1}[\-][0-9]{2}[\-][0-9]{6})"
  pcmLong = "([0-9]{1}[\-][0-9]{2}[\-][0-9]{6}[\-][0-9]{2}[\-][0-9]{3})"
  PCMNumberPattern = pcmShort & "|" & pcmLong
End Property

'---------------------------------------------------------------------------------------
' Procedure : NotPCMNumberPattern
' Descr.    : Provides the inverse regex for PCM Number
'---------------------------------------------------------------------------------------
Property Get NotPCMNumberPattern() As String
Dim pcmShort As String
Dim pcmLong As String

  pcmShort = "([0-9]{1}[\-][0-9]{2}[\-][0-9]{6})"
  pcmLong = "([0-9]{1}[\-][0-9]{2}[\-][0-9]{6}[\-][0-9]{2}[\-][0-9]{3})"
  NotPCMNumberPattern = "^((?!(" & pcmShort & "|" & pcmLong & "))[\s\S])*"
End Property

Property Get LettersPattern() As String
  LettersPattern = "([A-Za-z]+)"
End Property

Property Get NumbersPattern() As String
  NumbersPattern = "([0-9]+)"
End Property

Property Get VariableNamesPattern() As String
  VariableNamesPattern = "([0-9_]*\w*)"
End Property

