VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilRegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : RegEx
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Public Function Replace( _
ByVal pattern As String, _
ByVal searchValue As String, _
ByVal ReplaceValue As String, _
Optional ByVal IsGlobal As Boolean = True, _
Optional ByVal IsCaseSensitive As Boolean = True _
) As String
Dim RegEx As Object
  
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Global = IsGlobal
  RegEx.pattern = pattern
  RegEx.IgnoreCase = Not IsCaseSensitive
  Replace = RegEx.Replace(searchValue, ReplaceValue)
End Function

Public Function match( _
ByVal pattern As String, _
ByVal queryString As String, _
Optional ByVal IsGlobal As Boolean = True, _
Optional ByVal IsCaseSensitive As Boolean = True _
) As Variant
Dim RegEx As Object
Dim RegExeResults As Object
Dim RegExRes As Object
Dim result As ArrayList

  
  Set RegEx = CreateObject("vbscript.regexp")
  Set result = New ArrayList
  
  With RegEx
    .Global = IsGlobal
    .pattern = pattern
    .IgnoreCase = Not IsCaseSensitive
    
    If .Test(queryString) Then
      Set RegExeResults = .Execute(queryString)
      For Each RegExRes In RegExeResults
        result.Append RegExRes.submatches(0)
      Next
    End If
  End With
  
  match = result.Values
End Function


