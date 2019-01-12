VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeadColumnCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class     : HeadColumnCollection
' Author    : Ivanov, Bozhan
' Descr.    : Extends Collection. Represents a collecion of HeadColumn objects
'---------------------------------------------------------------------------------------
' Example:
' Set c = New HeadColumnDict
' c.setHeadColumnAliases "c1", "c2", "c3", "c4"
' c.HeadColumn("c1").Init ThisWorkbook.Worksheets(1), "Title1", "Title2", ...
' c.HeadColumn("c1").InitWithArr ThisWorkbook.Worksheets(1), AliasesArray
'---------------------------------------------------------------------------------------
'
'TODO exapnd all HC ranges to same columnd depth
Option Explicit

Private col As Collection 'Collection
Private colAliases As Collection

Private Sub Class_Initialize()
  Set col = New Collection
  Set colAliases = New Collection
End Sub

Private Sub Class_Terminate()
  Set col = Nothing
  Set colAliases = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Init
' Descr.    : initArray is expected to be in the following format
'   Array(Array(Alias, Range Or Worksheet or ListObject, Array(titles for HC Object), [Data Format]),
'         Array(Alias2, Range Or Worksheet  or ListObject, Array(titles for HC Object 2), [Data Format2]))
' Example   : hcc.Init Array(Array("Alias1", ThisWorkbook.Worksheets("test"), Array("Column Title1", "Column Title2"), "Date"), _
               Array("Alias2", ThisWorkbook.Worksheets("test").ListObjects("testTable").Range, Array("Column Title1", "Column Title2"), "Hyperlink"))
'---------------------------------------------------------------------------------------
Public Sub Init(ByRef initArray As Variant)
Dim domain As Object
On Error GoTo Init_Error
  
  If Not InitArrayIsConsistent(initArray) Then
    err.raise Number:=ErrCode.TypeMismatch, _
      Description:="Wrong format!" & vbLf _
      & "Expected: Array(Array(String, Object, Array(String)),Array(String, Object, Array(String)), [Variant])"
  End If
  
  Dim i As Integer
  For i = LBound(initArray) To UBound(initArray)
    addHeadColumn initArray(i)(0)
    Set domain = initArray(i)(1)
    Me.HeadColumn(initArray(i)(0)).Init domain, initArray(i)(2)
    
    If UBound(initArray(i)) > 2 Then
      If Not IsEmpty(initArray(i)(3)) Then
        Me.HeadColumn(initArray(i)(0)).dataType = initArray(i)(3)
      End If
    End If
  Next i

Init_Exit:
On Error Resume Next
Exit Sub

Init_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Init of module HeadColumnDict" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume Init_Exit

End Sub

'Public Sub Init_withHeadColumnDictType(ByRef InitCol As Collection)
'Dim hcto As InitHeadColumnDictType
'
'  For Each hcto In InitCol
'    If TypeOf hcto Is InitHeadColumnDictType Then
'
'      addHeadColumn hcto.Alias
'
'      Dim domain As Range
'      If TypeOf hcto.domain Is Excel.Worksheet Then
'        Set domain = hcto.domain.UsedRange
'      ElseIf TypeOf hcto.domain Is Excel.ListObject Then
'        Set domain = hcto.domain.Range
'      ElseIf TypeOf hcto.domain Is Excel.Range Then
'        Set domain = hcto.domain
'      Else
'        Set domain = Nothing
'      End If
'
'      Me.HeadColumn(hcto.Alias).init domain:=domain, headTitles:=hcto.Headers
'      If Not IsEmpty(hcto.DataFormat) Then
'        Me.HeadColumn(hcto.Alias).DataFormat = hcto.DataFormat
'      End If
'
'    End If
'  Next hcto
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitArrayIsConsistent
' Descr.    : Checks if the input is in the following format
'   Array(Variant) and more specific
'   Array(Array(String, Array(String)),
'         Array(String, Array(String)))
'---------------------------------------------------------------------------------------
Private Function InitArrayIsConsistent(initArray As Variant) As Boolean
On Error GoTo InitArrayIsConsistent_Error

InitArrayIsConsistent = False
  
  If IsArray(initArray) Then
    If Util.Arrays.IsEmpty(initArray) Then Exit Function
    
    If Util.Arrays.GetDataType(initArray) >= vbArray Then
      Dim i As Integer
      For i = LBound(initArray) To UBound(initArray)
        If InitArrayElementIsConsistant(initArray(i)) Then
          InitArrayIsConsistent = True
        Else
          InitArrayIsConsistent = False
          Exit Function
        End If
      Next i
    End If
  End If

InitArrayIsConsistent_Exit:
On Error Resume Next
Exit Function

InitArrayIsConsistent_Error:
  Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure InitArrayIsConsistent of module HeadColumnDict" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
  InitArrayIsConsistent = False
Resume InitArrayIsConsistent_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : InitArrayElementIsConsistant
' Descr.    : checks if the input is in the following format
'   Array(Variant) and more specific
'   Array(String, Array(String))
'---------------------------------------------------------------------------------------
Private Function InitArrayElementIsConsistant(initArrayElement As Variant) As Boolean
On Error GoTo InitArrayElementIsConsistant_Error

InitArrayElementIsConsistant = False

  If IsArray(initArrayElement) Then
    If Util.Arrays.IsEmpty(initArrayElement) Then Exit Function
          
    If UBound(initArrayElement) >= 2 Then
      If VarType(initArrayElement(0)) = vbString _
      And (TypeName(initArrayElement(1)) = "Range" _
      Or TypeName(initArrayElement(1)) = "Worksheet" _
      Or TypeOf initArrayElement(1) Is Excel.ListObject) _
      And VarType(initArrayElement(2)) >= vbArray + vbString Then
        InitArrayElementIsConsistant = True
        Exit Function
      End If
    End If
  End If

InitArrayElementIsConsistant_Exit:
On Error Resume Next
Exit Function

InitArrayElementIsConsistant_Error:
  Debug.Print "Error " & err.Number & " (" & err.Description & ") in procedure InitArrayElementIsConsistant of module HeadColumnDict" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
  InitArrayElementIsConsistant = False
Resume InitArrayElementIsConsistant_Exit
End Function

Property Get HeadColumn(ByVal alias As Variant) As HeadColumn
  Set HeadColumn = col(CStr(alias))
End Property

Property Get HeadColumns() As Collection
  Set HeadColumns = col
End Property

Property Get aliases() As Collection
  Set aliases = colAliases
End Property

Public Sub SetHeadColumnAliases(ByRef aliases As Variant)
  appendAliases aliases
End Sub

Public Sub AppendHeadColumnAliases(ByRef aliases As Variant)
  appendAliases aliases
End Sub

Private Sub appendAliases(aliases As Variant)
On Error GoTo appendAliases_Error

  If IsArray(aliases) Then
    If Util.Arrays.IsEmpty(aliases) Then
      err.raise Number:=ErrCode.TypeMismatch, _
        Description:="Empty aliases array provided"
    Else 'is an array and is not empty
      Dim l As Integer
      For l = LBound(aliases) To UBound(aliases)
        If VarType(aliases(l)) = vbString _
        And aliases(l) <> vbNullString Then
          addHeadColumn aliases(l)
        End If
      Next l
    End If
  ElseIf Len(CStr(aliases)) > 0 Then
    addHeadColumn CStr(aliases)
  Else
    err.raise Number:=ErrCode.TypeMismatch, _
      Description:="Provided parameter was not an array or convertable to string"
  End If

appendAliases_Exit:
On Error Resume Next
Exit Sub

appendAliases_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure appendAliases of module HeadColumnDict" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume appendAliases_Exit
End Sub

Public Function count() As Long
  count = col.count
End Function

Public Sub Remove(ByVal alias As Variant)
  col.Remove CStr(alias)
  colAliases.Remove CStr(alias)
End Sub

Public Sub Clear()
  Set col = New Collection
  Set colAliases = New Collection
End Sub

Public Function IsEveryHeadFound() As Boolean
Dim hc As HeadColumn
    
  IsEveryHeadFound = True
    
  For Each hc In col
    If hc.Head Is Nothing Then
      IsEveryHeadFound = False
      Exit Function
    End If
  Next hc
End Function

Public Function GetTitlesOfHeadsNotFound() As Variant
Dim hc As HeadColumn
Dim ae As ArrayList
Dim t As Variant

  GetTitlesOfHeadsNotFound = Array()
  Set ae = New ArrayList
  
  For Each hc In col
    If hc.Head Is Nothing Then ' head is not found & titles are set
      ae.Append hc.Headers
    End If
  Next hc
  
  GetTitlesOfHeadsNotFound = ae.Values
  Set ae = Nothing
End Function

Public Function GetTitlesOfHeadsFound() As Variant
Dim hc As Object
Dim ae As ArrayList
Dim t As Variant

  Set ae = New ArrayList
  
  For Each hc In col
    If Not hc.Head Is Nothing Then ' head is not found & titles are set
      ae.Append hc.Headers
    End If
  Next hc
  
  GetTitlesOfHeadsFound = ae.Values
  Set ae = Nothing
End Function

' TODO:
' Impl FindMatchingWorksheet(ParamArray ANDCond() as Variant, ParamArray ORCond() as Variant)
'   supple a string with condition: A AND B OR C
'TODO impl a find appropriate worksheet function by testing HCAliases

Public Function FindMatchingWorksheet( _
  ParamArray containingStrings() As Variant _
) As Worksheet

End Function

Public Function GetMaxUsedRowsFromAllHeaderFound() As Long
Dim a As Variant
Dim l As Long
  
  l = 0
  For Each a In Me.aliases
    If Me.HeadColumn(a).LastCell.Row > l Then
      l = Me.HeadColumn(a).LastCell.Row
    End If
  Next a
  
  GetMaxUsedRowsFromAllHeaderFound = l
End Function

Public Function GetMinStartRowFromAllHeaderFound() As Long
Dim a As Variant
Dim l As Long
  
  l = Me.GetMaxUsedRowsFromAllHeaderFound()
  For Each a In Me.aliases
    If Me.HeadColumn(a).Head.Row < l Then
      l = Me.HeadColumn(a).Head.Row
    End If
  Next a
  
  GetMinStartRowFromAllHeaderFound = l
End Function

Public Function GetFirstFound() As HeadColumn
Dim hc As HeadColumn

  Set GetFirstFound = Nothing
  
  For Each hc In col
    If Not hc.Head Is Nothing Then
      Set GetFirstFound = hc
      Exit Function
    End If
  Next hc
End Function

Public Function HeadersConsistant() As Boolean
Dim foundHeaders As ArrayList
Dim missingHeaders As ArrayList
 
  HeadersConsistant = True
   
  If col Is Nothing Then
    HeadersConsistant = False
    Exit Function
  End If
  
  If col.count = 0 Then
    HeadersConsistant = False
    Exit Function
  End If
  
  ' check for found headers
  Set foundHeaders = New ArrayList
  foundHeaders.Append Me.GetTitlesOfHeadsFound()
  
  If Not foundHeaders.IsAllocated _
  Or foundHeaders.IsEmpty Then 'no heads found
    HeadersConsistant = False
    Exit Function
  End If
  
  ' check for headers not found and report
  Set missingHeaders = New ArrayList
  missingHeaders.Append Me.GetTitlesOfHeadsNotFound()
  If missingHeaders.IsAllocated And Not missingHeaders.IsEmpty Then 'some heads not found
    HeadersConsistant = False
    Exit Function
  End If
End Function

Private Sub addHeadColumn(ByVal alias As Variant)
  col.Add New HeadColumn, CStr(alias)
  colAliases.Add CStr(alias), CStr(alias)
End Sub

Public Sub ReCalc()
Dim hc As HeadColumn
  For Each hc In col
    hc.ReCalc
  Next hc
End Sub
