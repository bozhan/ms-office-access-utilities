VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ArrayList
' Author    : Ivanov, Bozhan
' Purpose   : Extends the functionality of a single dimention array
' Note      : Since the extend is class extra array dimentions can be added by appending
'             an object of the ArrayList class as an element of another object of ArrayList
'---------------------------------------------------------------------------------------
'TODO Implement the RemoveDuplicate and getDuplicate functions from utilarrays
Option Explicit
Private a() As Variant

Private Sub Class_Initialize()
  a = Array()
End Sub

Property Get Values() As Variant
  Values = a
End Property

Property Let Values(ByRef inputArray As Variant)
  a = inputArray
End Property

Property Get Value(ByVal key As Long) As Variant
Attribute Value.VB_UserMemId = 0
On Error GoTo value_Error

  If key <= Me.ub Then
    If Me.VarsType = vbObject Or IsObject(a(key)) Then
      Set Value = a(key)
    Else
      Value = a(key)
    End If
  Else
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="Property Value - Input key value is grater than Upperbound!"
  End If

value_Exit:
Exit Property

value_Error:
  MsgBox "Error " & err.Number & " (" & _
  err.Description & ") in procedure value of module ArrayList" & vbLf & _
    InfoErrMsg()
Resume value_Exit
End Property

Property Let Value(ByVal key As Long, ByVal Value As Variant)
On Error GoTo value_Error

  If key <= Me.ub Then
    If Me.VarsType = vbObject Or IsObject(Value) Then
      Set a(key) = Value
    Else
      a(key) = Value
    End If
  Else
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="Property Value - Input key value is grater than Upperbound!"
  End If

value_Exit:
Exit Property

value_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure value of module ArrayList" & vbLf & _
    InfoErrMsg()
Resume value_Exit
End Property

Property Get ub() As Long
  If Me.IsAllocated Then
    ub = UBound(a)
  Else
    ub = -1
  End If
End Property

Property Get lb() As Long
  If Me.IsAllocated Then
    lb = LBound(a)
  Else
    lb = 0
  End If
End Property

Property Get IsAllocated() As Boolean
  IsAllocated = Util.Arrays.IsAllocated(a)
End Property

Property Get Dimensions() As Integer
  Dimensions = Util.Arrays.GetDimensions(a)
End Property

Property Get IsDynamic() As Boolean
  IsDynamic = Util.Arrays.IsDynamic(a)
End Property

Property Get IsEmpty() As Boolean
  IsEmpty = Util.Arrays.IsEmpty(a)
End Property

Property Get min() As Variant
Dim l As Long
Dim minVal As Variant

  If Not Util.Arrays.IsOfSimpleDataTypes(a) Then
    min = Null
    Exit Property
  End If

  minVal = a(LBound(a))
  For l = UBound(a) To LBound(a) Step -1
    If a(l) < minVal Then
      minVal = a(l)
    End If
  Next l

  min = minVal
End Property

Property Get max() As Variant
Dim l As Long
Dim maxVal As Variant

  If Not Util.Arrays.IsOfSimpleDataTypes(a) Then
    max = Null
    Exit Property
  End If

  maxVal = a(LBound(a))
  For l = UBound(a) To LBound(a) Step -1
    If a(l) > maxVal Then
      maxVal = a(l)
    End If
  Next l

  max = maxVal
End Property

Property Get VarsType() As VbVarType
  VarsType = Util.Arrays.GetDataType(a)
End Property

Public Function GetIndexOf(ByVal Value As Variant) As Long
  GetIndexOf = Util.Arrays.GetIndexOf(a, Value)
End Function

Public Function Contains( _
  ByVal Value As Variant, _
  Optional ByVal CompareMethod As VBA.VbCompareMethod = vbTextCompare, _
  Optional ByVal LookAt As XlLookAt = XlLookAt.xlWhole _
) As Boolean
  Contains = Util.Arrays.Contains(a, Value, CompareMethod, LookAt)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Append
' Descr.    : Accepts simple vartypes, objects, collections and arrays and appends it to Me.a
'---------------------------------------------------------------------------------------
' TODO explore dimentions expansion
' http://www.cpearson.com/excel/ArraysAndRanges.aspx
Public Sub Append(ByVal varName As Variant)
  
  If IsArray(varName) Then
    If Not Util.Arrays.IsEmpty(varName) Then
      Select Case Util.Arrays.GetDimensions(varName)
        Case 1
          Dim l As Long
          For l = LBound(varName) To UBound(varName)
            appendValue varName(l)
          Next l
        Case 2
          Dim p As Long
          For p = LBound(varName, 1) To UBound(varName, 1)
            appendValue varName(p, 1)
          Next p
      End Select
    End If
  Else ' VarName not an array
    appendValue varName
  End If
End Sub

Private Sub appendValue(ByRef Value As Variant)
  If IsObject(Value) Then
    If TypeOf Value Is Collection Then
      appendCollection Value
    Else
      Util.Arrays.IncreaseUpperBound a
      Set a(UBound(a)) = Value
    End If
  Else
    Util.Arrays.IncreaseUpperBound a
    a(UBound(a)) = Value
  End If
End Sub

Private Sub appendCollection(ByRef col As Variant)
Dim v As Variant
  For Each v In col
    Append v
  Next v
End Sub

Public Function ToString() As String
  ToString = Util.Arrays.ToString(a)
End Function

Public Sub Reverse()
  Util.Arrays.ReverseInPlace a
End Sub

Public Sub Intersect(ByRef arr As Variant)
  Dim tmp() As Variant

  If IsArray(arr) Then
    tmp = Util.Arrays.GetIntersection(a, arr)
  ElseIf VarType(arr) = vbObject And TypeName(arr) = TypeName(Me) Then
    tmp = Util.Arrays.GetIntersection(a, arr.ArrayVariant)
  End If

  If Util.Arrays.IsAllocated(tmp) Then
    a = tmp
  End If
End Sub

Public Sub SortInPlace(Optional ByVal Descending As Boolean = False)
  Call Util.Arrays.QSortInPlace(inputArray:=a, Descending:=Descending)
End Sub

Public Sub Clear()
  a = Array()
End Sub

Public Function Remove(ByVal Index As Long) As Boolean
  Remove = Util.Arrays.DeleteArrayElement(a, Index, True)
End Function



