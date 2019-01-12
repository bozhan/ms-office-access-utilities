VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Ivanov, Bozhan
' Purpose   : Provides array Utility to extend functionality
'   Handles only up to 2D arrays
'   Some of these Procedures are modified versions of the procedures in modArraySupport
'   By Chip Pearson, chip@cpearson.com, www.cpearson.com
'   For details on these functions, see www.cpearson.com/excel/VBAArrays.htm
'---------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

Public Enum SearchMethodEnum
  default = 1
  BinarySeach = 2
End Enum

'---------------------------------------------------------------------------------------
' Procedure : RemoveDuplicates
' Author    : Ivanov, Bozhan
' Purpose   : returns a sorted array with removed duplicates
'---------------------------------------------------------------------------------------
Public Function RemoveDuplicates( _
  ByVal a As Variant _
) As Variant
  If Not Me.IsAllocated(a) Then
    RemoveDuplicates = a
    Exit Function
  End If
  
  If Me.GetDimensions(a) > 1 Then
    RemoveDuplicates = a
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="The used array has more than 1 Dimentions, which is not supported!"
    Exit Function
  End If
  
  Me.QSortInPlace inputArray:=a, CompareMode:=vbTextCompare
  
  Dim l As Long, res As New ArrayList
  For l = LBound(a) To UBound(a)
    If l = UBound(a) Then 'array end is reached
      res.Append a(l)
      Exit For
    End If
    
    If Util.Strings.CompareValues(a(l), a(l + 1), vbTextCompare) <> 0 Then
      res.Append a(l)
    End If
  Next l
  RemoveDuplicates = res.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDuplicates
' Author    : Ivanov, Bozhan
' Purpose   : returns an array of (unique) dupliate values in an array
'---------------------------------------------------------------------------------------
Public Function GetDuplicates( _
  ByVal a As Variant, _
  Optional ByVal unique As Boolean = True _
) As Variant
  
  If Not IsArray(a) Then
    err.raise Number:=ErrCode.InvalidProcedureCall, _
      Description:="Array expected as argument in procedure UtilArrays:GetDuplicates()!"
  End If

  If Not Me.IsAllocated(a) Then
    GetDuplicates = Array()
    Exit Function
  End If
  
  If Me.GetDimensions(a) > 1 Then
    GetDuplicates = Array()
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="The used array has more than 1 Dimentions, which is not supported in procedure UtilArrays:GetDuplicates()!"
  End If
  
  Me.QSortInPlace inputArray:=a, CompareMode:=vbTextCompare
  
  Dim l As Long, res As New ArrayList
  For l = LBound(a) To UBound(a)
    If l = UBound(a) Then 'array end is reached
      Exit For
    End If
    
    If Util.Strings.CompareValues(a(l), a(l + 1), vbTextCompare) = 0 Then
      res.Append a(l + 1)
    End If
  Next l
  
  If unique Then
    GetDuplicates = Me.RemoveDuplicates(res.Values)
  Else
    GetDuplicates = res.Values
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDimensions
' Descr.    : Returns the number of dimensions of an array.
'             An unallocated dynamic array has 0 dimensions.
'             This condition can also be tested with IsEmpty
'             Loop, increasing the dimension index Ndx, until an error occurs.
'             An error will occur when Ndx exceeds the number of dimension in the array.
'             Return Ndx - 1.
'             Returns -1 dimenstions when non array was input
'---------------------------------------------------------------------------------------
'
Public Function GetDimensions( _
  ByRef arr As Variant _
) As Long

Dim Ndx As Long, res As Long
  GetDimensions = -1
  If Not IsArray(arr) Then Exit Function
  
On Error Resume Next
  Do
    Ndx = Ndx + 1
    res = UBound(arr, Ndx)
  Loop Until err.Number <> 0
  
  GetDimensions = Ndx - 1
  CheckDimentions Ndx - 1
On Error GoTo 0
End Function

Private Sub CheckDimentions( _
  ByVal dimentions As Long _
)
  If dimentions > 2 Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="The used array has more than 2 Dimentions, which is not supported!"
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsDynamic
' Descr.    : Returns TRUE or FALSE indicating whether Arr is dynamic or not.
'             Note that if you attempt to ReDim a static array in the same procedure
'             in which it is declared, you'll get a compiler error.
'---------------------------------------------------------------------------------------
'
Public Function IsDynamic( _
  ByRef arr As Variant _
) As Boolean

  If Not IsArray(arr) Then
    IsDynamic = False
    Exit Function
  End If
    
  If IsEmpty(arr) Then 'isEmpty() checks for allocation as well => dynamic
    IsDynamic = True
    Exit Function
  End If
    
  On Error Resume Next
  err.Clear
  
  Dim LUBound As Long
  LUBound = UBound(arr, 1)
  ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + 1)
  
  Select Case err.Number
    Case ErrCode.NoError ' Successfully increased the UBound
      ReDim Preserve arr(LBound(arr, 1) To LUBound)
      IsDynamic = True
    Case ErrCode.SubscriptOutOfRange ' Arr is a multi-dimensional dynamic
      IsDynamic = True
    Case ErrCode.ArrayIsFixedOrLocked ' Arr is a static single- or multi-dimensional
      IsDynamic = False
    Case Else ' Unexpected error
      IsDynamic = False
  End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsEmpty
' Descr.    : This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'             The VBA IsArray function indicates whether a variable is an array, but it does not
'             distinguish between allocated and unallocated arrays. It will return TRUE for both
'             allocated and unallocated arrays. This function tests whether the array has actually
'             been allocated.
'             This function is really the ReverseInPlace of IsAllocated.
'---------------------------------------------------------------------------------------
'
Public Function IsEmpty( _
  ByRef arr As Variant _
) As Boolean

  If Not IsAllocated(arr) Then
    IsEmpty = True
    Exit Function
  End If
  
  Dim dims As Long, c As Long, r As Long
  dims = GetDimensions(arr)
  IsEmpty = True 'if non empty element is found IsEmpty will become False
  
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      If IsArray(arr(r)) Then
        If Not IsEmpty(arr(r)) Then
          IsEmpty = False
          Exit Function
        End If
      ElseIf VarType(arr(r)) <> vbEmpty Then
        IsEmpty = False
        Exit Function
      End If
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        If IsArray(arr(r, c)) Then
          If Not IsEmpty(arr(r, c)) Then
            IsEmpty = False
            Exit Function
          End If
        ElseIf VarType(arr(r, c)) <> vbEmpty Then
          IsEmpty = False
          Exit Function
        End If
      Next c
    End If
  Next r
End Function
'---------------------------------------------------------------------------------------
' Procedure : IsAllocated
' Descr.    : Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
'             sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
'             been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
'             allocated.
'             The VBA IsArray function indicates whether a variable is an array, but it does not
'             distinguish between allocated and unallocated arrays. It will return TRUE for both
'             allocated and unallocated arrays. This function tests whether the array has actually
'             been allocated.
'---------------------------------------------------------------------------------------
'
Public Function IsAllocated( _
  ByRef arr As Variant _
) As Boolean

  If Not IsArray(arr) Then
    IsAllocated = False
    Exit Function
  End If
    
  Dim n As Long
  On Error Resume Next ' use error handler to determine allocation
  n = UBound(arr, 1)
  
  If (err.Number = 0) Then
    If LBound(arr, 1) <= UBound(arr, 1) Then
      IsAllocated = True
    Else
      IsAllocated = False
    End If
  Else
    IsAllocated = False ' error. unallocated array
  End If
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDataType
' Descr.    : Compares the varTypes of all array elements and if consistent the
'   corresponding vartype is returened if not then vbVariant is returned
'---------------------------------------------------------------------------------------
Public Function GetDataType( _
  ByRef arr As Variant _
) As VbVarType
  
  If Not IsAllocated(arr) Then
    GetDataType = vbNull
    Exit Function
  End If
  
  If IsEmpty(arr) Then
    GetDataType = vbEmpty
    Exit Function
  End If
  
  Dim dims As Long
  dims = GetDimensions(arr)
  
  Dim vType As VbVarType, vTypePre As VbVarType
  Dim r As Long, c As Long
  vType = vbNull
  vTypePre = vbNull
  
  For r = LBound(arr, 1) To UBound(arr, 1)
    vTypePre = vType
    If dims = 1 Then
      If IsArray(arr(r)) Then
        vType = vbArray
      ElseIf VarType(arr(r)) <> vbEmpty Then
        vType = VarType(arr(r))
      End If
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        If IsArray(arr(r, c)) Then
          vType = vbArray
        ElseIf VarType(arr(r, c)) <> vbEmpty Then
          vType = VarType(arr(r, c))
        End If
      Next c
    End If
    
    If vType <> vbNull _
    And vTypePre <> vbNull _
    And vTypePre <> vType Then
      GetDataType = vbVariant
      Exit Function
    End If
  Next r
  
  GetDataType = vType
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsOfDataType
' Purpose   : comapres all elemets's datatypes with the provided one
'   Returns True if all elements are of the provided data type
'   False otherwise.
'---------------------------------------------------------------------------------------
'
Public Function IsOfDataType( _
  ByRef arr As Variant, _
  ByVal vType As VbVarType _
) As Boolean
  
  If Not IsAllocated(arr) Then
    IsOfDataType = False
    Exit Function
  End If
  
  If IsEmpty(arr) And vType = vbEmpty Then
    IsOfDataType = True
    Exit Function
  End If
  
  Dim dims As Long
  dims = GetDimensions(arr)
  
  Dim flag As Boolean
  flag = True
  Dim r As Long, c As Long
  
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      If IsArray(arr(r)) Then
        flag = flag And (vbArray = vType)
      ElseIf VarType(arr(r)) <> vbEmpty Then
        flag = flag And (VarType(arr(r)) = vType)
      End If
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        If IsArray(arr(r, c)) Then
          flag = flag And (vbArray = vType)
        ElseIf VarType(arr(r, c)) <> vbEmpty Then
          flag = flag And (VarType(arr(r, c)) = vType)
        End If
      Next c
    End If
  Next r
  
  IsOfDataType = flag
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDataTypesOfElements
' Descr.    : Returns the datatypes for each element in the array
'---------------------------------------------------------------------------------------
Public Function GetDataTypesOfElements( _
  ByRef arr As Variant _
) As Variant
  
  Dim result As Variant
  result = Array()
    
  If IsEmpty(arr) Then
    GetDataTypesOfElements = result
    Exit Function
  End If
  
  Dim dims As Long
  dims = GetDimensions(arr)
  Dim r As Long, c As Long
  
  If dims = 1 Then
    ReDim result(LBound(arr, 1) To UBound(arr, 1))
  ElseIf dims = 2 Then
    ReDim result(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
  End If
  
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      result(r) = VarType(arr(r))
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        result(r, c) = VarType(arr(r, c))
      Next c
    End If
  Next r
  
  GetDataTypesOfElements = result
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDataTypeNamesOfElements
' Descr.    : Returns the datatype names for each element in the array
'---------------------------------------------------------------------------------------
Public Function GetDataTypeNamesOfElements( _
  ByRef arr As Variant _
) As Variant
  
  Dim result As Variant
  result = Array()
    
  If IsEmpty(arr) Then
    GetDataTypeNamesOfElements = result
    Exit Function
  End If
  
  Dim dims As Long
  dims = GetDimensions(arr)
  Dim r As Long, c As Long
  
  If dims = 1 Then
    ReDim result(LBound(arr, 1) To UBound(arr, 1))
  ElseIf dims = 2 Then
    ReDim result(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
  End If
  
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      result(r) = TypeName(arr(r))
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        result(r, c) = TypeName(arr(r, c))
      Next c
    End If
  Next r
  
  GetDataTypeNamesOfElements = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : IncreaseUpperBound
' Descr.    : allocate a given number of keys in provided dynamic array
'   if the array is not allocated it will be dynamically allocated (0 to 0) or (0 to 0, 0 to 0)
'   depending on the dimention worked on.
'   bound provides the new upper bound value for the provided dimention of the array.
'---------------------------------------------------------------------------------------
'
Public Sub IncreaseUpperBound( _
  ByRef arr As Variant, _
  Optional bound As Long = 1, _
  Optional dimention As Long = 1 _
)

  If bound < 0 Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, Description:="Increase bound value must be >= 0"
  End If
  
  If dimention <= 0 Or dimention > 2 Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, Description:="Dimention value must be > 0 and <= 2"
  End If
  
  If Not IsAllocated(arr) Then
    If dimention = 1 Then
      ReDim arr(0 To 0)
    ElseIf dimention = 2 Then
      ReDim arr(0 To 0, 0 To 0)
    End If
    
    IncreaseUpperBound arr, bound - 1, dimention
    Exit Sub
  End If
    
  If Not IsDynamic(arr) Then
    err.raise Number:=ErrCode.ArrayIsFixedOrLocked, Description:="Provided array is not dynamic!"
  End If
    
  Dim dims As Long
  dims = GetDimensions(arr)
  If dimention > dims Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, Description:="Requested dimention does not exist"
  End If
  
  If dimention = 1 And dims = 1 Then
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + bound)
  ElseIf dimention = 1 And dims = 2 Then
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + bound, LBound(arr, 2) To UBound(arr, 2))
  ElseIf dimention = 2 And dims = 2 Then
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + bound)
  End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsOfSimpleDataTypes
' Descr.    : tests all elements of an array for their data type simplicity
'   returns true if all elements of the array are of a simepl data type
'   returns false otherwise
'---------------------------------------------------------------------------------------
Public Function IsOfSimpleDataTypes( _
  ByRef arr As Variant _
) As Boolean
  
  Dim dims As Long, typeStatus As Boolean
  typeStatus = True
  dims = GetDimensions(arr)
  
  Dim r As Long, c As Long
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      typeStatus = typeStatus And IsSimpleDataType(arr(r))
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        typeStatus = typeStatus And IsSimpleDataType(arr(r, c))
      Next c
    End If
  Next r
  
  IsOfSimpleDataTypes = typeStatus
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsSimpleDataType
' Descr.    : This function returns FALSE if V is any one of the following variable types:
'   vbArray
'   vbDataObject
'   vbObject
'   vbUserDefinedType
'   or if it is an array of any type.
'   TRUE is returened otherwise.
'---------------------------------------------------------------------------------------
Private Function IsSimpleDataType(v As Variant) As Boolean
On Error Resume Next

  If IsArray(v) Or IsObject(v) Then
    IsSimpleDataType = False
    Exit Function
  End If
  
  Select Case VarType(v)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType
      IsSimpleDataType = False
    Case Else
      IsSimpleDataType = True
  End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetNumElements
' Purpose   : Returns the number of elements in the specified dimension (default 1)
'   of the provided array
'   returns 0 when:
'     Arr is not an array
'     Arr is an unallocated array
'     dimension is greater than the number of dimension of Arr
'     dimension is less than 1.
'---------------------------------------------------------------------------------------
'
Public Function GetNumElements( _
  arr As Variant, _
  Optional dimension = 1 _
) As Long

  If IsEmpty(arr) _
  Or dimension < 1 Then
    GetNumElements = 0
    Exit Function
  End If

  Dim dims As Long
  dims = GetDimensions(arr)
  If dimension > dims Then
    GetNumElements = 0
    Exit Function
  End If
  
  GetNumElements = UBound(arr, dimension) - LBound(arr, dimension) + 1
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIntersection
' Descr.    : returns the intersection of 2 arrays
'   Works only with array of simple data types, which are 1D
'---------------------------------------------------------------------------------------
Public Function GetIntersection( _
  ByVal a As Variant, _
  ByVal b As Variant _
) As Variant
Dim i As Long, j As Long
Dim result() As Variant

  result = Array()

  If Not IsAllocated(b) Or Not IsAllocated(a) Then
    'err.Raise Number:=ErrCode.SubscriptOutOfRange, Description:="Parameter is an empty array"
    GetIntersection = Array()
    Exit Function
  End If
  
  If Not IsOfSimpleDataTypes(b) Or Not IsOfSimpleDataTypes(a) Then
    err.raise Number:=ErrCode.TypeMismatch, Description:="Paramter not of simple data type"
  End If
  
  Me.QSortInPlace inputArray:=a, Descending:=True
  Me.QSortInPlace inputArray:=b, Descending:=True
  
  i = LBound(a)
  j = LBound(b)
  
  While i <= UBound(a) And j <= UBound(b)
    If Util.Strings.CompareValues(a(i), b(j), vbTextCompare) = 0 Then
      Me.IncreaseUpperBound result, 1
      result(UBound(result)) = a(i)
      i = i + 1
      j = j + 1
    ElseIf Util.Strings.CompareValues(a(i), b(j), vbTextCompare) = 1 Then
      i = i + 1
    Else
      j = j + 1
    End If
  Wend
  
  GetIntersection = result
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetDifference
' Description : The abosulte relative complement of A with respect to a set B,
'   also termed the difference of sets A and B, written B \ A, is the set of elements in
'   B but not in A, e.g. {1,2,3} \ {2,3,4} = {1}
'   Multiple occurrences of identical elements have no effect on the cardinality of a set or its contents.
'   Duplicate values are considers as the same term
'---------------------------------------------------------------------------------------
Public Function GetDifference( _
  ByVal a As Variant, _
  ByVal b As Variant _
) As Variant
Dim result() As Variant
  result = Array()

  If Not Me.IsAllocated(b) Or Not Me.IsAllocated(a) Then _
    err.raise 9, , "Parameter is an empty array"
  
  If Not Me.IsOfSimpleDataTypes(b) Or Not Me.IsOfSimpleDataTypes(a) Then _
    err.raise 13, , "Paramter not of simple data type"
  
  Me.QSortInPlace inputArray:=a
  Me.QSortInPlace inputArray:=b
  
  Dim i As Long
  For i = LBound(a) To UBound(a)
    If Not Me.Contains(b, a(i)) Then
      Me.IncreaseUpperBound result, 1
      result(UBound(result)) = a(i)
    End If
  Next i
  
  GetDifference = result
End Function


'---------------------------------------------------------------------------------------
' Procedure   : GetDifferenceSymmetric
' Description : returns the symetric difference of 2 arrays (A \ B) Union (B \ A)
'   Works only with array of simple data types
'---------------------------------------------------------------------------------------
Public Function GetDifferenceSymmetric( _
  ByVal a As Variant, _
  ByVal b As Variant _
) As Variant
Dim result() As Variant
  result = Array()

  If Not Me.IsAllocated(b) Or Not Me.IsAllocated(a) Then _
    err.raise 9, , "Parameter is an empty array"
  
  If Not Me.IsOfSimpleDataTypes(b) Or Not Me.IsOfSimpleDataTypes(a) Then _
    err.raise 13, , "Paramter not of simple data type"
  
  Me.QSortInPlace inputArray:=a
  Me.QSortInPlace inputArray:=b

  Dim i As Long
  For i = LBound(a) To UBound(a)
    If Not Me.Contains(b, a(i)) Then
      Me.IncreaseUpperBound result, 1
      result(UBound(result)) = a(i)
    End If
  Next i
  
  For i = LBound(b) To UBound(b)
    If Not Me.Contains(a, b(i)) Then
      Me.IncreaseUpperBound result, 1
      result(UBound(result)) = b(i)
    End If
  Next i
  
  GetDifferenceSymmetric = result
End Function

'---------------------------------------------------------------------------------------
' Procedure   : ConcatenateArrays
' Author      : Pearson Software Consulting Services
' Description :
' ConcatenateArrays
' This function appends ArrayToAppend to the end of ResultArray, increasing the size of ResultArray
' as needed. ResultArray must be a dynamic array, but it need not be allocated. ArrayToAppend
' may be either static or dynamic, and if dynamic it may be unallocted. If ArrayToAppend is
' unallocated, ResultArray is left unchanged.
'
' The data types of ResultArray and ArrayToAppend must be either the same data type or
' compatible numeric types. A compatible numeric type is a type that will not cause a loss of
' precision or cause an overflow. For example, ReturnArray may be Longs, and ArrayToAppend amy
' by Longs or Integers, but not Single or Doubles because information might be lost when
' converting from Double to Long (the decimal portion would be lost). To skip the compatability
' check and allow any variable type in ResultArray and ArrayToAppend, set the NoCompatabilityCheck
' parameter to True. If you do this, be aware that you may loose precision and you may will
' get an overflow error which will cause a result of 0 in that element of ResultArra.
'
' Both ReaultArray and ArrayToAppend must be one-dimensional arrays.
'
'---------------------------------------------------------------------------------------
Public Function ConcatenateArrays(ResultArray As Variant, ArrayToAppend As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean

Dim VTypeResult As VbVarType
Dim Ndx As Long
Dim res As Long
Dim NumElementsToAdd As Long
Dim AppendNdx As Long
Dim VTypeAppend As VbVarType
Dim ResultLB As Long
Dim ResultUB As Long
Dim ResultWasAllocated As Boolean

'''''''''''''''''''''''''''''''''
' Set the default result.
''''''''''''''''''''''''''''''''
ConcatenateArrays = False

'''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    Exit Function
End If
''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is an array.
''''''''''''''''''''''''''''''''''
If IsArray(ArrayToAppend) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are single
' dimensional.
''''''''''''''''''''''''''''''''''
If GetDimensions(ResultArray) > 1 Then
    Exit Function
End If
If GetDimensions(ArrayToAppend) > 1 Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic.
'''''''''''''''''''''''''''''''''''
If IsDynamic(arr:=ResultArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is allocated.
' If ArrayToAppend is not allocated,
' we have nothing to append, so
' exit with a True result.
''''''''''''''''''''''''''''''''''''
If IsAllocated(arr:=ArrayToAppend) = False Then
    ConcatenateArrays = True
    Exit Function
End If


If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure the array are compatible
    ' data types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(ResultArray, ArrayToAppend) = False Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The arrays are not compatible data types.
        '''''''''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
   
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(ResultArray) - vbArray = vbObject Then
        If IsAllocated(ArrayToAppend) = True Then
            For Ndx = LBound(ArrayToAppend) To UBound(ArrayToAppend)
                If IsObject(ArrayToAppend(Ndx)) = False Then
                    Exit Function
                End If
            Next Ndx
        End If
    End If
End If
   
   
'''''''''''''''''''''''''''''''''''''''
' Get the number of elements in
' ArrrayToAppend
'''''''''''''''''''''''''''''''''''''''
NumElementsToAdd = UBound(ArrayToAppend) - LBound(ArrayToAppend) + 1
''''''''''''''''''''''''''''''''''''''''
' Get the bounds for resizing the
' ResultArray. If ResultArray is allocated
' use the LBound and UBound+1. If
' ResultArray is not allocated, use
' the LBound of ArrayToAppend for both
' the LBound and UBound of ResultArray.
''''''''''''''''''''''''''''''''''''''''

If IsAllocated(arr:=ResultArray) = True Then
    ResultLB = LBound(ResultArray)
    ResultUB = UBound(ResultArray)
    ResultWasAllocated = True
    ReDim Preserve ResultArray(ResultLB To ResultUB + NumElementsToAdd)
Else
    ResultUB = UBound(ArrayToAppend)
    ResultWasAllocated = False
    ReDim ResultArray(LBound(ArrayToAppend) To UBound(ArrayToAppend))
End If

''''''''''''''''''''''''''''''''''''''''
' Copy the data from ArrayToAppend to
' ResultArray.
''''''''''''''''''''''''''''''''''''''''
If ResultWasAllocated = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was allocated, we
    ' have to put the data from ArrayToAppend
    ' at the end of the ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''
    AppendNdx = LBound(ArrayToAppend)
    For Ndx = ResultUB + 1 To UBound(ResultArray)
        If IsObject(ArrayToAppend(AppendNdx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        Else
            ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        End If
        AppendNdx = AppendNdx + 1
        If AppendNdx > UBound(ArrayToAppend) Then
            Exit For
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was not allocated, we simply
    ' copy element by element from ArrayToAppend
    ' to ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = LBound(ResultArray) To UBound(ResultArray)
        If IsObject(ArrayToAppend(Ndx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(Ndx)
        Else
            ResultArray(Ndx) = ArrayToAppend(Ndx)
        End If
    Next Ndx

End If
'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
ConcatenateArrays = True

End Function

'---------------------------------------------------------------------------------------
' Procedure : QSortInPlace
' Author    : Chip Pearson, www.cpearson.com, chip@cpearson.com
' Modifier  : Ivanov, Bozhan
' Descr.    : This function sorts the array InputArray in place.
'   It works with either string or numeric data. A part of the array may be sorted
'   by setting the optional LB and UB parameters to the first (LB) and last (UB) element
'   indexes that need sorting.
'   By default, the sort method is NOT case sensitive. Setting the CompareMode argument
'   to vbBinaryCompare (=0) will make the sort method case sensitive.
'   Returns TRUE if the array was successfully sorted and FALSE if error occured.
'   Any occuring error will be dipslayed in a MsgBox.
'   To suppress message boxes, set the NoAlerts parameter to TRUE.
' Modifying Notes: If you modify this code and you call "Exit Procedure",
'   you MUST decrment the RecursionLevel variable. E.g.,
'     If SomethingThatCausesAnExit Then
'       RecursionLevel = RecursionLevel - 1
'       Exit Function
'     End If
'---------------------------------------------------------------------------------------

Public Function QSortInPlace( _
  ByRef inputArray As Variant, _
  Optional ByVal lb As Long = -1&, _
  Optional ByVal ub As Long = -1&, _
  Optional ByVal Descending As Boolean = False, _
  Optional ByVal CompareMode As VbCompareMethod = vbTextCompare _
) As Boolean
    
Dim Temp As Variant
Dim buffer As Variant
Dim curLow As Long
Dim curHigh As Long
Dim curMidpoint As Long
Dim errMsg As String
Dim dimentions As Integer

Static RecursionLevel As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level of recursion
' (the function calling itself). RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure or recursively by itself.
' The variable is decremented when the procedure exits.
' We do the input parameter validation only when RecursionLevel is 1
' (when the function is called by another function, not when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo QSortInPlace_Error

  QSortInPlace = False
  errMsg = vbNullString

  RecursionLevel = RecursionLevel + 1
  
  If RecursionLevel = 1 Then ' performed only on the first recursive call
    If Not IsArray(inputArray) Then
      errMsg = "The InputArray parameter is not an array."
      GoTo QSortInPlace_Error
    End If
      
    If lb < 0 Then lb = LBound(inputArray)
    If ub < 0 Then ub = UBound(inputArray)
    
    dimentions = GetDimensions(inputArray)
    
    If dimentions = 0 Then
      errMsg = "The InputArray is an empty, unallocated array."
      GoTo QSortInPlace_Error
    ElseIf dimentions > 1 Then
      errMsg = "The InputArray is multi-dimensional." & _
        "QSortInPlace works only on single-dimensional arrays."
      GoTo QSortInPlace_Error
    End If
   
    If Not Me.IsOfSimpleDataTypes(inputArray) Then
      errMsg = "InputArray is not an array of simple data types."
      GoTo QSortInPlace_Error
    End If
    
    Select Case lb
      Case Is < LBound(inputArray)
        errMsg = "The LB lower bound parameter is less than the LBound of the InputArray"
        GoTo QSortInPlace_Error
      Case Is > UBound(inputArray)
        errMsg = "The LB lower bound parameter is greater than the UBound of the InputArray"
        GoTo QSortInPlace_Error
      Case Is > ub
        errMsg = "The LB lower bound parameter is greater than the UB upper bound parameter."
        GoTo QSortInPlace_Error
    End Select
    
    Select Case ub
      Case Is > UBound(inputArray)
        errMsg = "The UB upper bound parameter is greater than the upper bound of the InputArray."
        GoTo QSortInPlace_Error
      Case Is < LBound(inputArray)
        errMsg = "The UB upper bound parameter is less than the lower bound of the InputArray."
        GoTo QSortInPlace_Error
      Case Is < lb
        errMsg = "the UB upper bound parameter is less than the LB lower bound parameter."
        GoTo QSortInPlace_Error
    End Select
  
    If ub = lb Then
      QSortInPlace = True
      GoTo QSortInPlace_Exit
    End If
  End If ' RecursionLevel = 1
  
  If Not ((CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare)) Then
    CompareMode = vbTextCompare
  End If
  
  ' Begin the actual sorting process.
  curLow = lb
  curHigh = ub
  
  If lb = 0 Then
      curMidpoint = ((lb + ub) \ 2) + 1
  Else
      curMidpoint = (lb + ub) \ 2 ' note integer division (\) here
  End If
  
  Temp = inputArray(curMidpoint) ' mark current mid point of array
  
  Do While (curLow <= curHigh)
    Do While Util.Strings.CompareValues(v1:=inputArray(curLow), v2:=Temp, CompareMethod:=CompareMode) < 0
      curLow = curLow + 1
      If curLow = ub Then Exit Do
    Loop
    
    Do While Util.Strings.CompareValues(v1:=Temp, v2:=inputArray(curHigh), CompareMethod:=CompareMode) < 0
      curHigh = curHigh - 1
      If curHigh = lb Then Exit Do
    Loop

    If (curLow <= curHigh) Then
      buffer = inputArray(curLow)
      inputArray(curLow) = inputArray(curHigh)
      inputArray(curHigh) = buffer
      curLow = curLow + 1
      curHigh = curHigh - 1
    End If
  Loop
  
  If lb < curHigh Then
    QSortInPlace inputArray:=inputArray, lb:=lb, ub:=curHigh, _
      Descending:=Descending, CompareMode:=CompareMode
  End If
  
  If curLow < ub Then
    QSortInPlace inputArray:=inputArray, lb:=curLow, ub:=ub, _
      Descending:=Descending, CompareMode:=CompareMode
  End If
  
  ' If Descending is True, ReverseInPlace the order of the array, but only if the recursion level is 1
  If Descending And RecursionLevel = 1 Then
    ReverseInPlace inputArray, lb, ub
  End If
  
  QSortInPlace = True
  
QSortInPlace_Exit:
  RecursionLevel = RecursionLevel - 1
Exit Function
  
QSortInPlace_Error:
  If errMsg <> vbNullString Then
    MsgBox "Error occured in procedure QSortInPlace of module ArrayUtil." & vbCrLf & errMsg _
      , vbExclamation, "Error occured..."
  End If
  QSortInPlace = False
  GoTo QSortInPlace_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : QSortInPlaceOnRefValues
' Created By: Ivanov, Bozhan  <>
' Descr.    : Function based on the QSortInPlaceOnRefValues Function. In this Function the values
'   comapred for the sorting are referenced in a collection. Meaning that the actual array
'   values are used to reference a third value in the collection, which in turn is used as
'   the sorting comapre value when calling the CompareValues function
'   The collection is reffered to from the key string values where the corresponding
'   values are compared as mid, low and high
'---------------------------------------------------------------------------------------
'
Public Function QSortInPlaceOnRefValues( _
    ByRef inputArray As Variant, _
    ByRef RefValuesCollection As Collection, _
    Optional ByVal lb As Long = -1&, _
    Optional ByVal ub As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare _
) As Boolean

Dim Temp As Variant
Dim buffer As Variant
Dim curLow As Long
Dim curHigh As Long
Dim curMidpoint As Long
Dim errMsg As String
Dim dimentions As Integer

Static RecursionLevel As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level
' of recursion  (the function calling itself).
' RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure
' or recursively by itself. The variable is decremented
' when the procedure exits. We do the input parameter
' validation only when RecursionLevel is 1 (when
' the function is called by another function, not
' when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo QSortInPlaceOnRefValues_Error

  QSortInPlaceOnRefValues = False
  errMsg = vbNullString

  RecursionLevel = RecursionLevel + 1
  
  If RecursionLevel = 1 Then ' performed only on the first recursive call
    If Not IsArray(inputArray) Then
      errMsg = "The InputArray parameter is not an array."
      GoTo QSortInPlaceOnRefValues_Error
    End If
      
    If lb < 0 Then lb = LBound(inputArray)
    If ub < 0 Then ub = UBound(inputArray)
    
    dimentions = Me.GetDimensions(inputArray)
    
    If dimentions = 0 Then
      errMsg = "The InputArray is an empty, unallocated array."
      GoTo QSortInPlaceOnRefValues_Error
    ElseIf dimentions > 1 Then
      errMsg = "The InputArray is multi-dimensional." & _
        "QSortInPlaceOnRefValues works only on single-dimensional arrays."
      GoTo QSortInPlaceOnRefValues_Error
    End If
    
    If Not Me.IsOfSimpleDataTypes(inputArray) Then
      errMsg = "InputArray is not an array of simple data types."
      GoTo QSortInPlaceOnRefValues_Error
    End If
    
    Select Case lb
      Case Is < LBound(inputArray)
        errMsg = "The LB lower bound parameter is less than the LBound of the InputArray"
        GoTo QSortInPlaceOnRefValues_Error
      Case Is > UBound(inputArray)
        errMsg = "The LB lower bound parameter is greater than the UBound of the InputArray"
        GoTo QSortInPlaceOnRefValues_Error
      Case Is > ub
        errMsg = "The LB lower bound parameter is greater than the UB upper bound parameter."
        GoTo QSortInPlaceOnRefValues_Error
    End Select
    
    Select Case ub
      Case Is > UBound(inputArray)
        errMsg = "The UB upper bound parameter is greater than the upper bound of the InputArray."
        GoTo QSortInPlaceOnRefValues_Error
      Case Is < LBound(inputArray)
        errMsg = "The UB upper bound parameter is less than the lower bound of the InputArray."
        GoTo QSortInPlaceOnRefValues_Error
      Case Is < lb
        errMsg = "the UB upper bound parameter is less than the LB lower bound parameter."
        GoTo QSortInPlaceOnRefValues_Error
    End Select
  
    If ub = lb Then
      QSortInPlaceOnRefValues = True
      GoTo QSortInPlaceOnRefValues_Exit
    End If
  End If ' RecursionLevel = 1
  
  If Not ((CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare)) Then
    CompareMode = vbTextCompare
  End If
  
  ' Begin the actual sorting process.
  curLow = lb
  curHigh = ub
  
  If lb = 0 Then
    curMidpoint = ((lb + ub) \ 2) + 1
  Else
    curMidpoint = (lb + ub) \ 2 ' note integer division (\) here
  End If
  
  Temp = RefValuesCollection(inputArray(curMidpoint))
  
  Do While (curLow <= curHigh)
    Do While Util.Strings.CompareValues(v1:=RefValuesCollection(inputArray(curLow)), v2:=Temp, CompareMethod:=CompareMode) < 0
      curLow = curLow + 1
      If curLow = ub Then Exit Do
    Loop
    
    Do While Util.Strings.CompareValues(v1:=Temp, v2:=RefValuesCollection(inputArray(curHigh)), CompareMethod:=CompareMode) < 0
      curHigh = curHigh - 1
      If curHigh = lb Then Exit Do
    Loop
    
    If (curLow <= curHigh) Then
      buffer = inputArray(curLow)
      inputArray(curLow) = inputArray(curHigh)
      inputArray(curHigh) = buffer
      curLow = curLow + 1
      curHigh = curHigh - 1
    End If
  Loop
  
  If lb < curHigh Then
    QSortInPlaceOnRefValues inputArray:=inputArray, RefValuesCollection:=RefValuesCollection, _
      lb:=lb, ub:=curHigh, _
      Descending:=Descending, CompareMode:=CompareMode
  End If
  
  If curLow < ub Then
    QSortInPlaceOnRefValues inputArray:=inputArray, RefValuesCollection:=RefValuesCollection, _
      lb:=curLow, ub:=ub, _
      Descending:=Descending, CompareMode:=CompareMode
  End If
  
  ' If Descending is True, ReverseInPlace the order of the array, but only if the recursion level is 1
  If Descending And RecursionLevel = 1 Then
    ReverseInPlace inputArray, lb, ub
  End If
  
  QSortInPlaceOnRefValues = True
  
QSortInPlaceOnRefValues_Exit:
  RecursionLevel = RecursionLevel - 1
Exit Function
  
QSortInPlaceOnRefValues_Error:
  If errMsg <> vbNullString Then
    MsgBox "Error occured in procedure QSortInPlaceOnRefValues of module ArrayUtil." & vbCrLf & errMsg _
      , vbExclamation, "Error occured..."
  End If
  GoTo QSortInPlaceOnRefValues_Exit
End Function




'TODO Create Flip Matrix for 2D (flip horizontaly, vertically, diag left, diag right)
'---------------------------------------------------------------------------------------
' Procedure : ReverseInPlace
' Descr.    : Reverses the order of elements in a.
'   To ReverseInPlace the entire array, omit or set to less than 0 the LB and UB parameters.
'   To ReverseInPlace only part of tbe array, set LB and/or
'   UB to the LBound and UBound of the sub array to be reversed.
'   Procedure works only in 1D
'---------------------------------------------------------------------------------------
Public Sub ReverseInPlace( _
  ByRef a As Variant, _
  Optional lb As Long = -1, _
  Optional ub As Long = -1 _
)
  If Not IsArray(a) Then
    err.raise Number:=ErrCode.TypeMismatch, Description:="The a parameter is not an array."
  End If
  
  If Not IsAllocated(a) Then
    Exit Sub
  End If
  
  Dim dims As Integer
  dims = GetDimensions(a)
  If dims > 1 Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, Description:="Provided array is multi-dimensional." & _
      "ReverseInPlace works only on single-dimensional arrays."
  End If
  
  If lb < 0 Then lb = LBound(a)
  If ub < 0 Then ub = UBound(a)
  
  Dim l As Long, Temp As Variant
  'using the division factor of 2 (x\2) to find mid point of array
  For l = lb To (lb + ((ub - lb - 1) \ 2))
    If IsObject(a(l)) Then
      Set Temp = a(l)
      Set a(l) = a(ub - (l - lb))
      Set a(ub - (l - lb)) = Temp
    Else
      Temp = a(l)
      a(l) = a(ub - (l - lb))
      a(ub - (l - lb)) = Temp
    End If
  Next l
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetReversed
' Descr.    : Returns a reversed order version of elements in a.
'   To ReverseInPlace the entire array, omit or set to less than 0 the LB and UB parameters.
'   To ReverseInPlace only part of tbe array, set LB and/or
'   UB to the LBound and UBound of the sub array to be reversed.
'---------------------------------------------------------------------------------------
Public Function GetReversed( _
  ByVal a As Variant, _
  Optional lb As Long = -1, Optional ub As Long = -1 _
) As Variant
  
  ReverseInPlace a, lb, ub
  GetReversed = a
End Function

'---------------------------------------------------------------------------------------
' Procedure : areDataTypesCompatible
' Descr.    : This function determines if v2 is compatiable with v1. If the two
'             data types are the same, they are compatible. If the value of v2 can
'             be stored in v1 with no loss of precision or an overflow, they are compatible.
'             For example, if v1 is a Long and v2 is an Integer, they are compatible
'             because an integer can be stored in a Long with no loss of information. If v1
'             is a Long and v2 is a Double, they are not compatible because information
'             will be lost converting from a Double to a Long (the decimal portion will be lost).
'---------------------------------------------------------------------------------------
'
Public Function AreDataTypesCompatible( _
  v1 As Variant, _
  v2 As Variant _
) As Boolean
Dim v1Type As VbVarType
Dim v2Type As VbVarType

  If Not IsAllocated(v1) _
  And Not IsAllocated(v2) Then
    AreDataTypesCompatible = True
    Exit Function
  End If
  
  If Not IsAllocated(v1) _
  Xor Not IsAllocated(v2) Then
    AreDataTypesCompatible = False
    Exit Function
  End If
  
  AreDataTypesCompatible = (GetDataType(v1) = GetDataType(v2))
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIndexOf
' Descr.    : returns the index of the element if found in array and -1 otherwise
'---------------------------------------------------------------------------------------
Public Function GetIndexOf( _
  ByRef a As Variant, _
  ByVal Value As Variant, _
  Optional ByVal CompareMethod As VBA.VbCompareMethod = vbTextCompare, _
  Optional ByVal LookAt As XlLookAt = XlLookAt.xlWhole, _
  Optional ByVal searchMethod As SearchMethodEnum = SearchMethodEnum.default _
) As Long
  Select Case searchMethod
  Case SearchMethodEnum.default
    GetIndexOf = GetIndexOf_DefaultSearch(a, Value, CompareMethod, LookAt)
  Case SearchMethodEnum.BinarySeach
    GetIndexOf = GetIndexOf_BinarySearch(a, Value)
  End Select
End Function

'TODO refactor into smaller functions Single responsibility principle
'---------------------------------------------------------------------------------------
' Procedure : GetIndexOf_DefaultSearch
' Descr.    : returns the index of the element if found in array and -1 otherwise
'---------------------------------------------------------------------------------------
Public Function GetIndexOf_DefaultSearch( _
  ByRef a As Variant, _
  ByVal Value As Variant, _
  Optional ByVal CompareMethod As VBA.VbCompareMethod = vbTextCompare, _
  Optional ByVal LookAt As XlLookAt = XlLookAt.xlWhole _
) As Long
On Error GoTo GetIndexOf_Error

  GetIndexOf_DefaultSearch = -1
  If Not Me.IsAllocated(a) Or Me.IsEmpty(a) Then Exit Function
    
  Dim idx As Long
  If Me.GetDimensions(a) = 1 Then
    For idx = LBound(a) To UBound(a)
      If LookAt = xlPart Then
        If InStr(1, CStr(a(idx)), CStr(Value), CompareMethod) > 0 Then
          GetIndexOf_DefaultSearch = idx
          Exit Function
        End If
      ElseIf LookAt = xlWhole Then
        If StrComp(CStr(a(idx)), CStr(Value), CompareMethod) = 0 Then
          GetIndexOf_DefaultSearch = idx
          Exit Function
        End If
      End If
    Next idx
  Else
    For idx = LBound(a, 1) To UBound(a, 1)
      If LookAt = xlPart Then
        If InStr(1, CStr(a(idx, 1)), CStr(Value), CompareMethod) > 0 Then
          GetIndexOf_DefaultSearch = idx
          Exit Function
        End If
      ElseIf LookAt = xlWhole Then
        If StrComp(CStr(a(idx, 1)), CStr(Value), CompareMethod) = 0 Then
          GetIndexOf_DefaultSearch = idx
          Exit Function
        End If
      End If
    Next idx
  End If

GetIndexOf_Exit:
On Error Resume Next
Exit Function

GetIndexOf_Error:
  MsgBox "Error " & err.Number & _
  " (" & err.Description & ") in procedure GetIndexOf of module ArrayUtil" & vbLf & _
    InfoErrMsg()
Resume GetIndexOf_Exit
End Function

'TODO adjust for a new parameter xlLookAt for xlpart and xlwhole
'---------------------------------------------------------------------------------------
' Procedure : GetIndexOf_BinarySearch
' Descr.    : applies a binary search to a sorted array
'   If value is found returns the index, else returns -1
'   Pre: sortedArray is a sortet array in asc order preferably.
'   The array order will not be analyzed in this function
'---------------------------------------------------------------------------------------
Public Function GetIndexOf_BinarySearch( _
  ByRef sortedArray As Variant, _
  ByVal Value As Variant _
) As Long
Dim min As Long, max As Long, mid As Long, comp As Long
Dim inv As Boolean
Dim d As Integer
On Error GoTo GetIndexOf_BinarySearch_Error

  GetIndexOf_BinarySearch = -1
  
  If Not Me.IsAllocated(sortedArray) Or Me.IsEmpty(sortedArray) Then
    err.raise Number:=ErrCode.SubscriptOutOfRange, _
      Description:="No array found!"
  End If
  
  d = Me.GetDimensions(sortedArray)
  min = LBound(sortedArray, 1)
  max = UBound(sortedArray, 1)
  
  If d = 1 Then
    inv = (StrComp(CStr(sortedArray(min)), CStr(sortedArray(max)), vbTextCompare) = 1) '(sortedArray(min) > sortedArray(max))
  Else
    inv = (StrComp(CStr(sortedArray(min, 1)), CStr(sortedArray(max, 1)), vbTextCompare) = 1) '(sortedArray(min) > sortedArray(max))
  End If
   
  Do
    mid = (min + max) \ 2
    
    If d = 1 Then
      comp = StrComp(CStr(sortedArray(mid)), CStr(Value), vbTextCompare)
    Else
      comp = StrComp(CStr(sortedArray(mid, 1)), CStr(Value), vbTextCompare)
    End If
    
    If comp = 0 Then
      GetIndexOf_BinarySearch = mid
      Exit Do
    ElseIf ((comp = -1) Xor inv) Then 'sortedArray(mid) < value
      min = mid + 1
    Else
      max = mid - 1
    End If
  Loop Until min > max

GetIndexOf_BinarySearch_Exit:
Exit Function

GetIndexOf_BinarySearch_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetIndexOf_BinarySearch of module ArrayUtil" & vbLf & _
    InfoErrMsg()
Resume GetIndexOf_BinarySearch_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Contains
' Descr.    : returns true if the value is contained in the array
'---------------------------------------------------------------------------------------
Public Function Contains( _
  ByRef a As Variant, _
  ByVal val As Variant, _
  Optional ByVal CompareMethod As VBA.VbCompareMethod = vbTextCompare, _
  Optional ByVal LookAt As XlLookAt = XlLookAt.xlWhole _
) As Boolean
  If LookAt = xlPart Then
    Contains = (GetIndexOf(a, val, CompareMethod, LookAt, default) > -1)
  ElseIf LookAt = xlWhole Then
    Contains = False
    Dim i As Long
    For i = LBound(a) To UBound(a)
      If StrComp(a(i), val, CompareMethod) = 0 Then
        Contains = True
        Exit Function
      End If
    Next i
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ToString
' Descr.    : converts array to a CSV string, object will be displayed as thier types
'---------------------------------------------------------------------------------------
'
Public Function ToString( _
  ByRef arr As Variant _
) As String
Dim result As String, sep As String
Dim idx As Long

  result = vbNullString
  sep = ";"

  If IsEmpty(arr) Then
    result = "{" & vbNullString & "}"
    GoTo ToString_Exit
  End If

  Dim dims As Long
  dims = GetDimensions(arr)
  Dim r As Long, c As Long
    
  For r = LBound(arr, 1) To UBound(arr, 1)
    If dims = 1 Then
      If IsArray(arr(r)) Then
        result = result & sep & "{" & ToString(arr(r)) & "}"
      ElseIf IsObject(arr(r)) Or VarType(arr(r)) = vbVariant Then
        result = result & sep & "{" & TypeName(arr(r)) & "}"
      Else
        result = result & sep & CStr(Nz(arr(r), "Null"))
      End If
      result = Util.Strings.RemoveExcessSeperators(result, sep, False)
    ElseIf dims = 2 Then
      For c = LBound(arr, 2) To UBound(arr, 2)
        If IsArray(arr(r, c)) Then
          result = result & sep & "{" & ToString(arr(r, c)) & "}"
        ElseIf IsObject(arr(r, c)) Or VarType(arr(r, c)) = vbVariant Then
          result = result & sep & "{" & TypeName(arr(r, c)) & "}"
        Else
          result = result & sep & CStr(Nz(arr(r, c), "Null"))
        End If
        result = Util.Strings.RemoveExcessSeperators(result, sep, False)
      Next c
      result = "[" & result & "]"
    End If
  Next r
  
ToString_Exit:
  ToString = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetSplitOnIndex
' Descr.    : Splits an array into 2 arrays contained in an array (0 To 1)
'   The return(0) contains all elements up to and including the provided index,
'   the return(1) part contains the rest of the array elements
'---------------------------------------------------------------------------------------
Public Function GetSplitOnIndex( _
  ByRef a As Variant, _
  ByVal Index As Long _
) As Variant
Dim result(0 To 1) As Variant
Dim part0 As ArrayList
Dim part1 As ArrayList
Dim l As Long
On Error GoTo GetSplitOnIndex_Error

  If IsEmpty(a) Then
    GetSplitOnIndex = a
    Exit Function
  End If
  
  If Index > UBound(a) Or Index < LBound(a) Then
    GetSplitOnIndex = a
    Exit Function
  End If
  
  Set part0 = New ArrayList
  Set part1 = New ArrayList
  
  For l = LBound(a) To Index
    part0.Append a(l)
  Next l
  
  For l = Index + 1 To UBound(a)
    part1.Append a(l)
  Next l
  
  result(0) = part0.Values
  result(1) = part1.Values
  
  GetSplitOnIndex = result
  
GetSplitOnIndex_Exit:
On Error Resume Next
  Set part0 = Nothing
  Set part1 = Nothing
Exit Function

GetSplitOnIndex_Error:
  MsgBox "Error " & err.Number & _
  " (" & err.Description & ") in procedure GetSplitOnIndex of module ArrayUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume GetSplitOnIndex_Exit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CompareArrays
' compares two arrays, Array1 and Array2, element by element, and returns the results of
' the comparisons as a new array. Each element of the result array will have values -1, 0, or +1.
' -1: the element in Array1 was less than the corresponding element in Array2
'  0: the elements are equal
' +1: the element in Array1 is greater than Array2
' Array1 and Array2 must be allocated single-dimensional arrays
'
' An error will occur if Array1 or Array2 contains an Object or User Defined Type.
'
' When comparing elements, the procedure does the following:
' If both elements are numeric data types, they are compared arithmetically.
' If one element is a numeric data type and the other is a string and that string is numeric,
' then both elements are converted to Doubles and compared arithmetically.
' If the string is not numeric, both elements are converted to strings and compared using StrComp,
' with the compare mode set by CompareMode.
' If both elements are numeric strings, they are converted to Doubles and compared arithmetically.
' If either element is not a numeric string, the elements are converted and compared with StrComp.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CompareArrays( _
  ByRef Array1 As Variant, _
  ByRef Array2 As Variant, _
  Optional ByVal CompareMode As VbCompareMethod = vbTextCompare _
) As Variant
  CompareArrays = Array()
  
  ' Ensure arrays are not empty and allocated
  If IsEmpty(Array1) _
  Or IsEmpty(Array2) Then
    Exit Function
  End If
  
  ' Ensure we work with 1D arrays
  If GetDimensions(Array1) > 1 _
  Or GetDimensions(Array2) > 1 Then
    Exit Function
  End If
  
  ' Ensure the arrays are the same size.
  If (UBound(Array1) - LBound(Array1)) <> (UBound(Array2) - LBound(Array2)) Then
      Exit Function
  End If
  
  If ContainsObjects(Array1) Then Exit Function
  If ContainsObjects(Array2) Then Exit Function
  
  'allocate result
  Dim ResultArray As Variant
  ReDim ResultArray(LBound(Array1) To UBound(Array1))
  
  Dim idx1 As Long
  Dim idx2 As Long
  Dim resIdx As Long
  
  idx1 = LBound(Array1)
  idx2 = LBound(Array2)
  resIdx = LBound(Array1)
  
  Dim Done As Boolean
  Done = False
  
  Do Until Done = True
    If IsArray(Array1(idx1)) _
    And IsArray(Array2(idx2)) Then
      ResultArray(resIdx) = CompareArrays(Array1(idx1), Array2(idx2), CompareMode)
    Else
      ResultArray(resIdx) = Util.Strings.CompareValues(Array1(idx1), Array2(idx2), CompareMode)
    End If
    
    resIdx = resIdx + 1
    idx1 = idx1 + 1
    idx2 = idx2 + 1
    
    ' If idx1 is greater than UBound(Array1) end of the road
    If idx1 > UBound(Array1) Then
      Done = True
    End If
  Loop
  
  CompareArrays = ResultArray
End Function

'---------------------------------------------------------------------------------------
' Procedure : AreArraysEqual
' Purpose   : returns true if all elements of array1 and array2 compared 1 by 1 are the equal
'---------------------------------------------------------------------------------------
'
Public Function AreArraysEqual( _
  ByRef Array1 As Variant, _
  ByRef Array2 As Variant, _
  Optional ByVal CompareMode As VbCompareMethod = vbTextCompare _
) As Boolean
  
  Dim res As Variant
  res = CompareArrays(Array1, Array2, CompareMode)
  Dim ans As Boolean
  ans = True
  Dim l As Long, k As Long
  For l = LBound(res) To UBound(res)
    If IsArray(res(l)) Then
      For k = LBound(res(l)) To UBound(res(l))
        ans = ans And (res(l)(k) = 0)
      Next k
    Else
      ans = ans And (res(l) = 0)
    End If
  Next l
  
  AreArraysEqual = ans
End Function

'---------------------------------------------------------------------------------------
' Procedure : ContainsObjects
' Purpose   : searches for object or user defined elements in an array
'---------------------------------------------------------------------------------------
'
Public Function ContainsObjects( _
  ByRef arr As Variant _
) As Boolean
Dim types As Variant
  
  ContainsObjects = False
  If IsEmpty(arr) Then Exit Function
  
  types = GetDataTypesOfElements(arr)
  
  If Contains(types, 9) _
  Or Contains(types, 36) Then
    ContainsObjects = True
    Exit Function
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ContainsArrays
' Purpose   : searches for object or user defined elements in an array
'---------------------------------------------------------------------------------------
'
Public Function ContainsArrays( _
  ByRef arr As Variant _
) As Boolean
Dim types As Variant
  
  ContainsArrays = False
  If IsEmpty(arr) Then Exit Function
  
  types = GetDataTypesOfElements(arr)
    
  Dim i As Long
  For i = LBound(types) To UBound(types)
    If types(i) >= vbArray Then
      ContainsArrays = True
      Exit Function
    End If
  Next i
End Function

'TODO CONTINUE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ConcatenateArrays
' This function appends ArrayToAppend to the end of ResultArray, increasing the size of ResultArray
' as needed. ResultArray must be a dynamic array, but it need not be allocated. ArrayToAppend
' may be either static or dynamic, and if dynamic it may be unallocted. If ArrayToAppend is
' unallocated, ResultArray is left unchanged.
'
' The data types of ResultArray and ArrayToAppend must be either the same data type or
' compatible numeric types. A compatible numeric type is a type that will not cause a loss of
' precision or cause an overflow. For example, ReturnArray may be Longs, and ArrayToAppend amy
' by Longs or Integers, but not Single or Doubles because information might be lost when
' converting from Double to Long (the decimal portion would be lost). To skip the compatability
' check and allow any variable type in ResultArray and ArrayToAppend, set the NoCompatabilityCheck
' parameter to True. If you do this, be aware that you may loose precision and you may will
' get an overflow error which will cause a result of 0 in that element of ResultArra.
'
' Both ReaultArray and ArrayToAppend must be one-dimensional arrays.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConcatenateArrays( _
  ByRef Array1 As Variant, _
  ByRef Array2 As Variant _
) As Variant
  Dim res As ArrayList
  Set res = New ArrayList
  res.Append Array1
  res.Append Array2
  GetConcatenateArrays = res.Values
End Function


'----------------------------------------------------------------------------------------------------
'cPearson 's untouched precedures
'####################################################################################################


Public Function CopyArray(DestinationArray As Variant, SourceArray As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyArray
' This function copies the contents of SourceArray to the DestinationaArray. Both SourceArray
' and DestinationArray may be either static or dynamic and either or both may be unallocated.
'
' If DestinationArray is dynamic, it is resized to match SourceArray. The LBound and UBound
' of DestinationArray will be the same as SourceArray, and all elements of SourceArray will
' be copied to DestinationArray.
'
' If DestinationArray is static and has more elements than SourceArray, all of SourceArray
' is copied to DestinationArray and the right-most elements of DestinationArray are left
' intact.
'
' If DestinationArray is static and has fewer elements that SourceArray, only the left-most
' elements of SourceArray are copied to fill out DestinationArray.
'
' If SourceArray is an unallocated array, DestinationArray remains unchanged and the procedure
' terminates.
'
' If both SourceArray and DestinationArray are unallocated, no changes are made to either array
' and the procedure terminates.
'
' SourceArray may contain any type of data, including Objects and Objects that are Nothing
' (the procedure does not support arrays of User Defined Types since these cannot be coerced
' to Variants -- use classes instead of types).
'
' The function tests to ensure that the data types of the arrays are the same or are compatible.
' See the function AreDataTypesCompatible for information about compatible data types. To skip
' this compability checking, set the NoCompatabilityCheck parameter to True. Note that you may
' lose information during data conversion (e.g., losing decimal places when converting a Double
' to a Long) or you may get an overflow (storing a Long in an Integer) which will result in that
' element in DestinationArray having a value of 0.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim VTypeSource As VbVarType
Dim VTypeDest As VbVarType
Dim SNdx As Long
Dim DNdx As Long


'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
CopyArray = False

''''''''''''''''''''''''''''''''''
' Ensure both DestinationArray and
' SourceArray are arrays.
''''''''''''''''''''''''''''''''''
If IsArray(DestinationArray) = False Then
    Exit Function
End If
If IsArray(SourceArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure DestinationArray and
' SourceArray are single-dimensional.
' 0 indicates an unallocated array,
' which is allowed.
'''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=SourceArray) > 1 Then
    Exit Function
End If
If GetDimensions(arr:=DestinationArray) > 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' If SourceArray is not allocated,
' leave DestinationArray intact and
' return a result of True.
''''''''''''''''''''''''''''''''''''
If IsAllocated(arr:=SourceArray) = False Then
    CopyArray = True
    Exit Function
End If

If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure both arrays are the same
    ' type or compatible data types. See
    ' the function AreDataTypesCompatible
    ' for information about compatible
    ' types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(DestinationArray, SourceArray) = False Then
        CopyArray = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(DestinationArray) - vbArray = vbObject Then
        If IsAllocated(SourceArray) = True Then
            For SNdx = LBound(SourceArray) To UBound(SourceArray)
                If IsObject(SourceArray(SNdx)) = False Then
                    Exit Function
                End If
            Next SNdx
        End If
    End If
End If

If IsAllocated(arr:=DestinationArray) = True Then
    If IsAllocated(arr:=SourceArray) = True Then
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both arrays are allocated, copy from
        ' SourceArray to DestinationArray. If
        ' SourceArray is smaller that DesetinationArray,
        ' the right-most elements of DestinationArray
        ' are left unchanged. If SourceArray is larger
        ' than DestinationArray, the right most elements
        ' of SourceArray are not copied.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        DNdx = LBound(DestinationArray)
        On Error Resume Next
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(DNdx) = SourceArray(DNdx)
            Else
                DestinationArray(DNdx) = SourceArray(DNdx)
            End If
            DNdx = DNdx + 1
            If DNdx > UBound(DestinationArray) Then
                Exit For
            End If
        Next SNdx
        On Error GoTo 0
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' If SourceArray is not allocated, so we have
        ' nothing to copy. Exit with a result
        ' of True. Leave DestinationArray intact.
        '''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
        Exit Function
    End If

Else
    If IsAllocated(arr:=SourceArray) = True Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If Destination array is not allocated and
        ' SourceArray is allocated, Redim DestinationArray
        ' to the same size as SourceArray and copy
        ' the elements from SourceArray to DestinationArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        ReDim DestinationArray(LBound(SourceArray) To UBound(SourceArray))
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(SNdx) = SourceArray(SNdx)
            Else
                DestinationArray(SNdx) = SourceArray(SNdx)
            End If
        Next SNdx
        On Error GoTo 0
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both SourceArray and DestinationArray are
        ' unallocated, we have nothing to copy (this condition
        ' is actually detected above, but included here
        ' for Consistency), so get out with a result of True.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
        Exit Function
    End If
End If

'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
CopyArray = True

End Function



Public Function CopyArraySubSetToArray(inputArray As Variant, ResultArray As Variant, _
    FirstElementToCopy As Long, LastElementToCopy As Long, DestinationElement As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyArraySubSetToArray
' This function copies elements of InputArray to ResultArray. It takes the elements
' from FirstElementToCopy to LastElementToCopy (inclusive) from InputArray and
' copies them to ResultArray, starting at DestinationElement. Existing data in
' ResultArray will be overwrittten. If ResultArray is a dynamic array, it will
' be resized if needed. If ResultArray is a static array and it is not large
' enough to copy all the elements, no elements are copied and the function
' returns False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim SrcNdx As Long
Dim DestNdx As Long
Dim NumElementsToCopy As Long

'''''''''''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''''''''''
CopyArraySubSetToArray = False

''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray and ResultArray are
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    Exit Function
End If
If IsArray(ResultArray) = False Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=inputArray) <> 1 Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure ResultArray is unallocated or
' single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=ResultArray) > 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure the bounds and indexes are valid.
''''''''''''''''''''''''''''''''''''''''''''
If FirstElementToCopy < LBound(inputArray) Then
    Exit Function
End If
If LastElementToCopy > UBound(inputArray) Then
   Exit Function
End If
If FirstElementToCopy > LastElementToCopy Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Calc the number of elements we'll copy
' from InputArray to ResultArray.
'''''''''''''''''''''''''''''''''''''''''
NumElementsToCopy = LastElementToCopy - FirstElementToCopy + 1

If IsDynamic(arr:=ResultArray) = False Then
    If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is static and can't be resized.
        ' There is not enough room in the array to
        ' copy all the data.
        '''''''''''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
Else
    ''''''''''''''''''''''''''''''''''''''''''''
    ' ResultArray is dynamic and can be resized.
    ' Test whether we need to resize the array,
    ' and resize it if required.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsEmpty(arr:=ResultArray) = True Then
        '''''''''''''''''''''''''''''''''''''''
        ' ResultArray is unallocated. Resize it
        ' to DestinationElement + NumElementsToCopy - 1.
        ' This provides empty elements to the left
        ' of the DestinationElement and room to
        ' copy NumElementsToCopy.
        '''''''''''''''''''''''''''''''''''''''''
        ReDim ResultArray(1 To DestinationElement + NumElementsToCopy - 1)
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is allocated. If there isn't room
        ' enough in ResultArray to hold NumElementsToCopy
        ' starting at DestinationElement, we need to
        ' resize the array.
        '''''''''''''''''''''''''''''''''''''''''''''''''
        If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
            If DestinationElement + NumElementsToCopy > UBound(ResultArray) Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the ResultArray.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If NumElementsToCopy + DestinationElement > UBound(ResultArray) Then
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + DestinationElement - 1)
                Else
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy)
                End If
            Else
                ''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the array to hold NumElementsToCopy
                ' starting at DestinationElement.
                ''''''''''''''''''''''''''''''''''''''''''''
                ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy - DestinationElement + 2)
            End If
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' The ResultArray is large enough to hold
            ' NumberOfElementToCopy starting at DestinationElement.
            ' No need to resize the array.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Copy the elements from InputArray to ResultArray
' Note that there is no type compatibility checking
' when copying the elements.
'''''''''''''''''''''''''''''''''''''''''''''''''''
DestNdx = DestinationElement
For SrcNdx = FirstElementToCopy To LastElementToCopy
    If IsObject(inputArray(SrcNdx)) = True Then
        Set ResultArray(DestNdx) = inputArray(DestNdx)
    Else
        On Error Resume Next
        ResultArray(DestNdx) = inputArray(SrcNdx)
        On Error GoTo 0
    End If
    DestNdx = DestNdx + 1
Next SrcNdx

CopyArraySubSetToArray = True

End Function



Public Function CopyNonNothingObjectsToArray( _
  ByRef SourceArray As Variant, _
  ByRef ResultArray As Variant _
) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyNonNothingObjectsToArray
' This function copies all objects that are not Nothing from SourceArray
' to ResultArray. ResultArray MUST be a dynamic array of type Object or Variant.
' E.g.,
'       Dim ResultArray() As Object ' Or
'       Dim ResultArray() as Variant
'
' ResultArray will be Erased and then resized to hold the non-Nothing elements
' from SourceArray. The LBound of ResultArray will be the same as the LBound
' of SourceArray, regardless of what its LBound was prior to calling this
' procedure.
'
' This function returns True if the operation was successful or False if an
' an error occurs. If an error occurs, a message box is displayed indicating
' the error. To suppress the message boxes, set the NoAlerts parameter to
' True.
'
' This function uses the following procedures. They are declared as Private
' procedures at the end of this module.
'       IsDynamic
'       IsEmpty
'       GetDimensions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ResNdx As Long
Dim InNdx  As Long

'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = False

'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(SourceArray) = False Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case GetDimensions(arr:=SourceArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' Not Allowed.
        '''''''''''''''''''''''''''''
        Debug.Print "SourceArray is an unallocated array."
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        
        Debug.Print "SourceArray is a multi-dimensional array. This is not allowed."
        
        Exit Function
End Select
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    Debug.Print "ResultArray is not an array."
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an dynamic.
'''''''''''''''''''''''''''''''''''
If IsDynamic(arr:=ResultArray) = False Then
    Debug.Print "ResultArray is not a dynamic array."
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case GetDimensions(arr:=ResultArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        Debug.Print "SourceArray is a multi-dimensional array. This is not allowed."
        Exit Function
End Select

'''''''''''''''''''''''''''''''''
' Ensure that all the elements of
' SourceArray are in fact objects.
'''''''''''''''''''''''''''''''''
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If IsObject(SourceArray(InNdx)) = False Then
        Debug.Print "Element " & CStr(InNdx) & " of SourceArray is not an object."
        Exit Function
    End If
Next InNdx

''''''''''''''''''''''''''''''
' Erase the ResultArray. Since
' ResultArray is dynamic, this
' will relase the memory used
' by ResultArray and return
' the array to an unallocated
' state.
''''''''''''''''''''''''''''''
Erase ResultArray
''''''''''''''''''''''''''''''
' Now, size ResultArray to the
' size of SourceArray. After
' moving all the non-Nothing
' elements, we'll do another
' resize to get ResultArray
' to the used size. This method
' allows us to avoid Redim
' Preserve for every element.
'''''''''''''''''''''''''''''
ReDim ResultArray(LBound(SourceArray) To UBound(SourceArray))

ResNdx = LBound(SourceArray)
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If Not SourceArray(InNdx) Is Nothing Then
        Set ResultArray(ResNdx) = SourceArray(InNdx)
        ResNdx = ResNdx + 1
    End If
Next InNdx
''''''''''''''''''''''''''''''''''''''''''
' Now that we've copied all the
' non-Nothing elements from SourceArray
' to ResultArray, we call Redim Preserve
' to resize the ResultArray to the size
' actually used. Test ResNdx to see
' if we actually copied any elements.
''''''''''''''''''''''''''''''''''''''''''
If ResNdx > LBound(SourceArray) Then
    '''''''''''''''''''''''''''''''''''''''
    ' If ResNdx > LBound(SourceArray) then
    ' we copied at least one element out of
    ' SourceArray.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve ResultArray(LBound(ResultArray) To ResNdx - 1)
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Otherwise, we didn't copy any elements
    ' from SourceArray (all elements in SourceArray
    ' were Nothing). In this case, Erase ResultArray.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Erase ResultArray
End If
'''''''''''''''''''''''''''''
' No errors were encountered.
' Return True.
'''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = True


End Function

Public Function DeleteArrayElement(inputArray As Variant, ElementNumber As Long, _
    Optional ResizeDynamic As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DeleteArrayElement
' This function deletes an element from InputArray, and shifts elements that are to the
' right of the deleted element to the left. If InputArray is a dynamic array, and the
' ResizeDynamic parameter is True, the array will be resized one element smaller. Otherwise,
' the right-most entry in the array is set to the default value appropriate to the data
' type of the array (0, vbNullString, Empty, or Nothing). If the array is an array of Variant
' types, the default data type is the data type of the last element in the array.
' The function returns True if the elememt was successfully deleted, or False if an error
' occurrred. This procedure works only on single-dimensional
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim vType As VbVarType

''''''''''''''''''''''''''''''''
' Set the default result
''''''''''''''''''''''''''''''''
DeleteArrayElement = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array
'''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=inputArray) <> 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a valid ElementNumber
''''''''''''''''''''''''''''''''''''''''''''''
If (ElementNumber < LBound(inputArray)) Or (ElementNumber > UBound(inputArray)) Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' If we have a single element array, Erase it.
''''''''''''''''''''''''''''''''''''''''''''''
If LBound(inputArray) = UBound(inputArray) Then
    Erase inputArray
    Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''
' Get the variable data type of the element
' we're deleting.
''''''''''''''''''''''''''''''''''''''''''''''
vType = VarType(inputArray(UBound(inputArray)))
If vType >= vbArray Then
    vType = vType - vbArray
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the left
''''''''''''''''''''''''''''''''''''''''''''''
For Ndx = ElementNumber To UBound(inputArray) - 1
    inputArray(Ndx) = inputArray(Ndx + 1)
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''
' If ResizeDynamic is True, resize the array
' if it is dynamic.
''''''''''''''''''''''''''''''''''''''''''''''
If IsDynamic(arr:=inputArray) = True Then
    If ResizeDynamic = True Then
        ''''''''''''''''''''''''''''''''
        ' Resize the array and get out.
        ''''''''''''''''''''''''''''''''
        ReDim Preserve inputArray(LBound(inputArray) To UBound(inputArray) - 1)
        DeleteArrayElement = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''
' Set the last element of the
' InputArray to the proper
' default value.
'''''''''''''''''''''''''''''
Select Case vType
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
        inputArray(UBound(inputArray)) = 0
    Case vbString
        inputArray(UBound(inputArray)) = vbNullString
    Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
        inputArray(UBound(inputArray)) = Empty
    Case vbBoolean
        inputArray(UBound(inputArray)) = False
    Case vbObject
        Set inputArray(UBound(inputArray)) = Nothing
    Case Else
        inputArray(UBound(inputArray)) = 0
End Select

DeleteArrayElement = True

End Function

Public Function FirstNonEmptyStringIndexInArray(inputArray As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FirstNonEmptyStringIndexInArray
' This returns the index into InputArray of the first non-empty string.
' This is generally used when InputArray is the result of a sort operation,
' which puts empty strings at the beginning of the array.
' Returns -1 is an error occurred or if the entire array is empty strings.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

If IsArray(inputArray) = False Then
    FirstNonEmptyStringIndexInArray = -1
    Exit Function
End If

Select Case GetDimensions(arr:=inputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''
        ' indicates an unallocated dynamic array.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''
        ' single dimensional array. OK.
        '''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''
        ' multidimensional array. Invalid.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
        Exit Function
End Select

For Ndx = LBound(inputArray) To UBound(inputArray)
    If inputArray(Ndx) <> vbNullString Then
        FirstNonEmptyStringIndexInArray = Ndx
        Exit Function
    End If
Next Ndx

FirstNonEmptyStringIndexInArray = -1
End Function


Public Function InsertElementIntoArray(inputArray As Variant, Index As Long, _
    Value As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InsertElementIntoArray
' This function inserts an element with a value of Value into InputArray at locatation Index.
' InputArray must be a dynamic array. The Value is stored in location Index, and everything
' to the right of Index is shifted to the right. The array is resized to make room for
' the new element. The value of Index must be greater than or equal to the LBound of
' InputArray and less than or equal to UBound+1. If Index is UBound+1, the Value is
' placed at the end of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
InsertElementIntoArray = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure InputArray is dynamic.
''''''''''''''''''''''''''''''''
If IsDynamic(arr:=inputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is allocated.
'''''''''''''''''''''''''''''''''
If IsAllocated(arr:=inputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''
If GetDimensions(arr:=inputArray) <> 1 Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Ensure Index is a valid element index.
' We allow Index to be equal to
' UBound + 1 to facilitate inserting
' a value at the end of the array. E.g.,
' InsertElementIntoArray(Arr,UBound(Arr)+1,123)
' will insert 123 at the end of the array.
'''''''''''''''''''''''''''''''''''''''''
If (Index < LBound(inputArray)) Or (Index > UBound(inputArray) + 1) Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Resize the array
'''''''''''''''''''''''''''''''''''''''''''''
ReDim Preserve inputArray(LBound(inputArray) To UBound(inputArray) + 1)
'''''''''''''''''''''''''''''''''''''''''''''
' First, we set the newly created last element
' of InputArray to Value. This is done to trap
' an error 13, type mismatch. This last entry
' will be overwritten when we shift elements
' to the right, and the Value will be inserted
' at Index.
'''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
err.Clear
inputArray(UBound(inputArray)) = Value
If err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' an error 13, type mismatch.
    ' Redim the array back to its original
    ' size and exit the function.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve inputArray(LBound(inputArray) To UBound(inputArray) - 1)
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the right.
'''''''''''''''''''''''''''''''''''''''''''''
For Ndx = UBound(inputArray) To Index + 1 Step -1
    inputArray(Ndx) = inputArray(Ndx - 1)
Next Ndx

'''''''''''''''''''''''''''''''''''''''''''''
' Insert Value at Index
'''''''''''''''''''''''''''''''''''''''''''''
inputArray(Index) = Value


InsertElementIntoArray = True


End Function


Public Function IsNumericDataType(testVar As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsNumericDataType
'
' This function returns TRUE or FALSE indicating whether the data
' type of a variable is a numeric data type. It will return TRUE
' for all of the following data types:
'       vbCurrency
'       vbDecimal
'       vbDouble
'       vbInteger
'       vbLong
'       vbSingle
'
' It will return FALSE for any other data type, including empty Variants and objects.
' If TestVar is an allocated array, it will test data type of the array
' and return TRUE or FALSE for that data type. If TestVar is an allocated
' array, it tests the data type of the first element of the array. If
' TestVar is an array of Variants, the function will indicate only whether
' the first element of the array is numeric. Other elements of the array
' may not be numeric data types. To test an entire array of variants
' to ensure they are all numeric data types, use the IsVariantArrayNumeric
' function.
'
' It will return FALSE for any other data type. Use this procedure
' instead of VBA's IsNumeric function because IsNumeric will return
' TRUE if the variable is a string containing numeric data. This
' will cause problems with code like
'        Dim V1 As Variant
'        Dim V2 As Variant
'        V1 = "1"
'        V2 = "2"
'        If IsNumeric(V1) = True Then
'            If IsNumeric(V2) = True Then
'                Debug.Print V1 + V2
'            End If
'        End If
'
' The output of the Debug.Print statement will be "12", not 3,
' because V1 and V2 are strings and the '+' operator acts like
' the '&' operator when used with strings. This can lead to
' unexpected results.
'
' IsNumeric should only be used to test strings for numeric content
' when converting a string value to a numeric variable.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Element As Variant
    Dim NumDims As Long
    If IsArray(testVar) = True Then
        NumDims = GetDimensions(arr:=testVar)
        If NumDims > 1 Then
            '''''''''''''''''''''''''''''''''
            ' this procedure does not support
            ' multi-dimensional arrays.
            '''''''''''''''''''''''''''''''''
            IsNumericDataType = False
            Exit Function
        End If
        If IsAllocated(arr:=testVar) = True Then
            Element = testVar(LBound(testVar))
            Select Case VarType(Element)
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
                    Exit Function
                Case Else
                    IsNumericDataType = False
                    Exit Function
            End Select
        Else
            Select Case VarType(testVar) - vbArray
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
                    Exit Function
                Case Else
                    IsNumericDataType = False
                    Exit Function
            End Select
        End If
    End If
    Select Case VarType(testVar)
        Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            IsNumericDataType = True
        Case Else
            IsNumericDataType = False
    End Select
End Function



Public Function IsVariantArrayConsistent(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsVariantArrayConsistent
'
' This returns TRUE or FALSE indicating whether an array of variants
' contains all the same data types. Returns FALSE under the following
' circumstances:
'       Arr is not an array
'       Arr is an array but is unallocated
'       Arr is a multidimensional array
'       Arr is allocated but does not contain consistant data types.
'
' If Arr is an array of objects, objects that are Nothing are ignored.
' As long as all non-Nothing objects are the same object type, the
' function returns True.
'
' It returns TRUE if all the elements of the array have the same
' data type. If Arr is an array of a specific data types, not variants,
' (E.g., Dim V(1 To 3) As Long), the function will return True. If
' an array of variants contains an uninitialized element (VarType =
' vbEmpty) that element is skipped and not used in the comparison. The
' reasoning behind this is that an empty variable will return the
' data type of the variable to which it is assigned (e.g., it will
' return vbNullString to a String and 0 to a Double).
'
' The function does not support arrays of User Defined Types.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FirstDataType As VbVarType
Dim Ndx As Long
'''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not an array.
'''''''''''''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not allocated.
''''''''''''''''''''''''''''''''''''''''''
If IsAllocated(arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''
' Exit with false on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr) <> 1 Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Test if we have an array of a specific
' type rather than Variants. If so,
' return TRUE and get out.
''''''''''''''''''''''''''''''''''''''''''
If (VarType(arr) <= vbArray) And _
    (VarType(arr) <> vbVariant) Then
    IsVariantArrayConsistent = True
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Get the data type of the first element.
''''''''''''''''''''''''''''''''''''''''''
FirstDataType = VarType(arr(LBound(arr)))
''''''''''''''''''''''''''''''''''''''''''
' Loop through the array and exit if
' a differing data type if found.
''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(arr) + 1 To UBound(arr)
    If VarType(arr(Ndx)) <> vbEmpty Then
        If IsObject(arr(Ndx)) = True Then
            If Not arr(Ndx) Is Nothing Then
                If VarType(arr(Ndx)) <> FirstDataType Then
                    IsVariantArrayConsistent = False
                    Exit Function
                End If
            End If
        Else
            If VarType(arr(Ndx)) <> FirstDataType Then
                IsVariantArrayConsistent = False
                Exit Function
            End If
        End If
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''''''
' If we make it out of the loop,
' then the array is consistent.
''''''''''''''''''''''''''''''''''''''''''
IsVariantArrayConsistent = True

End Function



Public Function IsVariantArrayNumeric(testarray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsVariantArrayNumeric
'
' This function return TRUE if all the elements of an array of
' variants are numeric data types. They need not all be the same data
' type. You can have a mix of Integer, Longs, Doubles, and Singles.
' As long as they are all numeric data types, the function will
' return TRUE. If a non-numeric data type is encountered, the
' function will return FALSE. Also, it will return FALSE if
' TestArray is not an array, or if TestArray has not been
' allocated. TestArray may be a multi-dimensional array. This
' procedure uses the IsNumericDataType function to determine whether
' a variable is a numeric data type. If there is an uninitialized
' variant (VarType = vbEmpty) in the array, it is skipped and not
' used in the comparison (i.e., Empty is considered a valid numeric
' data type since you can assign a number to it).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim DimNdx As Long
Dim NumDims As Long

''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(testarray) = False Then
    IsVariantArrayNumeric = False
    Exit Function
End If
''''''''''''''''''''''''''''''''
' Ensure that TestArray has been
' allocated.
''''''''''''''''''''''''''''''''
If IsAllocated(arr:=testarray) = False Then
    IsVariantArrayNumeric = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is a one
' dimensional array. This procedure
' will not work on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''
'If GetDimensions(Arr:=TestArray) > 1 Then
'    IsVariantArrayNumeric = False
'    Exit Function
'End If

NumDims = GetDimensions(arr:=testarray)
If NumDims = 1 Then
    '''''''''''''''''''''''''''''''''''
    ' single dimensional array
    '''''''''''''''''''''''''''''''''''
    For Ndx = LBound(testarray) To UBound(testarray)
        If IsObject(testarray(Ndx)) = True Then
            IsVariantArrayNumeric = False
            Exit Function
        End If

        If VarType(testarray(Ndx)) <> vbEmpty Then
            If IsNumericDataType(testVar:=testarray(Ndx)) = False Then
                IsVariantArrayNumeric = False
                Exit Function
            End If
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''
    ' multi-dimensional array
    ''''''''''''''''''''''''''''''''''''
    For DimNdx = 1 To NumDims
        For Ndx = LBound(testarray, DimNdx) To UBound(testarray, DimNdx)
            If VarType(testarray(Ndx, DimNdx)) <> vbEmpty Then
                If IsNumericDataType(testVar:=testarray(Ndx, DimNdx)) = False Then
                    IsVariantArrayNumeric = False
                    Exit Function
                End If
            End If
        Next Ndx
    Next DimNdx
End If

'''''''''''''''''''''''''''''''''''''''
' If we made it out of the loop, then
' the array is entirely numeric.
'''''''''''''''''''''''''''''''''''''''
IsVariantArrayNumeric = True

End Function



Public Function MoveEmptyStringsToEndOfArray(inputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure takes the SORTED array InputArray, which, if sorted in
' ascending order, will have all empty strings at the front of the array.
' This procedure moves those strings to the end of the array, shifting
' the non-empty strings forward in the array.
' Note that InputArray MUST be sorted in ascending order.
' Returns True if the array was correctly shifted (if necessary) and False
' if an error occurred.
' This function uses the following functions, which are included as Private
' procedures at the end of this module.
'       FirstNonEmptyStringIndexInArray
'       GetDimensions
'       IsAllocated

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As String
Dim Ndx As Long
Dim Ndx2 As Long
Dim NonEmptyNdx As Long
Dim FirstNonEmptyNdx As Long


''''''''''''''''''''''''''''''''
' Ensure InpuyArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
    Exit Function
End If


''''''''''''''''''''''''''''''''''''
' Ensure that the array is allocated
' (not an empty array).
''''''''''''''''''''''''''''''''''''
If IsAllocated(arr:=inputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
    Exit Function
End If


FirstNonEmptyNdx = FirstNonEmptyStringIndexInArray(inputArray:=inputArray)
If FirstNonEmptyNdx <= LBound(inputArray) Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' No empty strings at the beginning of the
    ' array. Get out now.
    ''''''''''''''''''''''''''''''''''''''''''
    MoveEmptyStringsToEndOfArray = True
    Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the array, swapping vbNullStrings
' at the beginning with values at the end.
''''''''''''''''''''''''''''''''''''''''''''''''
NonEmptyNdx = FirstNonEmptyNdx
For Ndx = LBound(inputArray) To UBound(inputArray)
    If inputArray(Ndx) = vbNullString Then
        inputArray(Ndx) = inputArray(NonEmptyNdx)
        inputArray(NonEmptyNdx) = vbNullString
        NonEmptyNdx = NonEmptyNdx + 1
        If NonEmptyNdx > UBound(inputArray) Then
            Exit For
        End If
    End If
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set entires (Ndx+1) to UBound(InputArray) to
' vbNullStrings.
''''''''''''''''''''''''''''''''''''''''''''''''''''
For Ndx2 = Ndx + 1 To UBound(inputArray)
    inputArray(Ndx2) = vbNullString
Next Ndx2
MoveEmptyStringsToEndOfArray = True

End Function

Public Function ResetVariantArrayToDefaults(inputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ResetVariantArrayToDefaults
' This resets all the elements of an array of Variants back to their appropriate
' default values. The elements of the array may be of mixed types (e.g., some Longs,
' some Objects, some Strings, etc). Each data type will be set to the appropriate
' default value (0, vbNullString, Empty, or Nothing). It returns True if the
' array was set to defautls, or False if an error occurred. InputArray must be
' an allocated single-dimensional array. This function differs from the Erase
' function in that it preserves the original data types, while Erase sets every
' element to Empty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
ResetVariantArrayToDefaults = False

'''''''''''''''''''''''''''''''
' Ensure InputArray is an array
'''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional allocated array.
'''''''''''''''''''''''''''''''
If GetDimensions(arr:=inputArray) <> 1 Then
    Exit Function
End If

For Ndx = LBound(inputArray) To UBound(inputArray)
    'SetVariableToDefault InputArray(Ndx)
Next Ndx

ResetVariantArrayToDefaults = True

End Function

Public Function ReverseArrayInPlace( _
  inputArray As Variant _
) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is reversed. This works only on single-dimensional arrays
' of simple data types (String, Single, Double, Integer, Long). It will not work
' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    
    Debug.Print "The InputArray parameter is not an array."
    
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case GetDimensions(inputArray)
    Case 0
        Debug.Print "The input array is an empty, unallocated array."
        Exit Function
    Case 1
        ' ok
    Case Else
        
        Debug.Print "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        Exit Function
End Select

Ndx2 = UBound(inputArray)
''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(inputArray) To ((UBound(inputArray) - LBound(inputArray) + 1) \ 2)
    'swap the elements
    Temp = inputArray(Ndx)
    inputArray(Ndx) = inputArray(Ndx2)
    inputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayInPlace = True

End Function



Public Function ReverseArrayOfObjectsInPlace( _
  inputArray As Variant _
) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayOfObjectsInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is reversed. This works only with arrays of objects. It does
' not work on simple variables. Use ReverseArrayInPlace for simple variables. An error
' will occur if an element of the array is not an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    Debug.Print "The InputArray parameter is not an array."
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case GetDimensions(inputArray)
    Case 0
        Debug.Print "The input array is an empty, unallocated array."
        Exit Function
    Case 1
        ' ok
    Case Else
        
        Debug.Print "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        Exit Function
End Select

Ndx2 = UBound(inputArray)

'''''''''''''''''''''''''''''''''''''
' ensure the entire array consists
' of objects (Nothing objects are
' allowed).
'''''''''''''''''''''''''''''''''''''
For Ndx = LBound(inputArray) To UBound(inputArray)
    If IsObject(inputArray(Ndx)) = False Then
        Debug.Print "Array item " & CStr(Ndx) & " is not an object."
        Exit Function
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(inputArray) To ((UBound(inputArray) - LBound(inputArray) + 1) \ 2)
    Set Temp = inputArray(Ndx)
    Set inputArray(Ndx) = inputArray(Ndx2)
    Set inputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = True

End Function


Public Function SetObjectArrayToNothing(inputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetObjectArrrayToNothing
' This sets all the elements of InputArray to Nothing. Use this function
' rather than Erase because if InputArray is an array of Variants, Erase
' will set each element to Empty, not Nothing, and the element will cease
' to be an object.
'
' The function returns True if successful, False otherwise.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim n As Long

''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''''''''
If IsArray(inputArray) = False Then
    SetObjectArrayToNothing = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single-dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=inputArray) <> 1 Then
    SetObjectArrayToNothing = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the array is allocated and that each
' element is an object (or Nothing). If the
' array is not allocated, return True.
' We do this test before setting any element
' to Nothing so we don't end up with an array
' that is a mix of Empty and Nothing values.
' This means looping through the array twice,
' but it ensures all or none of the elements
' get set to Nothing.
''''''''''''''''''''''''''''''''''''''''''''''''
If IsAllocated(arr:=inputArray) = True Then
    For n = LBound(inputArray) To UBound(inputArray)
        If IsObject(inputArray(n)) = False Then
            SetObjectArrayToNothing = False
            Exit Function
        End If
    Next n
Else
    SetObjectArrayToNothing = True
    Exit Function
End If


'''''''''''''''''''''''''''''''''''''''''''''
' Set each element of InputArray to Nothing.
'''''''''''''''''''''''''''''''''''''''''''''
For n = LBound(inputArray) To UBound(inputArray)
    Set inputArray(n) = Nothing
Next n

SetObjectArrayToNothing = True

End Function

Public Function TransposeArray(InputArr As Variant, OutputArr As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransposeArray
' This transposes a two-dimensional array. It returns True if successful or
' False if an error occurs. InputArr must be two-dimensions. OutputArr must be
' a dynamic array. It will be Erased and resized, so any existing content will
' be destroyed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim RowNdx As Long
Dim ColNdx As Long
Dim LB1 As Long
Dim LB2 As Long
Dim UB1 As Long
Dim UB2 As Long

'''''''''''''''''''''''''''''''''''
' Ensure InputArr and OutputArr
' are arrays.
'''''''''''''''''''''''''''''''''''
If (IsArray(InputArr) = False) Or (IsArray(OutputArr) = False) Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure OutputArr is a dynamic
' array.
'''''''''''''''''''''''''''''''''''
If IsDynamic(arr:=OutputArr) = False Then
    TransposeArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure InputArr is two-dimensions,
' no more, no lesss.
''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=InputArr) <> 2 Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' Get the Lower and Upper bounds of
' InputArr.
'''''''''''''''''''''''''''''''''''''''
LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)

'''''''''''''''''''''''''''''''''''''''''
' Erase and ReDim OutputArr
'''''''''''''''''''''''''''''''''''''''''
Erase OutputArr
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Output array. Not the that the LBound and UBound
' values are preserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the elemetns of InputArr and put each value
' in the proper element of the tranposed array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
    For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
        OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
    Next ColNdx
Next RowNdx

'''''''''''''''''''''''''
' Success -- return True.
'''''''''''''''''''''''''
TransposeArray = True

End Function


Public Function VectorsToArray(arr As Variant, ParamArray Vectors()) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VectorsToArray
' This function takes 1 or more single-dimensional arrays and converts
' them into a single multi-dimensional array. Each array in Vectors
' comprises one row of the new array. The number of columns in the
' new array is the maximum of the number of elements in each vector.
' Arr MUST be a dynamic array of a data type compatible with ALL the
' elements in each Vector. The code does NOT trap for an error
' 13 - Type Mismatch.
'
' If the Vectors are of differing sizes, Arr is sized to hold the
' maximum number of elements in a Vector. The procedure Erases the
' Arr array, so when it is reallocated with Redim, all elements will
' be the reset to their default value (0 or vbNullString or Empty).
' Unused elements in the new array will remain the default value for
' that data type.
'
' Each Vector in Vectors must be a single dimensional array, but
' the Vectors may be of different sizes and LBounds.
'
' Each element in each Vector must be a simple data type. The elements
' may NOT be Object, Arrays, or User-Defined Types.
'
' The rows and columns of the result array are 0-based, regardless of
' the LBound of each vector and regardless of the Option Base statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Vector As Variant
Dim VectorNdx As Long
Dim NumElements As Long
Dim NumRows As Long
Dim NumCols As Long
Dim RowNdx As Long
Dim ColNdx As Long
Dim vType As VbVarType

'''''''''''''''''''''''''''''''''''
' Ensure we have an Array
''''''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    VectorsToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure we have a dynamic array
''''''''''''''''''''''''''''''''''
If IsDynamic(arr) = False Then
    VectorsToArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure that at least one vector
' was passed in Vectors
'''''''''''''''''''''''''''''''''
If IsMissing(Vectors) = True Then
    VectorsToArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''
' Loop through Vectors to determine the
' size of the result array. We do this
' loop first to prevent having to do
' a Redim Preserve. This requires looping
' through Vectors a second time, but this
' is still faster than doing Redim Preserves.
'''''''''''''''''''''''''''''''''''''''''''''''
For Each Vector In Vectors
    ''''''''''''''''''''''''''''
    ' Ensure Vector is single
    ' dimensional array. This
    ' will take care of the case
    ' if Vector is an unallocated
    ' array (GetDimensions = 0
    ' for an unallocated array).
    ''''''''''''''''''''''''''''
    If GetDimensions(Vector) <> 1 Then
        VectorsToArray = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Ensure that Vector is not an array.
    '''''''''''''''''''''''''''''''''''''
    If IsArray(Vector) = False Then
        VectorsToArray = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''
    ' Increment the number of rows.
    ' Each Vector is one row or the
    ' result array. Test the size
    ' of Vector. If it is larger
    ' than the existing value of
    ' NumCols, set NumCols to the
    ' new, larger, value.
    '''''''''''''''''''''''''''''''''
    NumRows = NumRows + 1
    If NumCols < UBound(Vector) - LBound(Vector) + 1 Then
        NumCols = UBound(Vector) - LBound(Vector) + 1
    End If
Next Vector
''''''''''''''''''''''''''''''''''''''''''''
' Redim Arr to the appropriate size. Arr
' is 0-based in both directions, regardless
' of the LBound of the original Arr and
' regardless of the LBounds of the Vectors.
''''''''''''''''''''''''''''''''''''''''''''
ReDim arr(0 To NumRows - 1, 0 To NumCols - 1)

'''''''''''''''''''''''''''''''
' Loop row-by-row.
For RowNdx = 0 To NumRows - 1
    ''''''''''''''''''''''''''''''''
    ' Loop through the columns.
    ''''''''''''''''''''''''''''''''
    For ColNdx = 0 To NumCols - 1
        ''''''''''''''''''''''''''''
        ' Set Vector (a Variant) to
        ' the Vectors(RowNdx) array.
        ' We declare Vector as a
        ' variant so it can take an
        ' array of any simple data
        ' type.
        ''''''''''''''''''''''''''''
        Vector = Vectors(RowNdx)
        '''''''''''''''''''''''''''''
        ' The vectors need not ber
        If ColNdx < UBound(Vector) - LBound(Vector) + 1 Then
            vType = VarType(Vector(LBound(Vector) + ColNdx))
            If vType >= vbArray Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Test for VType >= vbArray. The VarType of an array
                ' is vbArray + VarType(element of array). E.g., the
                ' VarType of an array of Longs equal vbArray + vbLong.
                ' Anything greater than or equal to vbArray is an
                ' array of some time.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                VectorsToArray = False
                Exit Function
            End If
            If vType = vbObject Then
                VectorsToArray = False
                Exit Function
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Vector(LBound(Vector) + ColNdx) is
            ' a simple data type. If Vector(LBound(Vector) + ColNdx)
            ' is not a compatible data type with Arr, then a Type
            ' Mismatch error will occur. We do NOT trap this error.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            arr(RowNdx, ColNdx) = Vector(LBound(Vector) + ColNdx)
        End If
    Next ColNdx
Next RowNdx

VectorsToArray = True

End Function




Public Function IsArraySorted(testarray As Variant, _
    Optional Descending As Boolean = False) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArraySorted
' This function determines whether a single-dimensional array is sorted. Because
' sorting is an expensive operation, especially so on large array of Variants,
' you may want to determine if an array is already in sorted order prior to
' doing an actual sort.
' This function returns True if an array is in sorted order (either ascending or
' descending order, depending on the value of the Descending parameter -- default
' is false = Ascending). The decision to do a string comparison (with StrComp) or
' a numeric comparison (with < or >) is based on the data type of the first
' element of the array.
' If TestArray is not an array, is an unallocated dynamic array, or has more than
' one dimension, or the VarType of TestArray is not compatible, the function
' returns NULL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StrCompResultFail As Long
Dim NumericResultFail As Boolean
Dim Ndx As Long
Dim NumCompareResult As Boolean
Dim StrCompResult As Long

Dim IsString As Boolean
Dim vType As VbVarType

''''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''''
If IsArray(testarray) = False Then
    IsArraySorted = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If GetDimensions(arr:=testarray) <> 1 Then
    IsArraySorted = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' The following code sets the values of
' comparison that will indicate that the
' array is unsorted. It the result of
' StrComp (for strings) or ">=" (for
' numerics) equals the value specified
' below, we know that the array is
' unsorted.
'''''''''''''''''''''''''''''''''''''''''''''
If Descending = True Then
    StrCompResultFail = -1
    NumericResultFail = False
Else
    StrCompResultFail = 1
    NumericResultFail = True
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Determine whether we are going to do a string
' comparison or a numeric comparison.
''''''''''''''''''''''''''''''''''''''''''''''
vType = VarType(testarray(LBound(testarray)))
Select Case vType
    Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
    '''''''''''''''''''''''''''''''''
    ' Unsupported types. Reutrn Null.
    '''''''''''''''''''''''''''''''''
        IsArraySorted = Null
        Exit Function
    Case vbString, vbVariant
    '''''''''''''''''''''''''''''''''
    ' Compare as string
    '''''''''''''''''''''''''''''''''
        IsString = True
    Case Else
    '''''''''''''''''''''''''''''''''
    ' Compare as numeric
    '''''''''''''''''''''''''''''''''
        IsString = False
End Select

For Ndx = LBound(testarray) To UBound(testarray) - 1
    If IsString = True Then
        StrCompResult = StrComp(testarray(Ndx), testarray(Ndx + 1))
        If StrCompResult = StrCompResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    Else
        NumCompareResult = (testarray(Ndx) >= testarray(Ndx + 1))
        If NumCompareResult = NumericResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    End If
Next Ndx


''''''''''''''''''''''''''''
' If we made it out of  the
' loop, then the array is
' in sorted order. Return
' True.
''''''''''''''''''''''''''''
IsArraySorted = True
End Function


Public Function CombineTwoDArrays(arr1 As Variant, _
    arr2 As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TwoArraysToOneArray
' This takes two 2-dimensional arrays, Arr1 and Arr2, and
' returns an array combining the two. The number of Rows
' in the result is NumRows(Arr1) + NumRows(Arr2). Arr1 and
' Arr2 must have the same number of columns, and the result
' array will have that many columns. All the LBounds must
' be the same. E.g.,
' The following arrays are legal:
'        Dim Arr1(0 To 4, 0 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The following arrays are illegal
'        Dim Arr1(0 To 4, 1 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The returned result array is Arr1 with additional rows
' appended from Arr2. For example, the arrays
'    a    b        and     e    f
'    c    d                g    h
' become
'    a    b
'    c    d
'    e    f
'    g    h
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr1.
'''''''''''''''''''''''''''''''''
Dim LBoundRow1 As Long
Dim UBoundRow1 As Long
Dim LBoundCol1 As Long
Dim UBoundCol1 As Long

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr2.
'''''''''''''''''''''''''''''''''
Dim LBoundRow2 As Long
Dim UBoundRow2 As Long
Dim LBoundCol2 As Long
Dim UBoundCol2 As Long

'''''''''''''''''''''''''''''''''''
' Upper and lower bounds of Result.
'''''''''''''''''''''''''''''''''''
Dim LBoundRowResult As Long
Dim UBoundRowResult As Long
Dim LBoundColResult As Long
Dim UBoundColResult As Long

'''''''''''''''''
' Index Variables
'''''''''''''''''
Dim RowNdx1 As Long
Dim ColNdx1 As Long
Dim RowNdx2 As Long
Dim ColNdx2 As Long
Dim RowNdxResult As Long
Dim ColNdxResult As Long


'''''''''''''
' Array Sizes
'''''''''''''
Dim NumRows1 As Long
Dim NumCols1 As Long

Dim NumRows2 As Long
Dim NumCols2 As Long

Dim NumRowsResult As Long
Dim NumColsResult As Long

Dim Done As Boolean
Dim result() As Variant
Dim ResultTrans() As Variant

Dim v As Variant


'''''''''''''''''''''''''''''''
' Ensure that Arr1 and Arr2 are
' arrays.
''''''''''''''''''''''''''''''
If (IsArray(arr1) = False) Or (IsArray(arr2) = False) Then
    CombineTwoDArrays = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are allocated
' two dimensional arrays.
''''''''''''''''''''''''''''''''''
If (GetDimensions(arr1) <> 2) Or (GetDimensions(arr2) <> 2) Then
    CombineTwoDArrays = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' Ensure that the LBound and UBounds
' of the second dimension are the
' same for both Arr1 and Arr2.
'''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''
' Get the existing bounds.
''''''''''''''''''''''''''
LBoundRow1 = LBound(arr1, 1)
UBoundRow1 = UBound(arr1, 1)

LBoundCol1 = LBound(arr1, 2)
UBoundCol1 = UBound(arr1, 2)

LBoundRow2 = LBound(arr2, 1)
UBoundRow2 = UBound(arr2, 1)

LBoundCol2 = LBound(arr2, 2)
UBoundCol2 = UBound(arr2, 2)

''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the total number of rows for the result
' array.
''''''''''''''''''''''''''''''''''''''''''''''''''
NumRows1 = UBoundRow1 - LBoundRow1 + 1
NumCols1 = UBoundCol1 - LBoundCol1 + 1
NumRows2 = UBoundRow2 - LBoundRow2 + 1
NumCols2 = UBoundCol2 - LBoundCol2 + 1

'''''''''''''''''''''''''''''''''''''''''
' Ensure the number of columns are equal.
'''''''''''''''''''''''''''''''''''''''''
If NumCols1 <> NumCols2 Then
    CombineTwoDArrays = Null
    Exit Function
End If

NumRowsResult = NumRows1 + NumRows2

'''''''''''''''''''''''''''''''''''''''
' Ensure that ALL the LBounds are equal.
''''''''''''''''''''''''''''''''''''''''
If (LBoundRow1 <> LBoundRow2) Or _
    (LBoundRow1 <> LBoundCol1) Or _
    (LBoundRow1 <> LBoundCol2) Then
    CombineTwoDArrays = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''
' Get the LBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
LBoundColResult = LBoundRow1
'''''''''''''''''''''''''''''''
' Get the UBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
UBoundColResult = UBoundCol1

UBoundRowResult = LBound(arr1, 1) + NumRows1 + NumRows2 - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Result array to have number of rows equal to
' number-of-rows(Arr1) + number-of-rows(Arr2)
' and number-of-columns equal to number-of-columns(Arr1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)

RowNdxResult = LBound(result, 1) - 1

Done = False
Do Until Done
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr1 to Result
    ''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx1 = LBound(arr1, 1) To UBound(arr1, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx1 = LBound(arr1, 2) To UBound(arr1, 2)
            v = arr1(RowNdx1, ColNdx1)
            result(RowNdxResult, ColNdx1) = v
        Next ColNdx1
    Next RowNdx1

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr2 to Result
    '''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx2 = LBound(arr2, 1) To UBound(arr2, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx2 = LBound(arr2, 2) To UBound(arr2, 2)
            v = arr2(RowNdx2, ColNdx2)
            result(RowNdxResult, ColNdx2) = v
        Next ColNdx2
    Next RowNdx2

    If RowNdxResult >= UBound(result, 1) + (LBoundColResult = 1) Then
        Done = True
    End If
'''''''''''''
' End Of Loop
'''''''''''''
Loop
'''''''''''''''''''''''''
' Return the Result
'''''''''''''''''''''''''
CombineTwoDArrays = result

End Function



Function ExpandArray(arr As Variant, WhichDim As Long, AdditionalElements As Long, _
        FillValue As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExpandArray
' This expands a two-dimensional array in either dimension. It returns the result
' array if successful, or NULL if an error occurred. The original array is never
' changed.
' Paramters:
' --------------------
' Arr                   is the array to be expanded.
'
' WhichDim              is either 1 for additional rows or 2 for
'                       additional columns.
'
' AdditionalElements    is the number of additional rows or columns
'                       to create.
'
' FillValue             is the value to which the new array elements should be
'                       initialized.
'
' You can nest calls to Expand array to expand both the number of rows and
' columns. E.g.,
'
' C = ExpandArray(ExpandArray(Arr:=A, WhichDim:=1, AdditionalElements:=3, FillValue:="R"), _
'    WhichDim:=2, AdditionalElements:=4, FillValue:="C")
' This first adds three rows at the bottom of the array, and then adds four
' columns on the right of the array.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim result As Variant
Dim RowNdx As Long
Dim ColNdx As Long
Dim ResultRowNdx As Long
Dim ResultColNdx As Long
Dim NumRows As Long
Dim NumCols As Long
Dim NewUBound As Long

Const ROWS_ As Long = 1
Const COLS_ As Long = 2


''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    ExpandArray = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''''
' Ensure Arr has two dimenesions.
'''''''''''''''''''''''''''''''''
If GetDimensions(arr:=arr) <> 2 Then
    ExpandArray = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''''
' Ensure the dimension is 1 or 2.
'''''''''''''''''''''''''''''''''
Select Case WhichDim
    Case 1, 2
    Case Else
        ExpandArray = Null
        Exit Function
End Select

''''''''''''''''''''''''''''''''''''
' Ensure AdditionalElements is > 0.
' If AdditionalElements  < 0, return NULL.
' If AdditionalElements  = 0, return Arr.
''''''''''''''''''''''''''''''''''''
If AdditionalElements < 0 Then
    ExpandArray = Null
    Exit Function
End If
If AdditionalElements = 0 Then
    ExpandArray = arr
    Exit Function
End If

NumRows = UBound(arr, 1) - LBound(arr, 1) + 1
NumCols = UBound(arr, 2) - LBound(arr, 2) + 1

If WhichDim = ROWS_ Then
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim result(LBound(arr, 1) To UBound(arr, 1) + AdditionalElements, LBound(arr, 2) To UBound(arr, 2))
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            result(RowNdx, ColNdx) = arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with FillValue.
    '''''''''''''''''''''''''''''''
    For RowNdx = UBound(arr, 1) + 1 To UBound(result, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            result(RowNdx, ColNdx) = FillValue
        Next ColNdx
    Next RowNdx
Else
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim result(LBound(arr, 1) To UBound(arr, 1), UBound(arr, 2) + AdditionalElements)
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            result(RowNdx, ColNdx) = arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with FillValue.
    '''''''''''''''''''''''''''''''
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = UBound(arr, 2) + 1 To UBound(result, 2)
            result(RowNdx, ColNdx) = FillValue
        Next ColNdx
    Next RowNdx

End If
''''''''''''''''''''
' Return the result.
''''''''''''''''''''
ExpandArray = result

End Function



Function SwapArrayRows(arr As Variant, Row1 As Long, Row2 As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SwapArrayRows
' This function returns an array based on Arr with Row1 and Row2 swapped.
' It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim v As Variant
Dim result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(arr) = False Then
    SwapArrayRows = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
result = arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If GetDimensions(arr:=arr) <> 2 Then
    SwapArrayRows = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Row1 > UBound(arr, 1)) Or (Row2 > UBound(arr, 1)) Then
    SwapArrayRows = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' If Row1 = Row2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Row1 = Row2 Then
    SwapArrayRows = arr
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim v(LBound(arr, 2) To UBound(arr, 2))
'''''''''''''''''''''''''''''''''''''''''
' Put Row1 in V
'''''''''''''''''''''''''''''''''''''''''
For ColNdx = LBound(arr, 2) To UBound(arr, 2)
    v(ColNdx) = arr(Row1, ColNdx)
    result(Row1, ColNdx) = arr(Row2, ColNdx)
    result(Row2, ColNdx) = v(ColNdx)
Next ColNdx

SwapArrayRows = result

End Function


Function SwapArrayColumns(arr As Variant, Col1 As Long, Col2 As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SwapArrayColumns
' This function returns an array based on Arr with Col1 and Col2 swapped.
' It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim v As Variant
Dim result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(arr) = False Then
    SwapArrayColumns = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
result = arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If GetDimensions(arr:=arr) <> 2 Then
    SwapArrayColumns = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Col1 > UBound(arr, 2)) Or (Col2 > UBound(arr, 2)) Then
    SwapArrayColumns = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' If Col1 = Col2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Col1 = Col2 Then
    SwapArrayColumns = arr
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim v(LBound(arr, 1) To UBound(arr, 1))
'''''''''''''''''''''''''''''''''''''''''
' Put Col2 in V
'''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(arr, 1) To UBound(arr, 1)
    v(RowNdx) = arr(RowNdx, Col1)
    result(RowNdx, Col1) = arr(RowNdx, Col2)
    result(RowNdx, Col2) = v(RowNdx)
Next RowNdx

SwapArrayColumns = result

End Function

Function GetColumn(arr As Variant, ResultArr As Variant, ColumnNumber As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetColumn
' This populates ResultArr with a one-dimensional array that is the
' specified column of Arr. The existing contents of ResultArr are
' destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RowNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If GetDimensions(arr) <> 2 Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsDynamic(ResultArr) = False Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(arr, 2) < ColumnNumber Then
    GetColumn = False
    Exit Function
End If
If LBound(arr, 2) > ColumnNumber Then
    GetColumn = False
    Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(arr, 1) To UBound(arr, 1))
For RowNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(RowNdx) = arr(RowNdx, ColumnNumber)
Next RowNdx

GetColumn = True


End Function


Function GetRow(arr As Variant, ResultArr As Variant, RowNumber As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetRow
' This populates ResultArr with a one-dimensional array that is the
' specified row of Arr. The existing contents of ResultArr are
' destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ColNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If GetDimensions(arr) <> 2 Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsDynamic(ResultArr) = False Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(arr, 1) < RowNumber Then
    GetRow = False
    Exit Function
End If
If LBound(arr, 1) > RowNumber Then
    GetRow = False
    Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(arr, 2) To UBound(arr, 2))
For ColNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(ColNdx) = arr(RowNumber, ColNdx)
Next ColNdx

GetRow = True

End Function



'-------------------------------------------------------------------------------
'TODO Needs to be moved to a new module UtilCollection in future with other col functions
'---------------------------------------------------------------------------------------
' Procedure : IsCollectionItem
' Author    : Ivanov, Bozhan
' Purpose   : returns true if item is part of the collection
'---------------------------------------------------------------------------------------
Public Function IsCollectionItem( _
  ByRef col As Collection, _
  ByVal key As Variant _
) As Boolean
On Error Resume Next
  IsCollectionItem = False
  Call col(CStr(key))
  If err.Number = 0 Then
    IsCollectionItem = True
    Exit Function
  ElseIf err.Number = 5 Then
    Exit Function
  ElseIf err.Number <> 0 Then
    Dim num As Long
    num = err.Number
    On Error GoTo 0
    err.raise num
  End If
End Function


'
''---------------------------------------------------------------------------------------
'' Procedure   : GetDifference
'' Description : The relative complement of A with respect to a set B,
''   also termed the difference of sets A and B, written B \ A, is the set of elements in
''   B but not in A, e.g. {1,2,3} \ {2,3,4} = {1}
''---------------------------------------------------------------------------------------
'Public Function GetDifference( _
'  ByVal A As Variant, _
'  ByVal WithRespectTo_B As Variant, _
'  Optional ByVal SortDescending As Boolean = False _
') As Variant
'Dim result() As Variant
'
'  result = Array()
'
'  If Not Me.IsAllocated(WithRespectTo_B) Or Not Me.IsAllocated(A) Then _
'    err.raise 9, , "Parameter is an empty array"
'
'  If Not Me.IsOfSimpleDataTypes(WithRespectTo_B) Or Not Me.IsOfSimpleDataTypes(A) Then _
'    err.raise 13, , "Paramter not of simple data type"
'
'  Me.QSortInPlace InputArray:=A, Descending:=SortDescending
'  Me.QSortInPlace InputArray:=WithRespectTo_B, Descending:=SortDescending
'
'  Dim i As Long, j As Long
'  i = LBound(A)
'  j = LBound(WithRespectTo_B)
'
'  While i <= UBound(A) And j <= UBound(WithRespectTo_B)
'    If CompareValues(A(i), WithRespectTo_B(j), vbTextCompare) <> 0 Then
'      Me.IncreaseUpperBound result, 1
'      result(UBound(result)) = A(i)
'    Else
'      j = j + 1
'    End If
'
'    i = i + 1
'  Wend
'
'  While i < UBound(A)
'    Me.IncreaseUpperBound result, 1
'    result(UBound(result)) = A(i)
'    i = i + 1
'  Wend
'
'  GetDifference = result
'End Function

'TODO Revise
'---------------------------------------------------------------------------------------
' Procedure   : BubbleSort1DArray
' Description :
' Source      : https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/Bubble_Sort_One_Dimensional_Arrays
'---------------------------------------------------------------------------------------
Function BubbleSort1DArray(vIn As Variant, bAscending As Boolean, Optional vRet As Variant) As Boolean
    ' Sorts the single dimension list array, ascending or descending
    ' Returns sorted list in vRet if supplied, otherwise in vIn modified
        
    Dim First As Long, Last As Long
    Dim i As Long, j As Long, bWasMissing As Boolean
    Dim Temp As Variant, vW As Variant
    
    First = LBound(vIn)
    Last = UBound(vIn)
    
    ReDim vW(First To Last, 1)
    vW = vIn
    
    If bAscending = True Then
        For i = First To Last - 1
            For j = i + 1 To Last
                If vW(i) > vW(j) Then
                Temp = vW(j)
                vW(j) = vW(i)
                vW(i) = Temp
                End If
            Next j
        Next i
    Else 'descending sort
        For i = First To Last - 1
            For j = i + 1 To Last
                If vW(i) < vW(j) Then
                Temp = vW(j)
                vW(j) = vW(i)
                vW(i) = Temp
                End If
            Next j
        Next i
    End If
  
   'find whether optional vRet was initially missing
    bWasMissing = IsMissing(vRet)
   
   'transfers
   If bWasMissing Then
     vIn = vW  'return in input array
   Else
     ReDim vRet(First To Last, 1)
     vRet = vW 'return with input unchanged
   End If
   
   BubbleSort1DArray = True

End Function

'TODO revise
'--------------------------------------------------------------------------------
' Procedure : Sort2DArr
' Source    : https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/Bubble_Sort_on_One_Key
' Purpose   : Bubblesorts a 2D array on 1 key, up or down, on any column or row.
'             Options include in-place, with the source changed, or
'             returned in vRet, with the source array intact.
'             Optional parameters default to: ROW SORT in place, ASCENDING,
'             using COLUMN ONE as the key.
'--------------------------------------------------------------------------------
Function SortArr2D1Key(ByRef vA As Variant, _
                       Optional ByVal bIsAscending As Boolean = True, _
                       Optional ByVal bIsRowSort As Boolean = True, _
                       Optional ByVal SortIndex As Long = -1, _
                       Optional ByRef vRet As Variant) As Boolean

    
    Dim condition1 As Boolean, vR As Variant
    Dim i As Long, j As Long, y As Long, t As Variant
    Dim loR As Long, hiR As Long, loC As Long, hiC As Long
    Dim r As Long, c As Long, bWasMissing As Boolean
    
    'find bounds of vA data input array
    loR = LBound(vA, 1): hiR = UBound(vA, 1)
    loC = LBound(vA, 2): hiC = UBound(vA, 2)
    
    'find whether optional vR was initially missing
    bWasMissing = IsMissing(vRet)
    'If Not bWasMissing Then Set vRet = Nothing
    
    'check input range of SortIndex
    If bIsRowSort And (SortIndex < loC Or SortIndex > hiC) Then
        MsgBox "SortIndex out of bounds in Sort2DArr; closing now"
        Exit Function
    Else:
    End If
    
    If Not bIsRowSort And (SortIndex < loR Or SortIndex > hiR) Then
        MsgBox "SortIndex out of bounds in Sort2DArr; closing now"
        Exit Function
    Else:
    End If
    
    'pass to a work variable
    vR = vA
    
    'steer input options
    If bIsRowSort Then GoTo ROWSORT Else GoTo COLSORT
    
ROWSORT:
    For i = loR To hiR - 1
        For j = loR To hiR - 1
            If bIsAscending Then
                condition1 = vR(j, SortIndex) > vR(j + 1, SortIndex)
            Else
                condition1 = vR(j, SortIndex) < vR(j + 1, SortIndex)
            End If
            If condition1 Then
                For y = loC To hiC
                    t = vR(j, y)
                    vR(j, y) = vR(j + 1, y)
                    vR(j + 1, y) = t
                Next y
            End If
        Next
    Next
    GoTo TRANSFERS
    
COLSORT:
    For i = loC To hiC - 1
        For j = loC To hiC - 1
            If bIsAscending Then
                condition1 = vR(SortIndex, j) > vR(SortIndex, j + 1)
            Else
                condition1 = vR(SortIndex, j) < vR(SortIndex, j + 1)
            End If
            If condition1 Then
                For y = loR To hiR
                    t = vR(y, j)
                    vR(y, j) = vR(y, j + 1)
                    vR(y, j + 1) = t
                Next y
            End If
        Next
    Next
    GoTo TRANSFERS
    
TRANSFERS:
    'decide whether to return in vA or vRet
    If Not bWasMissing Then
        'vRet was the intended return array
        'so return vRet leaving vA intact
        vRet = vR
    Else:
        'vRet is not intended return array
        'so reload vA with vR
        vA = vR
    End If
    
    'set return function value
    SortArr2D1Key = True
    
End Function

'TODO Revise
'--------------------------------------------------------------------------------------
' Procedure : Sort2DArr3Keys
' Purpose   : Bubblesorts a 2D array using 3 keys, up or down, on any column or row.
'             For example, sorting using up to three columns;
'             Eg; first sorts surnames, then sorts among same surnames for first names,
'             then among similar surnames with same first names for middle names.
'             Options include in-place, with the source changed, or
'             if supplied, returned in vR, with the source array intact.
'             Optional parameters default to: ROW SORT, ASCENDING.
'             Trailing key options that are not needed should be set to same as previous.
' Source    : https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/Bubble_Sort_on_Multiple_Keys
'---------------------------------------------------------------------------------------
Function SortArr2D3Keys(vA As Variant, _
                        Optional Key1 As Long = -1, _
                        Optional Key2 As Long = -1, _
                        Optional Key3 As Long = -1, _
                        Optional ByVal bIsAscending As Boolean = True, _
                        Optional ByVal bIsRowSort As Boolean = True, _
                        Optional ByRef vR As Variant) As Boolean


ASSIGNMENTS:
    Dim condition1 As Boolean, vW As Variant, Temp
    Dim i As Long, j As Long, y As Long, t As Variant
    Dim loR As Long, hiR As Long, loC As Long, hiC As Long
    Dim r As Long, c As Long, bWasMissing As Boolean
    Dim sCombo As String, reply
    Dim b1Used As Boolean, b2Used As Boolean, b3Used As Boolean
    
    'find bounds of vA data input array
    loR = LBound(vA, 1): hiR = UBound(vA, 1)
    loC = LBound(vA, 2): hiC = UBound(vA, 2)
    
    'find whether optional vR was initially missing
    bWasMissing = IsMissing(vR)
    If Not bWasMissing Then Set vR = Nothing
    
KEYCHECKS:
    If Key1 <> -1 Then
        b1Used = True
        'check key within bounds
        If bIsRowSort And (Key1 < loC Or Key1 > hiC) Then
            MsgBox "Sort key1 out of bounds"
            Exit Function
        End If
        If Not bIsRowSort And (Key1 < loR Or Key1 > hiR) Then
            MsgBox "Sort key1 out of bounds"
            Exit Function
        End If
    End If
    
    If Key2 <> -1 Then
        b2Used = True
        'check key within bounds
        If bIsRowSort And (Key2 < loC Or Key2 > hiC) Then
            MsgBox "Sort key2 out of bounds"
            Exit Function
        End If
        If Not bIsRowSort And (Key2 < loR Or Key2 > hiR) Then
            MsgBox "Sort key2 out of bounds"
            Exit Function
        End If
    End If
    
    If Key3 <> -1 Then
        b3Used = True
        'check key within bounds
        If bIsRowSort And (Key3 < loC Or Key3 > hiC) Then
            MsgBox "Sort key3 out of bounds"
            Exit Function
        End If
        If Not bIsRowSort And (Key3 < loR Or Key3 > hiR) Then
            MsgBox "Sort key3 out of bounds"
            Exit Function
        End If
    End If
    
    sCombo = CStr(Abs(b1Used)) & CStr(Abs(b2Used)) & CStr(Abs(b3Used))
    'MsgBox sCombo
    
    Select Case sCombo
    Case "000"
        'no keys selected
        If bIsRowSort Then
           reply = MsgBox("No keys selected." & vbCrLf & _
           "Use lower bound column for a single key?", vbCritical + vbQuestion + vbYesNo, "Please confirm your selection...")
           Select Case reply
           Case vbYes
               Key1 = loC
           Case Else
               Exit Function
           End Select
        Else
           reply = MsgBox("No keys selected." & vbCrLf & _
           "Use lower bound row for a single key?", vbCritical + vbQuestion + vbYesNo, "Please confirm your selection...")
           Select Case reply
           Case vbYes
               Key1 = loR
           Case Else
               Exit Function
           End Select
        End If
    Case "100", "110", "111"
        'proceed normally
    Case Else
        MsgBox "Only three combinations of sort keys are possible" & vbCrLf & _
        "Key1 alone, Key1 with Key2, or Key1 with Key2 and Key3."
        Exit Function
    End Select
    
WORKARRAY:
    'use a working array for sorting
    vW = vA
    
STEERING:
    'steer input options
    If bIsRowSort Then GoTo ROWSORT Else GoTo COLSORT
    
ROWSORT:
    'row sort using 3 intersort keys
    'Sort rows of array using first column index, Key1
    For i = loR To hiR - 1
        For j = i + 1 To hiR
            'set < for descending, and > for ascending
            If bIsAscending Then
                condition1 = vW(i, Key1) > vW(j, Key1)
            Else
                condition1 = vW(i, Key1) < vW(j, Key1)
            End If
            If condition1 Then
                For c = loC To hiC
                    Temp = vW(i, c)
                    vW(i, c) = vW(j, c)
                    vW(j, c) = Temp
                Next
            End If
        Next
    Next
    If b2Used Then
        'Sort rows of array using second column index, Key2
        For i = loR To hiR - 1
            For j = i + 1 To hiR
                'if-condition avoids independence of second sort
                'note that a third stage would have THREE terms
                If vW(i, Key1) = vW(j, Key1) Then
                    'set < for descending, and > for ascending
                    If bIsAscending Then
                        condition1 = vW(i, Key2) > vW(j, Key2)
                    Else
                        condition1 = vW(i, Key2) < vW(j, Key2)
                    End If
                    If condition1 Then
                        For c = loC To hiC
                            Temp = vW(i, c)
                            vW(i, c) = vW(j, c)
                            vW(j, c) = Temp
                        Next
                    End If
                End If
            Next
        Next
    Else
        GoTo TRANSFERS
    End If
    If b3Used Then
        'Sort rows of array using third column index, Key3
        For i = loR To hiR - 1
            For j = i + 1 To hiR
                'if-condition avoids independence of second sort
                'note that a third stage would have THREE terms
                If vW(i, Key1) = vW(j, Key1) And vW(i, Key2) = vW(j, Key2) Then
                    'set < for descending, and > for ascending
                    If bIsAscending Then
                        condition1 = vW(i, Key3) > vW(j, Key3)
                    Else
                        condition1 = vW(i, Key3) < vW(j, Key3)
                    End If
                    If condition1 Then
                        For c = loC To hiC
                            Temp = vW(i, c)
                            vW(i, c) = vW(j, c)
                            vW(j, c) = Temp
                        Next
                    End If
                End If
            Next
        Next
    End If
    GoTo TRANSFERS
   
COLSORT:
    'column sort using 3 intersort keys
    'Sort columns of array using first row index, Key1
    For i = loC To hiC - 1
        For j = i + 1 To hiC
            'set < for descending, and > for ascending
            If bIsAscending Then
                condition1 = vW(Key1, i) > vW(Key1, j)
            Else
                condition1 = vW(Key1, i) < vW(Key1, j)
            End If
            If condition1 Then
                For c = loR To hiR
                    Temp = vW(c, i)
                    vW(c, i) = vW(c, j)
                    vW(c, j) = Temp
                Next
            End If
        Next
    Next
    If b2Used Then
        'Sort columns of array using second row index, Key2
        For i = loC To hiC - 1
            For j = i + 1 To hiC
                'if-condition avoids independence of second sort
                'note that a third stage would have THREE terms
                If vW(Key1, i) = vW(Key1, j) Then
                    'set < for descending, and > for ascending
                    If bIsAscending Then
                        condition1 = vW(Key2, i) > vW(Key2, j)
                    Else
                        condition1 = vW(Key2, i) < vW(Key2, j)
                    End If
                    If condition1 Then
                        For c = loR To hiR
                            Temp = vW(c, i)
                            vW(c, i) = vW(c, j)
                            vW(c, j) = Temp
                        Next
                    End If
                End If
            Next
        Next
    Else
        GoTo TRANSFERS
    End If
    If b3Used Then
        'Sort columns of array using third  row index, Key2
        For i = loC To hiC - 1
            For j = i + 1 To hiC
                'if-condition avoids independence of second sort
                'note that a third stage would have THREE terms
                If vW(Key1, i) = vW(Key1, j) And vW(Key2, i) = vW(Key2, j) Then
                    'set < for descending, and > for ascending
                    If bIsAscending Then
                        condition1 = vW(Key3, i) > vW(Key3, j)
                    Else
                        condition1 = vW(Key3, i) < vW(Key3, j)
                    End If
                    If condition1 Then
                        For c = loR To hiR
                            Temp = vW(c, i)
                            vW(c, i) = vW(c, j)
                            vW(c, j) = Temp
                        Next
                    End If
                End If
            Next
        Next
    End If
    GoTo TRANSFERS
    
TRANSFERS:
    'decide whether to return in vA or vR
    If Not bWasMissing Then
        'vR was the intended return array
        'so return vR leaving vA intact
        vR = vW
    Else:
        'vR is not intended
        'so reload vA with vR
        vA = vW
    End If
    
    'set return function value
    SortArr2D3Keys = True
    
End Function

