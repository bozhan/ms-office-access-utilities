VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeadColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class     : HeadColumn
' Author    : Ivanov, Bozhan
' Descr.    : Represents a range of a whole column and a seperate range for the
'             column header
'---------------------------------------------------------------------------------------
Option Explicit
Private searchOptions As Collection
Private headerNames As Variant
Private ws As Worksheet
Private rngSearch As Range
Private rngHead As Range
Private rngAfterHead As Range
Private rngLastCell As Range
Private usedRangeRows As Long
Private usedRangeCols As Long
Private df As String
Private newHeaderName As String

Private Sub Class_Initialize()
  Set searchOptions = New Collection
  headerNames = Array()
  df = ""
End Sub

Private Sub Class_Terminate()
  Set searchOptions = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : Init
' Description :
'---------------------------------------------------------------------------------------
Public Sub Init( _
  ByRef SearchDomain As Object, _
  ByVal HeadTitles As Variant, _
  Optional LookAt As XlLookAt = xlWhole, _
  Optional LookIn As XlFindLookIn = xlValues, _
  Optional SearchOrder As XlSearchOrder = xlByRows, _
  Optional SearchDirection As XlSearchDirection = xlNext, _
  Optional MatchCase As Boolean = False, _
  Optional SearchFormat As Boolean = False, _
  Optional dataType As Variant = vbNullString, _
  Optional NewHeaderValue As Variant = vbNullString _
)
  If TypeOf SearchDomain Is Excel.Worksheet Then
    Me.HeadSearchRange = SearchDomain.UsedRange
  ElseIf TypeOf SearchDomain Is Excel.ListObject Then
    Me.HeadSearchRange = SearchDomain.Range
  ElseIf TypeOf SearchDomain Is Excel.Range Then
    Me.HeadSearchRange = SearchDomain
  Else
    err.raise Number:=ErrCode.TypeMismatch, _
      Description:="Invalid domain object or parameter is nothing"
  End If
  
  Me.Worksheet = Me.HeadSearchRange.Worksheet
  appendTitles HeadTitles
  Me.SetSearchOptions LookAt, LookIn, SearchOrder, SearchDirection, MatchCase, SearchFormat
  Me.dataType = dataType
  Me.newHeader = NewHeaderValue
End Sub

Property Get newHeader() As String
  newHeader = newHeaderName
End Property

Property Let newHeader(ByVal Name As String)
  newHeaderName = Name
End Property

Property Get UsedRange() As Range
  Set UsedRange = Me.Worksheet.Range( _
    Me.Worksheet.Cells(Me.Head.Row, Me.Head.Column), _
    Me.Worksheet.Cells(Me.LastCell.Row, Me.LastCell.Column))
End Property

Property Get HeadSearchRange() As Range
  Set HeadSearchRange = rngSearch
End Property

Property Let HeadSearchRange(ByRef rng As Range)
  Set rngSearch = rng
End Property

Property Get Worksheet() As Worksheet
  If ws Is Nothing _
  And Not Me.HeadSearchRange Is Nothing Then
    Set ws = Me.HeadSearchRange.Worksheet
  End If
  
  Set Worksheet = ws
End Property

Property Let Worksheet(w As Worksheet)
  Set ws = w
End Property

Property Get Head() As Range
  If rngHead Is Nothing _
  And Not Me.HeadSearchRange Is Nothing Then
    Dim i As Integer
    For i = LBound(Me.Headers) To UBound(Me.Headers)
      Set rngHead = findFirst(Me.Headers(i), Me.HeadSearchRange)
      If Not rngHead Is Nothing Then
        Exit For
      End If
    Next i
  End If
  
  If rngHead Is Nothing Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="None of the provided column headers """ & Util.Arrays.ToString(Me.Headers) & """" & _
      " was found in the range " & Me.HeadSearchRange.Address & _
      " in worksheet """ & Me.Worksheet.Name & """"
  End If
  
  Set Head = rngHead
End Property

'---------------------------------------------------------------------------------------
' Procedure   : AfterHeadRange
' Description : in case the no data exist but the headers
'   we set the afterhead as the next cell after the header as default
'---------------------------------------------------------------------------------------
Property Get AfterHeadRange() As Range
  If rngAfterHead Is Nothing Then
    Dim maxRow As Long
    maxRow = Me.LastCell.Row
    If maxRow = Me.Head.Row Then
      maxRow = maxRow + 1 'set the next cell for default AfterHead
    End If
    Set rngAfterHead = Me.Worksheet.Range( _
      Me.Worksheet.Cells(Me.Head.Row + 1, Me.Head.Column), _
      Me.Worksheet.Cells(maxRow, Me.LastCell.Column))
  End If
  
  'Set AfterHeadRange = rngAfterHead
  
  Set AfterHeadRange = Me.Worksheet.Range( _
      Me.Worksheet.Cells(Me.Head.Row + 1, Me.Head.Column), _
      Me.Worksheet.Cells(Me.LastCell.Row, Me.LastCell.Column))
End Property

Property Get dataType() As String
    dataType = df
End Property

Property Let dataType(ByVal dType As String)
  df = dType
End Property

'---------------------------------------------------------------------------------------
' Procedure : LastCell
' Descr.    : returns the last used cell in the headcolumn used range with column depth
'   equal to the max depth of the worksheet used range
'   Used later to set the column depth of the AfterHead range we use Worksheet.
'   We adjust the depth of the HeadClumn beacuase we want all headcolumns with the
'   same worksheet to have the same depth
'---------------------------------------------------------------------------------------
Property Get LastCell() As Range
  If rngLastCell Is Nothing Then
    Set rngLastCell = getLastCell
  End If
  
  'using previous row and column count from used range
  'to determine if the used range has changed from the last time
  'rngLastCell was set, and if it needs to be recalculated
  If usedRangeRows <> Me.Worksheet.UsedRange.rows.count _
  Or usedRangeCols < Me.Worksheet.UsedRange.Columns.count Then
    Set rngLastCell = getLastCell
    Debug.Print " - recalc last cell - "
  End If
  
  Set LastCell = rngLastCell
End Property

Private Function getLastCell() As Range
  Dim Last As Range
  Set Last = Util.Excel.GetLastRowWithData(Me.Worksheet, Me.Worksheet.UsedRange)
  usedRangeRows = Me.Worksheet.UsedRange.rows.count
  usedRangeCols = Me.Worksheet.UsedRange.Columns.count
  
  Set getLastCell = Me.Worksheet.Cells( _
      Last.rows(Last.rows.count).Row, _
      Me.Head.Columns(Me.Head.Columns.count).Column)
End Function

Property Get Headers() As Variant
  Headers = headerNames
End Property

Public Sub AddHeaders(HeadTitles As Variant)
  appendTitles HeadTitles
End Sub

Private Sub appendTitles(HeadTitles As Variant)
  Dim al As New ArrayList
  'IsAllocated() returns false also when no array is provided
  If Util.Arrays.IsAllocated(HeadTitles) Then
    If Util.Arrays.IsOfSimpleDataTypes(HeadTitles) Then
      Dim l As Integer
      For l = LBound(HeadTitles) To UBound(HeadTitles)
        If CStr(HeadTitles(l)) <> "" Then
          al.Append CStr(HeadTitles(l))
        End If
      Next l
    End If
  ElseIf Util.Strings.IsSimpleDataType(HeadTitles) Then
    al.Append Split(CStr(HeadTitles), ";")
  End If
  
  headerNames = al.Values
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Find
' Descr.    : searches in the headcolumn range for the provided parameter
'   returns the found range or nothing if not found
'---------------------------------------------------------------------------------------
Public Function Find(ByVal searchValue As Variant) As Range
  Set Find = findFirst(searchValue, Me.AfterHeadRange)
End Function

Public Function FindAll(ByVal searchValue As Variant) As Collection
  Set FindAll = Util.Excel.FindAllInRange( _
    Me.AfterHeadRange, _
    FindWhat:=searchValue, _
    LookIn:=searchOptions("LookIn"), _
    LookAt:=searchOptions("LookAt"), _
    SearchOrder:=searchOptions("SearchOrder"), _
    MatchCase:=searchOptions("MatchCase"))
End Function

'---------------------------------------------------------------------------------------
' Procedure : findFirst
' Descr.    : returns the full range (also merged area) of the first match of string
'---------------------------------------------------------------------------------------
Private Function findFirst( _
  ByVal searchValue As Variant, _
  Optional ByRef HeadSearchRange As Range _
) As Range
On Error GoTo findFirst_Error
Dim rng As Object
Dim result As Object

  Set findFirst = Nothing
  Set rng = HeadSearchRange
  If rng Is Nothing Then Set rng = Me.Worksheet.UsedRange
  Set result = Util.Excel.GetMergedRange( _
    rng.Find( _
    What:=searchValue, _
    LookIn:=searchOptions("LookIn"), _
    LookAt:=searchOptions("LookAt"), _
    SearchOrder:=searchOptions("SearchOrder"), _
    SearchDirection:=searchOptions("SearchDirection"), _
    MatchCase:=searchOptions("MatchCase"), _
    SearchFormat:=searchOptions("SearchFormat")))
  
  Set findFirst = result
  
findFirst_Exit:
On Error Resume Next
  Set rng = Nothing
  Exit Function
  
findFirst_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure findFirst of module HeadColumnR" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume findFirst_Exit
End Function

Public Sub SetSearchOptions( _
  Optional LookAt As XlLookAt = xlWhole, _
  Optional LookIn As XlFindLookIn = xlValues, _
  Optional SearchOrder As XlSearchOrder = xlByRows, _
  Optional SearchDirection As XlSearchDirection = xlNext, _
  Optional MatchCase As Boolean = False, _
  Optional SearchFormat As Boolean = False _
)
  
  Set searchOptions = New Collection 'reset previously set parameters
  searchOptions.Add LookAt, "LookAt"
  searchOptions.Add LookIn, "LookIn"
  searchOptions.Add SearchOrder, "SearchOrder"
  searchOptions.Add SearchDirection, "SearchDirection"
  searchOptions.Add MatchCase, "MatchCase"
  searchOptions.Add SearchFormat, "SearchFormat"
End Sub

Public Sub Sort( _
  Optional ByVal order As XlSortOrder = xlAscending, _
  Optional ByVal Header As XlYesNoGuess = xlYes, _
  Optional ByVal expandToUsedRange As Boolean = True, _
  Optional ByRef tableStartRange As Range _
)
Dim startRange As Range
  
On Error GoTo Sort_Error

  If tableStartRange Is Nothing Then
    Set startRange = Me.Worksheet.Cells(Me.Head.Row, Me.UsedRange.Columns(1).Column)
  Else
    Set startRange = tableStartRange
  End If
  
  If expandToUsedRange Then
    Util.Excel.SortWorksheetByRange _
      Me.Worksheet, _
      Me.Head, _
      startRange, _
      order, _
      Header
  Else
    Util.Excel.SortRange _
      Me.AfterHeadRange, _
      order, _
      Header
  End If

Sort_Exit:
On Error Resume Next
Exit Sub

Sort_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Sort of module HeadColumn" & vbLf & _
    InfoErrMsg()
Resume Sort_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Filter
' Descr.    : up to 2 criteria for filtering a HeadColumn range
'---------------------------------------------------------------------------------------
Public Sub Filter( _
  Optional ByVal Criteria1 As Variant = Null, _
  Optional Operator As XlAutoFilterOperator = XlAutoFilterOperator.xlAnd, _
  Optional ByVal Criteria2 As Variant = Null, _
  Optional VisibleDropDown As Boolean = False, _
  Optional ByRef tableStartRange As Range _
)
Dim startRange As Range
On Error GoTo Filter_Error

  If tableStartRange Is Nothing Then
    Set startRange = Me.Worksheet.Cells(Me.Head.Row, Me.UsedRange.Columns(1).Column)
  Else
    Set startRange = tableStartRange
  End If
  
  Util.Excel.FilterWorksheetToCriteriaOnRange _
    Me.Worksheet, _
    Me.Head.Column, _
    Criteria1, _
    startRange, _
    Operator, _
    Criteria2, _
    VisibleDropDown

Filter_Exit:
On Error Resume Next
Exit Sub

Filter_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Filter of module HeadColumn" & vbLf & _
    InfoErrMsg()
Resume Filter_Exit
End Sub

Public Sub ReCalc()
  Me.Head = Nothing
  Me.AfterHeadRange = Nothing
End Sub

'TODO FIx bug when the column is part of a list and the result ends up in the same list
' will not input only unique values
' fix name to a more descriptive one like ..InRange
Public Function GetRangeOfUniqueValues( _
  Optional ByRef insertAfter As Range _
) As Range
On Error GoTo GetUniqueValues_Error
Dim r As Range

  Set GetRangeOfUniqueValues = Nothing
  
  If insertAfter Is Nothing Then
    Dim lastCellAtHeadRow As Range
    'Debug.Print wu.GetLastColumnWithData(Me.Worksheet).Column
    'Debug.Print Me.Head.row + 1
    Set lastCellAtHeadRow = Me.Worksheet.Cells( _
      Me.Head.Row + 1, Util.Excel.GetLastColumnWithData(Me.Worksheet).Column + 1)
    'Debug.Print lastCellAtHeadRow.Address
    lastCellAtHeadRow.offset(-1).Value = Me.Head.Value '"DISTINCT " &
    Set r = lastCellAtHeadRow.Resize( _
      Me.AfterHeadRange.rows.count, Me.AfterHeadRange.Columns.count)
  Else
    insertAfter.offset(0, 1).Insert
    Set r = insertAfter.offset(0, 1).Resize( _
      Me.AfterHeadRange.rows.count, Me.AfterHeadRange.Columns.count)
  End If
  'Debug.Print r.Address
  r.Value = Me.AfterHeadRange.Value
  r.RemoveDuplicates 1, xlGuess
  Set GetRangeOfUniqueValues = r.Resize(Util.Excel.GetLastRowWithData(Me.Worksheet, r).Row)
  'Debug.Print r.Address
GetUniqueValues_Exit:
On Error Resume Next
  Set r = Nothing
Exit Function

GetUniqueValues_Error:
Debug.Print "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure GetUniqueValues of HeadColumn:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume GetUniqueValues_Exit
  
End Function

'TODO just use Util.Arrays
Public Function GetUniqueValues() As Variant
Dim wb As Workbook
Dim tempWs As Worksheet
Dim result As ArrayList
Dim v As Variant
Dim alertStatus As Boolean
Dim updatingStatus As Boolean
On Error GoTo GetUniqueValues_Error

  Set result = New ArrayList
  Set wb = Me.Worksheet.Parent
  
  alertStatus = Me.Worksheet.Application.DisplayAlerts
  updatingStatus = Me.Worksheet.Application.ScreenUpdating
  Me.Worksheet.Application.DisplayAlerts = False
  Me.Worksheet.Application.ScreenUpdating = False
        
  Set tempWs = wb.Worksheets.Add(after:=Me.Worksheet)
  Dim r As Range
  Set r = tempWs.Range("A1").Resize(Me.AfterHeadRange.rows.count, Me.AfterHeadRange.Columns.count)
  
  r.Value = Me.AfterHeadRange.Value
  r.RemoveDuplicates 1, xlNo
  Set r = r.Resize(Util.Excel.GetLastRowWithData(tempWs, r).Row)
  v = r.Value
  
  Dim l As Long
  For l = LBound(v, 1) To UBound(v, 1)
    result.Append v(l, 1)
  Next l
  
  GetUniqueValues = result.Values
  
GetUniqueValues_Exit:
On Error Resume Next
  tempWs.Delete
  Me.Worksheet.Application.DisplayAlerts = alertStatus
  Me.Worksheet.Application.ScreenUpdating = updatingStatus
  Set wb = Nothing
  Set tempWs = Nothing
  Set r = Nothing
Exit Function

GetUniqueValues_Error:
Debug.Print "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure GetUniqueValues of HeadColumn:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume GetUniqueValues_Exit
  
End Function



