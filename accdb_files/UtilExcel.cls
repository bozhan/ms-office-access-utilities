VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : UtilExcel
' Author    : Ivanov, Bozhan
' Purpose   : Provides workbook Utility to extend functionality and management of
'   excel workbooks and worksheets
'---------------------------------------------------------------------------------------
' TODO modFormat:setConditionalFormatingForRange
' TODO to be excel object independent import all xl enums from Excel lib as private enums with the same name
Option Explicit

Public Enum XlInsertShiftDirection
  xlShiftToRight = -4161
  xlShiftDown = -4121
End Enum

Public Enum XlBorderWeight
  xlHairline = 1
  xlMedium = -4138
  xlThick = 4
  xlThin = 2
End Enum

Public Enum XlHAlign
  xlHAlignCenter = -4108
  xlHAlignCenterAcrossSelection = 7
  xlHAlignDistributed = -4117
  xlHAlignFill = 5
  xlHAlignGeneral = 1
  xlHAlignJustify = -4130
  xlHAlignLeft = -4131
  xlHAlignRight = -4152
End Enum

Public Enum XlVAlign
  xlVAlignBottom = -4107
  xlVAlignCenter = -4108
  xlVAlignDistributed = -4117
  xlVAlignJustify = -4130
  xlVAlignTop = -4160
End Enum

Public Enum XlLineStyle
  xlContinuous = 1
  xlDash = -4115
  xlDashDot = 4
  xlDashDotDot = 5
  xlDot = -4118
  xlDouble = -4119
  xlLineStyleNone = -4142
  xlSlantDashDot = 13
End Enum

Public Enum XlBordersIndex
  xlDiagonalDown = 5
  xlDiagonalUp = 6
  xlEdgeBottom = 9
  xlEdgeLeft = 7
  xlEdgeRight = 10
  xlEdgeTop = 8
  xlInsideHorizontal = 12
  xlInsideVertical = 11
End Enum

Private Sub Class_Initialize()
On Error GoTo Class_Initialize_Error
  
Class_Initialize_Exit:
  On Error Resume Next
Exit Sub

Class_Initialize_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Class_Initialize of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
  Resume Class_Initialize_Exit
End Sub

'aqcuire the ribbon pointer on ribbon load
Private Function getRibbon(ByRef RibbonPointer As Long) As IRibbonUI
  Dim rib As IRibbonUI
  Util.Windows.CopyMemoryReferenceAdrFromValueAdr rib, Temp("pmUIRibbonPointer"), 4
  Set getRibbon = rib
End Function

Public Function CreateWorkbook( _
  ByRef app As Excel.Application, _
  ByVal path As String, _
  Optional ByVal title As String = vbNullString, _
  Optional ByVal Subject As String = vbNullString _
) As Workbook
Dim wb As Workbook
On Error GoTo CreateWorkbook_Error
  If Not Util.File.IsWorkbookPath(path) Then
    err.raise 76, , "Invalid workbook path"
  End If
  
  Set wb = app.Workbooks.Add()
  With wb
    .title = title
    .Subject = Subject
    .SaveAs path
  End With
  
  Set CreateWorkbook = wb
    
CreateWorkbook_Exit:
On Error Resume Next
Exit Function

CreateWorkbook_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure CreateWorkbook of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
  Resume CreateWorkbook_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetHiddenColumnsIndeces
' Descr.    : Returns an array of all hidden columns in the given worksheet
'---------------------------------------------------------------------------------------
Public Function GetHiddenColumnsIndeces( _
  ByRef ws As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Variant
On Error GoTo GetHiddenColumnsIndeces_Error
  GetHiddenColumnsIndeces = Array()
  
  Dim rngToSearch As Range
  If SearchRange Is Nothing Then
    Set rngToSearch = ws.UsedRange
  Else
    Set rngToSearch = SearchRange
  End If
  
  Dim hiddenColumns As ArrayList
  Set hiddenColumns = New ArrayList
  
  Dim l As Long
  For l = 1 To rngToSearch.Columns.count
    If ws.Columns(l).Hidden Then
      hiddenColumns.Append l
    End If
  Next l
  
  GetHiddenColumnsIndeces = hiddenColumns.Values
  
GetHiddenColumnsIndeces_Exit:
On Error Resume Next
Exit Function

GetHiddenColumnsIndeces_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetHiddenColumnsIndeces of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume GetHiddenColumnsIndeces_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetHiddenRowsIndeces
' Descr.    : Returns an array of all hidden rows in the given worksheet
'---------------------------------------------------------------------------------------
Public Function GetHiddenRowsIndeces( _
  ByRef ws As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Variant
On Error GoTo GetHiddenRowsIndeces_Error
  GetHiddenRowsIndeces = Array()
  
  Dim rngToSearch As Range
  If SearchRange Is Nothing Then
    Set rngToSearch = ws.UsedRange
  Else
    Set rngToSearch = SearchRange
  End If

  Dim hiddenRows As ArrayList
  Set hiddenRows = New ArrayList
  
  Dim l As Long
  For l = 1 To rngToSearch.rows.count
    If ws.rows(l).Hidden Then
      hiddenRows.Append l
    End If
  Next l
  
  GetHiddenRowsIndeces = hiddenRows.Values
  
GetHiddenRowsIndeces_Exit:
On Error Resume Next
Exit Function

GetHiddenRowsIndeces_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetHiddenRowsIndeces of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume GetHiddenRowsIndeces_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetGroupedColumnsIndeces
' Descr.    : Returns the indeces of columns which have an outline level > 1
'---------------------------------------------------------------------------------------
Public Function GetGroupedColumnsIndeces( _
  ByRef ws As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Variant
  
  GetGroupedColumnsIndeces = Array()
  
  Dim rngToSearch As Range
  If SearchRange Is Nothing Then
    Set rngToSearch = ws.UsedRange
  Else
    Set rngToSearch = SearchRange
  End If

  Dim groupedColumns As ArrayList
  Set groupedColumns = New ArrayList
  
  Dim l As Long
  For l = 1 To rngToSearch.Columns.count
    If ws.Columns(l).OutlineLevel > 1 Then
      groupedColumns.Append l
    End If
  Next l
  
  GetGroupedColumnsIndeces = groupedColumns.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetGroupedRowsIndeces
' Descr.    : Returns the indeces of rows which have an outline level > 1
'---------------------------------------------------------------------------------------
Public Function GetGroupedRowsIndeces( _
  ByRef ws As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Variant
  
  GetGroupedRowsIndeces = Array()
  
  Dim rngToSearch As Range
  If SearchRange Is Nothing Then
    Set rngToSearch = ws.UsedRange
  Else
    Set rngToSearch = SearchRange
  End If

  Dim groupedRows As ArrayList
  Set groupedRows = New ArrayList
  
  Dim l As Long
  For l = 1 To rngToSearch.rows.count
    If ws.rows(l).OutlineLevel > 1 Then
      groupedRows.Append l
    End If
  Next l
  
  GetGroupedRowsIndeces = groupedRows.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : HideColumnsWithIndeces
' Descr.    : Hides columns in the refernced sheet with the given column indeces
'---------------------------------------------------------------------------------------
Public Sub HideColumnsWithIndeces( _
  ByRef ws As Worksheet, _
  ByRef columnIndeces As Variant _
)
Dim l As Long
Dim ae As ArrayList

  Set ae = New ArrayList
  ae.Append columnIndeces
  
  If Not ae.IsEmpty Then
    For l = ae.lb To ae.ub
      ws.Columns(ae(l)).Hidden = True
    Next l
  End If
  
  Set ae = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowColumnsWithIndeces
' Descr.    : Shows columns in the refernced sheet with the given column indeces
'---------------------------------------------------------------------------------------
Public Sub ShowColumnsWithIndeces( _
  ByRef ws As Worksheet, _
  ByRef columnIndeces As Variant _
)
Dim l As Long
Dim ae As ArrayList

  Set ae = New ArrayList
  ae.Append columnIndeces
  
  If Not ae.IsEmpty Then
    For l = ae.lb To ae.ub
      ws.Columns(ae(l)).Hidden = False
    Next l
  End If
  
  Set ae = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideWorksheet
' Descr.    : Hides a worksheet by name in the provided workbook
'---------------------------------------------------------------------------------------
Public Sub HideWorksheet( _
  ByRef wb As Workbook, _
  ByVal WorksheetName As String _
)
  wb.Worksheets(WorksheetName).Visible = XlSheetVisibility.xlSheetHidden
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : HideWorksheets
' Description : Hides worksheets matching provided names
'---------------------------------------------------------------------------------------
Public Sub HideWorksheets( _
  ByRef wb As Workbook, _
  ByVal worksheetNames As Variant _
)
  Dim al As New ArrayList
  al.Append worksheetNames
  
  Dim ws As Worksheet
  For Each ws In wb.Worksheets
    If al.Contains(ws.Name) Then ws.Visible = XlSheetVisibility.xlSheetHidden
  Next ws
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowWorksheet
' Descr.    : Shows a worksheet by name in the provided workbook
'---------------------------------------------------------------------------------------
Public Sub ShowWorksheet( _
  ByRef wb As Workbook, _
  ByVal WorksheetName As String _
)
  wb.Worksheets(WorksheetName).Visible = XlSheetVisibility.xlSheetVisible
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : ShowWorksheets
' Description : Shows worksheets matching provided names
'---------------------------------------------------------------------------------------
Public Sub ShowWorksheets( _
  ByRef wb As Workbook, _
  ByVal worksheetNames As Variant _
)
Dim al As New ArrayList
  al.Append worksheetNames
  
  Dim ws As Worksheet
  For Each ws In wb.Worksheets
    If al.Contains(ws.Name) Then ws.Visible = XlSheetVisibility.xlSheetVisible
  Next ws
End Sub

Public Function IsWorksheetVisible( _
  ByRef wb As Workbook, _
  ByRef WorksheetName As String _
) As Boolean
  IsWorksheetVisible = wb.Worksheets(WorksheetName).Visible
End Function

'---------------------------------------------------------------------------------------
' Procedure : HideRowsWithIndeces
' Descr.    : Hides rows in the refernced sheet with the given row indeces
'---------------------------------------------------------------------------------------
Public Sub HideRowsWithIndeces( _
  ByRef SearchWorksheet As Worksheet, _
  ByRef RowIndeces As Variant _
)
Dim l As Long
Dim ae As ArrayList

  Set ae = New ArrayList
  ae.Append RowIndeces
  
  If Not ae.IsEmpty Then
    For l = ae.lb To ae.ub
      SearchWorksheet.rows(ae(l)).Hidden = True
    Next l
  End If
  
  Set ae = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowRowsWithIndeces
' Descr.    : Shows rows in the refernced sheet with the given row indeces
'---------------------------------------------------------------------------------------
Public Sub ShowRowsWithIndeces( _
  ByRef SearchWorksheet As Worksheet, _
  ByRef RowIndeces As Variant _
)
Dim l As Long
Dim ae As ArrayList

  Set ae = New ArrayList
  ae.Append RowIndeces
  
  If Not ae.IsEmpty Then
    For l = ae.lb To ae.ub
      SearchWorksheet.rows(ae(l)).Hidden = False
    Next l
  End If
  
  Set ae = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : GetVisibleComments
' Description : Returns all visible comments in provided worksheet
'---------------------------------------------------------------------------------------
Public Function GetVisibleComments( _
  ByRef xlWorksheet As Worksheet _
) As Collection
  Dim res As New Collection
  Dim c As Comment
  For Each c In xlWorksheet.Comments
    If c.Visible Then res.Add c
  Next c
  Set GetVisibleComments = res
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetHiddenComments
' Description : Returns all hidden comments in provided worksheet
'---------------------------------------------------------------------------------------
Public Function GetHiddenComments( _
  ByRef xlWorksheet As Worksheet _
) As Collection
  Dim res As New Collection
  Dim c As Comment
  For Each c In xlWorksheet.Comments
    If Not c.Visible Then res.Add c
  Next c
  Set GetHiddenComments = res
End Function

'---------------------------------------------------------------------------------------
' Procedure   : ShowComments
' Description : Sets provided comments in collection to visible
'---------------------------------------------------------------------------------------
Public Sub ShowComments( _
  ByRef CommentsCollection As Collection _
)
  Dim c As Comment
  For Each c In CommentsCollection
    c.Visible = True
  Next c
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : HideComments
' Description : Sets provided comments in collection to not visible
'---------------------------------------------------------------------------------------
Public Sub HideComments( _
  ByRef CommentsCollection As Collection _
)
  Dim c As Comment
  For Each c In CommentsCollection
    c.Visible = False
  Next c
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetLastColumnWithData
' Descr.    : returns a range of the last column where data was found
'---------------------------------------------------------------------------------------
Public Function GetLastColumnWithData( _
  ByRef SearchWorksheet As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Range
On Error GoTo GetLastColumnWithData_Error
Dim rngByRows As Range
Dim rngByColumns As Range
Dim rngToSearch As Range
Dim c2 As Long
Dim c1 As Long

  If SearchRange Is Nothing Then
    Set rngToSearch = SearchWorksheet.Cells
  Else
    Set rngToSearch = SearchRange
  End If
  
  If SearchWorksheet Is Nothing Then err.raise 91, , "Worksheet object is not set"
  
  Set GetLastColumnWithData = rngToSearch.Cells(1, 1) ' set default return value
  
  Set rngByColumns = rngToSearch.Find( _
    What:="*", _
    after:=rngToSearch.Cells(1, 1), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False, _
    SearchFormat:=False)
  
  Set rngByRows = rngToSearch.Find( _
    What:="*", _
    after:=rngToSearch.Cells(1, 1), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False, _
    SearchFormat:=False)
  
  If Not rngByRows Is Nothing _
  And Not rngByColumns Is Nothing Then
    c1 = rngByColumns.Column
    c2 = rngByRows.Column
    
    If c1 < c2 Then 'check if a higher column number exists on the row variant
      Set GetLastColumnWithData = SearchWorksheet.Columns(c2)
    Else
      Set GetLastColumnWithData = SearchWorksheet.Columns(c1)
    End If
  End If
GetLastColumnWithData_Exit:
  On Error Resume Next
Exit Function

GetLastColumnWithData_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetLastColumnWithData of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
  Resume GetLastColumnWithData_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLastRowWithData
' Descr.    : returns a range of the last row where data was found
'---------------------------------------------------------------------------------------
Public Function GetLastRowWithData( _
  ByRef SearchWorksheet As Worksheet, _
  Optional ByRef SearchRange As Range _
) As Range
On Error GoTo GetLastRowWithData_Error
Dim rngByRows As Range
Dim rngByColumns As Range
Dim rngToSearch As Range
Dim r1 As Long
Dim r2 As Long
  
  If SearchRange Is Nothing Then
    Set rngToSearch = SearchWorksheet.Cells
  Else
    Set rngToSearch = SearchRange
  End If
  
 If SearchWorksheet Is Nothing Then err.raise 91, , "Worksheet object is not set"
  
  Set GetLastRowWithData = rngToSearch.Cells(1, 1) ' set default return value

  Set rngByRows = rngToSearch.Find( _
    What:="*", _
    after:=rngToSearch.Cells(1, 1), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False)
  
  Set rngByColumns = rngToSearch.Find( _
    What:="*", _
    after:=rngToSearch.Cells(1, 1), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False)

  If Not rngByRows Is Nothing _
  And Not rngByColumns Is Nothing Then
    r1 = rngByRows.Row
    r2 = rngByColumns.Row
    
    If r1 < r2 Then 'check if a higher row number exists on the column variant
      Set GetLastRowWithData = SearchWorksheet.rows(r2)
    Else
      Set GetLastRowWithData = SearchWorksheet.rows(r1)
    End If
  End If
GetLastRowWithData_Exit:
  On Error Resume Next
Exit Function

GetLastRowWithData_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetLastRowWithData of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
  Resume GetLastRowWithData_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetUsedRange
' Descr.    : returns the working range for the supplied worksheet
' Examples  : wu.GetUsedRange(Worksheet, rng) - the range from rng to max row and col
'   wu.GetUsedRange(Worksheet, Worksheet.Rows(rng.Row)) - range from A(rng.row) to max row col
'---------------------------------------------------------------------------------------
Public Function getUsedRange( _
  ByRef ws As Worksheet, _
  Optional ByRef fromStartRange As Range _
) As Range
Dim maxRow As Long, maxCol As Long
On Error GoTo GetUsedRange_Error

  Set getUsedRange = Nothing
  
  If ws Is Nothing Then
    err.raise 91, , "Worksheet or fromStartRange are not set"
  End If
  
  If fromStartRange Is Nothing Then
    Set fromStartRange = ws.Cells(1, 1)
  End If
  
  maxRow = Me.GetLastRowWithData(ws).Row
  maxCol = Me.GetLastColumnWithData(ws).Column
  
  If maxRow > fromStartRange.Row And maxCol > fromStartRange.Column Then
    Set getUsedRange = ws.Range(fromStartRange.Cells(1, 1), ws.Cells(maxRow, maxCol))
  ElseIf maxRow > fromStartRange.Row And maxCol <= fromStartRange.Column Then
    Set getUsedRange = ws.Range(fromStartRange.Cells(1, 1), ws.Cells(maxRow, 1))
  ElseIf maxRow <= fromStartRange.Row And maxCol > fromStartRange.Column Then
    Set getUsedRange = ws.Range(fromStartRange.Cells(1, 1), ws.Cells(1, maxCol))
  ElseIf maxRow <= fromStartRange.Row And maxCol <= fromStartRange.Column Then
    Set getUsedRange = ws.Cells(1, 1)
  End If
  

GetUsedRange_Exit:
On Error Resume Next
Exit Function

GetUsedRange_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetUsedRange of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
Resume GetUsedRange_Exit
  
End Function

Public Sub SortWorksheetByRange( _
  ByRef ws As Worksheet, _
  ByRef key As Range, _
  Optional ByRef tableStartRange As Range, _
  Optional ByVal order As XlSortOrder = xlAscending, _
  Optional ByVal Header As XlYesNoGuess = xlYes _
)
Dim wr As Range

  Set wr = Me.getUsedRange(ws, tableStartRange)
'  Debug.Print wr.Address
  wr.Sort _
    Key1:=key, _
    order1:=order, _
    Header:=Header
End Sub

Public Sub SortRange( _
  ByRef rngToSort As Range, _
  Optional ByVal order As XlSortOrder = xlAscending, _
  Optional ByVal Header As XlYesNoGuess = xlYes _
)
  rngToSort.Sort _
    Key1:=rngToSort.Cells(1, 1), _
    order1:=order, _
    Header:=Header
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FilterWorksheetToCriteriaOnRange
' Descr.    :
'---------------------------------------------------------------------------------------
Public Sub FilterWorksheetToCriteriaOnRange( _
  ByRef ws As Worksheet, _
  ByVal fieldNum As Long, _
  Optional ByVal Criteria1 As Variant = Null, _
  Optional ByRef tableStartRange As Range, _
  Optional op As XlAutoFilterOperator = XlAutoFilterOperator.xlAnd, _
  Optional Criteria2 As Variant = Null, _
  Optional VisibleDropDown As Boolean = False _
)
Dim rng As Range
  
On Error GoTo FilterWorksheetToCriteriaOnRange_Error

  Set rng = Me.getUsedRange(ws, tableStartRange)
  
  rng.AutoFilter _
    Field:=fieldNum, _
    Criteria1:=Criteria1, _
    Operator:=op, _
    Criteria2:=Criteria2, _
    VisibleDropDown:=VisibleDropDown

FilterWorksheetToCriteriaOnRange_Exit:
On Error Resume Next
Exit Sub

FilterWorksheetToCriteriaOnRange_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure FilterWorksheetToCriteriaOnRange of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
Resume FilterWorksheetToCriteriaOnRange_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : GetWorksheetsByName
' Description : returns all worksheets in the provided workbook that match the given
'   name(s). Parameter names can be a string or an array of strings
'---------------------------------------------------------------------------------------
Public Function GetWorksheetsByName( _
  ByRef wb As Workbook, _
  ByVal names As Variant _
) As Collection
Dim result As New Collection
Dim ws As Worksheet
Dim al As New ArrayList
  al.Append names
  
  For Each ws In wb.Worksheets
    If al.Contains(ws.Name) Then result.Add ws, ws.Name
  Next ws
  
  Set GetWorksheetsByName = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetWorksheetNames
' Descr.    : returns an array of strings - the names of all available worksheets in
'             a supplied workbook.
'---------------------------------------------------------------------------------------
Public Function GetWorksheetNames( _
  ByRef wb As Workbook _
) As Variant
Dim result As New ArrayList
Dim ws As Worksheet

  For Each ws In wb.Worksheets
    result.Append ws.Name
  Next ws
  
  GetWorksheetNames = result.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetListObjects
' Descr.    : returns a collection of all available list objects in a supplied workbook
'---------------------------------------------------------------------------------------
Public Function GetListObjects( _
  ByRef wb As Workbook _
) As Collection
Dim result As New Collection
Dim ws As Worksheet
Dim lo As ListObject
  
  For Each ws In wb.Worksheets
    For Each lo In ws.ListObjects
      result.Add lo, lo.Name
    Next lo
  Next ws
  
  Set GetListObjects = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetListObjectsInSheetName
' Descr.    : returns a collection of all available list objects in a supplied workbook
'   that are in a given worksheet by name
'---------------------------------------------------------------------------------------
Public Function GetListObjectsInSheetName( _
  ByRef wb As Workbook, _
  ByVal sheetName As Variant _
) As Collection
Dim result As New Collection
Dim ws As Worksheet, wss As Collection
Dim lo As ListObject
  
  Set wss = GetWorksheetsByName(wb, sheetName)
  
  For Each ws In wss
    For Each lo In ws.ListObjects
      result.Add lo, lo.Name
    Next lo
  Next ws
  
  Set GetListObjectsInSheetName = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetListObjectNames
' Descr.    : returns an array of all available list object names in a supplied workbook
'---------------------------------------------------------------------------------------
Public Function GetListObjectNames( _
  ByRef wb As Workbook _
) As Variant
Dim result As New ArrayList
Dim los As Collection
Dim lo As ListObject
  
  Set los = GetListObjects(wb)
    
  For Each lo In los
    result.Append lo.Name
  Next lo
  
  GetListObjectNames = result.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetListObjectNames
' Descr.    : returns an array of all available list object names in a supplied workbook
'             The parameter sheetName can also be an array of multiple names
'---------------------------------------------------------------------------------------
Public Function GetListObjectNamesInSheetName( _
  ByRef wb As Workbook, _
  ByVal sheetName As Variant _
) As Variant
Dim result As New ArrayList
Dim los As Collection
Dim lo As ListObject
  
  Set los = GetListObjectsInSheetName(wb, sheetName)
    
  For Each lo In los
    result.Append lo.Name
  Next lo
  
  GetListObjectNamesInSheetName = result.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetListObjectByName
' Description : returns a list object matching the provided name
'   Function with Workbook parameter is necessary,
'   beacause ListObject are children of the Worksheet class
'---------------------------------------------------------------------------------------
Public Function GetListObjectByName( _
  ByRef wb As Workbook, _
  ByVal ListName As String _
) As ListObject
  
  Set GetListObjectByName = Nothing
      
  Dim los As Collection, lo As ListObject
  Set los = GetListObjects(wb)
  For Each lo In los
    If StrComp(lo.Name, ListName, vbTextCompare) = 0 Then
      Set GetListObjectByName = lo
      Exit Function
    End If
  Next lo
End Function

Public Function IsListObjectName( _
  ByRef wb As Workbook, _
  ByVal listObjectName As String _
) As Boolean
  IsListObjectName = Not (GetListObjectByName(wb, listObjectName) Is Nothing)
End Function

Public Function GetListObjectColumnName( _
  ByRef lo As ListObject, _
  ByVal columnIndex As Long _
) As String
  GetListObjectColumnName = ""
  If lo Is Nothing Then Exit Function
  
  Dim lc As ListColumn
  For Each lc In lo.ListColumns
    If lc.Index = columnIndex Then
      GetListObjectColumnName = lc.Name
      Exit Function
    End If
  Next lc
End Function

Public Function GetListObjectColumnIndex( _
  ByRef lo As ListObject, _
  ByVal columnName As String _
) As Long
  GetListObjectColumnIndex = 0
  If lo Is Nothing Then Exit Function
  
  Dim lc As ListColumn
  For Each lc In lo.ListColumns
    If StrComp(columnName, lc.Name, vbTextCompare) = 0 Then
      GetListObjectColumnIndex = lc.Index
      Exit Function
    End If
  Next lc
End Function

Public Function GetWorkbooksFromPaths( _
  ByRef xlApp As Object, _
  ByVal workbookPaths As Variant _
) As Variant
Dim wbs As ArrayList
Dim paths As ArrayList
Dim l As Long

  Set wbs = New ArrayList
  Set paths = New ArrayList
  
  If Not Util.Arrays.IsAllocated(workbookPaths) Then Exit Function
  
  paths.Append workbookPaths
  
  For l = paths.lb To paths.ub
    If Util.File.FileExists(paths(l)) Then
      wbs.Append Me.GetWorkbookFromPath(xlApp, paths(l)) ' xlApp.Workbooks.Open(paths(l), False, False, , , , True)
    End If
  Next l

  GetWorkbooksFromPaths = wbs.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetWorkbookFromPath
' Author    : Ivanov, Bozhan
' Purpose   : opens and returns a workbook from the provided path and assignes it to the
'   provided application object
'---------------------------------------------------------------------------------------
Public Function GetWorkbookFromPath( _
  ByRef xlApp As Object, _
  ByVal workbookPath As String, _
  Optional ByVal Password As Variant, _
  Optional ByVal writeResPassword As Variant _
) As Workbook

  Set GetWorkbookFromPath = xlApp.Workbooks.Open( _
    fileName:=workbookPath, _
    Password:=Password, _
    writeResPassword:=writeResPassword, _
    UpdateLinks:=False, _
    readOnly:=False, _
    IgnoreReadOnlyRecommended:=True)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : IsWorkbookOpen
' Description : Returns TRUE if a workbook is open in provided application
'   Accepts workbook name or workbook object as input
'---------------------------------------------------------------------------------------
Public Function IsWorkbookOpen( _
  ByRef xlApp As Excel.Application, _
  ByVal xlWorkbook As Variant _
) As Boolean
  
  If xlApp Is Nothing Then
    IsWorkbookOpen = False
    Exit Function
  End If
  
  Dim wb As Workbook
  If VarType(xlWorkbook) = vbObject Then
    On Error Resume Next
    Set wb = xlApp.Workbooks(xlWorkbook.Name)
    On Error GoTo 0
  ElseIf Util.Strings.IsSimpleDataType(xlWorkbook) Then
    On Error Resume Next
    Set wb = xlApp.Workbooks(CStr(xlWorkbook))
    On Error GoTo 0
  End If
  
  IsWorkbookOpen = (Not wb Is Nothing)
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsRangeName
' Descr.    : returns true if range name exist in workbook
'---------------------------------------------------------------------------------------
Public Function IsRangeName( _
  ByRef wb As Workbook, _
  rangeName As String _
) As Boolean
Dim n As Name
  
  IsRangeName = False
  
  For Each n In wb.names
    If StrComp(n.Name, rangeName, vbTextCompare) = 0 Then
      IsRangeName = True
      Exit Function
    End If
  Next n
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRangeNameIndex
' Descr.    : returns index of the range name, but -1 if range name was not found
'---------------------------------------------------------------------------------------
Public Function GetRangeNameIndex( _
  ByRef wb As Workbook, _
  ByVal rangeName As String _
) As Long
  
  GetRangeNameIndex = -1
  If Me.IsRangeName(wb, rangeName) Then
    GetRangeNameIndex = wb.names(rangeName).Index
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRangeByName
' Descr.    : returns range of a range name
'---------------------------------------------------------------------------------------
Public Function GetRangeByName( _
  ByRef wb As Workbook, _
  ByVal rangeName As String _
) As Range

  Set GetRangeByName = Nothing
  
  If Me.IsRangeName(wb, rangeName) Then
    Set GetRangeByName = wb.names(rangeName).RefersToRange
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : MergeRanges
' Descr.    : A Union operation that accepts parameters that are Nothing
'---------------------------------------------------------------------------------------
Public Function MergeRanges( _
  ParamArray rngs() As Variant _
) As Range
Dim l As Long
Dim rng As Range
On Error GoTo MergeRanges_Error
  
  For l = LBound(rngs) To UBound(rngs)
    If IsObject(rngs(l)) _
    And Not rngs(l) Is Nothing _
    And TypeOf rngs(l) Is Excel.Range Then
      If rng Is Nothing Then
        Set rng = rngs(l)
      Else
        Set rng = rng.Application.Union(rng, rngs(l))
      End If
    End If
  Next l
  
  Set MergeRanges = rng

MergeRanges_Exit:
  On Error Resume Next
Exit Function

MergeRanges_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure MergeRanges of module WorkbookUtil" & vbLf & _
    InfoErrMsg()
  Resume MergeRanges_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : CloseWorbooksInApplication
' Descr.    : Closes all open workbook in an excel application
'---------------------------------------------------------------------------------------
Public Sub CloseWorbooksInApplication( _
  ByRef xlApp As Excel.Application, _
  Optional ByVal bSave As Boolean = False _
)
Dim wb As Workbook
On Error Resume Next

  For Each wb In xlApp.Workbooks
    wb.Close bSave
  Next wb
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CloseWorkbooks
' Descr.    : Closes all provided as parameters workbooks
'   Workbooks can be an array of Excel.Workbook elements of a collection of Workbooks
'---------------------------------------------------------------------------------------
Public Sub CloseWorkbooks( _
  ByRef Workbooks As Variant, _
  Optional ByVal bSave As Boolean = False _
)
  If IsArray(Workbooks) Then
    Dim l As Long
    For l = LBound(Workbooks) To UBound(Workbooks)
      If TypeOf Workbooks(l) Is Excel.Workbook _
      And Not Workbooks(l) Is Nothing Then
        Workbooks(l).Close bSave
      End If
    Next l
  ElseIf IsObject(Workbooks) Then
    If TypeOf Workbooks Is Collection Then
      Dim wb As Workbook
      For Each wb In Workbooks
        If TypeOf wb Is Excel.Workbook _
        And Not wb Is Nothing Then
          wb.Close bSave
        End If
      Next wb
    End If
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertCheckBoxAndGetIndex
' Descr.    : returns the index of the inserted checkboxes
'   On error of any kind -1 will be returned
'---------------------------------------------------------------------------------------
Public Function InsertCheckBoxAndGetIndex( _
  ByRef rng As Range, _
  Optional ByVal Caption As String = vbNullString, _
  Optional ByVal text As String = vbNullString, _
  Optional ByVal checked As Long = Excel.xlOn, _
  Optional ByVal namePrefix As String = vbNullString _
) As Long
Dim chk As MSforms.CheckBox
Dim ws As Worksheet
On Error GoTo InsertCheckBoxAndGetIndex_Error

  InsertCheckBoxAndGetIndex = -1

  Set ws = rng.Worksheet
  Set chk = ws.CheckBoxes.Add( _
    rng.Left, _
    rng.Top, _
    rng.Width, _
    17.25)

  On Error Resume Next
  With chk
    .Caption = Caption
    .text = text
    .Value = checked
    .Display3DShading = False
    .Name = namePrefix & .Name
    '.LinkedCell = rng.Address(external:=True)
  End With

  On Error GoTo InsertCheckBoxAndGetIndex_Error
  InsertCheckBoxAndGetIndex = chk.Index

InsertCheckBoxAndGetIndex_Exit:
On Error Resume Next
  Set ws = Nothing
  Set chk = Nothing
Exit Function

InsertCheckBoxAndGetIndex_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure InsertCheckBoxAndGetIndex of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume InsertCheckBoxAndGetIndex_Exit
End Function

Public Function GetCaptionsOfCheckedCheckBoxesInWorksheet( _
  ByRef ws As Worksheet _
) As Variant
Dim chk As MSforms.CheckBox
Dim ae As ArrayList
  
  Set ae = New ArrayList
  
  For Each chk In ws.CheckBoxes
    If chk.Value = 1 Then
      ae.Append chk.Caption
    End If
  Next chk
  
  GetCaptionsOfCheckedCheckBoxesInWorksheet = ae.Values
  Set ae = Nothing
End Function

Public Sub DeleteCheckboxesInWorksheet(ByRef ws As Worksheet)
Dim chk As MSforms.CheckBox
  For Each chk In ws.CheckBoxes
    chk.Delete
  Next chk
End Sub

Public Sub SelectCheckboxesInWorksheet(ByRef ws As Worksheet)
Dim chk As MSforms.CheckBox
  For Each chk In ws.CheckBoxes
    chk.Value = 1
  Next chk
End Sub

Public Sub DeselectCheckboxesInWorksheet(ByRef ws As Worksheet)
Dim chk As MSforms.CheckBox
  For Each chk In ws.CheckBoxes
    chk.Value = 0
  Next chk
End Sub

Public Function GetMergedRange(rng As Range) As Range
On Error GoTo GetMergedRange_Exit
  
  Set GetMergedRange = rng
  If rng.MergeCells Then
    Set GetMergedRange = rng.MergeArea
  End If
  
GetMergedRange_Exit:
End Function

Public Sub TurnOffScreenUpdatingAndAlerts( _
  Optional ByRef xlApp As Excel.Application _
)
  If xlApp Is Nothing Then Set xlApp = ThisWorkbook.Application
  With xlApp
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .AskToUpdateLinks = False
    .Cursor = xlNorthwestArrow
  End With
End Sub

Public Sub TurnOnScreenUpdatingAndAlerts( _
  Optional ByRef xlApp As Excel.Application _
)
  If xlApp Is Nothing Then Set xlApp = ThisWorkbook.Application
  With xlApp
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .AskToUpdateLinks = True
    .Cursor = xlDefault
  End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsShape
' Descr.    : returns true if shape exists in worksheet
'---------------------------------------------------------------------------------------
Public Function IsShape(ws As Worksheet, shapeName As String) As Boolean
Dim s As Shape

On Error Resume Next
  IsShape = False
  Set s = ws.Shapes(shapeName)

  If Not s Is Nothing Then
    IsShape = True
    Set s = Nothing
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetShapeText
' Descr.    : returns text value for a given shape name
'---------------------------------------------------------------------------------------
Public Function GetShapeText(ByRef ws As Worksheet, shapeName As String) As String
  GetShapeText = vbNullString
  
  If Me.IsShape(ws, shapeName) Then
    GetShapeText = CStr(ws.Shapes(shapeName).TextFrame.Characters.text)
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetShapeText
' Descr.    : sets text value for a given shape name
'---------------------------------------------------------------------------------------
Public Sub SetShapeText(ByRef ws As Worksheet, ByVal shapeName As String, ByVal text As String)
  If Me.IsShape(ws, shapeName) Then
    ws.Shapes(shapeName).TextFrame.Characters.text = text
  End If
End Sub

Public Function FindFirstInWorkbook( _
  ByRef SearchWorkbook As Workbook, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Range
Dim ws As Worksheet
  
  For Each ws In SearchWorkbook.Worksheets
    Set FindFirstInWorkbook = FindFirstInWorksheet( _
    SearchWorksheet:=ws, _
    FindWhat:=FindWhat, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    MatchCase:=MatchCase, _
    BeginsWith:=BeginsWith, _
    EndsWith:=EndsWith, _
    SubStringCompareMethod:=SubStringCompareMethod)
    
    If Not FindFirstInWorkbook Is Nothing Then
      Exit Function
    End If
  Next ws
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindFirstInWorksheets
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindFirstInWorksheets( _
  ByRef WorksheetsArray As Variant, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Range
On Error GoTo FindFirstInWorksheets_Error
  
  If Not Util.Arrays.IsAllocated(WorksheetsArray) Then err.raise Number:=ErrCode.SubscriptOutOfRange, _
    Description:="Array not allocated"
  If Util.Arrays.IsEmpty(WorksheetsArray) Then err.raise Number:=ErrCode.SubscriptOutOfRange, _
    Description:="Array is empty"
  
  Dim l As Long
  For l = LBound(WorksheetsArray) To UBound(WorksheetsArray)
    If IsObject(WorksheetsArray(l)) _
    And TypeOf WorksheetsArray(l) Is Excel.Worksheet Then
      
      Dim ws As Worksheet
      Set ws = WorksheetsArray(l)
      
      Set FindFirstInWorksheets = FindFirstInWorksheet( _
        SearchWorksheet:=ws, _
        FindWhat:=FindWhat, _
        LookIn:=LookIn, _
        LookAt:=LookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase, _
        BeginsWith:=BeginsWith, _
        EndsWith:=EndsWith, _
        SubStringCompareMethod:=SubStringCompareMethod)
    
      If Not FindFirstInWorksheets Is Nothing Then
        Exit Function
      End If
    End If
    
  Next l

FindFirstInWorksheets_Exit:
On Error Resume Next
  Set ws = Nothing
Exit Function

FindFirstInWorksheets_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure FindFirstInWorksheets of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume FindFirstInWorksheets_Exit
End Function

Public Function FindFirstInWorksheet( _
  ByRef SearchWorksheet As Worksheet, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Range
Dim UsedRange As Range

  Set UsedRange = Me.getUsedRange(SearchWorksheet)
  
  Set FindFirstInWorksheet = FindFirstInRange( _
    SearchRange:=UsedRange, _
    FindWhat:=FindWhat, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    MatchCase:=MatchCase, _
    BeginsWith:=BeginsWith, _
    EndsWith:=EndsWith, _
    SubStringCompareMethod:=SubStringCompareMethod)
  
  Set UsedRange = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindFirstInRange
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindFirstInRange( _
  ByRef SearchRange As Range, _
  ByVal FindWhat As Variant, _
  Optional ByVal after As Range, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Range
On Error GoTo FindFirstInRange_Error
  Set FindFirstInRange = Nothing
 
  If BeginsWith <> vbNullString _
  Or EndsWith <> vbNullString Then
    LookAt = xlPart
  End If
  
  If FindWhat = vbNullString _
  Or IsNull(FindWhat) _
  Or IsArray(FindWhat) Then
    err.raise Number:=ErrCode.TypeMismatch, _
      Description:="FindWhat was either nothing or of wrong type {" & VarType(FindWhat) & "}"
  End If
  
  If after Is Nothing Then
    Set after = Me.GetLastCellInRange(SearchRange)
  End If

  Set FindFirstInRange = SearchRange.Find( _
    What:=FindWhat, _
    after:=after, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    MatchCase:=MatchCase)

FindFirstInRange_Exit:
On Error Resume Next
Exit Function

FindFirstInRange_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure FindFirstInRange of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume FindFirstInRange_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLastCellInRange
' Descr.    : Returns the very last cell in a given range independent of the presence of
'  data content.
'---------------------------------------------------------------------------------------
Public Function GetLastCellInRange(SearchRange As Range) As Range
Dim area As Range
Dim maxRow As Long
Dim maxCol As Long
  ' Get the max rows and columns of the range to find the last cell
  For Each area In SearchRange.Areas
    With area
      If .Cells(.Cells.count).Row > maxRow Then
        maxRow = .Cells(.Cells.count).Row
      End If
      If .Cells(.Cells.count).Column > maxCol Then
        maxCol = .Cells(.Cells.count).Column
      End If
    End With
  Next area
  
  Set GetLastCellInRange = SearchRange.Worksheet.Cells(maxRow, maxCol)
End Function

' TODO refine/refactor
'Private Function findLastCellWithDataInRangeByRow(ByRef rng As Range) As Range
'
'End Function
'
'Private Function findLastCellWithDataInRangeByColumn(ByRef rng As Range) As Range
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : FindAllInRange
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindAllInRange( _
  ByRef SearchRange As Range, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Collection
On Error GoTo FindAllInRange_Error
Dim addrs As ArrayList
Dim result As Collection
  
  Set result = New Collection
  Set addrs = New ArrayList
  If (FindWhat = "") Then GoTo FindAllInRange_Exit
  Dim nextFound As Range
  
  Do
    Set nextFound = FindFirstInRange( _
      SearchRange:=SearchRange, _
      FindWhat:=FindWhat, _
      after:=nextFound, _
      LookIn:=LookIn, _
      LookAt:=LookAt, _
      SearchOrder:=SearchOrder, _
      MatchCase:=MatchCase, _
      BeginsWith:=BeginsWith, _
      EndsWith:=EndsWith, _
      SubStringCompareMethod:=SubStringCompareMethod)
    
    If nextFound Is Nothing Then Exit Do
    If addrs.Contains(nextFound.Address) Then Exit Do
    
    result.Add nextFound, nextFound.Address
    addrs.Append nextFound.Address
  Loop While True
  
FindAllInRange_Exit:
On Error Resume Next
  Set FindAllInRange = result
  Set result = Nothing
  Set addrs = Nothing
Exit Function

FindAllInRange_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure FindAllInRange of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume FindAllInRange_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindAllInWorksheet
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindAllInWorksheet( _
  ByRef SearchWorksheet As Worksheet, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Collection
Dim UsedRange As Range
Dim result As Collection
On Error GoTo FindAllInWorksheet_Exit
  Set FindAllInWorksheet = Nothing

  Set UsedRange = Me.getUsedRange(SearchWorksheet)
    
  Set result = FindAllInRange( _
    SearchRange:=UsedRange, _
    FindWhat:=FindWhat, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    MatchCase:=MatchCase, _
    BeginsWith:=BeginsWith, _
    EndsWith:=EndsWith, _
    SubStringCompareMethod:=SubStringCompareMethod)
  
FindAllInWorksheet_Exit:
On Error Resume Next
  Set FindAllInWorksheet = result
  Set UsedRange = Nothing
  Set result = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindAllInWorksheets
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindAllInWorksheets( _
  ByRef WorksheetsArray As Variant, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Collection
On Error GoTo FindAllInWorksheets_Error
Dim result As Collection
Dim wsRanges As Collection

  If Not Util.Arrays.IsAllocated(WorksheetsArray) Then err.raise Number:=ErrCode.SubscriptOutOfRange, _
    Description:="Array not allocated"
  If Util.Arrays.IsEmpty(WorksheetsArray) Then err.raise Number:=ErrCode.SubscriptOutOfRange, _
    Description:="Array is empty"
  
  Set result = New Collection
  
  Dim l As Long
  For l = LBound(WorksheetsArray) To UBound(WorksheetsArray)
    If IsObject(WorksheetsArray(l)) _
    And TypeOf WorksheetsArray(l) Is Excel.Worksheet Then
      
      Dim ws As Worksheet
      Set ws = WorksheetsArray(l)
  
      Set wsRanges = FindAllInWorksheet( _
        SearchWorksheet:=ws, _
        FindWhat:=FindWhat, _
        LookIn:=LookIn, _
        LookAt:=LookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase, _
        BeginsWith:=BeginsWith, _
        EndsWith:=EndsWith, _
        SubStringCompareMethod:=SubStringCompareMethod)
        
        If wsRanges.count > 0 Then
          Set result = GetJoinCollections(result, wsRanges)
        End If
    End If
  Next l
  
FindAllInWorksheets_Exit:
On Error Resume Next
  Set FindAllInWorksheets = result
  Set ws = Nothing
  Set wsRanges = Nothing
  Set result = Nothing
Exit Function

FindAllInWorksheets_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure FindAllInWorksheets of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume FindAllInWorksheets_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindAllInWorkbook
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function FindAllInWorkbook( _
  ByRef SearchWorkbook As Workbook, _
  ByVal FindWhat As Variant, _
  Optional ByVal LookIn As XlFindLookIn = xlValues, _
  Optional ByVal LookAt As XlLookAt = xlWhole, _
  Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
  Optional ByVal MatchCase As Boolean = False, _
  Optional ByVal BeginsWith As String = vbNullString, _
  Optional ByVal EndsWith As String = vbNullString, _
  Optional ByVal SubStringCompareMethod As VbCompareMethod = vbTextCompare _
) As Collection
Dim ws As Worksheet
Dim result As Collection
Dim wsRanges As Collection

  Set result = New Collection
  
  For Each ws In SearchWorkbook.Worksheets
    Set wsRanges = FindAllInWorksheet( _
      SearchWorksheet:=ws, _
      FindWhat:=FindWhat, _
      LookIn:=LookIn, _
      LookAt:=LookAt, _
      SearchOrder:=SearchOrder, _
      MatchCase:=MatchCase, _
      BeginsWith:=BeginsWith, _
      EndsWith:=EndsWith, _
      SubStringCompareMethod:=SubStringCompareMethod)
      
      If wsRanges.count > 0 Then
        Set result = GetJoinCollections(result, wsRanges)
      End If
  Next ws
  
FindAllInWorksheet_Exit:
On Error Resume Next
  Set FindAllInWorkbook = result
  Set ws = Nothing
  Set wsRanges = Nothing
  Set result = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetColumnLetter
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function GetColumnLetter(ByRef rng As Range) As String
Dim ws As Worksheet
On Error GoTo GetColumnLetter_Error

  GetColumnLetter = vbNullString
  
  Set ws = rng.Worksheet
  GetColumnLetter = Split(ws.Cells(1, rng.Column).Address(True, False), "$")(0)
GetColumnLetter_Exit:
On Error Resume Next
Exit Function

GetColumnLetter_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetColumnLetter of module WorkbookUtil" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume GetColumnLetter_Exit
End Function
'---------------------------------------------------------------------------------------
' Procedure : GetRangeFromListObjectOnCondition
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function GetRangeFromListObjectOnCondition( _
  ByRef sourceTableRange As Range, _
  ByVal conditionColumnHeader As String, _
  ByVal conditionValue As Variant, _
  ByVal searchColumnHeader As String _
) As Range
Dim condHC As HeadColumn, searchHC As HeadColumn
  
On Error GoTo GetRangeFromListObjectOnCondition_Error

  Set GetRangeFromListObjectOnCondition = Nothing
  Set condHC = New HeadColumn
  Set searchHC = New HeadColumn
  
  condHC.Init sourceTableRange, conditionColumnHeader
  searchHC.Init sourceTableRange, searchColumnHeader
  
  If Not condHC.Find(conditionValue) Is Nothing Then
    Set GetRangeFromListObjectOnCondition = searchHC.AfterHeadRange(condHC.Find(conditionValue).Row - condHC.Head.Row)
  Else
    err.raise Number:=ErrCode.ObjectVariableNotSet, _
      Description:="No cell matching the condition value " & _
      """" & conditionValue & """" & _
      " was found in worksheet " & _
      """" & sourceTableRange.Worksheet.Name & """" & _
      " in the range " & _
      """" & sourceTableRange.Address & """"
  End If

GetRangeFromListObjectOnCondition_Exit:
On Error Resume Next
Exit Function

GetRangeFromListObjectOnCondition_Error:
  Set GetRangeFromListObjectOnCondition = Nothing
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetRangeFromListObjectOnCondition of module UtilExcel" _
    & vbLf & InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume GetRangeFromListObjectOnCondition_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetRangeInListObjectOnCondition
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function SetRangeInListObjectOnCondition( _
  ByRef sourceTableRange As Range, _
  ByVal conditionColumnHeader As String, _
  ByVal conditionValue As Variant, _
  ByVal searchColumnHeader As String, _
  ByVal updateRangeValue As Variant _
) As Boolean
Dim rng As Range
On Error GoTo SetRangeInListObjectOnCondition_Error
  
  SetRangeInListObjectOnCondition = False
  
  Set rng = Me.GetRangeFromListObjectOnCondition( _
    sourceTableRange, _
    conditionColumnHeader, _
    conditionValue, _
    searchColumnHeader)
    
  If Not rng Is Nothing Then
    rng.Value = updateRangeValue
    SetRangeInListObjectOnCondition = True
  End If

SetRangeInListObjectOnCondition_Exit:
Exit Function

SetRangeInListObjectOnCondition_Error:
  SetRangeInListObjectOnCondition = False
Resume SetRangeInListObjectOnCondition_Exit
End Function

Public Sub CopyHyperlinks( _
  ByRef src As Range, _
  ByRef dst As Range _
)
  If src.Hyperlinks.count > 0 Then
    Dim hLink As Hyperlink
    For Each hLink In src.Hyperlinks
      dst.Hyperlinks.Add _
        Anchor:=dst, _
        Address:=hLink.Address
    '    TextToDisplay:=src.value
    Next hLink
    
  End If
End Sub

'todo make desc and clear variable names
Public Sub ImportDataFromDSVFile( _
  ByVal FilePath As String, _
  ByVal colDelimiter As String, _
  ByRef resultRange As Range, _
  Optional ByVal rowDelimiter As String = vbCrLf _
)
Dim c As String
Dim v As Variant

  c = Util.File.GetFileContent(FilePath)
  v = Util.Strings.getSplitDSV(c, colDelimiter, rowDelimiter)
  Set resultRange = resultRange.Resize(UBound(v, 1) + 1, UBound(v, Util.Arrays.GetDimensions(v)) + 1)
  resultRange = v
End Sub
Public Sub ImportDataFromDSVFileOnCondition( _
  ByVal FilePath As String, _
  ByVal colDelimiter As String, _
  ByRef resultRange As Range, _
  ByVal condition As Variant, _
  Optional ByVal rowDelimiter As String = vbCrLf _
)
Dim c As String
Dim v As Variant

  c = Util.File.GetFileContent(FilePath)
  v = Util.Strings.getSplitByColumnsDSVOnCondition(c, colDelimiter, condition, rowDelimiter)
  Set resultRange = resultRange.Resize(UBound(v, 1) + 1, UBound(v, Util.Arrays.GetDimensions(v)) + 1)
  resultRange = v
End Sub
Public Sub CropWorksheet(ByRef ws As Excel.Worksheet)
Dim wst As Worksheet
Dim wb As Workbook
  Set wb = ws.Parent
  For Each wst In wb.Worksheets
    If wst.Name <> ws.Name Then wst.Delete
  Next wst
End Sub

Public Function CreateWorksheet( _
  ByRef wb As Workbook, _
  Optional ByVal sheetName As String = vbNullString _
) As Worksheet
Dim ws As Worksheet

  Set CreateWorksheet = Nothing
  
  If sheetName = vbNullString _
  Or IsWorksheetName(wb, sheetName) Then
    sheetName = sheetName & "_" & Util.DateTime.GetDateTimeStamp()
  End If
  
  Set ws = wb.Worksheets.Add()
  ws.Name = sheetName
  Set CreateWorksheet = ws
  Set ws = Nothing
End Function

Public Function IsWorksheetName( _
  ByRef wb As Workbook, _
  ByVal sheetName As String _
) As Boolean
IsWorksheetName = (GetWorksheetsByName(wb, sheetName).count > 0)
End Function

Public Function SortCustomListObjectsByColumn( _
  ByRef los As Collection, _
  ByVal sortColumnName As String _
) As Variant
Dim lo As ListObject
Dim lc As ListColumn
  TurnOffScreenUpdatingAndAlerts
  For Each lo In los
    On Error Resume Next
    Set lc = lo.ListColumns(sortColumnName)
    If Not lc Is Nothing Then
      lo.Sort.SortFields.Clear
      lo.Sort.SortFields.Add _
        lc.DataBodyRange, _
        XlSortOn.xlSortOnValues, _
        XlSortOrder.xlAscending
      lo.Sort.Apply
    Else
      Debug.Print "Error: no list column with name " & sortColumnName & " was found in list object " & lo.Name
    End If
    On Error GoTo 0
  Next lo
  TurnOnScreenUpdatingAndAlerts
End Function

Public Function SortListObjectsByColumn( _
  ByRef los As ListObjects, _
  ByVal sortColumnName As String _
) As Variant
Dim lo As ListObject
Dim lc As ListColumn
  TurnOffScreenUpdatingAndAlerts
  For Each lo In los
    On Error Resume Next
    Set lc = lo.ListColumns(sortColumnName)
    If Not lc Is Nothing Then
      lo.Sort.SortFields.Clear
      lo.Sort.SortFields.Add _
        lc.DataBodyRange, _
        XlSortOn.xlSortOnValues, _
        XlSortOrder.xlAscending
      lo.Sort.Apply
    Else
      Debug.Print "Error: no list column with name " & sortColumnName & " was found in list object " & lo.Name
    End If
    On Error GoTo 0
  Next lo
  TurnOnScreenUpdatingAndAlerts
End Function

Public Function GetOLEObjectCount( _
  ByRef ws As Worksheet, _
  ByVal objType As XlOLEType _
) As Long
Dim count As Long
Dim ole As oleObject
  count = 0
  For Each ole In ws.OLEObjects
    If ole.OLEType = objType Then count = count + 1
  Next ole
  GetOLEObjectCount = count
End Function

Public Function GetAssignedOLEObjectToRange( _
  ByRef ws As Worksheet, _
  ByVal targetAddress As String _
) As oleObject
Dim ole As oleObject
  Set GetAssignedOLEObjectToRange = Nothing
  For Each ole In ws.OLEObjects
    If StrComp(ole.LinkedCell, targetAddress, vbTextCompare) = 0 Then
      Set GetAssignedOLEObjectToRange = ole
      Exit Function
    End If
  Next ole
End Function

Public Function ExistsAssignedOLEObjectToRange( _
  ByRef ws As Worksheet, _
  targetAddress As String _
) As Boolean
  Dim ole As oleObject
  Set ole = GetAssignedOLEObjectToRange(ws, targetAddress)
  If ole Is Nothing Then
    ExistsAssignedOLEObjectToRange = False
  Else
    ExistsAssignedOLEObjectToRange = True
  End If
End Function

Public Function CreateComboBoxForCell( _
  ByRef rng As Range, _
  ByRef cboSettings As OLEObjectSetting _
) As oleObject
  Set CreateComboBoxForCell = Nothing
  If rng.count <> 1 Then _
    err.raise Number:=ErrCode.SubscriptOutOfRange, Description:="Provided range argument must be a single cell."
  
  Dim ws As Worksheet, cbo As oleObject
  Set ws = rng.Worksheet
  Set cbo = GetAssignedOLEObjectToRange(ws, rng.Address)
  
  If cbo Is Nothing Then
    Set cbo = ws.OLEObjects.Add( _
      ClassType:=cboSettings.ClassType, _
      fileName:=cboSettings.IconFileName, _
      Link:=cboSettings.FileLink, _
      DisplayAsIcon:=cboSettings.DisplayAsIcon, _
      IconFileName:=cboSettings.IconFileName, _
      IconIndex:=cboSettings.IconIndex, _
      IconLabel:=cboSettings.IconLabel, _
      Left:=cboSettings.Left, _
      Top:=cboSettings.Top, _
      Width:=cboSettings.Width, _
      Height:=cboSettings.Height)
  End If
  
  With cbo
    .LinkedCell = cboSettings.LinkedCell
    .Height = cboSettings.Height
    .Width = cboSettings.Width
    .Top = cboSettings.Top
    .Left = cboSettings.Left
    .Enabled = cboSettings.Enabled
    .Visible = cboSettings.Visible
    .ListFillRange = cboSettings.ListFillRange
  End With
  
  Set CreateComboBoxForCell = cbo
End Function

Public Sub PromptMacroSecuritySettings(ByRef app As Excel.Application)
  app.CommandBars.ExecuteMso ("MacroSecurity")
End Sub

'TODO Revise Subprocedure
Public Sub ImportTextFileContentIntoNewWorksheet(FName As String, sep As String)
Dim RowNdx As Long
Dim ColNdx As Integer
Dim TempVal As Variant
Dim WholeLine As String
Dim pos As Integer
Dim NextPos As Integer
Dim SaveColNdx As Integer
Dim wb As Workbook, ws As Worksheet
Dim FilePath As String
Dim wu As UtilExcel

  Set wb = ThisWorkbook
  Set ws = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count))
  Set wu = New UtilExcel
  
  TurnOffScreenUpdatingAndAlerts
  
  Open FName For Input Access Read As #1
  RowNdx = 1
  
  While Not EOF(1)
    Line Input #1, WholeLine
    
    If Right(WholeLine, 1) <> sep Then
      WholeLine = WholeLine & sep
    End If
    
    ColNdx = 1
    pos = 1
    NextPos = InStr(pos, WholeLine, sep)
    
    While NextPos >= 1
      TempVal = mid(WholeLine, pos, NextPos - pos)
      ws.Cells(RowNdx, ColNdx).Value = TempVal
      
      pos = NextPos + 1
      ColNdx = ColNdx + 1
      NextPos = InStr(pos, WholeLine, sep)
    Wend
    RowNdx = RowNdx + 1
  Wend
  Close #1
  
  TurnOnScreenUpdatingAndAlerts
End Sub

'TODO move to a UtilCollection class Module
Private Function GetJoinCollections( _
  ParamArray collections() As Variant _
) As Collection
Dim col As Variant
Dim obj As Variant
  
  Set GetJoinCollections = New Collection
  
  For Each col In collections
    If TypeOf col Is Collection Then
      For Each obj In col
        GetJoinCollections.Add obj
      Next obj
    End If
  Next col
End Function

''TODO revise call of GetListObjectNames constraints will be an array("", ..)
'Public Function GetListObjectsByNameMatching( _
'  ByRef wb As Workbook, _
'  ByVal match As String, _
'  Optional matchType As XlLookAt, _
'  Optional ByVal worksheetNameFilter As Variant _
') As Collection
'Dim los As Collection
'Dim col As New Collection
'
'  set los = GetListObjects(wb
'
'  loNames = GetListObjectNames(wb, worksheetNameFilter)
'
'  Dim i As Integer
'  For i = LBound(loNames) To UBound(loNames)
'    If matchType = xlPart Then
'      If InStr(1, loNames(i), match, vbTextCompare) > 0 Then
'        col.Add GetListObjectByName(wb, loNames(i)), loNames(i)
'      End If
'    ElseIf matchType = xlWhole Then
'      If StrComp(loNames(i), match, vbTextCompare) = 0 Then
'        col.Add GetListObjectByName(wb, loNames(i)), loNames(i)
'      End If
'    End If
'  Next i
'
'  Set GetListObjectsByNameMatching = col
'End Function
