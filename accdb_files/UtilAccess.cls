VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event linkCreateTable(ByVal tableName As String)
Public Event linkCreateTables(ByVal Index As Long, ByVal count As Long)
Public Event relinkTable(ByVal tableName As String)
Public Event relinkTables(ByVal Index As Long, ByVal count As Long)

Public Enum AcObjectType
  Form = -32768
  Macro = -32766
  Reports = -32764
  Module = -32761
  users = -32758
  Database_Document = -32757
  Data_Access_Pages = -32756
  Table_Local_Access = 1
  Database = 2
  Containers = 3
  Table_Linked_ODBC = 4
  queries = 5
  Table_Linked_Access = 6
  SubDataSheets = 8
End Enum

Public Enum AcLookAt
  Whole = 1
  Part = 2
End Enum

Dim db As Database

Private Sub Class_Initialize()
  Set db = CurrDb()
End Sub

Private Sub Class_Terminate()
  Set db = Nothing
End Sub

'TODO implement a EMultiLookup FROM provided recordset
'TODO Replace the global EmultiLooup with this version

'---------------------------------------------------------------------------------------
' Procedure : EMultiLookup
' Purpose   : returns recordset result in an array form
'   based on constructed query from provided paramters
'   Returns single variant(string) value if one record is returned for one simple field
'   Returns a 1D array of variant values if 1 record is returened with 1 multi-value field
'   Returns a 2D array in case of multiple records and/or fields
'   1D are the fields and 2D are the rows or records
'   The first index of the result contains the names of the fields
'   e.g. result(1,0) = FieldName1, result(2,0) = FieldName2
'   The field indes of the result starts from 0 so:
'   e.g. result(1,1) will return field 1 value in record 1
'        result(3,3) will return value of field 3 in record 3
'   Multivalued fields will be returned as an array of variant values
'   e.g. if result(2,3) is multi-valued field then result(2,3)(0) will be 1st value
'   Return Type: String or 1D: Array(Variant) or 2D: Array(Variant, Variant)
'     Mutil-valued fields return: Array(Variant, Array(), ...)
'---------------------------------------------------------------------------------------
'Public Function EMultiLookup( _
'  fieldNames As Variant, _
'  domain As String, _
'  Optional Criteria As String = "", _
'  Optional Top1 As Boolean = True, _
'  Optional OrderClause As String = "", _
'  Optional OrderAsc As Boolean = True _
') As Variant
'Dim sql As String
'Dim result As Variant
'On Error GoTo EMultiLookup_Error
'  result = Null
'  Dim db As Database, rs As Recordset
'  Set db = CurrDb
'  sql = getSQLFromParameters(fieldNames, domain, Criteria, Top1, OrderClause, OrderAsc)
'  Set rs = db.OpenRecordset(sql)
'
'  If (rs.EOF And rs.BOF) Then GoTo EMultiLookup_Exit
'
'  Dim multiNames As New ArrayList
'  multiNames.Append getMultiValueFieldNamesFromRecordset(rs)
'
'  rs.MoveLast
'  rs.MoveFirst
'  'case 1 - one record, one field
'  If rs.RecordCount = 1 And rs.fields.count = 1 Then
'    If multiNames.Contains(rs(0)) Then
'      result = getMultiValueFieldAsArray(rs(0))
'    Else
'      result = rs(0)
'    End If
'  'case 2 - multiple records, one or many fields
'  Else
'    ReDim result(1 To rs.fields.count, 0 To rs.RecordCount)
'    Dim fi As Integer, ri As Long, rsIdx As Integer
'    ri = 1
'
'    'assign field names to index 0 of result
'    For fi = LBound(result, 1) To UBound(result, 1)
'      rsIdx = fi - LBound(result, 1)
'      result(fi, 0) = rs(rsIdx).name
'    Next fi
'
'    Do While Not rs.EOF
'      For fi = LBound(result, 1) To UBound(result, 1)
'        rsIdx = fi - LBound(result, 1)
'        If multiNames.Contains(rs(rsIdx).name) Then
'          result(fi, ri) = getMultiValueFieldAsArray(rs(rsIdx))
'        Else
'          result(fi, ri) = rs(rsIdx).value
'        End If
'      Next fi
'      rs.MoveNext
'      ri = ri + 1
'    Loop
'  End If
'
'EMultiLookup_Exit:
'On Error Resume Next
'  rs.Close
'  Set rs = Nothing
'  Set db = Nothing
'  EMultiLookup = result
'Exit Function
'
'EMultiLookup_Error:
'err.Raise Number:=err.Number, _
'    Description:="Error in procedure EMultiLookup of UtilAccess:Module" & vbLf & _
'    err.Description & vbLf & InfoErrMsg()
'Resume EMultiLookup_Exit
'End Function

'---------------------------------------------------------------------------------------
' Procedure : getMultiValueFieldAsArray
' Purpose   : returns a field value as a 1D array of variant
'---------------------------------------------------------------------------------------
Public Function getMultiValueFieldAsArray(ByRef fld As DAO.Field) As Variant
Dim mv As New ArrayList
Dim mrs As Recordset
On Error GoTo getMultiValueFieldAsArray_Error

  Set mrs = fld.Value
  If Not (mrs.EOF And mrs.BOF) Then
    Do While Not mrs.EOF
      mv.Append mrs(0).Value
      mrs.MoveNext
    Loop
  End If
  
  If Not mv.IsEmpty Then
    getMultiValueFieldAsArray = mv.Values
  Else
    getMultiValueFieldAsArray = Null
  End If
  
getMultiValueFieldAsArray_Exit:
On Error Resume Next
  mrs.Close
  Set mrs = Nothing
Exit Function

getMultiValueFieldAsArray_Error:
err.raise Number:=err.Number, _
    Description:="Error in procedure getMultiValueFieldAsArray of UtilAccess:Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume getMultiValueFieldAsArray_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getMultiValueFieldNamesFromRecordset
' Purpose   : returns array of names of field with complex data type (multi-value fields)
'---------------------------------------------------------------------------------------
Public Function getMultiValueFieldNamesFromRecordset(ByRef rs As Recordset) As Variant
Dim res As ArrayList
On Error GoTo getMultiValueFieldNamesFromRecordset_Error

  Set res = New ArrayList
  If (rs.EOF And rs.BOF) Then GoTo getMultiValueFieldNamesFromRecordset_Exit
  
  rs.MoveFirst
  Dim f As DAO.Field
  For Each f In rs.fields
    'multi-value complex types are 102 to 109
    If f.Type > 101 And f.Type < 110 Then
      res.Append f.Name
    End If
  Next f
  
getMultiValueFieldNamesFromRecordset_Exit:
On Error Resume Next
  getMultiValueFieldNamesFromRecordset = res.Values
Exit Function

getMultiValueFieldNamesFromRecordset_Error:
err.raise Number:=err.Number, _
    Description:="Error in procedure getMultiValueFieldNamesFromRecordset of UtilAccess:Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume getMultiValueFieldNamesFromRecordset_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getSQLFromParameters
' Purpose   : returns a query built from the provided function paramters
'---------------------------------------------------------------------------------------
Public Function getSQLFromParameters( _
  ByVal fieldNames As Variant, _
  ByVal domain As String, _
  Optional ByVal WhereClause As String = "", _
  Optional ByVal Top1 As Boolean = True, _
  Optional ByVal OrderClause As String = "", _
  Optional ByVal OrderAsc As Boolean = True _
) As String
Dim sql As String, fns As Variant
On Error GoTo getSQLFromParameters_Error
  
  getSQLFromParameters = ""
  sql = "SELECT "
  If Top1 Then sql = sql & "TOP 1 "
  
  'constructiong input variant
  If VarType(fieldNames) = vbString Then
    fns = Array(fieldNames)
  ElseIf VarType(fieldNames) >= vbArray Then
    fns = fieldNames
  Else
    err.raise Number:=ErrCode.InvalidProcedureCall, Description:="Faulty fieldNames parameter"
  End If
  
  If Util.Arrays.IsEmpty(fns) Then
    err.raise Number:=ErrCode.ArgumentNotOptional, Description:="Missing fieldNames"
  End If
  
  'constructing and safe bracketizing field names in select, where and order clauses
  Dim i As Integer
  For i = LBound(fns) To UBound(fns)
    sql = sql & Util.Strings.bracketizeSquare(fns(i))
    sql = sql & ", "
    WhereClause = Util.Strings.bracketizeSquare_MatchInText(WhereClause, fns(i))
    OrderClause = Util.Strings.bracketizeSquare_MatchInText(OrderClause, fns(i))
  Next i
  sql = Util.Strings.RemoveExcessSeperators(sql, ",")
  
  'adding in the domain (table)
  If Len(domain) <= 0 Then
    err.raise Number:=ErrCode.ArgumentNotOptional, Description:="Missing domain"
  End If
  
  sql = sql & " FROM "
  sql = sql & Util.Strings.bracketizeSquare(domain)
  WhereClause = Util.Strings.bracketizeSquare_MatchInText(WhereClause, domain)
  OrderClause = Util.Strings.bracketizeSquare_MatchInText(OrderClause, domain)
  
  'adding in if any defined WhereClause and order clauses
  If (Len(WhereClause) > 0) Then sql = sql & " WHERE " & WhereClause
  If (Len(OrderClause) > 0) Then sql = sql & " ORDER BY " & OrderClause
  sql = sql & ";"
  
  getSQLFromParameters = sql
  
getSQLFromParameters_Exit:
On Error Resume Next
Exit Function

getSQLFromParameters_Error:
err.raise Number:=err.Number, _
    Description:="Error in procedure getSQLFromParameters of UtilAccess:Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume getSQLFromParameters_Exit
End Function

Private Sub squareBracketizeNewSQLElement( _
  ByRef sql As String, _
  ByVal text As String, _
  ByRef WhereClause As String, _
  ByRef OrderClause As String _
)

  

End Sub


Public Sub listTables(Optional includeSysTables As Boolean = False)
Dim tdf As DAO.TableDef
  For Each tdf In CurrDb().TableDefs
    If Not ((Left(tdf.Name, Len("MSys")) = "MSys") _
      And Not includeSysTables) Then
      Debug.Print tdf.Name
    End If
  Next tdf
End Sub

Public Function deleteRecords( _
  ByVal fields As String, _
  ByVal domain As String, _
  ByVal condition As String _
) As Boolean
On Error GoTo deleteRecords_Error
  deleteFromWhere fields, domain, condition
  deleteRecords = True
deleteRecords_Exit:
On Error Resume Next
Exit Function

deleteRecords_Error:
  deleteRecords = False
  MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
    "in procedure deleteRecords of module GlobalProcedures:Module" & vbLf & _
    InfoErrMsg(), vbExclamation, "Error occurred..."
Resume deleteRecords_Exit
End Function

Public Function isFieldNameInRecordset( _
  ByRef rst As Recordset, _
  ByVal fieldName As String _
) As Boolean
Dim f As Field
  
  isFieldNameInRecordset = False
  For Each f In rst.fields
    If StrComp(f.Name, fieldName, vbTextCompare) = 0 Then
      isFieldNameInRecordset = True
      Exit Function
    End If
  Next f
End Function

Public Function areFieldNamesInRecordset( _
  ByRef rst As Recordset, _
  ByVal fieldNames As Variant _
) As Boolean
Dim i As Integer
Dim names As Variant
    
    areFieldNamesInRecordset = True
    names = getRecordsetFieldNames(rst)
    
    For i = LBound(fieldNames) To UBound(fieldNames)
      If Not Util.Arrays.Contains(names, fieldNames(i), vbTextCompare) Then
        areFieldNamesInRecordset = False
        Exit Function
      End If
    Next i
End Function

Public Function getMissingRecordsetFieldNames( _
  ByRef rst As Recordset, _
  ByVal fieldNames As Variant _
) As Variant
Dim i As Integer
Dim names As Variant
Dim al As ArrayList
    
    getMissingRecordsetFieldNames = Array()
    Set al = New ArrayList
    names = getRecordsetFieldNames(rst)
    
    For i = LBound(fieldNames) To UBound(fieldNames)
      If Not Util.Arrays.Contains(names, fieldNames(i), vbTextCompare) Then
        al.Append fieldNames(i)
      End If
    Next i
    
    getMissingRecordsetFieldNames = al.Values
End Function

Public Function getRecordsetFieldsValues( _
  ByRef rst As Recordset, _
  ByVal fieldNames As Variant _
) As Collection
Dim fields As Collection
Dim fieldValues As Collection
Dim f As Field
  
  Set getRecordsetFieldsValues = New Collection
  If Util.Arrays.IsEmpty(fieldNames) Then Exit Function
  
  Set fields = getRecordsetFieldsByName(rst, fieldNames)
  Set fieldValues = New Collection
  
  For Each f In fields
    fieldValues.Add New ArrayList, f.Name
  Next f
  
  With rst
    If Not .EOF And .BOF Then Exit Function
    
    .MoveFirst
    Do Until .EOF
      For Each f In fields
        fieldValues(f.Name).Append .fields(f.Name).Value
      Next f
      .MoveNext
    Loop
  End With
  
  Set getRecordsetFieldsValues = fieldValues
End Function

Public Function getRecordsetFieldValues( _
  ByRef rst As Recordset, _
  ByVal fieldName As String _
) As Variant
Dim al As ArrayList
Dim f As Field

  getRecordsetFieldValues = Array()
  Set al = New ArrayList
  Set f = getRecordsetFieldByName(rst, fieldName)
  
  With rst
    If .EOF And .BOF Then Exit Function
    .MoveFirst
    Do Until .EOF
      al.Append .fields(f.Name).Value
      .MoveNext
    Loop
  End With
    
  getRecordsetFieldValues = al.Values
End Function

'Public Function getRecordsetFieldValuesOnCondition( _
'  ByRef rst As Recordset, _
'  ByVal fieldName As String, _
'  ByVal conditionFieldName As String, _
'  ByVal conditionFieldValue As String, _
'  Optional ByVal LookAt As AcLookAt = AcLookAt.Whole, _
'  Optional ByVal negate As Boolean = False, _
'  Optional ByVal compMethod As VbCompareMethod = VbCompareMethod.vbTextCompare _
') As Variant
'Dim al As ArrayList
'Dim f As Field
'Dim cf As Field
'
'  getRecordsetFieldValuesOnCondition = Array()
'  Set al = New ArrayList
'  Set f = getRecordsetFieldByName(rst, fieldName)
'  Set cf = getRecordsetFieldByName(rst, conditionFieldName)
'
'  With rst
'    If .EOF And .BOF Then Exit Function
'    .MoveFirst
'    Do Until .EOF
'      If Util.Strings.getCompareResultOnConditions( _
'        .fields(cf.name).value, _
'        conditionFieldValue, _
'        LookAt, _
'        negate, _
'        compMethod) Then
'          al.Append .fields(f.name).value
'      End If
'      .MoveNext
'    Loop
'  End With
'
'  getRecordsetFieldValuesOnCondition = al.Values
'End Function
  
Public Function getRecordsetFieldsByName( _
  ByRef rst As Recordset, _
  ByVal fieldNames As Variant _
) As Collection
Dim fields As Collection
Dim i As Integer

  Set getRecordsetFieldsByName = New Collection
  Set fields = New Collection
  
  For i = LBound(fieldNames) To UBound(fieldNames)
    fields.Add getRecordsetFieldByName(rst, fieldNames(i))
  Next i
  
  Set getRecordsetFieldsByName = fields
End Function

Public Function getRecordsetFieldByName( _
  ByRef rst As Recordset, _
  ByVal fieldName As String _
) As Field
Dim f As Field

  Set getRecordsetFieldByName = Nothing
  If Not Util.Access.isFieldNameInRecordset(rst, fieldName) Then _
    err.raise Number:=ErrCode.PropertyOrMethodNotFound, _
      Description:="The fieldname """ & fieldName & """ was not found in the provided Recordset"
  
  For Each f In rst.fields
    If StrComp(f.Name, fieldName, vbTextCompare) = 0 Then
      Set getRecordsetFieldByName = f
      Exit Function
    End If
  Next f
End Function

Public Function getRecordsetFieldNames( _
  ByRef rst As Recordset _
) As Variant
Dim f As Field
Dim al As ArrayList
  
  getRecordsetFieldNames = Array()
  Set al = New ArrayList
  
  For Each f In rst.fields
    al.Append f.Name
  Next f
  
  getRecordsetFieldNames = al.Values
End Function

Public Function exportRecordsetToExcel( _
  ByRef rst As Recordset, _
  ByVal exportFieldNames As Variant, _
  ByVal path As String, _
  Optional ByRef xlApp As Excel.Application, _
  Optional ByVal title As String = vbNullString, _
  Optional ByVal Subject As String = vbNullString, _
  Optional ByVal sheetName As String = vbNullString, _
  Optional ByVal Show As Boolean = True, _
  Optional ByVal cropSheet As Boolean = True _
) As Worksheet
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
On Error GoTo exportRecordsetToExcel_Error

  Set exportRecordsetToExcel = Nothing
  
  checkProvidedFieldNames rst, exportFieldNames
  checkExcelExportWorkbookPath path
  
  If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = False
  Util.Excel.TurnOffScreenUpdatingAndAlerts xlApp
  
  Set wb = getNewWorkbook(xlApp, path, title, Subject)
  Set ws = Util.Excel.CreateWorksheet(wb, sheetName)
  
  If cropSheet Then Util.Excel.CropWorksheet ws
  writeRecordsToWorksheet rst, exportFieldNames, ws
  
  wb.Save
  If Show Then
    xlApp.Visible = True
    Util.Excel.TurnOnScreenUpdatingAndAlerts xlApp
  End If
  
  Set exportRecordsetToExcel = ws
  
exportRecordsetToExcel_Exit:
On Error Resume Next
  Set ws = Nothing
Exit Function

exportRecordsetToExcel_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure exportRecordsetToExcel of AccessUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume exportRecordsetToExcel_Exit
End Function

Public Function exportRecordsetToWorkbook( _
  ByRef rst As Recordset, _
  ByVal exportFieldNames As Variant, _
  ByRef wb As Excel.Workbook, _
  Optional ByVal sheetName As String = vbNullString, _
  Optional ByVal Show As Boolean = True, _
  Optional ByVal cropSheet As Boolean = True _
) As Worksheet
Dim ws As Excel.Worksheet
Dim xlApp As Excel.Application
On Error GoTo exportRecordsetToWorkbook_Error

  Set exportRecordsetToWorkbook = Nothing
  Set xlApp = wb.Application
  xlApp.Visible = False
  Util.Excel.TurnOffScreenUpdatingAndAlerts xlApp
  
  Set ws = Util.Excel.CreateWorksheet(wb, sheetName)
  If cropSheet Then Util.Excel.CropWorksheet ws
  writeRecordsToWorksheet rst, exportFieldNames, ws
  
  wb.Save
  If Show Then
    xlApp.Visible = True
    Util.Excel.TurnOnScreenUpdatingAndAlerts xlApp
  End If
  Set exportRecordsetToWorkbook = ws
  
exportRecordsetToWorkbook_Exit:
On Error Resume Next
  Set ws = Nothing
Exit Function

exportRecordsetToWorkbook_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure exportRecordsetToWorkbook of AccessUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume exportRecordsetToWorkbook_Exit
End Function

Private Sub writeRecordsToWorksheet( _
  ByRef rst As Recordset, _
  ByVal exportFieldNames As Variant, _
  ByRef ws As Worksheet _
)
Dim f As Field
  Dim startRange As Range
  Set startRange = ws.Cells(1, 1)
  setExcelExportColumnHeaders ws, startRange, exportFieldNames
  setExcelExportRecordsetData rst, ws, startRange.offset(1, 0), exportFieldNames
End Sub

Private Sub setExcelExportRecordsetData( _
  ByRef rst As Recordset, _
  ByRef ws As Worksheet, _
  ByRef startRange As Range, _
  ByVal exportFieldNames As Variant _
)
Dim i As Integer
  With rst
    If .EOF And .BOF Then Exit Sub
    .MoveFirst
    
    Do Until .EOF
    
      For i = LBound(exportFieldNames) To UBound(exportFieldNames)
        ws.Cells(startRange.offset(rst.AbsolutePosition, 0).Row, startRange.offset(0, i).Column).Value _
          = rst.fields(exportFieldNames(i)).Value
      Next i
      .MoveNext
    Loop
  End With
End Sub

Private Sub setExcelExportColumnHeaders( _
  ByRef ws As Worksheet, _
  ByRef startRange As Range, _
  ByVal exportFieldNames As Variant _
)
Dim i As Integer
  For i = LBound(exportFieldNames) To UBound(exportFieldNames)
    ws.Cells(startRange.Row, startRange.offset(0, i).Column).Value = exportFieldNames(i)
  Next i
End Sub

Private Function getNewWorkbook( _
  ByRef xlApp As Excel.Application, _
  ByVal path As String, _
  Optional ByVal title As String = vbNullString, _
  Optional ByVal Subject As String = vbNullString _
) As Workbook
Dim wb As Workbook
  
  Set getNewWorkbook = Nothing
  Set wb = xlApp.Workbooks.Add()
  With wb
    .title = title
    .Subject = Subject
    .SaveAs path
  End With
  Set getNewWorkbook = wb
End Function

Private Sub checkExcelExportWorkbookPath(ByVal path As String)
  If Not Util.File.IsWorkbookPath(path) Then
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Invalid workbook path"
  End If
End Sub

Private Sub checkProvidedFieldNames( _
  ByRef rst As Recordset, _
  ByVal fieldNames As Variant _
)
Dim missingNames As Variant
  missingNames = getMissingRecordsetFieldNames(rst, fieldNames)
  If Not Util.Arrays.IsEmpty(missingNames) Then
    err.raise Number:=ErrCode.PropertyOrMethodNotFound, _
      Description:="The following field names were not found in the provided Recodset:" & Util.Arrays.ToString(missingNames)
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getWorksheets
' Descr.    : returns an array of all available worksheets in a supplied workbook
'             The paremeter "Filter" can be used to limit the result to a list of
'             worksheet names to be searched for and returned
'---------------------------------------------------------------------------------------
Private Function GetWorksheets( _
  ByRef wb As Workbook, _
  Optional ByVal nameConstraints As String = vbNullString, _
  Optional ByVal constraintDelimiter As String = "," _
) As Variant
Dim result As ArrayList
Dim ws As Worksheet

  Set result = New ArrayList
  For Each ws In wb.Worksheets
    If nameConstraints = vbNullString _
    Or Util.Arrays.Contains(Split(nameConstraints, constraintDelimiter), ws.Name) Then
      result.Append ws
    End If
  Next ws
  
  GetWorksheets = result.Values
  Set result = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : copyRecord
' Purpose   : manualFieldSet is in the form of
'   Array(Array(fieldName1,fieldName2, ...),Array(fieldValue1,fieldValue2, ...))
' returns the primary key of the new record
' Note: make sure that both recordsets are at the current record you would like to copy
'   usually a bookmark set or record find / move / seek is required if you are not cycling through
'---------------------------------------------------------------------------------------
'
Public Function getCopyRecord( _
  ByRef sourceRecord As Recordset, _
  ByRef newRecord As Recordset, _
  ByVal manualFieldSet As Variant, _
  ByVal skipFieldNames As Variant, _
  Optional ByVal primKeyName As String = "id" _
) As Variant
Dim f As Field
Dim rc As Recordset
Dim ra As Recordset
Dim manFieldNames As ArrayList
Dim manFieldValues As ArrayList
  
  getCopyRecord = Null
  Set manFieldNames = New ArrayList
  Set manFieldValues = New ArrayList
  manFieldNames.Append manualFieldSet(0)
  manFieldValues.Append manualFieldSet(1)
  
  With newRecord
    .AddNew
    'Debug.Print "New id:" & newRecord("id").value
    For Each f In .fields
      If Not Util.Arrays.Contains(manFieldNames.Values, f.Name) _
      And Not Util.Arrays.Contains(skipFieldNames, f.Name) Then
        If VarType(f) = vbObject Then 'if multivalued field
          Set rc = sourceRecord.fields(f.Name).Value
          Set ra = f.Value
          Do While rc.EOF = False
            ra.AddNew
            ra!Value = rc!Value
            rc.MoveNext
            ra.Update
          Loop
        Else
          f.Value = sourceRecord.fields(f.Name).Value
        End If
      End If
    Next f
    
    Dim i As Integer
    For i = manFieldNames.lb To manFieldNames.ub
      .fields(manFieldNames.Value(i)).Value = manFieldValues.Value(i)
    Next i
    
    .Update
    .Bookmark = .LastModified
    On Error Resume Next
    getCopyRecord = newRecord.fields(primKeyName).Value
  End With
  
  Set manFieldNames = Nothing
  Set manFieldValues = Nothing
End Function

Public Function executeDeleteRecordAction( _
  ByRef frm As Form, _
  ByVal msgDeleteWhat As String, _
  ByVal msgTitle As String, _
  ByVal fields As String, _
  ByVal domain As String, _
  ByVal condition As String _
) As Boolean
Dim msgResponse As VbMsgBoxResult
  
  executeDeleteRecordAction = False
  msgResponse = MsgBox( _
    ErrHandler.getWarningMsg(WarningMsgType.DeleteFromList, msgDeleteWhat), _
    vbOKCancel, msgTitle)
  
  If msgResponse = vbOK Then
    executeDeleteRecordAction = deleteRecords(fields, domain, condition)
  End If
  
  frm.Requery
End Function

Public Sub hideRibbon()
  CurrentProject.Application.DoCmd.ShowToolbar "Ribbon", acToolbarNo
End Sub

Public Sub showRibbon()
  CurrentProject.Application.DoCmd.ShowToolbar "Ribbon", acToolbarYes
End Sub

Public Sub minimizeRibbon()
  If CommandBars("ribbon").Height > 100 Then
    CommandBars.ExecuteMso "MinimizeRibbon"
  End If
End Sub

Public Sub maximizeRibbon()
  If CommandBars("ribbon").Height <= 100 Then
    CommandBars.ExecuteMso "MinimizeRibbon"
  End If
End Sub

Public Sub hideNavigationPane()
  DoCmd.NavigateTo "acNavigationCategoryObjectType"
  DoCmd.RunCommand acCmdWindowHide
End Sub

Public Sub showNavigationPane()
  Call DoCmd.SelectObject(acTable, , True)
End Sub

Public Function existsTableWithName(ByVal tableName As String) As Boolean
  existsTableWithName = False
  Dim db As Database
  Set db = CurrDb()
  Dim tdef As TableDef
  db.TableDefs.Refresh
  For Each tdef In db.TableDefs
    If StrComp(tdef.Name, tableName, vbTextCompare) = 0 Then
      existsTableWithName = True
      Exit Function
    End If
  Next tdef
End Function

'Accepts table name or tabledef as paramater
Public Function isLinkedTable(ByRef tbl As Variant) As Boolean
  isLinkedTable = False
  
  Dim tdf As TableDef
  Dim db As Database
  Set db = CurrDb()
  
  On Error Resume Next 'in case item not found in TableDefs
  If VarType(tbl) = VbVarType.vbObject Then
    Set tdf = tbl
  ElseIf VarType(tbl) = VbVarType.vbString Then
    Set tdf = db.TableDefs(tbl)
  End If
  On Error GoTo 0
  
  If tdf Is Nothing Then Exit Function
  If Left(tdf.Name, 1) = "~" Then Exit Function
  If (InStr(1, tdf.Connect, ";DATABASE=") <> 0) Then isLinkedTable = True
  Set db = Nothing
End Function

Public Function getLinkedTableConnection(ByRef tbl As Variant) As String
  getLinkedTableConnection = ""
  
  Dim tdf As TableDef
  Dim db As Database
  Set db = CurrDb()
  
  On Error Resume Next 'in case item not found in TableDefs
  If VarType(tbl) = VbVarType.vbObject Then
    Set tdf = tbl
  ElseIf VarType(tbl) = VbVarType.vbString Then
    Set tdf = db.TableDefs(tbl)
  End If
  On Error GoTo 0
  
  If tdf Is Nothing Then Exit Function
  If Left(tdf.Name, 1) = "~" Then Exit Function
  If (InStr(1, tdf.Connect, ";DATABASE=") <> 0) Then getLinkedTableConnection = tdf.Connect
  Set db = Nothing
End Function

Public Function getLinkedTableDatabasePath(ByRef tbl As Variant) As String
  getLinkedTableDatabasePath = ""
  Dim con As String
  con = Me.getLinkedTableConnection(tbl)
  If Len(con) = 0 Then Exit Function
  
  Dim v As Variant
  v = Split(con, ";")
  Dim i As Integer, dbPart As String
  For i = LBound(v) To UBound(v)
    If InStr(1, v(i), "DATABASE=", vbTextCompare) Then
      dbPart = trim$(v(i))
    End If
  Next i

  getLinkedTableDatabasePath = Right(dbPart, Len(dbPart) - Len("DATABASE="))
End Function

Public Sub relinkTables( _
  ByVal tableNames As Variant, _
  ByVal path As String, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "" _
)
Dim i As Integer
  For i = LBound(tableNames) To UBound(tableNames)
    If Len(trim(tableNames(i))) > 0 Then
      RaiseEvent relinkTables(i, UBound(tableNames))
      relinkTable tableNames(i), path, databaseType, Password
    End If
  Next i
End Sub

Public Sub relinkTable( _
  ByRef tbl As Variant, _
  ByVal path As String, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "" _
)
  Dim tdf As TableDef
  Dim db As Database
  Set db = CurrDb()
  
  If VarType(tbl) = VbVarType.vbObject Then
    Set tdf = tbl
  ElseIf VarType(tbl) = VbVarType.vbString Then
    Set tdf = db.TableDefs(tbl)
  End If
  
  RaiseEvent relinkTable(tdf.Name)
  
  Dim con As String
  con = databaseType
  If Len(Password) > 0 Then con = con & ";PWD=" & Password
  con = con & ";DATABASE=" & path
  
  tdf.Connect = con
  tdf.RefreshLink
  Set db = Nothing
  Set tdf = Nothing
End Sub

Public Sub linkTransferTable( _
  ByVal dbPath As String, _
  ByVal sourceName As String, _
  ByVal destinationName As String, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "" _
)
  If Len(Password) > 0 Then
    Dim dbs As DAO.Database
    Set dbs = CurrentProject.Application. _
      DBEngine.OpenDatabase(dbPath, False, False, ";pwd=" & Password)
  End If
  
  DoCmd.TransferDatabase acLink, databaseType, dbPath, acTable, sourceName, destinationName

On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Public Sub linkTransferTables( _
  ByVal dbPath As String, _
  ByVal tableNames As Variant, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "" _
)
  If Len(Password) > 0 Then
    Dim dbs As DAO.Database
    Set dbs = CurrentProject.Application. _
      DBEngine.OpenDatabase(dbPath, False, False, ";pwd=" & Password)
  End If
  
  Dim i As Integer
  For i = LBound(tableNames) To UBound(tableNames)
    If Len(trim(tableNames(i))) > 0 Then
      DoCmd.TransferDatabase acLink, databaseType, dbPath, acTable, tableNames(i), tableNames(i)
    End If
  Next i

On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Public Sub linkTransferTablesFromDb( _
  ByRef db As DAO.Database, _
  ByVal tableNames As Variant _
)
  Dim i As Integer
  For i = LBound(tableNames) To UBound(tableNames)
    If Len(trim(tableNames(i))) > 0 Then
      DoCmd.TransferDatabase _
        TransferType:=acLink, _
        databaseType:="Microsoft Access", _
        DatabaseName:=db.Name, _
        ObjectType:=acTable, _
        source:=tableNames(i), _
        Destination:=tableNames(i), _
        StoreLogin:=True
    End If
  Next i
End Sub

Public Sub linkCreateTables( _
  ByVal dbPath As String, _
  ByVal tableNames As Variant, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "", _
  Optional ByVal overwrite As Boolean = False _
)
  Dim i As Integer, destName As String, db As Database
  For i = LBound(tableNames) To UBound(tableNames)
    If Len(trim(tableNames(i))) > 0 Then
      RaiseEvent linkCreateTables(i, UBound(tableNames))
      destName = tableNames(i)
      
      If Me.existsTableWithName(tableNames(i)) Then
        If overwrite Then
          Set db = CurrDb
          db.Execute "DROP TABLE " & tableNames(i), dbFailOnError
          Set db = Nothing
        Else
          destName = getNewAvailableTableName(tableNames(i))
        End If
      End If
      
      linkCreateTable _
        dbPath, tableNames(i), destName, databaseType, Password
    End If
  Next i
End Sub

Public Function getNewAvailableTableName(ByVal tblName As String)
Dim i As Integer
  i = 1
  While Me.existsTableWithName(tblName)
    tblName = tblName & CStr(i)
    i = i + 1
  Wend
  getNewAvailableTableName = tblName
End Function

Public Sub linkCreateTable( _
  ByVal dbPath As String, _
  ByVal sourceName As String, _
  ByVal destinationName As String, _
  Optional ByVal databaseType As String = "MS Access", _
  Optional ByVal Password As String = "" _
)
Dim tblDef As TableDef
Dim dbs As Database
Dim conn As String

  conn = "MS Access"
  conn = conn & ";DATABASE=" & dbPath
  If Len(Password) > 0 Then
    conn = conn & ";PWD=" & Password
  End If
  
  RaiseEvent linkCreateTable(destinationName)
  
  Set dbs = CurrDb
  Set tblDef = dbs.CreateTableDef(destinationName)
  tblDef.Connect = conn
  tblDef.SourceTableName = sourceName
  dbs.TableDefs.Append tblDef
  dbs.TableDefs.Refresh
  RefreshDatabaseWindow
  Set tblDef = Nothing
  Set dbs = Nothing
End Sub

Public Sub deleteAllLinkedTables()
Dim tdf As TableDef
Dim tbls As Collection
  Set tbls = New Collection
  
  For Each tdf In CurrDb().TableDefs
    If isLinkedTable(tdf) Then
      tbls.Add tdf
    End If
  Next tdf
  
  For Each tdf In tbls
    DoCmd.DeleteObject acTable, tdf.Name
  Next tdf
  
  CurrDb().TableDefs.Refresh
End Sub

Public Function isQueryDef( _
  ByVal qryName As String _
) As Boolean
Dim qf As QueryDef
  
  isQueryDef = False
  CurrDb().QueryDefs.Refresh
  For Each qf In CurrDb().QueryDefs
    If StrComp(qf.Name, qryName, vbTextCompare) = 0 Then
      isQueryDef = True
      Exit Function
    End If
  Next qf
End Function

Public Function isTableDef( _
  ByVal tableName As String _
) As Boolean
Dim tf As TableDef
  
  isTableDef = False
  
  CurrDb().TableDefs.Refresh
  For Each tf In CurrDb().TableDefs
    If StrComp(tf.Name, tableName, vbTextCompare) = 0 Then
      isTableDef = True
      Exit Function
    End If
  Next tf
End Function






'=====================================================================================================================
' FROM UtilAccessForm
'=====================================================================================================================




Public Sub setFormProperty( _
    formName As String, _
    propertyName As String, _
    Value As Variant _
)
  If IsNull(getFormProperty(formName, propertyName)) Then
    CurrentProject.AllForms(formName).Properties.Add propertyName, Nz(Value)
  Else
    CurrentProject.AllForms(formName).Properties(propertyName).Value = Value
  End If
End Sub

Public Function getFormProperty( _
    formName As String, _
    propertyName As String _
) As Variant
On Error GoTo getFormProperty_Error

  getFormProperty = CurrentProject.AllForms(formName).Properties(propertyName).Value

getFormProperty_Exit:
On Error Resume Next
Exit Function

getFormProperty_Error:
  If err.Number = 2455 Then
      getFormProperty = Null
      Resume Next
  Else
    MsgBox "Error " & err.Number & " The property " & propertyName & "was not created! " _
    & "(" & err.Description & ") in procedure getFormProperty of module FormUtil" _
    & vbLf & InfoErrMsg, _
    vbExclamation, _
    "Error occurred..."
  End If
  
Resume getFormProperty_Exit
End Function

Public Function getFilterString( _
  fieldName As String, _
  fieldValue As Variant _
) As String
On Error GoTo getFilterString_Error

  getFilterString = ""

  If IsNull(fieldValue) Then _
    err.raise 94, , "parameter is null"
  
  getFilterString = fieldName & "=" & fieldValue

getFilterString_Exit:
On Error Resume Next
Exit Function

getFilterString_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure getFilterString of module Form_add_role_attributes" _
    & vbLf & InfoErrMsg, _
    vbExclamation, _
    "Error occurred..."
Resume getFilterString_Exit
End Function

Public Sub setFilter(ByRef frm As Form, ByVal Filter As String)
  frm.Filter = Filter
  frm.FilterOn = True
End Sub

'TODO not applicable anymore make generic and move to dbutils
Public Function getFilterControls( _
    ByRef ctls As controls, _
    Optional ByVal filterTag = "filter" _
) As Collection
Dim col As Collection
Dim c As Control

  Set getFilterControls = Nothing
  Set col = New Collection
  
  For Each c In ctls
    If StrComp(c.Tag, filterTag, vbTextCompare) = 0 Then
      col.Add c, c.Name
    End If
  Next c
  
  Set getFilterControls = col
  Set col = Nothing
End Function

Public Function getControlsWithTags( _
    ByRef ctls As controls, _
    ByVal tagsText As Variant, _
    Optional ByVal matchAll As Boolean = False, _
    Optional ByVal tagSeperator As String = "," _
) As Collection
Dim col As Collection
Dim c As Control
Dim controlTags As Variant
Dim compareTags As Variant
Dim match As Boolean
Dim comp As Variant

  Set getControlsWithTags = Nothing
  Set col = New Collection
  match = False
  
  If tagsText = vbNullString Then _
    err.raise Number:=5, Description:="Provided Tags text was empty"
  
  For Each c In ctls
    If Not hasControlProperty(c, "Tag") Then GoTo next_control
    If Len(c.Tag) = 0 Then GoTo next_control
    
    controlTags = Split(c.Tag, tagSeperator)
    compareTags = Split(tagsText, tagSeperator)
    
    comp = Util.Arrays.GetIntersection(controlTags, compareTags)
    If Util.Arrays.IsEmpty(comp) Then GoTo next_control
    
    If matchAll Then
      match = (UBound(comp) = UBound(compareTags))
    Else
      match = (UBound(comp) >= 0)
    End If
    
    If match Then
      match = False
      col.Add c, c.Name
    End If
next_control:
  Next c
  
  Set getControlsWithTags = col
  Set col = Nothing
End Function

Public Function getControlsWithTag( _
  ByRef ctls As controls, _
  ByVal tagText As String, _
  ByVal LookAt As AcLookAt _
) As Collection
Dim col As Collection
Dim c As Control
Dim filterTag As String

  filterTag = tagText
  Set getControlsWithTag = Nothing
  Set col = New Collection
  
  For Each c In ctls
    If LookAt = AcLookAt.Whole Then
      If StrComp(c.Tag, filterTag, vbTextCompare) = 0 Then
        col.Add c, c.Name
      End If
    ElseIf LookAt = AcLookAt.Part Then
      If InStr(1, c.Tag, filterTag, vbTextCompare) > 0 Then
        col.Add c, c.Name
      End If
    End If
  Next c
  
  Set getControlsWithTag = col
  Set col = Nothing
End Function

Public Function getControlsNames( _
  ByRef ctls As controls _
) As Variant
Dim names As ArrayList
Dim c As Control

  getControlsNames = Array()
  Set names = New ArrayList
  For Each c In ctls
    names.Append c.Name
  Next c
  
  getControlsNames = names.Values
End Function

'TODO
Public Function getFilterBasedOnControls( _
    ByRef filterControls As Collection _
) As String
Dim f As String
Dim c As Control
Dim filterValues As String
  
  f = vbNullString
  filterValues = vbNullString
  
  For Each c In filterControls
    filterValues = getFilterValuesFromControl(c, Nothing)
    If Not filterValues = vbNullString Then
      If Not f = vbNullString Then f = f & " and "
      f = f & "(" & getFilterValuesFromControl(c, Nothing) & ")"
    End If
  Next c
  getFilterBasedOnControls = f
End Function

'TODO create a generic function
' get filter controls
' assign each filter control a ref id field name
' give the result as input to getFilterBasedOnControls respectively getFilterValuesFromControl
' need to get bound field if rowsource exists or control source directly if txt
Private Function getFilterValuesFromControl( _
    ByRef c As Control, _
    ByVal fieldRefs As Variant _
) As String
Dim attrFieldName As String
Dim f As String
Dim v As Variant
Dim i As Integer
  
  getFilterValuesFromControl = vbNullString
  v = c.Value
  
  If IsNull(v) Then Exit Function
  
  For i = LBound(v) To UBound(v)
    f = f & attrFieldName & " = " & v(i)
    If i < UBound(v) Then
      f = f & " or "
    End If
  Next i

  
  getFilterValuesFromControl = f
End Function

Public Sub adjustControlPropertyToStringLength( _
  ByRef c As Object, _
  ByVal propertyName As String, _
  ByVal length As Long _
)
  If hasControlProperty(c, propertyName) _
  And hasControlProperty(c, "FontSize") Then
    setControlPropertyByName _
      c, _
      propertyName, _
      Util.Strings.getStringWidthFromFontSize(c.fontSize, length)
  End If
End Sub

Public Sub adjustControlWidthToStringLength( _
  ByRef c As Control, _
  ByVal length As Long _
)
  adjustControlPropertyToStringLength c, "Width", length
End Sub

Public Sub adjustControlListWidthToStringLength( _
  ByRef c As Control, _
  ByVal length As Long _
)
  adjustControlPropertyToStringLength c, "ListWidth", length
End Sub

Public Function hasControlProperty( _
  ByRef c As Object, _
  ByVal propertyName As String _
) As Boolean
Dim p As Property

  hasControlProperty = False
  
  On Error Resume Next
  For Each p In c.Properties
    If StrComp(p.Name, propertyName, vbTextCompare) = 0 Then
      hasControlProperty = True
      Exit Function
    End If
  Next p
End Function

Public Sub setControlPropertyByName( _
  ByRef c As Object, _
  ByVal propertyName As String, _
  ByVal Value As Variant _
)
Dim p As Property
  
  On Error Resume Next
  For Each p In c.Properties
    If StrComp(p.Name, propertyName, vbTextCompare) = 0 Then
      p.Value = Value
      Exit Sub
    End If
  Next p
End Sub

Public Function getMaxLenghtOfControlSource( _
  ByRef c As Control _
) As Long
  getMaxLenghtOfControlSource = 0
End Function

Public Function getRowSourceBoundColumnName( _
  ByRef c As Control _
) As String

  
  
End Function

Public Function getRowSourceColumnNames( _
  ByRef c As Control _
) As Variant
  
  getRowSourceColumnNames = Array()
  If hasControlProperty(c, "RowSource") Then
    getRowSourceColumnNames = getFieldsFromDataSource(c.RowSource)
  End If
End Function

Public Function getRowSourceVisibleColumnNames( _
  ByRef c As Control _
) As Variant

  

End Function

'TODO move to AccessUtil with all dependencies
Public Function getFieldsFromDataSource( _
  ByVal source As String _
) As Variant
  getFieldsFromDataSource = Array()
  
  If isQueryDef(source) Then
    getFieldsFromDataSource = getFieldsFromQueryDef(source)
  ElseIf isTableDef(source) Then
    getFieldsFromDataSource = getFieldsFromTableDef(source)
  Else
    getFieldsFromDataSource = getFieldsFromSQL(source)
  End If
End Function

' TODO adjust the field return to not process syntax words after select such as DISTINCT, TOP, etc.
Public Function getFieldsFromSQL( _
  ByVal sql As String _
) As Variant
Dim v As Variant

  getFieldsFromSQL = Array()
  
  If InStr(1, sql, "SELECT ", vbTextCompare) > 0 _
  And InStr(1, sql, "FROM", vbTextCompare) > 0 Then
    sql = Replace$(sql, "SELECT ", "", , , vbTextCompare)
    sql = Left$(sql, InStr(1, sql, "FROM", vbTextCompare) - 1)
    v = Util.Strings.SplitTrim(sql, ",")
    
    Dim l As Long
    For l = LBound(v) To UBound(v)
      If InStr(1, v(l), ".", vbTextCompare) Then
        v(l) = Split(v(l), ".")(1)
      End If
    Next l
    getFieldsFromSQL = v
  End If
End Function

Public Function getFieldsFromQueryDef( _
  ByVal qryName As String _
) As Variant
  
  getFieldsFromQueryDef = Array()
  If isQueryDef(qryName) Then
    getFieldsFromQueryDef = getFieldsFromSQL(CurrDb().QueryDefs(qryName).sql)
  End If
End Function

Public Function getFieldsFromTableDef( _
  ByVal tableName As String _
) As Variant
Dim f As Field
Dim fields As ArrayList
Dim tf As TableDef
Dim db As Database

On Error GoTo getFieldsFromTableDef_Error

  getFieldsFromTableDef = Array()
  Set db = CurrDb()
  Set fields = New ArrayList
  
  If isTableDef(tableName) Then
    Set tf = db.TableDefs(tableName)
    
    For Each f In tf.fields
      fields.Append f.Name
    Next f
    
    getFieldsFromTableDef = fields.Values
  End If

getFieldsFromTableDef_Exit:
On Error Resume Next
  Set db = Nothing
  Set fields = Nothing
  Set tf = Nothing
Exit Function

getFieldsFromTableDef_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure getFieldsFromTableDef of FormUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume getFieldsFromTableDef_Exit
End Function

Public Function getSourceTableNameOfBoundControl( _
  ByRef frm As Form, _
  ByRef c As Control _
) As String
On Error GoTo getSourceTableNameOfBoundControl_Error

  getSourceTableNameOfBoundControl = vbNullString
  
  If hasControlProperty(c, "ControlSource") Then
    With frm
      getSourceTableNameOfBoundControl = .Recordset.fields(.controls(c.Name).ControlSource).SourceTable
    End With
  End If
  
getSourceTableNameOfBoundControl_Exit:
On Error Resume Next
Exit Function

getSourceTableNameOfBoundControl_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure getSourceTableNameOfBoundControl of FormUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume getSourceTableNameOfBoundControl_Exit

End Function

Public Function getWhereClauseFromSelectionsInControls( _
  ByRef controls As Collection, _
  ByVal boundName As String, _
  Optional ByVal logOpBetweenContols = "or", _
  Optional ByVal logOpBetweenValues = "or", _
  Optional ByVal negateWholeClauses As Boolean = False, _
  Optional ByVal negateSubClauses As Boolean = False _
) As String
Dim f As String
Dim c As Control
Dim ctlClause As String
  f = vbNullString
  getWhereClauseFromSelectionsInControls = vbNullString
  
  If controls.count <= 0 Then Exit Function
  
  For Each c In controls
    If IsNull(c.Value) Then GoTo next_value
    If f <> "" Then f = f & " " & logOpBetweenContols & " "
    ctlClause = getWhereClauseFromSelectionsInSingleControl(c, boundName, logOpBetweenValues, negateSubClauses)
    
    If negateWholeClauses Then
      f = f & "NOT" & "(" & ctlClause & ")"
    Else
      f = f & "(" & ctlClause & ")"
    End If
next_value:
  Next c
  
  getWhereClauseFromSelectionsInControls = f
End Function

Public Function getWhereClauseFromSelectionsInSingleControl( _
  ByRef c As Control, _
  ByVal boundName As String, _
  Optional ByVal logOpBetweenValues As String = "or", _
  Optional ByVal negateSubClauses As Boolean = False _
) As String
Dim f As String
Dim i As Integer
Dim v As Variant

  f = vbNullString
  getWhereClauseFromSelectionsInSingleControl = vbNullString
  v = c.Value
  
  If IsNull(v) Then Exit Function
  
  For i = LBound(v) To UBound(v)
    If negateSubClauses Then
      f = f & "NOT" & "(" & boundName & "=" & v(i) & ")"
    Else
      f = f & "(" & boundName & "=" & v(i) & ")"
    End If
    If i < UBound(v) Then f = f & " " & logOpBetweenValues & " "
  Next i
  
  getWhereClauseFromSelectionsInSingleControl = f
End Function

'TODO device a function that returns the bound argument from the rowsource
Public Function getWhereClauseFromValuesInControls( _
  ByRef controls As Collection, _
  ByVal boundName As String, _
  Optional ByVal logOpBetweenContols As String = "or", _
  Optional ByVal logOpBetweenValues As String = "or", _
  Optional ByVal negateWholeClauses As Boolean = False, _
  Optional ByVal negateSubClauses As Boolean = False _
) As String
Dim result As String, cValues As String, sep As String
Dim c As Control
  result = ""
  sep = " " & logOpBetweenContols & " "
  
  If controls.count <= 0 Then Exit Function
  
  For Each c In controls
    cValues = getWhereClauseFromValuesInSingleControl(c, boundName, logOpBetweenValues, negateSubClauses)
    If Len(cValues) > 0 Then
      If negateWholeClauses Then
        result = result & "NOT" & cValues & sep
      Else
        result = result & cValues & sep
      End If
    End If
  Next c

  If Right(result, Len(sep)) = sep Then
    result = Left(result, Len(result) - Len(sep))
  End If

  getWhereClauseFromValuesInControls = result
End Function

Public Function getWhereClauseFromValuesInSingleControl( _
  ByRef ctl As Control, _
  ByVal boundName As String, _
  Optional ByVal logOpBetweenValues As String = "or", _
  Optional ByVal negateSubClauses As Boolean = False _
) As String
Dim result As String, sep As String
Dim v As Variant, vals As Collection

  On Error Resume Next
  If Len(ctl.RowSource) = 0 Then
    getWhereClauseFromValuesInSingleControl = ""
    Exit Function
  End If
  On Error GoTo 0
  
  result = ""
  sep = " " & logOpBetweenValues & " "
  Set vals = EMultiLookup(boundName, ctl.RowSource)

  For Each v In vals
    If Not IsNull(v) Then
      If negateSubClauses Then
        result = result & "NOT(" & boundName & "=" & v & ")"
      Else
        result = result & "(" & boundName & "=" & v & ")"
      End If
      result = result & sep
    End If
  Next v

  'removing last " or " between control values
  If Right(result, Len(sep)) = sep Then
    result = Left(result, Len(result) - Len(sep))
  End If
  If Not result = "" Then result = "(" & result & ")"
  getWhereClauseFromValuesInSingleControl = result
End Function

Private Function getLeftOfWhereClause( _
  ByVal sql As String _
) As String
Dim searchRes As Variant
  searchRes = InStr(1, sql, "WHERE ", vbTextCompare)
  
  If searchRes = 0 Or IsNull(searchRes) Then
    getLeftOfWhereClause = Replace(sql, ";", "") 'eliminate end of query for easy condition additions
  Else
    getLeftOfWhereClause = Left(sql, searchRes - 1) ' + Len("WHERE "))
  End If
End Function

Public Sub setQueryWhereClause( _
  ByRef qryName As String, _
  ByVal WhereClause As String, _
  Optional defaultWhereClause As String = "TRUE" _
)
Dim qdf As QueryDef
Dim sql As String, beforeWhere As String
  
  If Not isQueryDef(qryName) Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, Description:="Query name not found:" & qryName
  End If
  
  Set qdf = CurrDb().QueryDefs(qryName)
  sql = qdf.sql
  If Len(WhereClause) = 0 Then WhereClause = defaultWhereClause
  beforeWhere = getLeftOfWhereClause(sql)
  sql = beforeWhere & " WHERE " & WhereClause & ";"
  qdf.sql = Replace(sql, "  ", " ") 'clean up any double space
End Sub


'=====================================================================================================================
'=====================================================================================================================
'=====================================================================================================================
'=====================================================================================================================
' FROM UtilAccessQueryFactory
'=====================================================================================================================

Public Function SelectFromWhere() As Recordset

End Function

Public Sub deleteFromWhere( _
  ByVal fields As Variant, _
  ByVal tableName As String, _
  Optional ByVal WhereClause As String = vbNullString _
)
  db.Execute getQueryDelete(fields, tableName, WhereClause)
End Sub

Public Function getQueryDelete( _
  ByVal fields As Variant, _
  ByVal tableName As String, _
  Optional ByVal WhereClause As String = vbNullString _
) As String
Dim sql As String
Dim bracketizedTableName As String
Dim i As Integer

  getQueryDelete = vbNullString
  bracketizedTableName = Util.Strings.bracketizeSquare(tableName)
  
  If Not IsArray(fields) _
  And VarType(fields) = vbString _
  And fields <> vbNullString Then
    sql = "DELETE " & bracketizedTableName
    sql = sql & "." & Util.Strings.bracketizeSquare(fields)
    sql = sql & " FROM " & bracketizedTableName
    
    If WhereClause <> vbNullString Then
      sql = sql & " WHERE " & "(" & WhereClause & ");"
    End If
  Else
    If Util.Arrays.IsAllocated(fields) Then
      sql = "DELETE "
      
      For i = LBound(fields) To UBound(fields)
        sql = sql & bracketizedTableName
        sql = sql & "." & Util.Strings.bracketizeSquare(fields(i))
        sql = sql & ", "
      Next i
      
      sql = sql & " FROM " & bracketizedTableName
      
      If WhereClause <> vbNullString Then
        sql = sql & " WHERE " & "(" & WhereClause & ");"
      End If
    End If
  End If
  
  getQueryDelete = sql
End Function

Public Sub createTable( _
)

End Sub

Public Function getQueryCreateTable( _
  ByVal fields As Variant, _
  ByVal tableName As String, _
  Optional ByVal WhereClause As String = vbNullString _
) As String
Dim sql As String
Dim bracketizedTableName As String
Dim i As Integer

  getQueryCreateTable = vbNullString
  bracketizedTableName = Util.Strings.bracketizeSquare(tableName)
  
  If Not IsArray(fields) _
  And VarType(fields) = vbString _
  And fields <> vbNullString Then
    sql = "DELETE " & bracketizedTableName
    sql = sql & "." & Util.Strings.bracketizeSquare(fields)
    sql = sql & " FROM " & bracketizedTableName
    
    If WhereClause <> vbNullString Then
      sql = sql & " WHERE " & "(" & WhereClause & ");"
    End If
  Else
    If Util.Arrays.IsAllocated(fields) Then
      sql = "DELETE "
      
      For i = LBound(fields) To UBound(fields)
        sql = sql & bracketizedTableName
        sql = sql & "." & Util.Strings.bracketizeSquare(fields(i))
        sql = sql & ", "
      Next i
      
      sql = sql & " FROM " & bracketizedTableName
      
      If WhereClause <> vbNullString Then
        sql = sql & " WHERE " & "(" & WhereClause & ");"
      End If
    End If
  End If
  
  getQueryCreateTable = sql
End Function

'TODO use to create tables refine to method
Private Sub CreateTableX3()
   Dim dbs As Database
  
  ' Modify this line to include the path to Northwind on your computer.
  Set dbs = OpenDatabase("Northwind.mdb")
  
  ' Create a table with three fields and a primary key.
  dbs.Execute "CREATE TABLE NewTable " _
      & "(FirstName CHAR, LastName CHAR, " _
      & "SSN INTEGER CONSTRAINT MyFieldConstraint " _
      & "PRIMARY KEY);"
  dbs.Close
End Sub

'---------------------------------------------------------------------------------------
' Procedure : hideControlsInFormByName
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub hideControlsInFormByName(ByVal form_name As String, ByVal controlNames As Variant)
  If Not CurrentProject.AllForms.item(form_name).IsLoaded Then Exit Sub
  Dim formControlNames As Variant
  Dim frm As Form
  Set frm = Forms.item(form_name)
  formControlNames = getControlsNames(frm.controls)
  
  Dim namesToHide As Variant
  namesToHide = Util.Arrays.GetIntersection(formControlNames, controlNames)
  Dim i As Integer
  For i = LBound(namesToHide) To UBound(namesToHide)
    frm.controls(namesToHide(i)).Visible = False
  Next i
End Sub


