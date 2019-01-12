VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function getFieldNamesForResultsExcelExport() As Variant
  getFieldNamesForResultsExcelExport = Util.Strings.SplitTrim( _
    Expression:=GlobalSetting(GLOBAL_SETTING.ExcelExportRecordsetFieldNames), _
    Delimiter:=GlobalSetting(GLOBAL_SETTING.ValueSeperator), _
    Compare:=vbTextCompare)
End Function

Public Function getFieldNamesForCriteriaExcelExport() As Variant
  getFieldNamesForCriteriaExcelExport = Util.Strings.SplitTrim( _
    Expression:=GlobalSetting(GLOBAL_SETTING.ExcelExportCriteriaRecordsetFieldNames), _
    Delimiter:=GlobalSetting(GLOBAL_SETTING.ValueSeperator), _
    Compare:=vbTextCompare)
End Function

Public Sub exportRoleFilterResultToExcel( _
  ByRef rst As DAO.Recordset, _
  ByRef ctrls As controls _
)
Dim ws As Excel.Worksheet
Dim saveDir As String
Dim fileName As String
  
  fileName = Nz(GlobalSetting(GLOBAL_SETTING.ExcelExportFileName), "ContactEmployees") & _
    "_" & Util.DateTime.GetDateTimeStamp & ".xlsx"
  saveDir = Util.File.GetFolderPathViaDialogWindow(Nz(getDefaultExcelExportDir(), ""), "Save file to...", "Save")
  
  Set ws = Util.Access.exportRecordsetToExcel( _
    rst:=rst, _
    exportFieldNames:=getFieldNamesForResultsExcelExport(), _
    path:=saveDir & "\" & fileName, _
    Subject:=GlobalSetting(GLOBAL_SETTING.ExcelExportSubject), _
    sheetName:=GlobalSetting(GLOBAL_SETTING.ExcelExportWorksheetNameData), _
    Show:=False, _
    cropSheet:=True)
  
  formatResultsExcelExportWorksheet ws
  
  Set ws = exportControlsValuesToWorkbook( _
    ctrls:=ctrls, _
    wb:=ws.Parent, _
    sheetName:=GlobalSetting(GLOBAL_SETTING.ExcelExportWorksheetNameCriteria), _
    Show:=False, _
    cropSheet:=False)
    
  formatCriteriaExcelExportWorksheet ws
  ws.Application.Visible = True
  Util.Excel.TurnOnScreenUpdatingAndAlerts ws.Application
End Sub

Private Function exportControlsValuesToWorkbook( _
  ByRef ctrls As controls, _
  ByRef wb As Excel.Workbook, _
  Optional ByVal sheetName As String = vbNullString, _
  Optional ByVal Show As Boolean = True, _
  Optional ByVal cropSheet As Boolean = True _
) As Excel.Worksheet
Dim ws As Excel.Worksheet
Dim xlApp As Excel.Application

  Set exportControlsValuesToWorkbook = Nothing
  Set xlApp = wb.Application
  
  xlApp.Visible = False
  Util.Excel.TurnOffScreenUpdatingAndAlerts xlApp
  Set ws = Util.Excel.CreateWorksheet(wb, sheetName)
  If cropSheet Then Util.Excel.CropWorksheet ws
  
  writeControlValuesToWorksheet getFilterCriteriaControls(ctrls), ws
  
  wb.Save
  If Show Then
    xlApp.Visible = True
    Util.Excel.TurnOnScreenUpdatingAndAlerts xlApp
  End If
  Set exportControlsValuesToWorkbook = ws
  Set xlApp = Nothing
End Function

Private Sub writeControlValuesToWorksheet( _
  ByRef ctrls As Collection, _
  ByRef ws As Excel.Worksheet _
)
Dim fieldNames As Variant, v As Variant
Dim valueFieldName As String, idFieldName As String
Dim i As Integer
Dim offset As XlRangeOffset
Dim startRange As Range

  Set startRange = ws.Cells(1, 1)
  offset.Row = 0
  offset.col = 0
  
  Dim c As Control
  For Each c In ctrls
  
    startRange.offset(offset.Row, offset.col).Value = c.ControlSource
    
    If Not IsNull(c.Value) Then
      fieldNames = Util.Access.getFieldsFromDataSource(c.RowSource)
      idFieldName = getIdFieldName(fieldNames)
      valueFieldName = getValueFieldName(fieldNames)
      v = c.Value
      For i = LBound(v) To UBound(v)
        offset.col = offset.col + 1
        startRange.offset(offset.Row, offset.col).Value = ELookup(valueFieldName, c.RowSource, idFieldName & "=" & v(i))
      Next i
    Else
      startRange.offset(offset.Row, offset.col + 1).Value = "N/A"
    End If
    
    offset.Row = offset.Row + 1
    offset.col = 0
  Next c
End Sub

Private Function getIdFieldName(ByVal fieldNames As Variant) As String
Dim i As Integer
  getIdFieldName = ""
  For i = LBound(fieldNames) To UBound(fieldNames)
    If InStr(1, fieldNames(i), "id", vbTextCompare) > 0 Then
      getIdFieldName = fieldNames(i)
      Exit Function
    End If
  Next i
End Function

Private Function getValueFieldName(ByVal fieldNames As Variant) As String
Dim i As Integer
  getValueFieldName = ""
  For i = LBound(fieldNames) To UBound(fieldNames)
    If InStr(1, fieldNames(i), "name", vbTextCompare) > 0 Then
      getValueFieldName = fieldNames(i)
      Exit Function
    End If
  Next i
End Function

Private Function getFilterCriteriaControls(ByRef ctrls As controls) As Collection
Dim filterControls As Collection
Dim filterFields As Variant
  
  Set getFilterCriteriaControls = New Collection
  Set filterControls = Util.Access.getControlsWithTag(ctrls, "filter_", AcLookAt.Part)
  filterFields = getFieldNamesForCriteriaExcelExport()
  
  Dim c As Control
  For Each c In filterControls
    If Util.Access.hasControlProperty(c, "ControlSource") Then
      If Util.Arrays.Contains(filterFields, c.ControlSource) Then
        getFilterCriteriaControls.Add c, c.Name
      End If
    End If
  Next c
End Function

Public Function getDefaultExcelExportDir() As String
Dim path As String
  
  getDefaultExcelExportDir = ""
  path = GlobalSetting(GLOBAL_SETTING.ExcelExportDefaultSaveLocation)
  If Len(Environ$(path)) > 0 Then
    getDefaultExcelExportDir = Util.File.GetFolderPath(Environ$(path))
  Else
    getDefaultExcelExportDir = Util.File.GetFolderPath(path)
  End If
End Function

Private Sub formatCriteriaExcelExport_SetHeaders(ByRef ws As Worksheet)
Dim headNames As Variant, v As Variant
Dim sep As Variant
Dim r As Range

  v = GlobalSetting(GLOBAL_SETTING.ExcelExportCriteriaColumnHeadersToFields)
  sep = GlobalSetting(GLOBAL_SETTING.ValueSeperator)
  headNames = Split(v, sep)
  
  Dim i As Integer
  Set r = ws.UsedRange
  For i = LBound(headNames) To UBound(headNames)
    Call r.Replace(trim(Split(headNames(i), "=")(0)), trim(Split(headNames(i), "=")(1)))
  Next i
End Sub

Private Sub formatCriteriaExcelExport_InsertColumnIndices(ByRef ws As Worksheet)
Dim r As Range
Dim i As Integer

  Set r = ws.rows(1)
  r.Insert XlInsertShiftDirection.xlShiftDown
  Set r = ws.Cells(1, 1)
  r.Cells(1, 1).Value = "Title"
  Dim co As Long, used As Range
  Set used = Util.Excel.getUsedRange(ws)
  co = 1
  For i = 1 To used.Columns.count - 1
    r.Cells(1, 1).offset(0, co).Value = "#" & co
    co = co + 1
  Next i
  
End Sub

Private Sub formatCriteriaExcelExportWorksheet(ByRef ws As Worksheet)
  formatCriteriaExcelExport_SetHeaders ws
  formatCriteriaExcelExport_InsertColumnIndices ws
  formatExcelExport_FormatHeaders ws.Range(ws.Cells(1, 1), ws.Cells(1, Util.Excel.GetLastColumnWithData(ws).Column))
  formatExcelExport_AlignData ws
  formatExcelExport_SetAutoFilter ws
  formatExcelExport_SetColumnsWidth ws
  formatExcelExport_SetDateStamp ws, 2, 0
End Sub

Private Sub formatResultsExcelExportWorksheet(ByRef ws As Worksheet)
  formatExcelExport_SetHeaders ws
  formatExcelExport_InsertIndexColumn ws
  formatExcelExport_FormatHeaders ws.Range(ws.Cells(1, 1), ws.Cells(1, Util.Excel.GetLastColumnWithData(ws).Column))
  formatExcelExport_AlignData ws
  formatExcelExport_SetAutoFilter ws
  formatExcelExport_SetColumnsWidth ws
  formatExcelExport_SetDateStamp ws, 2, 1
End Sub

Private Sub formatExcelExport_SetHeaders(ByRef ws As Worksheet)
Dim headNames As Variant, v As Variant
Dim sep As Variant
Dim r As Range

  v = GlobalSetting(GLOBAL_SETTING.ExcelExportColumnHeaderToFieldNames)
  sep = GlobalSetting(GLOBAL_SETTING.ValueSeperator)
  headNames = Split(v, sep)
  
  Dim i As Integer
  Set r = ws.UsedRange
  For i = LBound(headNames) To UBound(headNames)
    Call r.Replace(trim(Split(headNames(i), "=")(0)), trim(Split(headNames(i), "=")(1)))
  Next i
End Sub

Private Sub formatExcelExport_SetAutoFilter(ByRef ws As Worksheet)
Dim Headers As Range
  Set Headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, Util.Excel.GetLastColumnWithData(ws).Column))
  Headers.AutoFilter
End Sub

Private Sub formatExcelExport_SetDateStamp( _
  ByRef ws As Worksheet, _
  Optional ByVal rOffset As Long = 0, _
  Optional ByVal cOffset As Long = 0 _
)
Dim lastRow As Long
  lastRow = Util.Excel.GetLastRowWithData(ws).Row
  ws.Cells(lastRow, 1).offset(rOffset, cOffset).Value = "Selection from Database executed on:"
  ws.Cells(lastRow, 1).offset(rOffset + 1, cOffset) = Format(Now(), "yyyy-Mm-Dd Hh:Nn:Ss")
End Sub

Private Sub formatExcelExport_FormatHeaders( _
  ByRef rng As Excel.Range _
)
  formatExcelExport_SetAlign rng, xlVAlignCenter, xlHAlignCenter
  formatExcelExport_SetColor rng, RGB(238, 236, 225), True
  formatExcelExport_SetBorder rng, ColorIndex.Black, XlBorderWeight.xlMedium
End Sub
  

Private Sub formatExcelExport_SetAlign( _
  ByRef rng As Excel.Range, _
  Optional ByVal VerticalAlignment As XlVAlign = XlVAlign.xlVAlignBottom, _
  Optional ByVal HorizontalAlignment As XlHAlign = XlHAlign.xlHAlignLeft _
)
  With rng
    .VerticalAlignment = VerticalAlignment
    .HorizontalAlignment = HorizontalAlignment
  End With
End Sub

Private Sub formatExcelExport_SetColor( _
  ByRef rng As Excel.Range, _
  Optional ByVal Color As Long = 16777215, _
  Optional ByVal bold As Boolean = True _
)
  With rng
    .Interior.Color = Color
    .Font.bold = bold
  End With
End Sub

Private Sub formatExcelExport_SetBorder( _
  ByRef rng As Excel.Range, _
  Optional ByVal borderColor As ColorIndex = ColorIndex.Black, _
  Optional ByVal borderWeight As XlBorderWeight = XlBorderWeight.xlMedium _
)
  With rng.Borders(XlBordersIndex.xlEdgeBottom)
    .LineStyle = XlLineStyle.xlContinuous
    .ColorIndex = ColorIndex.Black
    .Weight = XlBorderWeight.xlMedium
  End With
End Sub

Private Sub formatExcelExport_AlignData( _
  ByRef ws As Worksheet, _
  Optional ByVal VerticalAlignment As XlVAlign = XlVAlign.xlVAlignBottom, _
  Optional ByVal HorizontalAlignment As XlHAlign = XlHAlign.xlHAlignLeft _
)
  Dim data As Range
  Set data = ws.Range( _
    ws.Cells(2, 1), _
    ws.Cells(Util.Excel.GetLastRowWithData(ws).Row, Util.Excel.GetLastColumnWithData(ws).Column))
  With data
    .Font.bold = False
    .VerticalAlignment = VerticalAlignment
    .HorizontalAlignment = HorizontalAlignment
    With .Borders(XlBordersIndex.xlInsideHorizontal)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
    With .Borders(XlBordersIndex.xlInsideVertical)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
    With .Borders(XlBordersIndex.xlEdgeBottom)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
    With .Borders(XlBordersIndex.xlEdgeLeft)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
    With .Borders(XlBordersIndex.xlEdgeRight)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
    With .Borders(XlBordersIndex.xlEdgeTop)
      .LineStyle = XlLineStyle.xlContinuous
      .ColorIndex = ColorIndex.Black
      .Weight = XlBorderWeight.xlThin
    End With
  End With
  
  Dim Index As Range
  Set Index = ws.Range(ws.Cells(1, 1), ws.Cells(Util.Excel.GetLastRowWithData(ws).Row, 1))
  With Index
    .HorizontalAlignment = HorizontalAlignment
  End With

End Sub

Private Sub formatExcelExport_SetColumnsWidth(ByRef ws As Worksheet)
Dim r As Range
  Set r = ws.Range(ws.Cells(1, 1), ws.Cells(Util.Excel.GetLastRowWithData(ws).Row, Util.Excel.GetLastColumnWithData(ws).Column))
  'r.ColumnWidth = 22.29
  r.Columns.AutoFit
End Sub

Private Sub formatExcelExport_InsertIndexColumn(ByRef ws As Worksheet)
Dim r As Range
Dim i As Integer
Dim lastRow As Long

  lastRow = Util.Excel.GetLastRowWithData(ws).Row
  
  Set r = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1))
  r.Insert XlInsertShiftDirection.xlShiftToRight
  Set r = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1))
  'r.ColumnWidth = 4.29
  
  
  r.Cells(1, 1).Value = "#"
  For i = 1 To lastRow - 1
    r.Cells(1, 1).offset(i).Value = i
  Next i
End Sub

