VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ProjectSaveOption
  pjDoNotSave = 0
  pjPromptSave = 2
  pjSave = 1
End Enum

Public Enum ProjectPoolOpen
  pjPromptPool = 0
  pjPoolReadOnly = 1
  pjPoolReadWrite = 2
  pjPoolAndSharers = 3
  pjDoNotOpenPool = 4
End Enum

'---------------------------------------------------------------------------------------
' Module    : modProject
' Author    : Ivanov, Bozhan
' Date      : 12.09.2013
' LastUpdate: 09.04.2014
' Descr.    : Handles MS Project data related operations
'---------------------------------------------------------------------------------------

'Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

''---------------------------------------------------------------------------------------
'' Procedure : funGetProject
'' Type      : Function
'' Author    : Ivanov, Bozhan
'' Date      : 04.07.2013
'' Descr.    : returns the destination workbook as an object
''             as parameter accepts the excel application object and the workbook file path
''---------------------------------------------------------------------------------------
''
'Public Function funGetProject(ByRef mpApp As Object, ByVal path As String) As Object
'Dim bOpenStatus As Boolean
'
'On Error GoTo funGetProject_Error
'
'    Set funGetProject = Nothing ' set default return value
'
'    '---------------------------------------------------------------------------------------
'    ' Check if path is an MSProjec data
'    '---------------------------------------------------------------------------------------
'    If InStr(1, Right(Trim(path), 4), ".mpp", vbTextCompare) = 0 Then
'        GoTo funGetProject_Exit
'    End If
'
'    '---------------------------------------------------------------------------------------
'    ' Try openning the mpp file
'    '---------------------------------------------------------------------------------------
'    bOpenStatus = mpApp.FileOpenEx(path, False, , , , , , , , , , pjPoolReadWrite, , , True)
'
'    ' Check MS Proj opened the file
'    If Not bOpenStatus Then GoTo funGetProject_Exit
'
'    Set funGetProject = mpApp.ActiveProject
'
'funGetProject_Exit:
'    On Error Resume Next
'    Exit Function
'
'funGetProject_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure funGetProject of Modul modProject"
'    Resume funGetProject_Exit
'End Function

'Public Sub testProjOpen()
'Dim mpApp As Object
'Dim prj As Object
'Dim p As Object
'Dim t As Object 'Task
'
'    Set mpApp = CreateObject("Object")
'    Set prj = funGetProject( _
'      mpApp, _
'      "C:\Users\ivano00b\Local Software Projects\Excel\Creon, Ralf\LoadPlan Scheduling\Import_P.mpp")
'
'    If prj Is Nothing Then Exit Sub
'
'    Debug.Print prj.ProjectNamePrefix
'
'    ' Open file for output and print all tasks
'    Open "C:\Users\ivano00b\Local Software Projects\Excel\Creon, Ralf\LoadPlan Scheduling\test.txt" _
'        For Output As #1
'
'    Print #1, prj.Name
'    Print #1, "---------------------------------------------------"
'
'    For Each t In prj.Tasks
'        Print #1, t.Id & " | " & t.Text1 & " | " & t.Name & " | " & t.start & " | " & t.finish
'
'    Next t
'
'    Print #1, "---------------------------------------------------"
'    Print #1, "Total of " & prj.Tasks.Count & " tasks"
'    Print #1, "---------------------------------------------------"
'
'    Close #1 ' Close File
'
'    mpApp.FileCloseAllEx (pjDoNotSave)
'    mpApp.Quit
'
'    Set prj = Nothing
'    Set mpApp = Nothing
'
'    Debug.Print "Done!"
    
'End Sub

'Public Sub testProjCreate()
'Dim mpApp As Object 'Object
'Dim prj As Project 'Object
''Dim p As Object
'Dim t As Task 'Object
'
'    Set mpApp = CreateObject("Object")
'    Set prj = mpApp.projects.Add
'
'    If prj Is Nothing Then Exit Sub
'
'    mpApp.visible = False
'    mpApp.DisplayAlerts = False
'    mpApp.ScreenUpdating = False
'
'    prj.Name = "Test Crate"
'
'    With prj.Tasks.Add("t1")
'        .ActualStart = Format("24.09.1984", "dd.mm.yyyy")
'    End With
'
'    With prj.Tasks.Add("t2")
'        .ActualStart = Now
'        .ActualFinish = Now + 20
'    End With
'
'    With prj.Tasks.Add("t3")
'        .ActualStart = Format("30.09.2013", "dd.mm.yyyy")
'        .ActualFinish = .ActualStart + 1
'        '.LinkPredecessors prj.Tasks.item("t2"), pjFinishToStart
'        .TaskDependencies.Add prj.Tasks.item("t2"), pjFinishToStart
'    End With
'
'
'    prj.SaveAs "C:\Users\ivano00b\Local Software Projects\Excel\Creon, Ralf\LoadPlan Scheduling\tmp.mpp"
'
'    mpApp.FileCloseAllEx (pjSave)
'
'    mpApp.DisplayAlerts = True
'    mpApp.ScreenUpdating = True
'    mpApp.Quit
'
'    Set prj = Nothing
'    Set mpApp = Nothing
'
'    Debug.Print "Done!"
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : getApplication
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function getApplication( _
  Optional ByVal Visible As Boolean = False _
) As Object
On Error GoTo getApplication_Error
  
  Set getApplication = Nothing
  
On Error Resume Next
  Set getApplication = GetObject(, "Object")
  
  If err.Number <> 0 Then
    err.Clear
    On Error GoTo getApplication_Error
    Set getApplication = CreateObject("Object")
  End If
  
  If Not getApplication Is Nothing Then
    getApplication.Visible = Visible
  End If
  
getApplication_Exit:
Exit Function

getApplication_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure getApplication of module ProjectUtil" _
    & vbLf & INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getApplication_Exit
End Function

Public Function getNewApplication( _
  Optional ByVal Visible As Boolean = False _
) As Object
On Error GoTo getApplication_Error
  
  Set getNewApplication = CreateObject("Object")
  getNewApplication.Visible = Visible
  
getApplication_Exit:
Exit Function

getApplication_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure getApplication of module ProjectUtil" _
    & vbLf & INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getApplication_Exit
End Function

Public Function getNewProject( _
  Optional ByRef app As Object, _
  Optional ByVal Visible As Boolean = False _
) As Object
On Error GoTo getNewProject_Error

  Set getNewProject = Nothing
  
  If app Is Nothing Then
    Set app = getNewApplication()
  End If
  
  With app
    .Visible = Visible
    Set getNewProject = .Projects.Add
  End With

getNewProject_Exit:
On Error Resume Next
Exit Function

getNewProject_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure getNewProject of module ProjectUtil" _
    & vbLf & INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getNewProject_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getProjectFromPath
' Descr.    :
'---------------------------------------------------------------------------------------
'Public Function getProjectFromPath( _
'  ByVal path As String, _
'  Optional app As Object, _
'  Optional readOnly As Boolean = False, _
'  Optional ByVal visible As Boolean = False _
') As Object
'On Error GoTo getProjectFromPath_Error
'Dim fm As FileManager
'Dim openStatus As Boolean
'
'  Set getProjectFromPath = Nothing
'  Set fm = New FileManager
'  openStatus = False
'
'  If Not fm.FileExists(path) Then
'    Err.Raise 13, , "Filepath does not exist or acces is restricted: " & path
'    GoTo getProjectFromPath_Exit
'  End If
'
'  If Not fm.IsMSProjectPath(path) Then
'    Err.Raise 13, , "Provided path was not an MS Project file: " & path
'    GoTo getProjectFromPath_Exit
'  End If
'
'  If app Is Nothing Then
'    Set app = getNewApplication(visible)
'  End If
'
'  openStatus = app.FileOpenEx(path, readOnly, , , , , , , , , , ProjectPoolOpen.pjPoolReadWrite, , , True)
'
'  If Not openStatus Then GoTo getProjectFromPath_Exit
'
'  Set getProjectFromPath = app.ActiveProject
'
'getProjectFromPath_Exit:
'On Error Resume Next
'Exit Function
'
'getProjectFromPath_Error:
'  Debug.Print "Error " & Err.Number & " (" & Err.Description & ") " & _
'    "in procedure getProjectFromPath of module ProjectUtil" _
'    & vbLf & INFO_ERR_MSG, _
'    vbExclamation, _
'    "Error occurred..."
'Resume getProjectFromPath_Exit
'End Function

Public Function getProjectByName( _
  ByVal Name As String, _
  ByRef app As Object _
) As Object
  
On Error GoTo getProjectByName_Error

  Set getProjectByName = Nothing
  
  Dim prjTemp As Object
  For Each prjTemp In app.Projects
    If StrComp(prjTemp.Name, Name, vbTextCompare) = 0 Then
      Set getProjectByName = prjTemp
      GoTo getProjectByName_Exit
    End If
  Next prjTemp

getProjectByName_Exit:
On Error Resume Next
Exit Function

getProjectByName_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure getProjectByName of module ProjectUtil:Klassenmodul" _
  & vbLf & INFO_ERR_MSG, _
  vbExclamation, _
  "Error occurred..."
Resume getProjectByName_Exit
  
End Function

'TODO app object has to be discarded somehow if newone is created
Public Sub closeProject( _
  ByVal prj As Object, _
  Optional ByRef app As Object, _
  Optional ByVal Save As ProjectSaveOption = ProjectSaveOption.pjDoNotSave _
)
On Error GoTo closeProject_Error

  If prj Is Nothing Then _
    err.raise 91, , "Provided project name is empty"

  'If app Is Nothing Then Set app = prj.Application
  prj.Application.FileCloseEx pjSave
  'Dim prjTemp As Object
'  For Each prjTemp In app.Projects
'    If StrComp(prjTemp.name, prj.name, vbTextCompare) = 0 Then
'      prjTemp.Activate
'      app.FileCloseEx save:=save, NoAuto:=True
'
'    End If
'  Next prjTemp

closeProject_Exit:
On Error Resume Next
Exit Sub

closeProject_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
    "in procedure closeProject of module ProjectUtil" _
    & vbLf & INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume closeProject_Exit
End Sub

''Gives var not defined error
Public Sub closeProjectByName( _
  ByVal Name As String, _
  ByRef app As Object, _
  Optional ByVal Save As ProjectSaveOption = ProjectSaveOption.pjDoNotSave _
)
  closeProject getProjectByName(Name, app), app, Save
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getTablesContainingFields
' Descr.    : looks up the field names and titles of each tasktable in project and
'   returns all tasktables that contains all the provided fields in matchValues array
'   matchValues has the datastructure of (Array->Array)
'   Each sub array contains the versions of a field's name
'---------------------------------------------------------------------------------------
Public Function getTablesContainingFields( _
  ByRef prj As Object, _
  ByRef matchValues As Variant _
) As Collection
Dim ids As ArrayList
Dim tb As Object
Dim col As Collection
Dim fieldsNotFound As Collection
Dim i As Integer, j As Integer
On Error GoTo getTablesContainingFields_Error

  Set getTablesContainingFields = New Collection
  Set col = New Collection
  Set ids = New ArrayList
  Set fieldsNotFound = New Collection
  
  For Each tb In prj.TaskTables
    ids.Append (getAllTableFieldsIdentifiers(tb))
    
    For i = LBound(matchValues) To UBound(matchValues)
      
    Next i
    ids.Intersect matchValues
    
    If ids.ub = UBound(matchValues) Then
      col.Add tb, tb.Name
    End If
    
    ids.Clear
  Next tb
  
  Set getTablesContainingFields = col
  
getTablesContainingFields_Exit:
On Error Resume Next
Exit Function

getTablesContainingFields_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getTablesContainingFields of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getTablesContainingFields_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getAllTableFieldsIdentifiers
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function getAllTableFieldsIdentifiers( _
  ByRef tb As Object _
) As Variant
Dim result As ArrayList
Dim f As Object
On Error GoTo getAllTableFieldsIdentifiers_Error

  Set result = New ArrayList
  
  For Each f In tb.TableFields
    result.Append getFieldIdentifiers(f)
  Next f
  
  getAllTableFieldsIdentifiers = result.Values

getAllTableFieldsIdentifiers_Exit:
On Error Resume Next
Exit Function

getAllTableFieldsIdentifiers_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getAllTableFieldsIdentifiers of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getAllTableFieldsIdentifiers_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFieldIdentifiers
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function getFieldIdentifiers( _
  ByRef f As Object _
) As Variant
Dim result As ArrayList
Dim fieldId As Object 'PjField
Dim fieldName As String
Dim prj As Object 'Object
On Error GoTo getFieldIdentifiers_Error

  Set result = New ArrayList
  Set prj = f.Parent.Parent
  
  If f.Field >= 0 Then
    fieldName = prj.Application.FieldConstantToFieldName(f.Field)
    
    On Error Resume Next
    fieldId = prj.Application.FieldNameToFieldConstant(fieldName)
      
    result.Append f.Index
    result.Append f.title
    result.Append fieldName
    result.Append f.Field
    result.Append prj.Application.CustomFieldGetName(fieldId)
    result.Append "|"
    On Error GoTo getFieldIdentifiers_Error
  End If
  'Debug.Print result.ToString
  
  getFieldIdentifiers = result.Values
  
getFieldIdentifiers_Exit:
On Error Resume Next
  Set result = Nothing
  Set prj = Nothing
Exit Function

getFieldIdentifiers_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getFieldIdentifiers of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getFieldIdentifiers_Exit
End Function

Public Function getFieldName( _
  ByRef f As Object _
) As String
Dim prj As Object 'Object
On Error GoTo getFieldName_Error

  getFieldName = vbNullString
  Set prj = f.Parent.Parent
  
  If f.Field >= 0 Then
    getFieldName = prj.Application.FieldConstantToFieldName(f.Field)
  End If
  
getFieldName_Exit:
On Error Resume Next
  Set prj = Nothing
Exit Function

getFieldName_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getFieldName of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getFieldName_Exit
End Function

Public Function getCustomFieldName( _
  ByRef f As Object _
) As String
Dim fieldName As String
Dim fieldId As Object 'PjField
Dim prj As Object 'Object
On Error GoTo getCustomFieldName_Error

  getCustomFieldName = vbNullString
  Set prj = f.Parent.Parent
  
  If f.Field >= 0 Then
    fieldName = prj.Application.FieldConstantToFieldName(f.Field)
    fieldId = prj.Application.FieldNameToFieldConstant(fieldName)
    getCustomFieldName = prj.Application.CustomFieldGetName(fieldId)
  End If
  
getCustomFieldName_Exit:
On Error Resume Next
  Set prj = Nothing
Exit Function

getCustomFieldName_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getCustomFieldName of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getCustomFieldName_Exit
End Function


'---------------------------------------------------------------------------------------
' Procedure : getFieldWithIdentifier
' Descr.    : searches for a machting identifier with a provided value
'   search will be performed on index, title, field name, custom given name
'---------------------------------------------------------------------------------------
Public Function getFieldWithIdentifier( _
  ByRef tb As Object, _
  ByVal searchValue As Variant _
) As Object
Dim ids As Variant
Dim f As Object 'Object
On Error GoTo getFieldWithIdentifier_Error
  
  Set getFieldWithIdentifier = Nothing
  
  For Each f In tb.TableFields
    ids = getFieldIdentifiers(f)
    If Util.Arrays.Contains(ids, searchValue) Then
      Set getFieldWithIdentifier = f
      GoTo getFieldWithIdentifier_Exit
    End If
  Next f
  
getFieldWithIdentifier_Exit:
On Error Resume Next
ids = Null
Set f = Nothing
Exit Function

getFieldWithIdentifier_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getFieldWithIdentifier of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getFieldWithIdentifier_Exit
End Function

Public Function getTableByName( _
  ByRef prj As Object, _
  ByVal tableName As String, _
  Optional ByVal CompareMethod As VBA.VbCompareMethod = vbTextCompare _
) As Object
On Error GoTo getTableByName_Error
Dim tb As Object
  
  Set getTableByName = Nothing
  
  For Each tb In prj.TaskTables
    If StrComp(tb.Name, tableName, CompareMethod) = 0 Then
      Set getTableByName = tb
      GoTo getTableByName_Exit
    End If
  Next tb

getTableByName_Exit:
On Error Resume Next
Exit Function

getTableByName_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getTableByName of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getTableByName_Exit
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : getTableMatchingCriteria
' Descr.    : returns the first matching table containing all required fields and/or matching
'   prefered table name
'   requiredFields has the datastructure of (Array->Array)
'   Each sub array contains the versions of a field's name
'---------------------------------------------------------------------------------------
Public Function getTableMatchingCriteria( _
  ByRef prj As Object, _
  Optional ByVal prefferedTableName As String = vbNullString, _
  Optional ByVal requiredFields As Variant = Null _
) As Object 'Object
Dim tb As Object 'Object
Dim col As Collection
On Error GoTo getTableMatchingCriteria_Error
  
  Set getTableMatchingCriteria = Nothing
   
  If Not prefferedTableName = vbNullString Then
    Set tb = getTableByName(prj, prefferedTableName)
    
    If Not tb Is Nothing Then
      Set getTableMatchingCriteria = tb
      GoTo getTableMatchingCriteria_Exit
    End If
  End If
   
  If Util.Arrays.IsEmpty(requiredFields) Then
    Set col = getTaskTables(prj)
  Else
    Set col = getTablesContainingFields(prj, requiredFields)
  End If
  
  If col Is Nothing Then GoTo getTableMatchingCriteria_Exit
 
  Set getTableMatchingCriteria = col.item(1)
  
getTableMatchingCriteria_Exit:
On Error Resume Next
Exit Function

getTableMatchingCriteria_Error:
  Debug.Print "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getTableMatchingCriteria of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getTableMatchingCriteria_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getTaskTables
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function getTaskTables( _
  ByRef prj As Object _
) As Collection
Dim col As Collection
Dim tb As Object 'Object

On Error GoTo getTaskTables_Error

  Set getTaskTables = New Collection
  Set col = New Collection
    
  For Each tb In prj.TaskTables
    col.Add tb, tb.Name
  Next tb
  
  Set getTaskTables = col

getTaskTables_Exit:
On Error Resume Next
Exit Function

getTaskTables_Error:
  Debug.Print "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getTaskTables of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getTaskTables_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFieldsFromTable
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function getFieldsFromTable( _
  ByRef tb As Object, _
  Optional ByVal fieldNames As Variant = Null _
) As Collection
Dim col As Collection
Dim colNotFound As Collection
Dim i As Integer
On Error GoTo getFieldsFromTable_Error

  Set getFieldsFromTable = New Collection
  Set col = New Collection
  Set colNotFound = New Collection
    
  Dim f As Object
  If Util.Arrays.IsEmpty(fieldNames) Then
    For Each f In tb.TableFields
      If f.Field >= 0 Then
        col.Add f, tb.Application.FieldConstantToFieldName(f.Field)
      End If
    Next f
  Else
    For i = LBound(fieldNames) To UBound(fieldNames)
      Set f = getFieldWithIdentifier(tb, fieldNames(i))
  
      If f Is Nothing Then
        colNotFound.Add fieldNames(i)
      Else
        col.Add f, fieldNames(i)
      End If
    Next i
  End If
  
  
  If colNotFound.count > 0 Then
    Dim msg As String, v As Variant
    
    For Each v In colNotFound
      msg = msg & v & ";"
    Next v
    
    err.raise 9, , "Fields not found: " & msg & " in table: " & tb.Name
  Else
    Set getFieldsFromTable = col
  End If
   
getFieldsFromTable_Exit:
On Error Resume Next
Exit Function

getFieldsFromTable_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure getFieldsFromTable of module ProjectUtil" & vbLf & _
    INFO_ERR_MSG, _
    vbExclamation, _
    "Error occurred..."
Resume getFieldsFromTable_Exit

End Function

'TODO
Public Function AreFieldIdsInTable( _
  ByRef tbl As Object, _
  ByVal fieldIds As Variant _
)

End Function

'TODO
Public Function AreFieldNamesInTable( _
  ByRef tbl As Object, _
  ByVal fieldIds As Variant _
)

End Function

'
'Public Sub HideTasksWithIndeces( _
'  ByRef prj As Object, _
'  ByRef taskIndeces As Variant _
')
'Dim l As Long
'Dim ae As ArrayList
'
'  Set ae = New ArrayList
'  ae.Append taskIndeces
'
'  If Not ae.IsEmpty Then
'    For l = ae.LB To ae.UB
'      SearchWorksheet.Rows(ae(l)).Hidden = True
'    Next l
'  End If
'
'  Set ae = Nothing
'End Sub
'
'Public Sub ShowTasksWithIndeces( _
'  ByRef prj As Object, _
'  ByRef taskIndeces As Variant _
')
'Dim l As Long
'Dim ae As ArrayList
'
'  Set ae = New ArrayList
'  ae.Append taskIndeces
'
'  If Not ae.IsEmpty Then
'    For l = ae.LB To ae.UB
'      SearchWorksheet.Rows(ae(l)).Hidden = False
'    Next l
'  End If
'
'  Set ae = Nothing
'End Sub
'
'Public Function GetHiddenTasksIndeces( _
'  ByRef prj As Object _
') As Variant
'On Error GoTo GetHiddenTasksIndeces_Error
'  GetHiddenTasksIndeces = Array()
'
'  Dim rngToSearch As Range
'  If searchRange Is Nothing Then
'    Set rngToSearch = SearchWorksheet.UsedRange
'  Else
'    Set rngToSearch = searchRange
'  End If
'
'  Dim hiddenRows As ArrayList
'  Set hiddenRows = New ArrayList
'
'  Dim l As Long
'  For l = 1 To rngToSearch.Rows.Count
'    If SearchWorksheet.Rows(l).Hidden Then
'      hiddenRows.Append l
'    End If
'  Next l
'
'  GetHiddenTasksIndeces = hiddenRows.ArrayVariant
'
'GetHiddenTasksIndeces_Exit:
'On Error Resume Next
'Exit Function
'
'GetHiddenTasksIndeces_Error:
'  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetHiddenTasksIndeces of module WorkbookUtil" _
'    & vbLf & INFO_ERR_MSG, _
'    vbExclamation, _
'    "Error occurred..."
'Resume GetHiddenTasksIndeces_Exit
'End Function

'pjApp.FileOpen "Clocks.mpp"
'    pjApp.ActiveProject.Tasks.Add "Hang clocks"
'    pjApp.FileSaveAs "Clocks.mpp"
'    pjApp.FileClose
'    pjApp.Quit

Public Sub copyData(src As Object, dst As Object)
Dim t As Object 'Task
Dim Asgn As Object 'Assignment
Dim ColumnCount As Integer
Dim Columns As Integer
Dim Tcount As Integer
Dim xlCol As Range


  For Each t In src.Tasks
    If Not t Is Nothing Then
'        If t.OutlineLevel > ColumnCount Then
'            ColumnCount = t.OutlineLevel
'        End If
      Set t = dst.Tasks.Add(t.Name)
      
    End If
  Next t

Tcount = 0
For Each t In src.Tasks
    If Not t Is Nothing Then
        'Set xlCol = xlRow.offset(0, t.OutlineLevel)
        xlCol = t.Name
            If t.Summary Then
                'make bold
            End If
        For Each Asgn In t.ASSIGNMENTS
        
'            Asgn.ResourceName
'            (Asgn.Work / 480) & " Days"
'            (Asgn.ActualWork / 480) & " Days"
        Next Asgn
        Tcount = Tcount + 1
    End If
Next t

End Sub

