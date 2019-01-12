Attribute VB_Name = "GlobalDeclarations"

Public Const DEFAULT_DEVELOPER_NAME As String = "Bozhan Ivanov"
Public Const DEFAULT_DEVELOPER_CONTACT As String = ""

' Default error message
Public Const INFO_ERR_MSG As String = "For additional support, pleace contact " & _
  DEFAULT_DEVELOPER_NAME & _
  " at " & _
  DEFAULT_DEVELOPER_CONTACT

Public Type FrameClass
  Access As String
  Excel As String
  FrontPage As String
  Outlook As String
  PowerPoint_95 As String
  PowerPoint_97 As String
  PowerPoint_2000 As String
  PowerPoint_XP As String
  PowerPoint_2010 As String
  Project As String
  Word As String
  UserForm_97 As String
  UserForm_2000 As String
  VBE As String
End Type

Public Type AlertsAndUpdatingStatus
  ScreenUpdating As Boolean
  DisplayAlerts As Boolean
  EnableEvents As Boolean
End Type
'
Public Type OLEObjectSetting
  ClassType As Variant '"Forms.CommandButton.1", "Forms.ComboBox.1", ... (you must specify either ClassType or FileName). A string that contains the programmatic identifier for the object to be created. If ClassType is specified, FileName and Link are ignored.
  Height  As Variant '(you must specify either ClassType or FileName). A string that specifies the file to be used to create the OLE object.
  FileLink  As Variant 'True to have the new OLE object based on FileName be linked to that file. If the object isn't linked, the object is created as a copy of the file. The default value is False.
  DisplayAsIcon As Variant 'True to display the new OLE object either as an icon or as its regular picture. If this argument is True, IconFileName and IconIndex can be used to specify an icon.
  IconFileName   As Variant 'A string that specifies the file that contains the icon to be displayed. This argument is used only if DisplayAsIcon is True. If this argument isn't specified or the file contains no icons, the default icon for the OLE class is used.
  IconIndex As Variant 'The number of the icon in the icon file. This is used only if DisplayAsIcon is True and IconFileName refers to a valid file that contains icons. If an icon with the given index number doesn't exist in the file specified by IconFileName, the first icon in the file is used.
  IconLabel  As Variant 'A string that specifies a label to display beneath the icon. This is used only if DisplayAsIcon is True. If this argument is omitted or is an empty string (""), no caption is displayed.
  Left As Variant 'The initial coordinates of the new object, in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.
  Width  As Variant 'The initial size of the new object, in points.
  Top  As Variant 'The initial coordinates of the new object in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.
  LinkedCell  As Variant
  Enabled  As Variant
  Visible  As Variant
  ListFillRange As Variant
End Type

Public Type XlRangeOffset
  Row As Long
  col As Long
End Type

Public Const DEFAULT_ALLOWED_EXCEL_FILE_EXTENTIONS = "*.xls; *.xlsx; *.xlsm"
Public Const DEFAULT_ALLOWED_MS_PROJECT_FILE_EXTENTIONS = "*.mpp"
Public LAST_PATH As String

Private thisDb As DAO.Database
Private utils As Utility
Private dbutils As DbUtility
Private sett As Settings
Private errh As ErrorHandler
Private cu As DbUser

Public Property Get CurrUser() As DbUser
  If cu Is Nothing Then
    Set cu = New DbUser
    cu.username = Util.Windows.username()
  End If
  Set CurrUser = cu
End Property

Public Property Let CurrUser(ByRef User As DbUser)
  Set cu = User
End Property

Public Property Get Temp() As Collection
  If tmp Is Nothing Then Set tmp = New Collection
  Set Temp = tmp
End Property

Public Property Get CurrDb(Optional ByVal Refresh As Boolean = False) As DAO.Database
  If thisDb Is Nothing Or Refresh Then Set thisDb = CurrentDb()
  Set CurrDb = thisDb
End Property

Public Property Let CurrDb(Optional ByVal Refresh As Boolean = False, ByRef db As DAO.Database)
  Set thisDb = db
End Property

Public Property Get CurrDir()
  CurrDir = Util.File.GetFolderPath(CurrentProject.FullName)
End Property

Public Property Get Util() As Utility
  If utils Is Nothing Then Set utils = New Utility
  Set Util = utils
End Property

Public Property Get ClassGen() As ClassGenerator
  If cgen Is Nothing Then Set cgen = New ClassGenerator
  Set ClassGen = cgen
End Property

Public Property Get dbUtil() As DbUtility
  If dbutils Is Nothing Then Set dbutils = New DbUtility
  Set dbUtil = dbutils
End Property

Public Property Get Setting() As Settings
  If sett Is Nothing Then Set sett = New Settings
  Set Setting = sett
End Property

Property Get GlobalSetting(ByVal varID As GLOBAL_SETTING) As Variant
  GlobalSetting = Setting.GlobalSetting(varID)
End Property

Property Let GlobalSetting( _
  ByVal varID As GLOBAL_SETTING, _
  ByVal varValue As Variant _
)
  Setting.GlobalSetting(varID) = varValue
End Property

Property Get LocalSetting(ByVal varID As LOCAL_SETTING) As Variant
  LocalSetting = Setting.LocalSetting(varID)
End Property

Property Let LocalSetting( _
  ByVal varID As LOCAL_SETTING, _
  ByVal varValue As Variant _
)
  Setting.LocalSetting(varID) = varValue
End Property

Public Property Get ErrHandler() As ErrorHandler
  If errh Is Nothing Then Set errh = New ErrorHandler
  Set ErrHandler = errh
End Property


Public Function InfoErrMsg() As String
  InfoErrMsg = dbUtil.Servicer.InfoErrMsg()
End Function

Public Function getCurrentUserUsername() As String
  getCurrentUserUsername = CurrUser.username
End Function
'---------------------------------------------------------------------------------------
' Procedure : ELookup
' Author    : Allen Browne. allen@allenbrowne.com,
'   extended by Ivanov, Bozhan
' Date      : 2012-05-30
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ELookup(Expr As String, domain As String, Optional Criteria As Variant, _
    Optional OrderClause As Variant) As Variant
On Error GoTo Err_ELookup
    'Purpose:   Faster and more flexible replacement for DLookup()
    'Arguments: Same as DLookup, with additional Order By option.
    'Return:    Value of the Expr if found, else Null.
    '           Delimited list for multi-value field.
    'Author:    Allen Browne. allen@allenbrowne.com
    'Co-author: Ivanov, Bozhan  - Wrote customized solution
    'Updated:   December 2006, to handle multi-value fields (Access 2007 and later.)
    'Examples:
    '           1. To find the last value, include DESC in the OrderClause, e.g.:
    '               ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
    '           2. To find the lowest non-null value of a field, use the Criteria, e.g.:
    '               ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
    ' Added by Ivanov, Bozhan :
    '           3. To obtain the values of multiple fields specify the fields as in a query, e.g.:
    '               ELookup("[Surname], [FirstName]", "tblClient", , "ClientID DESC")
    '              The result of this ELookup has to e assigned to an array of type Variant, e.g.:
    '               Dim array() As Variant
    '               array = ELookup("[Surname], [FirstName]", "tblClient", , "ClientID DESC")
    '              Obtaining the result of array could be done with a simple for loop, e.g.:
    '               For i = 1 To UBound(array)
    '                   Debug.Print "array(" & TypeName(array(i)) & "): " & array(i)
    '               Next i
    'Note:      Requires a reference to the DAO library.
    Dim db As DAO.Database          'This database.
    Dim rs As DAO.Recordset         'To retrieve the value to find.
    Dim rsMVF As DAO.Recordset      'Child recordset to use for multi-value fields.
    Dim varResult As Variant        'Return value for function.
    Dim strSQL As String            'SQL statement.
    Dim strOut As String            'Output string to build up (multi-value field.)
    Dim lngLen As Long              'Length of string.
    Const strcSep = ","             'Separator between items in multi-value list.

    'Initialize to null.
    varResult = Null

    'Build the SQL string.
    strSQL = "SELECT TOP 1 " & Expr & " FROM " & domain
    If (Not IsMissing(Criteria)) Then
      If (Len(Criteria) > 0) Then strSQL = strSQL & " WHERE " & Criteria
    End If
    If (Not IsMissing(OrderClause)) Then
        If (Len(OrderClause) > 0) Then strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"

    'Lookup the value.
    Set db = DBEngine(0)(0)
    Set rs = db.OpenRecordset(strSQL, dbOpenForwardOnly)
    If rs.RecordCount > 0 Then
        'Will be an object if multi-value field.
        If VarType(rs(0)) = vbObject Then
            Set rsMVF = rs(0).Value
            Do While Not rsMVF.EOF
                If rs(0).Type = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!fileName & strcSep
                Else
                    strOut = strOut & rsMVF![Value].Value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            'Remove trailing separator.
            lngLen = Len(strOut) - Len(strcSep)
            If lngLen > 0& Then
                varResult = Left(strOut, lngLen)
            End If
            Set rsMVF = Nothing
        ElseIf rs.fields.count > 1 Then ' cutomized part of the function
            ReDim varResult(rs.fields.count - 1)
            ' Not a multi-value field, but requests values of multiple fields
            For lngLen = 0 To rs.fields.count - 1
                varResult(lngLen) = rs(lngLen)
            Next lngLen
            ' An Alternative would be varResult = rs.GetRows(rs.RecordCount)
            ' but in that case the row has to be specified because varResult
            ' will be redimed automatically to 2 dimentional array
        Else
            'Not a multi-value field: just return the value.
            varResult = rs(0)
        End If
    End If
    rs.Close
    If Not rsMVF Is Nothing Then rsMVF.Close

    'Assign the return value.
    ELookup = varResult

Exit_ELookup:
    Set rs = Nothing
    Set rsMVF = Nothing
    Set db = Nothing
    Exit Function

Err_ELookup:
    MsgBox err.Description, vbExclamation, "ELookup Error " & err.Number
    Resume Exit_ELookup
End Function

'---------------------------------------------------------------------------------------
' Procedure : EMultiLookup
' Author    : Ivanov, Bozhan
' Date      : 2016-12-08
'Purpose:   Faster and more flexible replacement for DLookup()
'Arguments: Same as DLookup, with additional Order By option.
'Return:    Value of the Expr if found, else Null.
'           Delimited list for multi-value field.
'Co-author: Ivanov, Bozhan  - Wrote customized solution
'Updated:   December 2006, to handle multi-value fields (Access 2007 and later.)
'Examples:
'           1. To find the last value, include DESC in the OrderClause, e.g.:
'               ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
'           2. To find the lowest non-null value of a field, use the Criteria, e.g.:
'               ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
' Added by Ivanov, Bozhan :
'           3. To obtain the values of multiple fields specify the fields as in a query, e.g.:
'               ELookup("[Surname], [FirstName]", "tblClient", , "ClientID DESC")
'              The result of this ELookup has to e assigned to an array of type Variant, e.g.:
'               Dim array() As Variant
'               array = ELookup("[Surname], [FirstName]", "tblClient", , "ClientID DESC")
'              Obtaining the result of array could be done with a simple for loop, e.g.:
'               For i = 1 To UBound(array)
'                   Debug.Print "array(" & TypeName(array(i)) & "): " & array(i)
'               Next i
'Note:      Requires a reference to the DAO library.
'---------------------------------------------------------------------------------------
'
Public Function EMultiLookup( _
  ByVal Expr As String, _
  ByVal domain As String, _
  Optional ByVal Criteria As Variant, _
  Optional ByVal OrderClause As Variant _
) As Collection
On Error GoTo Err_ELookup

    Dim db As DAO.Database          'This database.
    Dim rs As DAO.Recordset         'To retrieve the value to find.
    Dim rsMVF As DAO.Recordset      'Child recordset to use for multi-value fields.
    Dim varResult As Variant        'Return value for function.
    Dim strSQL As String            'SQL statement.
    Dim strOut As String            'Output string to build up (multi-value field.)
    Dim lngLen As Long              'Length of string.
    Dim res As Collection
    Const strcSep = ","             'Separator between items in multi-value list.


    'Initialize to null.
    varResult = Null
    Set res = New Collection

    'Build the SQL string.
    If InStr(1, domain, "SELECT ", vbTextCompare) > 0 _
    And InStr(1, domain, "FROM ", vbTextCompare) > 0 Then
      domain = "(" & Replace(domain, ";", "") & ")"
    End If
    strSQL = "SELECT " & Expr & " FROM " & domain
    If (Not IsMissing(Criteria)) Then
      If (Len(Criteria) > 0) Then strSQL = strSQL & " WHERE " & Criteria
    End If
    If (Not IsMissing(OrderClause)) Then
        If (Len(OrderClause) > 0) Then strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"

    'Lookup the value.
    Set db = DBEngine(0)(0)
    Set rs = db.OpenRecordset(strSQL, dbOpenForwardOnly)

    Do While Not rs.EOF
        'Will be an object if multi-value field.
        If VarType(rs(0)) = vbObject Then
            Set rsMVF = rs(0).Value
            Do While Not rsMVF.EOF
                If rs(0).Type = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!fileName & strcSep
                Else
                    strOut = strOut & rsMVF![Value].Value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            'Remove trailing separator.
            lngLen = Len(strOut) - Len(strcSep)
            If lngLen > 0& Then
                varResult = Left(strOut, lngLen)
            End If
            Set rsMVF = Nothing
        ElseIf rs.fields.count > 1 Then ' cutomized part of the function
            ReDim varResult(rs.fields.count - 1)
            ' Not a multi-value field, but requests values of multiple fields
            For lngLen = 0 To rs.fields.count - 1
                varResult(lngLen) = rs(lngLen)
            Next lngLen
            ' An Alternative would be varResult = rs.GetRows(rs.RecordCount)
            ' but in that case the row has to be specified because varResult
            ' will be redimed automatically to 2 dimentional array
        Else
            'Not a multi-value field: just return the value.
            varResult = rs(0)
        End If
        res.Add varResult
        rs.MoveNext
    Loop

    rs.Close
    If Not rsMVF Is Nothing Then rsMVF.Close
    Set EMultiLookup = res

Exit_ELookup:
    Set rs = Nothing
    Set rsMVF = Nothing
    Set db = Nothing
    Exit Function

Err_ELookup:
    MsgBox err.Description, vbExclamation, "ELookup Error " & err.Number
    Set EMultiLookup = New Collection
    Resume Exit_ELookup
End Function

