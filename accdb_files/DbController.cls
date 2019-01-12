VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TODO add status update
'SysCmd acSysCmdSetStatus, "Processing table [" & tdf.Name & "]..."
'TODO Evaluate all function if need to be public or private
'TODO sync global settings with local settings in initializeDatabase
Option Compare Database
Option Explicit

Private Const ACCESS_CONTROL_PASSWORD As String = "gW1GTQ8ylN5lCO"
Private Const BACKEND_PASSWORD As String = "XJ3zJMAPIoBPf4"

Public Enum dbPathTypeEnum
  Server = 1
  Working = 2
End Enum

Dim eh As ErrorHandler
Dim acdb As DAO.Database
Dim bedb As DAO.Database
Dim pbc As ProgressBarController

Private Sub Class_Initialize( _
)
  Set eh = New ErrorHandler
  Set pbc = New ProgressBarController
End Sub

Private Sub Class_Terminate( _
)
  On Error Resume Next
  Set eh = Nothing
  Set pbc = Nothing
  acdb.Close
  bedb.Close
  Set acdb = Nothing
  Set bedb = Nothing
End Sub

Public Sub initializeDatabase( _
)
  pbc.ShowForm
  pbc.Caption = "Initializing Database Resources..."
  pbc.TaskCount = 6
  
  pbc.IncrementTaskProgress StatusText:="Connecting to last working database..."
  If Not Len(getDatabaseDirPath(Working)) = 0 Then 'new installation
    ConnectToLastDB
  End If
  
  pbc.IncrementTaskProgress StatusText:="Checking User Clearance"
  checkUserPrivilegesAndUpdateLinks
  
  'backend will update only if working dir <> server dir and new version is available
  pbc.IncrementTaskProgress StatusText:="Checking for Backend updates"
  checkBackEndVersion
  
  'need to check user clearance again with the lates backend version information
  'if it is the same all operations will relinking be skipped in checkUserPrivilegesAndUpdateLinks
  pbc.IncrementTaskProgress StatusText:="Checking User Clearance"
  checkUserPrivilegesAndUpdateLinks
  
  pbc.IncrementTaskProgress StatusText:="Checking for Frontend updates"
  checkFrontEndVersion
      
  pbc.IncrementTaskProgress StatusText:="Re/Linking Backend Tables"
  reLinkToBackEndPost
  
  pbc.CloseForm
End Sub

Private Sub ConnectToLastDB()
  pbc.SubTaskCount = 2
  'Link to DB with the latest working path setting
  pbc.IncrementSubTaskProgress StatusText:="Re/Linking Access Control Tables"
  reLinkToAccessControl
  
  pbc.IncrementSubTaskProgress StatusText:="Preloading Backend Resources"
  reLinkToBackEndPre
End Sub

Private Sub checkUserPrivilegesAndUpdateLinks()
'check connection condition (user privileges) and set working path accordingly
'access to user clerance levels are only possible after reLinkToBackEndPre
  If updateWorkingPath(CurrUser.hasAdminPrivileges()) Then
    pbc.SubTaskCount = 2
    
    'relink with current working path if neccessarry
    pbc.IncrementSubTaskProgress StatusText:="Re/Linking Access Control Tables"
    reLinkToAccessControl
    
    pbc.IncrementSubTaskProgress StatusText:="Preloading Backend Resources"
    reLinkToBackEndPre
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : updateWorkingPath
' Purpose   : updates working path to the server is user is admin
'   updates working path to current project path if user is not admin
'---------------------------------------------------------------------------------------
Private Function updateWorkingPath( _
  ByVal connCondition As Boolean _
) As Boolean
  If connCondition Then
    If StrComp(Me.WorkingDirPath, Me.ServerDirPath, vbTextCompare) = 0 Then
      updateWorkingPath = False
    Else
      Me.WorkingDirPath = Me.ServerDirPath
      updateWorkingPath = True
    End If
  Else
    If StrComp(Me.WorkingDirPath, CurrDir(), vbTextCompare) = 0 Then
      updateWorkingPath = False
    Else
      Me.WorkingDirPath = CurrDir()
      updateWorkingPath = True
    End If
  End If
End Function

Property Get AccessControlDB( _
) As DAO.Database
  If acdb Is Nothing Then Set acdb = getAccessControlDatabase()
  If StrComp(acdb.Name, Me.AccessControlFilePath(Working)) <> 0 Then Set acdb = getAccessControlDatabase()
  Set AccessControlDB = acdb
End Property

Public Function getAccessControlDatabase() As DAO.Database
  Set getAccessControlDatabase = Nothing
  checkAccessControlAvailability Working
  
  Set getAccessControlDatabase = DBEngine.OpenDatabase( _
    Me.AccessControlFilePath(Working), False, False, ";pwd=" & getAccessControlCredential())
End Function

Property Let AccessControlDB( _
  ByRef db As DAO.Database _
)
  Set acdb = db
End Property

Property Get BackEndDB( _
) As DAO.Database
  If bedb Is Nothing Then Set bedb = getBackEndDatabase()
  If StrComp(bedb.Name, Me.BackEndFilePath(Working)) <> 0 Then Set bedb = getBackEndDatabase()
  Set BackEndDB = bedb
End Property

Private Function getBackEndDatabase() As DAO.Database
  Set getBackEndDatabase = Nothing
  checkBackEndAvailability Working
  Set getBackEndDatabase = DBEngine.OpenDatabase( _
    Me.BackEndFilePath(Working), False, False, ";pwd=" & getBackEndCredential())
End Function

Property Let BackEndDB( _
  ByRef db As DAO.Database _
)
  Set bedb = db
End Property

Property Get BackEndFilePath( _
  ByVal pType As dbPathTypeEnum _
) As String
  BackEndFilePath = getDatabaseDirPath(pType) & "\" & LocalSetting(LOCAL_SETTING.BackEndFileName)
End Property

Property Let BackEndFilePath( _
  ByVal pType As dbPathTypeEnum, _
  ByVal path As String _
)
  setDatabaseDirPath pType, path
  
  Dim fileName As String
  fileName = Util.File.GetFileName(path)
  If Len(fileName) > 0 Then
    LocalSetting(LOCAL_SETTING.BackEndFileName) = fileName
  End If
End Property

Property Get AccessControlFilePath( _
  ByVal pType As dbPathTypeEnum _
) As String
  AccessControlFilePath = getDatabaseDirPath(pType) & "\" & LocalSetting(LOCAL_SETTING.AccessControlFileName)
End Property

Property Let AccessControlFilePath( _
  ByVal pType As dbPathTypeEnum, _
  ByVal path As String _
)
  setDatabaseDirPath pType, path
  
  Dim fileName As String
  fileName = Util.File.GetFileName(path)
  If Len(fileName) > 0 Then
    LocalSetting(LOCAL_SETTING.AccessControlFileName) = fileName
  End If
End Property

Property Get LocalVersion() As Integer
  LocalVersion = CInt(Nz(LocalSetting(LOCAL_SETTING.version), 0))
End Property

Property Let LocalVersion(ByVal newVersion As Integer)
  LocalSetting(LOCAL_SETTING.version) = newVersion
End Property

Property Get GlobalVersion() As Integer
  GlobalVersion = CInt(Nz(GlobalSetting(GLOBAL_SETTING.version), 0))
End Property

Property Let GlobalVersion(ByVal newVersion As Integer)
  GlobalSetting(GLOBAL_SETTING.version) = newVersion
End Property

Property Get LocalBackEndVersion() As Integer
  LocalBackEndVersion = CInt(Nz(LocalSetting(LOCAL_SETTING.BackEndVersion), 0))
End Property

Property Let LocalBackEndVersion(ByVal newVersion As Integer)
  LocalSetting(LOCAL_SETTING.BackEndVersion) = newVersion
End Property

Property Get GlobalBackEndVersion() As Integer
  GlobalBackEndVersion = CInt(Nz(GlobalSetting(GLOBAL_SETTING.BackEndVersion), 0))
End Property

Property Let GlobalBackEndVersion(ByVal newVersion As Integer)
  GlobalSetting(GLOBAL_SETTING.BackEndVersion) = newVersion
End Property

Public Function getServerBackEndVersion() As Integer
  Dim serverdb As DAO.Database
  checkBackEndAvailability Server
  Set serverdb = DBEngine.OpenDatabase( _
    Me.BackEndFilePath(Server), False, False, ";pwd=" & getBackEndCredential())
  
  Dim rs As DAO.Recordset
  Set rs = serverdb.OpenRecordset("SELECT [varValue] FROM [settings] WHERE [varName] = 'BackEndVersion'")
  getServerBackEndVersion = CInt(Nz(rs(0), 0))
  Set serverdb = Nothing
End Function

Public Function getServerFrontEndVersion() As Integer
  Dim serverdb As DAO.Database
  checkBackEndAvailability Server
  Set serverdb = DBEngine.OpenDatabase( _
    Me.BackEndFilePath(Server), False, False, ";pwd=" & getBackEndCredential())
  
  Dim rs As DAO.Recordset
  Set rs = serverdb.OpenRecordset("SELECT [varValue] FROM [settings] WHERE [varName] = 'Version'")
  getServerFrontEndVersion = CInt(Nz(rs(0), 0))
  Set serverdb = Nothing
End Function

Property Get ServerDirPath( _
) As String
  ServerDirPath = getDatabaseDirPath(Server)
End Property

Property Let ServerDirPath(ByVal path As String)
  setDatabaseDirPath Server, path
End Property

Property Get WorkingDirPath( _
) As String
  WorkingDirPath = getDatabaseDirPath(Working)
End Property

Property Let WorkingDirPath(ByVal path As String)
  setDatabaseDirPath Working, path
End Property

Public Function getDatabaseDirPath( _
  ByVal pType As dbPathTypeEnum _
) As String
  Dim actDir As String
  Select Case pType
  Case dbPathTypeEnum.Server
    actDir = CStr(Nz(LocalSetting(LOCAL_SETTING.ServerDirPath), ""))
    If Len(actDir) <= 0 Then
      'TODO request new server dir
    End If
  Case dbPathTypeEnum.Working
    actDir = CStr(Nz(LocalSetting(LOCAL_SETTING.WorkingDirPath), ""))
    If Len(actDir) <= 0 Then
      'TODO request new server dir
    End If
  End Select
  getDatabaseDirPath = actDir
End Function

'---------------------------------------------------------------------------------------
' Procedure : setDatabaseDirPath
' Purpose   : extracts and proves the folder path from the provided parameter
'   and sets the path type respectively to the result
'---------------------------------------------------------------------------------------
Public Sub setDatabaseDirPath( _
  ByVal pType As dbPathTypeEnum, _
  ByVal path As String _
)
Dim actDir As String
On Error GoTo setDatabaseDirPath_Error

  actDir = Util.File.GetFolderPath(path)
  
  If Len(actDir) <= 0 Then
    err.raise Number:=ErrCode.FileNotFound, _
      Description:="Path argument is not a valid path"
  End If
  
  Select Case pType
  Case dbPathTypeEnum.Server
    LocalSetting(LOCAL_SETTING.ServerDirPath) = actDir
  Case dbPathTypeEnum.Working
    LocalSetting(LOCAL_SETTING.WorkingDirPath) = actDir
  End Select

setDatabaseDirPath_Exit:
On Error Resume Next
Exit Sub

setDatabaseDirPath_Error:
err.raise Number:=err.Number, _
    Description:="Error in procedure setDatabaseDirPath of DbControllerNew:Class Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume setDatabaseDirPath_Exit
End Sub

Public Sub updateWorkingAccessControl( _
)
  'before calling server path update is required if not available
  checkAccessControlAvailability Server
  
  Dim srcPath As String, dstPath As String
  srcPath = Me.AccessControlFilePath(Server)
  dstPath = Me.AccessControlFilePath(Working)
  
  'check if server and current project are in the same dir -> copy conflict
  If StrComp(srcPath, dstPath, vbTextCompare) = 0 Then
    Exit Sub
  End If
  
  'release database
  Me.AccessControlDB.Close
  Me.AccessControlDB = Nothing
  
  Dim copyResult As Boolean
  copyResult = Util.File.CopyFile(srcPath, dstPath, True)
  'Alternative: if file.copyFile is faulty quit app and use powershell cmtdlet copy-item -force
    
  If Not copyResult Then
    err.raise Number:=ErrCode.ErrorLoadingFile, _
      Description:="could not copy " & vbLf & _
      srcPath & vbLf & _
      "to destination" & vbLf & _
      dstPath
  End If
End Sub

Public Sub checkAccessControlAvailability( _
  ByVal pType As dbPathTypeEnum _
)
  If Not isAccessControlAvailable(pType) Then
    promptForNewAccessControlPath pType
  End If
End Sub

Public Function isAccessControlAvailable( _
  ByVal pType As dbPathTypeEnum _
) As Boolean
  isAccessControlAvailable = isDbFileAvailable(Me.AccessControlFilePath(pType))
End Function

Public Sub promptForNewAccessControlPath( _
  ByVal pType As dbPathTypeEnum _
)
  Dim ans As VbMsgBoxResult
  Dim newACPath As Variant

  ans = MsgBox("The current """ & getDBUtil_PathTypeLabel(pType) & """ Access Control path to file " & """" & _
    LocalSetting(LOCAL_SETTING.AccessControlFileName) & """" & _
    "is invalid or not accessible!" & vbLf & _
    "Would you like to update the """ & getDBUtil_PathTypeLabel(pType) & """ Access Control filepath yourself?", _
    vbYesNo, "Access Control file/connection missing...")
    
  If ans = vbYes Then
    newACPath = requestAccessControlPath(pType)
    If Util.File.FileExists(newACPath) Then
      Me.AccessControlFilePath(pType) = newACPath
    Else
      err.raise Number:=ErrCode.FileNotFound, Description:=getDBUtil_PathTypeLabel(pType) & " Access Control path is invalid!" & vbLf & _
        "Reinstalling the database might correct the problem." & vbLf & _
        "Otherwise, please contact your administrator!" & vbLf & _
        LocalSetting(LOCAL_SETTING.SupportName) & " at " & LocalSetting(LOCAL_SETTING.SupportContact)
    End If
  End If
End Sub

Public Function requestAccessControlPath( _
  ByVal pType As dbPathTypeEnum _
) As String
  Dim path As String
  path = Me.AccessControlFilePath(pType)
  requestAccessControlPath = Util.File.GetFilePathViaDialogWindow( _
    InitialFilename:=getLastUsedFolder(path), _
    title:="Select """ & getDBUtil_PathTypeLabel(pType) & """ Access Control File...", _
    FilterDesc:=Util.File.GetFileName(path), _
    FilterExt:="*." & Util.File.GetFileExtension(path))
End Function

Public Sub updateWorkingBackEnd( _
)
  'before calling server path update is required if not available
  checkBackEndAvailability Server
  
  Dim srcPath As String, dstPath As String
  srcPath = Me.BackEndFilePath(Server)
  dstPath = Me.BackEndFilePath(Working)
  
  'check if server and current project are in the same dir -> copy conflict
  If StrComp(srcPath, dstPath, vbTextCompare) = 0 Then
    Exit Sub
  End If
  
  'release database
  Me.BackEndDB.Close
  Me.BackEndDB = Nothing
  
  Dim copyResult As Boolean
  copyResult = Util.File.CopyFile(srcPath, dstPath, True)
  'Alternative: if file.copyFile is faulty quit app and use powershell cmtdlet copy-item -force
  
  'update local backend version
  Dim gver As Variant
  gver = getServerBackEndVersion()
  Me.GlobalBackEndVersion = gver
  Me.LocalBackEndVersion = gver
    
  If Not copyResult Then
    err.raise Number:=ErrCode.ErrorLoadingFile, _
      Description:="could not copy " & vbLf & _
      srcPath & vbLf & _
      "to destination" & vbLf & _
      dstPath
  End If
End Sub

Public Sub checkBackEndAvailability( _
  ByVal pType As dbPathTypeEnum _
)
  If Not isBackEndAvailable(pType) Then
    promptForNewBackEndPath pType
  End If
End Sub

Public Function isBackEndAvailable( _
  ByVal pType As dbPathTypeEnum _
) As Boolean
  isBackEndAvailable = isDbFileAvailable(Me.BackEndFilePath(pType))
End Function

Public Sub promptForNewBackEndPath( _
  ByVal pType As dbPathTypeEnum _
)
  Dim ans As VbMsgBoxResult
  Dim newBEPath As Variant

  ans = MsgBox("The current """ & getDBUtil_PathTypeLabel(pType) & """ BackEnd path to file " & """" & _
    LocalSetting(LOCAL_SETTING.BackEndFileName) & """" & _
    "is invalid or not accessible!" & vbLf & _
    "Would you like to update the """ & getDBUtil_PathTypeLabel(pType) & """ BackEnd filepath yourself?", _
    vbYesNo, "BackEnd file/connection missing...")
    
  If ans = vbYes Then
    newBEPath = requestBackEndPath(pType)
    If Util.File.FileExists(newBEPath) Then
      Me.BackEndFilePath(pType) = newBEPath
    Else
      err.raise Number:=ErrCode.FileNotFound, Description:=getDBUtil_PathTypeLabel(pType) & " BackEnd path is invalid!" & vbLf & _
        "Reinstalling the database might correct the problem." & vbLf & _
        "Otherwise, please contact your administrator!" & vbLf & _
        LocalSetting(LOCAL_SETTING.SupportName) & " at " & LocalSetting(LOCAL_SETTING.SupportContact)
    End If
  End If
End Sub

Public Function requestBackEndPath( _
  ByVal pType As dbPathTypeEnum _
) As String
  Dim path As String
  path = Me.BackEndFilePath(pType)
  requestBackEndPath = Util.File.GetFilePathViaDialogWindow( _
    InitialFilename:=getLastUsedFolder(path), _
    title:="Select """ & getDBUtil_PathTypeLabel(pType) & """ BackEnd File...", _
    FilterDesc:=Util.File.GetFileName(path), _
    FilterExt:="*." & Util.File.GetFileExtension(path))
End Function

Private Function isDbFileAvailable( _
  ByVal path As Variant _
) As Boolean
  isDbFileAvailable = (Not IsNull(path) And Util.File.FileExists(path))
End Function

Public Sub checkBackEndVersion()
  If Not isBackEndUpToDate Then
    updateWorkingAccessControl
    updateWorkingBackEnd
  End If
End Sub

Public Function isBackEndUpToDate() As Boolean
  isBackEndUpToDate = False
  If Me.LocalBackEndVersion = getServerBackEndVersion() Then isBackEndUpToDate = True
End Function

Public Sub reLinkToAccessControl()
On Error GoTo reLinkToAccessControl_error
  DoCmd.Hourglass True
  DoCmd.SetWarnings False
  
  link_reLink_Tables _
    db:=Me.AccessControlDB, _
    tableList:=LocalSetting(LOCAL_SETTING.AccessControlLinkTables), _
    pass:=getAccessControlCredential()
  
reLinkToAccessControl_Exit:
On Error Resume Next
DoCmd.Hourglass False
DoCmd.SetWarnings True
Exit Sub

reLinkToAccessControl_error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure reLinkToAccessControl of DbController:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume reLinkToAccessControl_Exit
End Sub

Public Function getAccessControlCredential() As String
  'TODO
  'check for user existance and encr password
  'if user does not exist err msg to contact admin for access rights
  'userform to apply for access rights
  getAccessControlCredential = ACCESS_CONTROL_PASSWORD
End Function

Public Sub reLinkToBackEndPre()
On Error GoTo reLinkToBackEndPre_error
  DoCmd.Hourglass True
  
  link_reLink_Tables _
   Me.BackEndDB, _
   LocalSetting(LOCAL_SETTING.BackEndLinkTablesPre), _
   getBackEndCredential()
  
reLinkToBackEndPre_Exit:
On Error Resume Next
DoCmd.Hourglass False
Exit Sub

reLinkToBackEndPre_error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure reLinkToBackEndPre of DbController:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume reLinkToBackEndPre_Exit
End Sub

Public Sub reLinkToBackEndPost()
On Error GoTo reLinkToBackEndPost_error
  DoCmd.Hourglass True
  
  link_reLink_Tables _
   Me.BackEndDB, _
   LocalSetting(LOCAL_SETTING.BackEndLinkTablesPost), _
   getBackEndCredential()
  
reLinkToBackEndPost_Exit:
On Error Resume Next
DoCmd.Hourglass False
Exit Sub

reLinkToBackEndPost_error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure reLinkToBackEndPost of DbController:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume reLinkToBackEndPost_Exit
End Sub

Public Function getBackEndCredential() As String
  'TODO impelent check in access control integration
  'If Me.AccessControlDB Is Nothing Then Me.AccessControlDB = Me.getAccessControlDatabase()
  'getBackEndCredential = vbNullString
  '  Dim pass As Variant
  '  pass = ELookup("password", "access_control", "username = '" & Environ$("Username") & "'")
  '  If IsNull(pass) Then
  '    MsgBox "User not authorized!", vbExclamation, "Authentication Required"
  '  End If
  getBackEndCredential = BACKEND_PASSWORD
End Function

Public Sub checkFrontEndVersion()
  If Not isFrontEndUpToDate Then
    updateFrontEnd
  End If
End Sub

Public Function isFrontEndUpToDate() As Boolean
  isFrontEndUpToDate = False
  If Me.LocalVersion = getServerFrontEndVersion() Then isFrontEndUpToDate = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : updateFrontEnd
' Purpose   : updates current working frontend with the server version
'---------------------------------------------------------------------------------------
Public Sub updateFrontEnd()
Dim updateScriptPath As String
Dim serverDir As String
Dim destinationDir As String
Dim fename As String
Dim support As String

On Error GoTo updateFrontEnd_Error

  updateScriptPath = Nz(GlobalSetting(GLOBAL_SETTING.UpdateScriptFilePath), "")
  If Not Util.File.FileExists(updateScriptPath) Then _
    err.raise Number:=ErrCode.FileNotFound, _
      Description:="Update Script File not found. Please contact your support team."
  
  serverDir = LocalSetting(LOCAL_SETTING.ServerDirPath)
  If Not Util.File.FolderExists(serverDir) Then _
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="ServerDirPath is invalid. Please contact your support team."
  
  destinationDir = CurrDir()
  fename = CurrentProject.Name
  support = GlobalSetting(GLOBAL_SETTING.SupportContact) & _
    " or " & GlobalSetting(GLOBAL_SETTING.DeveloperContact)
  
  'give admin the choice to cancel update of frontend
  If CurrUser.hasAdminPrivileges() Then
    Dim ans As VbMsgBoxResult
    ans = MsgBox("The current FrontEnd is out of date!" & vbLf & _
      "Would you like to replace it with the newest version from:" & _
      Nz(LocalSetting(LOCAL_SETTING.ServerDirPath), ""), _
      vbYesNo, "FrontEnd update...")
    If ans = vbNo Then Exit Sub
  Else
    MsgBox "A new version of the software has been issues." & vbLf & _
      "The databse will update and restart automatically. This process may take a minute!", _
      vbInformation, "Update procedure..."
  End If
  
  Me.GlobalVersion = getServerFrontEndVersion()
  
  'TODO
  'check powershell availability and version
  Call shell( _
    "powershell" & " " & _
    "-ExecutionPolicy Bypass" & " " & _
    "-file """ & updateScriptPath & """" & " " & _
    "-serverDir """ & serverDir & """" & " " & _
    "-destinationDir """ & destinationDir & """" & " " & _
    "-FEName """ & fename & """" & " " & _
    "-support """ & support & """" _
    , vbNormalFocus)
  
  CurrentProject.Application.Quit

updateFrontEnd_Exit:
On Error Resume Next
Exit Sub

updateFrontEnd_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure updateFrontEnd of DbController:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume updateFrontEnd_Exit
End Sub

Private Function getLastUsedFolder(ByVal path As Variant) As String
  Dim dirPath As String
  dirPath = Util.File.GetFolderPath(path)
  
  If Util.File.FolderExists(dirPath) Then
    getLastUsedFolder = dirPath
  Else
    getLastUsedFolder = Environ("USERPROFILE")
  End If
End Function

Private Sub link_reLink_Tables( _
  ByRef db As DAO.Database, _
  ByVal tableList As Variant, _
  ByVal pass As String _
  )
Dim tblsToRelink As Variant, tblsToLink As Variant, tblList As Variant
  
  If IsNull(tableList) Then Exit Sub
  tblList = Split(tableList, ";")
  
  tblsToRelink = getTableNamesToRelink(tblList, db.Name)
  tblsToLink = getTableNamesToLink(tblList)
  
  Util.Access.linkCreateTables dbPath:=db.Name, tableNames:=tblsToLink, Password:=pass
  Util.Access.relinkTables path:=db.Name, tableNames:=tblsToRelink, Password:=pass
End Sub

'return table names that have a different DATABASE attribute in their connect value than the path provided
Private Function getTableNamesToRelink( _
  ByVal tblList As Variant, _
  ByVal path As String _
) As Variant
Dim al As ArrayList
  Set al = New ArrayList
  getTableNamesToRelink = Array()
  
  Dim i As Integer, dbCon As String
  For i = LBound(tblList) To UBound(tblList)
    dbCon = Util.Access.getLinkedTableDatabasePath(tblList(i))
    If Len(dbCon) > 0 Then
      If StrComp(dbCon, path, vbTextCompare) <> 0 Then
        al.Append tblList(i)
      End If
    End If
  Next i
  
  getTableNamesToRelink = al.Values
End Function

'returns the names of tables not found in the tabledef collection
Private Function getTableNamesToLink(ByVal tblList As Variant) As Variant
Dim al As ArrayList
  Set al = New ArrayList
  getTableNamesToLink = Array()
  
  Dim i As Integer
  For i = LBound(tblList) To UBound(tblList)
    If Len(tblList(i)) > 0 Then ' necessary check cause isTableDef will return false in case tablename is ""
      If Not Util.Access.isTableDef(tblList(i)) Then
        al.Append tblList(i)
      End If
    End If
  Next i
  getTableNamesToLink = al.Values
End Function

Public Function getDBUtil_PathTypeLabel( _
  ByVal pType As dbPathTypeEnum _
) As String
  Select Case pType
  Case dbPathTypeEnum.Server
    getDBUtil_PathTypeLabel = "Server"
  Case dbPathTypeEnum.Working
    getDBUtil_PathTypeLabel = "Local"
  End Select
End Function

'TODO control ribbon depending on user
Private Sub RibbonState()
'On Error Resume Next
'    If Application.CommandBars("Ribbon").Visible Then
'        DoCmd.ShowToolbar "Ribbon", acToolbarNo
'    End If
'
'If Application.CommandBars("Ribbon").height > 100 Then Application.CommandBars.ExecuteMso "MinimizeRibbon"

'Dim accVer As String
'
'    Select Case SysCmd(acSysCmdAccessVer)
'        Case 7: accVer = "95"
'        Case 8: accVer = "97"
'        Case 9: accVer = "2000"
'        Case 10: accVer = "2002"
'        Case 11: accVer = "2003"
'        Case 12: accVer = "2007"
'        Case 13: accVer = "Pirated!"
'        Case 14: accVer = "2010"
'        Case 15: accVer = "2013"
'        Case Else: accVer = "Unknown"
'    End Select
'
'    RibbonState = (Application.CommandBars("Ribbon").Controls(1).Height < 100)
'
'    Select Case RibbonState
'        Case True
'            'Do nothing, already minimized
'        Case False
'            If accVer > 13 Then
'                CommandBars.ExecuteMso "MinimizeRibbon"
'            Else
'                SendKeys "^{F1}", False
'            End If
'    End Select

End Sub

