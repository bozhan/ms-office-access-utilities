VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const SETTINGS_TABLE_NAME As String = "settings"
Private Const LOCAL_SETTINGS_TABLE_NAME As String = "local_settings"
Private Const ENVIRONMENT_FIELD_NAME As String = "IsEnvironment"

Public Enum GLOBAL_SETTING
  DeveloperName = 1
  DeveloperContact = 2
  version = 3
  LoggingDirEnviron = 4
  MsgWarningDeleteFromList = 5
  MsgWarningUndefinedOperation = 6
  TextReplaceString = 7
  AllowedExcelExtentions = 8
  AllowedMSProjectExtentions = 9
  PPI = 10
  TPI = 11
  SupportName = 12
  SupportContact = 13
  MeetingLocation = 15
  MeetingSubject = 16
  MeetingDuration = 17
  EmailSubject = 18
  EmailBodyText = 19
  EmailAttachmentPath = 20
  ValueSeperator = 21
  ExcelExportFileName = 22
  ExcelExportSubject = 23
  ExcelExportWorksheetNameData = 24
  ExcelExportWorksheetNameCriteria = 25
  ExcelExportRecordsetFieldNames = 26
  ExcelExportColumnHeaderToFieldNames = 27
  ExcelExportDefaultSaveLocation = 28
  HelpFilePath = 31
  UpdateScriptFilePath = 32
  CorruptDataWarningLabelContent = 34
  ExcelExportCriteriaRecordsetFieldNames = 35
  ExcelExportCriteriaColumnHeadersToFields = 36
  ActivityLoggingEnabled = 37
  SCDIdentitySearchServerAddress = 38
  BackEndVersion = 39
End Enum

Public Enum LOCAL_SETTING
  version = 1
  SupportName = 2
  SupportContact = 3
  BackEndFileName = 4
  AccessControlFileName = 5
  AccessControlLinkTables = 8
  BackEndLinkTablesPre = 9
  ServerDirPath = 10
  BackEndLinkTablesPost = 11
  WorkingDirPath = 12
  BackEndVersion = 13
  TestMode = 14
End Enum

Property Get GlobalSetting( _
  ByVal varID As GLOBAL_SETTING _
) As Variant
  GlobalSetting = getSetting(SETTINGS_TABLE_NAME, varID)
End Property

Property Let GlobalSetting( _
  ByVal varID As GLOBAL_SETTING, _
  ByVal varValue As Variant _
)
  setSetting SETTINGS_TABLE_NAME, varID, varValue
End Property

Property Get LocalSetting( _
  ByVal varID As LOCAL_SETTING _
) As Variant
  LocalSetting = getSetting(LOCAL_SETTINGS_TABLE_NAME, varID)
End Property

Property Let LocalSetting( _
  ByVal varID As LOCAL_SETTING, _
  ByVal varValue As Variant _
)
  setSetting LOCAL_SETTINGS_TABLE_NAME, varID, varValue
End Property

Private Function getSetting( _
  ByVal settingsTableName As String, _
  ByVal varID As Variant _
) As Variant
  If ELookup(ENVIRONMENT_FIELD_NAME, settingsTableName, "SettingID = " & varID) Then
    getSetting = Environ$(ELookup("varValue", settingsTableName, "SettingID = " & varID))
  Else
    getSetting = ELookup("varValue", settingsTableName, "SettingID = " & varID)
  End If
End Function

Private Sub setSetting( _
  ByVal settingsTableName As String, _
  ByVal settingId As Variant, _
  ByVal varValue As Variant _
)
  On Error GoTo 0
  If Not settingExists(settingsTableName, settingId) Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="The setting table """ & settingsTableName & """ " & _
      "or variable name with SettingID """ & settingId & """ " & _
      "you are trying to set does not exist!"
    Exit Sub
  End If
  
  Dim sql As String
  sql = "UPDATE " & settingsTableName & " SET varValue = '" & varValue & "' WHERE SettingID = " & settingId
  
  Dim db As Database
  Set db = CurrDb()
  
  On Error GoTo Err_Execute
  db.Execute sql, dbFailOnError
  On Error GoTo 0
  
Exit_Execute:
  Set db = Nothing
  Exit Sub

Err_Execute:
  Dim errLoop As Error
   If DBEngine.Errors.count > 0 Then
      For Each errLoop In DBEngine.Errors
         MsgBox "Error number: " & errLoop.Number & vbCr & _
            errLoop.Description
      Next errLoop
   End If
   Resume Exit_Execute
End Sub
Private Function settingExists( _
  ByVal settingsTableName As String, _
  ByVal settingId As GLOBAL_SETTING _
) As Boolean
  settingExists = Not IsNull(ELookup("varName", settingsTableName, "SettingID = " & settingId))
End Function

