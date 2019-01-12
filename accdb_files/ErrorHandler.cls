VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum ErrCode
  NoError = 0
  InvalidProcedureCall = 5
  Overflow = 6
  OutOfMemory = 7
  SubscriptOutOfRange = 9
  ArrayIsFixedOrLocked = 10
  DivisionByZero = 11
  TypeMismatch = 13
  OutOfStringSpace = 14
  CannotPerformRequestedOperation = 17
  OutOfStackSspace = 28
  FileNotFound = 53
  BadFileMode = 54
  FileAlreadyOpen = 55
  DeviceIOError = 57
  FileAlreadyExists = 58
  PermissionDenied = 70
  PathFileAccessError = 75
  PathNotFound = 76
  ObjectVariableNotSet = 91
  InvalidFileFormat = 321
  CannotAccessRegistry = 335
  PropertyOrMethodNotFound = 423
  ObjectRequired = 424
  ObjectDoesnotSupportProperty = 438
  NamedArgumentNotFound = 448
  ArgumentNotOptional = 449
  ErrorSavingFile = 31036
  ErrorLoadingFile = 31037
End Enum

Public Enum CustomErrCode
  AccessDenied = 1
End Enum

Public Enum WarningMsgType
  DeleteFromList = 0
End Enum
'***
'* Zweck:   Fehlerroutine
'* Infos:   strErrNo = Fehlernummer
'*          strErrMsg = Fehlermeldung
'*          strModul = Name des Moduls, in dem der Fehler aufgetreten ist
'***
Sub prc_ErrorHandler(strErrNo As String, strErrMsg As String, strModul As String)

On Error GoTo Err_prc_ErrorHandler

    '*** Fehlermeldung ausgeben
    MsgBox strErrNo & ": " & strErrMsg, vbCritical, "Fehler in " & strModul

Exit_prc_ErrorHandler:
    Exit Sub

Err_prc_ErrorHandler:
    MsgBox err.Number & ": " & err.Description, vbCritical, "Fehler in General:prc_ErrorHandler"
    Resume Exit_prc_ErrorHandler

End Sub

Public Function getWarningMsg( _
  ByVal msgType As WarningMsgType, _
  ByVal insertText As String _
) As String
Dim msg As String
Dim repStr As String

  Select Case msgType
    Case WarningMsgType.DeleteFromList
      msg = Nz(GlobalSetting(GLOBAL_SETTING.MsgWarningDeleteFromList), "")
    Case Else
      msg = Nz(GlobalSetting(GLOBAL_SETTING.MsgWarningUndefinedOperation), "")
  End Select
  repStr = Nz(GlobalSetting(GLOBAL_SETTING.TextReplaceString), "")
  getWarningMsg = Replace(msg, repStr, insertText)
End Function

'---------------------------------------------------------------------------------------
' Procedure : logErrorToFile
' Purpose   : Logs the occured error with a timestamp to an external log file
'---------------------------------------------------------------------------------------
'
Public Sub logErrorToFile(err As ErrObject, errSource As String)
Dim logFilePath As String
Dim logString As String
Dim logDir As String
On Error GoTo logErrorToFile_Error

    logDir = Nz(GlobalSetting(GLOBAL_SETTING.LoggingDirEnviron), "C:\Temp")
    
    If Not Util.File.FolderExists(logDir) Then _
      err.raise Number:=ErrCode.PathNotFound, _
        Description:="log path was not found or invalid"
    
    logFilePath = logDir & "\" & Util.DateTime.GetDateStamp() & "_PIdb_ErrorLog.txt"
      
    'Open and wrtie title for log
    Open logFilePath For Append As #1
    logString = Now() & ":" & _
            CurrentUser() & VBA.vbTab & _
            "Error " & err.Number & " (" & err.Description & ")" & " in " & _
            errSource
    Print #1, logString
    Close #1
    
logErrorToFile_Exit:
On Error Resume Next
  Exit Sub

logErrorToFile_Error:
  'logErrorToDb err, "logErrorToFile of Module modGlobal"
  Resume logErrorToFile_Exit
End Sub

