VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const LOG_TABLE_NAME As String = "log"

Public Sub logActivity(ByVal activity As String, ByVal duration As Double)
  If Not IsLoggingEnabled Then Exit Sub
  Dim db As Database
  Dim logTableName As String
  Set db = CurrDb()
  
  db.Execute "INSERT INTO " & LOG_TABLE_NAME & _
    " (username, activity, duration) VALUES " & _
    "(" & "'" & CurrUser.username & "'" & ", " & "'" & activity & "'" & ", " & "'" & CStr(Round(duration, 3)) & "'" & ")"
  
  Set db = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsLoggingEnabled
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function IsLoggingEnabled() As Boolean
On Error GoTo IsLoggingEnabled_Error

  IsLoggingEnabled = CBool(Nz(GlobalSetting(GLOBAL_SETTING.ActivityLoggingEnabled), False))

IsLoggingEnabled_Exit:
On Error Resume Next
Exit Function

IsLoggingEnabled_Error:
  IsLoggingEnabled = True
  'log error
  GoTo IsLoggingEnabled_Exit
End Function
