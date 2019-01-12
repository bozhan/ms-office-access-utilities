VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : UtilRegistry
' Author      : Ivanov, Bozhan
' Description : Windows registry management functions
'  Return Types for ReadRef:
'  REG_SZ - string
'  REG_DWORD - integer
'  REG_BINARY - VBArray of integers
'  REG_EXPAND_SZ - string
'  REG_MULTI_SZ - VBArray of strings
'---------------------------------------------------------------------------------------

Option Explicit

Public Enum RegestryValueType
  REG_SZ = 0
  REG_MULTI_SZ = 1
  REG_EXPAND_SZ = 2
  REG_DWORD = 3
  REG_QWORD = 4
  REG_BINARY = 5
End Enum

Private sh As Object 'WshShell

Private Property Get shell() As Object 'WshShell
  If sh Is Nothing Then Set sh = CreateObject("wscript.shell")
  Set shell = sh
End Property

Private Sub Class_Terminate()
  Set shell = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsKeyAvailable
' Purpose   : Checks if a regitry key exists and is accesible and returns true if so
'---------------------------------------------------------------------------------------
Public Function IsKeyAvailable( _
  ByVal KeyPath As String _
) As Boolean
  Dim result As Variant
  
  On Error Resume Next
  result = shell.RegRead(KeyPath)
  On Error GoTo 0
  
  IsKeyAvailable = Not (IsNull(result) Or IsEmpty(result))
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsValueAvailable
' Purpose   : Checks if a regitry value exists and is accesible and returns true if so
'---------------------------------------------------------------------------------------
Public Function IsValueAvailable( _
  ByVal KeyName As String, _
  ByVal ValueName As String _
) As Boolean

  If StrComp(Right(KeyName, 1), "\", vbBinaryCompare) <> 0 Then
    KeyName = KeyName & "\"
  End If
  
  IsValueAvailable = IsKeyAvailable(KeyName & ValueName)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadKey
' Purpose   : Returns regestry value if available, else returns empty string
'---------------------------------------------------------------------------------------
Public Function ReadKey( _
  ByVal KeyPath As String _
) As Variant
  
  Dim result As Variant
  On Error Resume Next
  result = shell.RegRead(KeyPath)
  On Error GoTo 0
  
  ReadKey = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadValue
' Purpose   : Returns regestry value if available, else returns empty string
'---------------------------------------------------------------------------------------
Public Function ReadValue( _
  ByVal KeyName As String, _
  ByVal ValueName As String _
) As Variant

  If StrComp(Right(KeyName, 1), "\", vbBinaryCompare) <> 0 Then
    KeyName = KeyName & "\"
  End If
  
  ReadValue = ReadKey(KeyName & ValueName)
End Function

'---------------------------------------------------------------------------------------
' Procedure : writeValue
' Purpose   : Writes a registry value
'---------------------------------------------------------------------------------------
Public Function writeValue( _
  ByVal KeyName As String, _
  ByVal ValueName As String, _
  ByVal Value As String, _
  ByVal ValueType As RegestryValueType _
) As Boolean
Dim shell As Object
Dim iResult As Integer
Dim valType As String

  iResult = 1 ' set iResult to 1 to aid error handling
  writeValue = True
  
  Set shell = CreateObject("wscript.shell")
  
  If StrComp(Right(KeyName, 1), "\", vbBinaryCompare) <> 0 Then ' KeyName dows not end with "\"
    KeyName = KeyName & "\"
  End If
  
  Select Case ValueType
    Case REG_SZ
      valType = "REG_SZ"
    Case REG_MULTI_SZ
      valType = "REG_MULTI_SZ"
    Case REG_EXPAND_SZ
      valType = "REG_EXPAND_SZ"
    Case REG_DWORD
      valType = "REG_DWORD"
    Case REG_QWORD
      valType = "REG_QWORD"
    Case REG_BINARY
      valType = "REG_BINARY"
  End Select
  
  On Error Resume Next
  'iResult will be 0 if writeValue was successful and <> 0 when not
  iResult = shell.RegWrite(KeyName & ValueName, Value, valType)
  On Error GoTo 0
  
  ' set return to true if the reg value was successfully written
  If iResult <> 0 Then writeValue = False
    
  Set shell = Nothing
End Function


