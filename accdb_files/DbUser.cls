VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim usrName As String
Dim restForms As Variant 'form names array will be initaited at first read
Dim restControls As Collection

Property Get username() As String
  username = usrName
End Property

Property Let username(ByVal nameValue As String)
  usrName = nameValue
End Property

Property Get ID() As Long
  ID = dbUtil.Permissioner.getUserId(usrName)
End Property

'avoid db read every property call, update only on obj creation or username change
Property Get ClearanceLevel() As Single
  ClearanceLevel = dbUtil.Permissioner.getClearanceLevelByUser(usrName)
End Property

Property Get RestrictedFormNames() As Variant
  If Util.Arrays.IsEmpty(restForms) Then
    restForms = dbUtil.Permissioner.getFormNamesRestrictedForClearanceLevel(Me.ClearanceLevel)
  End If
  RestrictedFormNames = restForms
End Property

Property Get RestrictedControlNamesByForm(ByVal formName As String) As Variant
  If restControls.count <= 0 Then initRestrictedControlNamesCollection
  RestrictedControlNamesByForm = restControls(formName)
End Property

Private Sub Class_Initialize()
  usrName = ""
  setDefaults
End Sub

'IMPORTANT: there might not be a form restriction, but only a control restriction
' therefore we have to obtain the form name for each restricted control using getFromNamesOfRestricedContols()
' from DbPermissioner
Private Sub initRestrictedControlNamesCollection()
  Dim formNames As Variant, l As Long
  formNames = Me.RestrictedFormNames
  For l = LBound(formNames) To UBound(formNames)
    restControls.Add _
      dbUtil.Permissioner.getControlNamesRestrictedForClearanceLevel(Me.ClearanceLevel, formNames(l)), _
      CStr(formNames(l))
  Next l
End Sub

'---------------------------------------------------------------------------------------
' Procedure : hasAdminPrivileges
' Purpose   : retrieves current user clearance level and compares with admin level
'   on error the function will return false, making sure the user does not get clearance
'---------------------------------------------------------------------------------------
Public Function hasAdminPrivileges()
On Error GoTo hasAdminPrivileges_Error

  hasAdminPrivileges = (Me.ClearanceLevel >= dbUtil.Permissioner.getClearanceLevelByLabel("Admin"))

hasAdminPrivileges_Exit:
On Error Resume Next
Exit Function

hasAdminPrivileges_Error:
Debug.Print "Clearance level could not be obtained for: " & Me.username & vbLf & "Admin Privileges have been denied."
hasAdminPrivileges = False
Resume hasAdminPrivileges_Exit
End Function

Private Sub setDefaults()
  restForms = Array()
  Set restControls = New Collection
End Sub

Public Sub Refresh()
  setDefaults
End Sub
