VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbServicer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const DEFAULT_DEVELOPER_NAME As String = "Bozhan Ivanov"
Private Const DEFAULT_CONTACT_DEV As String = ""
Private Const DEFAULT_INFO_ERR_MSG As String = "For additional support, pleace contact " & _
  DEFAULT_DEVELOPER_NAME & _
  " at " & _
  DEFAULT_CONTACT_DEV

Public Function InfoErrMsg() As String
On Error GoTo InfoErrMsg_Error
Dim errMsg As String
  
  If IsNull(LocalSetting(LOCAL_SETTING.SupportName)) _
  Or IsNull(LocalSetting(LOCAL_SETTING.SupportContact)) Then _
    GoTo InfoErrMsg_Error
  
  errMsg = "For additional support, pleace contact " & _
    LocalSetting(LOCAL_SETTING.SupportName) & _
    " at " & _
    LocalSetting(LOCAL_SETTING.SupportContact)
  
InfoErrMsg_Exit:
  InfoErrMsg = errMsg
On Error Resume Next
Exit Function

InfoErrMsg_Error:
  errMsg = DEFAULT_INFO_ERR_MSG
  Debug.Print err.Number, err.Description
Resume InfoErrMsg_Exit
End Function

Public Sub displayGeneralMessage( _
  ByVal title As String, _
  ByVal message As String, _
  Optional ByVal form_caption As String = " " _
)
Dim frm As Form_general_message

  DoCmd.OpenForm "general_message", acNormal, , , , acHidden
  If CurrentProject.AllForms("general_message").IsLoaded Then
    Set frm = Forms.item("general_message")
    With frm
      .Caption = form_caption
      .lblTitle.Caption = title
      .lblMessage.Caption = message
    End With
    frm.Visible = True
  End If
End Sub

Public Sub ScheduleMeeting( _
  ByRef rst As Recordset _
)
  Util.Outlook.createOutlookAppointment _
    getEmailsFromEmployeesRecordsetOfRequiredParticipants(rst), _
    getEmailsFromEmployeesRecordsetOfOptionalParticipants(rst), _
    Array(), _
    Nz(GlobalSetting(GLOBAL_SETTING.MeetingSubject), ""), _
    Nz(GlobalSetting(GLOBAL_SETTING.MeetingLocation), ""), _
    Nz(CInt(GlobalSetting(GLOBAL_SETTING.MeetingDuration)), 0)
End Sub

Public Sub createEmail( _
  ByRef rst As Recordset _
)
  Util.Outlook.createEmail _
    getEmailsFromEmployeesRecordsetOfRequiredParticipants(rst), _
    getEmailsFromEmployeesRecordsetOfOptionalParticipants(rst), _
    Array(), _
    Nz(GlobalSetting(GLOBAL_SETTING.EmailSubject), ""), _
    Nz(GlobalSetting(GLOBAL_SETTING.EmailBodyText), ""), _
    OlImportance.olImportanceNormal, _
    Nz(GlobalSetting(GLOBAL_SETTING.EmailAttachmentPath), "")
End Sub

Public Function getEmailsFromEmployeesRecordset( _
  rst As Recordset _
) As Variant
  getEmailsFromEmployeesRecordset = Util.Access.getRecordsetFieldValues(rst, "email")
End Function

Public Function getEmailsFromEmployeesRecordsetOfRequiredParticipants( _
  rst As Recordset _
) As Variant
  
  rst.Filter = "loa_name not like ""*optional*"""
  Dim rstFiltered As Recordset
  Set rstFiltered = rst.OpenRecordset
  
  getEmailsFromEmployeesRecordsetOfRequiredParticipants = _
    Util.Access.getRecordsetFieldValues(rstFiltered, "email")
  
'  getEmailsFromEmployeesRecordsetOfRequiredParticipants = Util.Access.getRecordsetFieldValuesOnCondition( _
    rst, _
    "email", _
    "loa_name", _
    "optional", _
    AcLookAt.Part, _
    True)
End Function

Public Function getEmailsFromEmployeesRecordsetOfOptionalParticipants( _
  rst As Recordset _
) As Variant
  
  rst.Filter = "loa_name like ""*optional*"""
  Dim rstFiltered As Recordset
  Set rstFiltered = rst.OpenRecordset
  
  getEmailsFromEmployeesRecordsetOfOptionalParticipants = _
    Util.Access.getRecordsetFieldValues(rstFiltered, "email")

'  getEmailsFromEmployeesRecordsetOfOptionalParticipants = Util.Access.getRecordsetFieldValuesOnCondition( _
    rst, _
    "email", _
    "loa_name", _
    "optional", _
    AcLookAt.Part, _
    False)
End Function



