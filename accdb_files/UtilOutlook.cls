VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : OutlookUtil
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit

Public Enum OlItemType
  olAppointmentItem = 1
  olContactItem = 2
  olDistributionListItem = 7
  olJournalItem = 4
  olMailItem = 0
  olNoteItem = 5
  olPostItem = 6
  olTaskItem = 3
End Enum

Public Enum OlMeetingStatus
  olMeeting = 1
  olMeetingCanceled = 5
  olMeetingReceived = 3
  olMeetingReceivedAndCanceled = 7
  olNonMeeting = 0
End Enum

Public Enum OlMailRecipientType
  olBCC = 3
  olCC = 2
  olOriginator = 0
  olTo = 1
End Enum

Public Enum OlMeetingRecipientType
  olOptional = 2
  olOrganizer = 0
  olRequired = 1
  olResource = 3
End Enum

Public Enum OlImportance
  olImportanceLow = 0
  olImportanceNormal = 1
  olImportanceHigh = 2
End Enum

Public Function IsOutlookInstalled( _
) As Boolean
On Error Resume Next
'IsOutlookInstalled = (Not CreateObject("Outlook.Application") Is Nothing)
IsOutlookInstalled = (Not GetOutlookApp() Is Nothing)
End Function

Public Function GetOutlookApp( _
) As Object
Dim olApp As Object

  On Error Resume Next
  Set olApp = GetObject(, "Outlook.Application")
  If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
  End If
  
  Set GetOutlookApp = olApp
End Function

'---------------------------------------------------------------------------------------
' Procedure : createOutlookAppointment
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub createOutlookAppointment( _
  ByVal requiredAttendies As Variant, _
  Optional ByVal optionalAttendies As Variant, _
  Optional ByVal resourceAttendies As Variant, _
  Optional ByVal Subject As String = "", _
  Optional ByVal location As String = "", _
  Optional ByVal duration As Integer = 0 _
)
Dim olApp As Object 'Outlook.Application
Dim appointment As Object 'Outlook.AppointmentItem
Dim reqRecipients As Collection
Dim optRecipients As Collection
Dim resRecipients As Collection
On Error GoTo createOutlookAppointment_Error

  Set olApp = GetOutlookApp()
  If olApp Is Nothing Then _
   err.raise Number:=91, Description:="No MS Outlook installation was detected"
  
  Set appointment = olApp.CreateItem(OlItemType.olAppointmentItem)
  
  With appointment
    .MeetingStatus = OlMeetingStatus.olMeeting
    .Subject = Subject
    .location = location
    .start = Format(DateTime.Date, "dd.mm.yyyy") & " " & Format(DateTime.Time, "hh:mm") '#9/24/2009 1:30:00 PM#
    .duration = duration
    
    'TODO check if all recipients meet the requrements (the match the email pattern)
    Set reqRecipients = assignParticipantsToAppointment( _
      appointment, requiredAttendies, OlMeetingRecipientType.olRequired)
    Set optRecipients = assignParticipantsToAppointment( _
      appointment, optionalAttendies, OlMeetingRecipientType.olOptional)
    Set resRecipients = assignParticipantsToAppointment( _
      appointment, resourceAttendies, OlMeetingRecipientType.olResource)

    .recipients.ResolveAll
    .Display
  End With

createOutlookAppointment_Exit:
On Error Resume Next
  Set olApp = Nothing
  Set appointment = Nothing
  Set reqRecipients = Nothing
  Set optRecipients = Nothing
  Set resRecipients = Nothing
Exit Sub

createOutlookAppointment_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure createOutlookAppointment of OutlookUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume createOutlookAppointment_Exit
End Sub

Private Function assignParticipantsToAppointment( _
  ByRef appointment As Object, _
  ByRef attendies As Variant, _
  ByRef recipType As OlMeetingRecipientType _
) As Collection
Dim recipients As Collection
Dim i As Integer

  Set recipients = New Collection
  For i = LBound(attendies) To UBound(attendies)
    recipients.Add appointment.recipients.Add(attendies(i))
    recipients(recipients.count).Type = recipType
  Next i
  Set assignParticipantsToAppointment = recipients
End Function

'---------------------------------------------------------------------------------------
' Procedure : createEmail
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub createEmail( _
  ByVal requiredRecipients As Variant, _
  Optional ByVal ccRecipients As Variant, _
  Optional ByVal bccRecipients As Variant, _
  Optional ByVal Subject As String = "", _
  Optional ByVal body As String = "", _
  Optional ByVal importance As OlImportance = OlImportance.olImportanceNormal, _
  Optional ByVal attachemntPath As String = "", _
  Optional ByVal sendDirectly As Boolean = False _
)
Dim olApp As Object 'Outlook.Application
Dim email As Object 'Outlook.MailItem
Dim reqRecips As Collection
Dim ccRecips As Collection
Dim bccRecips As Collection
Dim attachemnt As Object 'Outlook.Attachment
On Error GoTo createEmail_Error

  Set olApp = GetOutlookApp()
  If olApp Is Nothing Then _
   err.raise Number:=91, Description:="No MS Outlook installation was detected"
  
  Set email = olApp.CreateItem(OlItemType.olMailItem)

  With email
    .Subject = Subject
    .body = body
    .importance = importance
    
    Set reqRecips = assignParticipantsToEmail( _
      email, requiredRecipients, OlMailRecipientType.olTo)
    
    If Not IsMissing(ccRecipients) Then
      Set ccRecips = assignParticipantsToEmail(email, ccRecipients, OlMailRecipientType.olCC)
    End If
    
    If Not IsMissing(bccRecipients) Then
      Set bccRecips = assignParticipantsToEmail(email, bccRecipients, OlMailRecipientType.olBCC)
    End If
    
  
    If Util.File.FileExists(attachemntPath) Then
      Set attachemnt = .Attachments.Add(attachemntPath)
    End If
    
    .recipients.ResolveAll
   
    If sendDirectly Then
      .Save
      .Send
    Else
      .Display
    End If
  End With

createEmail_Exit:
On Error Resume Next
  Set olApp = Nothing
  Set email = Nothing
  Set reqRecips = Nothing
  Set ccRecips = Nothing
  Set bccRecips = Nothing
  Set attachemnt = Nothing
Exit Sub

createEmail_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure createEmail of OutlookUtil:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume createEmail_Exit
End Sub

Private Function assignParticipantsToEmail( _
  ByRef email As Object, _
  ByRef recipients As Variant, _
  ByRef recipType As OlMailRecipientType _
) As Collection
Dim recips As Collection
Dim i As Integer

  Set recips = New Collection
  For i = LBound(recipients) To UBound(recipients)
    recips.Add email.recipients.Add(recipients(i))
    recips(recips.count).Type = recipType
  Next i
  Set assignParticipantsToEmail = recips
End Function
'---------------------------------------------------------------------------------------
' Procedure : resolveUser
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function resolveUser(ByVal searchText As String) As Object 'ExchangeUser
On Error GoTo resolveUser_Error
Dim app As Object 'Outlook.Application
Dim ti As Object 'Outlook.TaskItem
Dim recp As Object 'Outlook.Recipient
  
  Set resolveUser = Nothing
  Set app = GetOutlookApp()
  Set ti = app.CreateItem(olTaskItem)
  Set recp = ti.recipients.Add(searchText)
  recp.Resolve
  If recp.Resolved Then Set resolveUser = recp.AddressEntry.GetExchangeUser
  'resolveUser = usr.name & " <" & usr.PrimarySmtpAddress & ">"
resolveUser_Exit:
On Error Resume Next
Set app = Nothing
Exit Function

resolveUser_Error:
resolveUser = "Cannot resolve with Outlook! Please enter manually.Debug.Print "
Resume resolveUser_Exit
End Function
