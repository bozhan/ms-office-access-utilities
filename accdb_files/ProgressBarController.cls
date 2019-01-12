VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBarController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    :
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private Type ProgressComponent
  Bar As Control
  Box As Control
  Label As Control
End Type

Dim compTop As ProgressComponent, compBot As ProgressComponent
Private Const formName = "progressbar"
Dim frm As Form_progressbar 'Object
Private WithEvents cmd As CommandButton
Attribute cmd.VB_VarHelpID = -1
Dim tc As Integer, taskIdx As Integer, subTc As Integer, subTaskIdx As Integer

Private Sub Class_Initialize()
  CloseForm
  DoCmd.OpenForm formName, acNormal, , , , acHidden
  
  If CurrentProject.AllForms(formName).IsLoaded Then Set frm = Forms.item(formName)
  If frm Is Nothing Then
    err.raise _
      Number:=ErrCode.SubscriptOutOfRange, _
      Description:="No Form with the name " & formName & " was found!"
  End If
  
  InitComponents
End Sub

Private Sub Class_Terminate()
  Me.CloseForm
  Set frm = Nothing
End Sub

Private Property Get Tasks() As ProgressComponent
  Tasks = compTop
End Property

Private Property Get SubTasks() As ProgressComponent
  SubTasks = compBot
End Property

Public Sub InitComponents( _
  Optional ByVal progressBGColor As ColorDec = ColorDec.Blue_Accent1, _
  Optional ByVal formCaption As Variant = "" _
)
  With compBot
    Set .Bar = Me.Form.barProgressBot
    Set .Box = Me.Form.boxProgressBot
    Set .Label = Me.Form.lblProgressBot
    .Bar.Width = 0
    .Bar.BackColor = progressBGColor
    .Bar.Visible = False
    .Box.Visible = False
    .Label.Visible = False
  End With
  
  With compTop
    Set .Bar = Me.Form.barProgressTop
    Set .Box = Me.Form.boxProgressTop
    Set .Label = Me.Form.lblProgressTop
    .Bar.Width = 0
    .Bar.BackColor = progressBGColor
    .Bar.Visible = False
    .Box.Visible = False
    .Label.Visible = False
  End With
  
  Me.Caption = formCaption
  Me.Form.Repaint
  DoEvents
  
  Me.TaskIndex = 1 'sets task index to at least 1 task
  Me.TaskCount = 1 'inits task count to 1
  Me.SubTaskCount = 1
  Me.SubTaskIndex = 1
End Sub

Public Property Get Form() As Form_progressbar
  Set Form = frm
End Property

Public Property Let Form(ByRef f As Form_progressbar)
  Set frm = f
End Property

Public Property Get Caption() As Variant
  Caption = Me.Form.Caption
End Property

Public Property Let Caption(ByVal Value As Variant)
  Me.Form.Caption = CStr(Value)
End Property

Public Property Get SubTaskProgressMax() As Single
  SubTaskProgressMax = SubTasks.Box.Width
End Property

Public Property Get SubTaskProgress() As Single
  SubTaskProgress = SubTasks.Bar.Width
End Property

Public Property Let SubTaskProgress(ByVal Value As Single)
  If Value > Me.SubTaskProgressMax Then Exit Property
  showComponent SubTasks
  SubTasks.Bar.Width = Value
  SubTasks.Label.Caption = Format((Me.SubTaskProgress / Me.SubTaskProgressMax) * 100, "0") & "%"
  DoEvents
End Property

Public Property Get TaskProgressMax() As Single
  TaskProgressMax = Tasks.Box.Width
End Property

Public Property Get TaskProgress() As Single
  TaskProgress = Tasks.Bar.Width
End Property

Public Property Let TaskProgress(ByVal Value As Single)
  If Value > Me.TaskProgressMax Then Exit Property
  showComponent Tasks
  Tasks.Bar.Width = Value
  Tasks.Label.Caption = Format((Me.TaskProgress / Me.TaskProgressMax) * 100, "0") & "%"
  DoEvents
End Property

Private Sub showComponent(ByRef comp As ProgressComponent)
   With comp
    .Box.Visible = True
    .Bar.Visible = True
    .Label.Visible = True
  End With
End Sub

Private Sub hideComponent(ByRef comp As ProgressComponent)
   With comp
    .Box.Visible = False
    .Bar.Visible = False
    .Label.Visible = False
  End With
End Sub

Public Property Get SubTaskProgressLeft() As Single
  SubTaskProgressLeft = Me.SubTaskProgressMax - Me.SubTaskProgress
End Property

Public Property Get TaskProgressLeft() As Single
  TaskProgressLeft = Me.TaskProgressMax - Me.TaskProgress
End Property

Public Property Get Status() As Variant
  Status = Me.Form.lblStatus.Caption
End Property

Public Property Let Status(ByVal Value As Variant)
  Me.Form.lblStatus.Caption = CStr(Value)
End Property

Property Get TaskCount() As Integer
  TaskCount = tc
End Property

Property Let TaskCount(ByVal Value As Integer)
  tc = Value
End Property

Property Get TaskIndex() As Integer
  TaskIndex = taskIdx
End Property

Property Let TaskIndex(ByVal Value As Integer)
  taskIdx = Value
End Property

Property Get SubTaskCount() As Integer
  SubTaskCount = subTc
End Property

Property Let SubTaskCount(ByVal Value As Integer)
  subTc = Value
End Property

Property Get SubTaskIndex() As Integer
  SubTaskIndex = subTaskIdx
End Property

Property Let SubTaskIndex(ByVal Value As Integer)
  subTaskIdx = Value
End Property

Public Sub ShowForm()
  If Not CurrentProject.AllForms(formName).IsLoaded Then
    Class_Initialize
  End If
  Me.Form.Visible = True
End Sub

Public Sub HideForm()
  Me.Form.Visible = False
End Sub

Public Sub CloseForm()
  If CurrentProject.AllForms(formName).IsLoaded Then
    DoCmd.Close acForm, formName
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IncrementTaskProgress
' Purpose   : uses the predefined taskCount and step to increment the overall prrogress
' shown on the SubTaskProgress bar form
' Optional a status text change may be provided, this will make it easy to change SubTaskProgress
' in the form with one-line
'---------------------------------------------------------------------------------------
Public Sub IncrementTaskProgress( _
  Optional ByVal StatusText As Variant, _
  Optional ByVal Step As Integer = 1 _
)
  'reset sub task progress when new main task starts
  Me.SubTaskCount = 1
  Me.SubTaskIndex = 1
  
  'set sub task progress to max if end of task progress was reached
  If Me.TaskIndex = Me.TaskCount Then
    Me.SubTaskProgress = Me.SubTaskProgressMax
  Else
    Me.SubTaskProgress = 0
  End If
  
  'if IncrementTaskProgress was called with overflow TaskIndex we increase the TaskCount
  If Me.TaskIndex > Me.TaskCount Then
    Me.TaskIndex = Me.TaskCount
    Me.TaskCount = Me.TaskCount + 1
  End If
  
  Me.TaskProgress = ((Me.TaskIndex + (Step - 1)) / Me.TaskCount) * Me.TaskProgressMax
  Me.TaskIndex = Me.TaskIndex + 1
    
  If Not IsMissing(StatusText) Then Me.Status = StatusText
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IncrementSubTaskProgress
' Purpose   : uses the predefined subTaskCount and step with subTaskIndex to increment
' the overall prrogress shown on the SubTaskProgress bar form
' Optional a status text change may be provided, this will make it easy to change SubTaskProgress
' in the form with one-line
'---------------------------------------------------------------------------------------
Public Sub IncrementSubTaskProgress( _
  Optional ByVal StatusText As Variant, _
  Optional ByVal Step As Integer = 1 _
)
  'if IncrementTaskProgress was called with overflow SubTaskIndex we increase the SubTaskCount
  If Me.SubTaskIndex > Me.SubTaskCount Then
    Me.SubTaskIndex = Me.SubTaskCount
    Me.SubTaskCount = Me.SubTaskCount + 1
  End If
  
  Me.SubTaskProgress = ((Me.SubTaskIndex + (Step - 1)) / Me.SubTaskCount) * Me.SubTaskProgressMax
  Me.SubTaskIndex = Me.SubTaskIndex + 1
  If Not IsMissing(StatusText) Then Me.Status = StatusText
End Sub







