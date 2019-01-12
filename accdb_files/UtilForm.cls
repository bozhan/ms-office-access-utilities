VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FormUtil
' Author    : Ivanov, Bozhan
' Purpose   :
' Dependancy: Requires Microsoft Forms 2.0 Object Library
'   C:\Windows\SysWOW64\FM20.dll
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : IsFormLoaded
' Descr.    :
'---------------------------------------------------------------------------------------
Public Function IsFormLoaded( _
  ByVal userFormObj As Object _
) As Boolean
  
  IsFormLoaded = False
  
  Dim frm As Variant
  For Each frm In VBA.UserForms
    If frm.Name = userFormObj.Name Then
      IsFormLoaded = True
      Exit Function
    End If
  Next
End Function

Public Function IsFormLoadedByName( _
  ByVal userFormName As String _
) As Boolean
  
  IsFormLoadedByName = False
  
  Dim frm As Variant
  For Each frm In VBA.UserForms
    If frm.Name = userFormName Then
      IsFormLoadedByName = True
      Exit Function
    End If
  Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateListChoiceForm
' Descr.    : Displays a list box with the provided Initial List and returns the selected
'   entries from the list in a form of a string array
'   if cancel was pressed an empty array will be returned
'---------------------------------------------------------------------------------------
'
Public Function GenerateListChoiceForm( _
  initList As Variant, _
  Optional Caption As String = "Select from list ...", _
  Optional text As String = "", _
  Optional Hint As String = "", _
  Optional lMultiSelect As MSforms.fmMultiSelect _
) As Variant
  GenerateListChoiceForm = Array()
      
  If Util.Arrays.GetDimensions(initList) > 1 Then
    err.raise Number:=ErrCode.InvalidProcedureCall, _
      Description:="Provided initial list array argument is a multi dimentional array!" & vbLf & _
      "This function accepts only single dimentional array lists."
  End If
  
  If Not Util.Arrays.IsOfSimpleDataTypes(initList) Then
    err.raise Number:=ErrCode.InvalidProcedureCall, _
      Description:="Provided values in the initial list array argument are not of simpel data types!" & vbLf & _
      "This function accepts only simple data typed array arguments that can be displayes as strings."
  End If
  
  'Initialize form with parameters
  Dim frm As New frmListChoice
  frm.Caption = Caption
  frm.lblText.Caption = text
  frm.lblHint.Caption = Hint
  frm.lstSelections.MultiSelect = lMultiSelect
  If lMultiSelect = fmMultiSelectExtended Then
    frm.lblHint.Caption = "Hint: You can select multiple values by using Ctrl or Shift"
  End If
  If lMultiSelect = fmMultiSelectMulti Then
    frm.lblHint.Caption = "Hint: You can select multiple values"
  End If
  frm.lstSelections.Clear
  
  ' Fill list box with initial data
  Dim i As Long
  For i = LBound(initList) To UBound(initList)
    frm.lstSelections.AddItem CStr(initList(i))
  Next i
  
  ' show and wait for button event checked by isFormLoaded()
  frm.Show
  
  ' get the value of all selected entities form list box
  Dim result As New ArrayList
  If IsFormLoaded(frm) Then 'ok was pressed
    For i = 0 To frm.lstSelections.ListCount - 1
      If frm.lstSelections.Selected(i) = True Then
        result.Append frm.lstSelections.List(i)
      End If
    Next i
  Else 'cancel was pressed
    result.Clear
  End If

  GenerateListChoiceForm = result.Values
End Function

Public Sub SetControlSizePosition( _
ByRef crl As Object, _
ByVal Top As Single, _
ByVal Height As Single, _
ByVal Left As Single, _
ByVal Width As Single _
)
  crl.Top = Top
  crl.Height = Height
  crl.Left = Left
  crl.Width = Width
End Sub

Public Sub SetControlPositionToRelativeObject( _
  ByRef crl As Object, _
  ByRef RelativeObject As Object, _
  ByVal TopOffset As Single, _
  ByVal LeftOffset As Single, _
  Optional ByVal Overlay As Boolean = False _
)
 If Overlay Then
    crl.Left = getLeftOffsetWithOverlay(RelativeObject, LeftOffset)
    crl.Top = getTopOffsetWithOverlay(RelativeObject, TopOffset)
  Else
    crl.Left = getLeftOffsetWithoutOverlay(crl, RelativeObject, LeftOffset)
    crl.Top = getTopOffsetWithoutOverlay(crl, RelativeObject, TopOffset)
  End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : addLabel
' Description : adds a label to a provided form obejct
'   if a relateive object is provided then the label position is se to the relative object
'   taking into consideration Top and Left offsets w.r.t the relative object
'   The overlay parameter allows the new label to overlay the relative object or not
'---------------------------------------------------------------------------------------
Public Function AddLabel( _
  ByRef Form As Object, _
  ByVal Name As String, _
  Optional ByVal text As String = "", _
  Optional ByVal Top As Single = 0, _
  Optional ByVal Height As Single = 12, _
  Optional ByVal Left As Single = 0, _
  Optional ByVal Width As Single = 50, _
  Optional ByVal IsVisible As Boolean = True, _
  Optional ByVal RelativeObject As Object = Nothing, _
  Optional ByVal TopOffset As Single = 0, _
  Optional ByVal LeftOffset As Single = 0, _
  Optional ByVal Overlay As Boolean = False _
) As Object
  Dim ctl As Object
  Set ctl = Form.controls.Add("Forms.Label.1", Name, IsVisible)
  ctl.Caption = text
  SetControlSizePosition ctl, Top, Height, Left, Width
  
  If Not RelativeObject Is Nothing Then
    SetControlPositionToRelativeObject ctl, RelativeObject, TopOffset, LeftOffset, Overlay
  End If
  
  Set AddLabel = Form.controls(Name)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : AddTextBox
' Description : adds a text box to a provided form obejct
'   if a relateive object is provided then the label position is se to the relative object
'   taking into consideration Top and Left offsets w.r.t the relative object
'   The overlay parameter allows the new label to overlay the relative object or not
'---------------------------------------------------------------------------------------
Public Function AddTextBox( _
  ByRef Form As Object, _
  ByVal Name As String, _
  Optional ByVal text As String = "", _
  Optional ByVal Top As Single = 0, _
  Optional ByVal Height As Single = 18, _
  Optional ByVal Left As Single = 0, _
  Optional ByVal Width As Single = 72, _
  Optional ByVal IsVisible As Boolean = True, _
  Optional ByVal RelativeObject As Object = Nothing, _
  Optional ByVal TopOffset As Single = 0, _
  Optional ByVal LeftOffset As Single = 0, _
  Optional ByVal Overlay As Boolean = False _
) As Object
  Dim ctl As Object
  Set ctl = Form.controls.Add("Forms.TextBox.1", Name, IsVisible)
  SetControlSizePosition ctl, Top, Height, Left, Width
    
  If Not RelativeObject Is Nothing Then
    SetControlPositionToRelativeObject ctl, RelativeObject, TopOffset, LeftOffset, Overlay
  End If
  
  ctl.text = text
  
  Set AddTextBox = Form.controls(Name)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : AddButton
' Description : adds a button to a provided form obejct
'   if a relateive object is provided then the label position is se to the relative object
'   taking into consideration Top and Left offsets w.r.t the relative object
'   The overlay parameter allows the new label to overlay the relative object or not
'---------------------------------------------------------------------------------------
Public Function AddButton( _
  ByRef Form As Object, _
  ByVal Name As String, _
  Optional ByVal text As String = "", _
  Optional ByVal Top As Single = 0, _
  Optional ByVal Height As Single = 24, _
  Optional ByVal Left As Single = 0, _
  Optional ByVal Width As Single = 72, _
  Optional ByVal IsVisible As Boolean = True, _
  Optional ByVal RelativeObject As Object = Nothing, _
  Optional ByVal TopOffset As Single = 0, _
  Optional ByVal LeftOffset As Single = 0, _
  Optional ByVal Overlay As Boolean = False _
) As Object
  Dim ctl As Object
  Set ctl = Form.Add("Forms.CommandButton.1", Name, IsVisible)
  SetControlSizePosition ctl, Top, Height, Left, Width
  
  If Not RelativeObject Is Nothing Then
    SetControlPositionToRelativeObject ctl, RelativeObject, TopOffset, LeftOffset, Overlay
  End If
  
  ctl.Caption = text
  
  Set AddButton = Form.controls(Name)
End Function

'---------------------------------------------------------------------------------------
' Procedure : getLeftOffsetWithOverlay
' Descr.    : returns the calculated Left offset for the current control object
'             relative to another control object
'             The left value of the control is calculated
'             in such a way that overlay of controls is allowed on the form and the
'             supplied offset value is calc form the inner edges of the relative control
'---------------------------------------------------------------------------------------
'
Private Function getLeftOffsetWithOverlay( _
  ByVal relObj As Object, _
  ByVal offset As Long _
) As Long
  If relObj.Left + offset < 0 Then
    getLeftOffsetWithOverlay = 0
  Else
    getLeftOffsetWithOverlay = relObj.Left + offset
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : getTopOffsetWithOverlay
' Descr.    : returns the calculated Top offset for the current control object
'             relative to another control object
'             The top value of the control is calculated
'             in such a way overlay of controls is allowed on the form and the
'             supplied offset value is calc form the inner edges of the relative control
'---------------------------------------------------------------------------------------
'
Private Function getTopOffsetWithOverlay( _
  ByVal relObj As Object, _
  ByVal offset As Long _
) As Long
  If relObj.Top + offset < 0 Then
    getTopOffsetWithOverlay = 0
  Else
    getTopOffsetWithOverlay = relObj.Top + offset
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : getLeftOffsetWithoutOverlay
' Descr.    : returns the calculated Left offset for the current control object
'             relative to another control object
'             The left value of the control is calculated
'             in such a way that no overlay of controls is allowed on the form and the
'             supplied offset value is calc form the outher edges of the relative control
'---------------------------------------------------------------------------------------
'
Private Function getLeftOffsetWithoutOverlay( _
  ByVal obj As Object, _
  ByVal relObj As Object, _
  ByVal offset As Long _
) As Long

  If offset > 0 Then
    getLeftOffsetWithoutOverlay = relObj.Left + relObj.Width + offset
  ElseIf ((offset < 0) And (relObj.Left + offset - obj.Width >= 0)) Then
    getLeftOffsetWithoutOverlay = relObj.Left + offset - obj.Width
  Else
    If (relObj.Top >= obj.Top + obj.Height) _
    Or (relObj.Top + relObj.Height <= obj.Top) Then 'NO vertical overlap occurs
      getLeftOffsetWithoutOverlay = relObj.Left + offset
    Else
      getLeftOffsetWithoutOverlay = relObj.Left + relObj.Width
    End If
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : getTopOffsetWithoutOverlay
' Descr.    : returns the calculated Left offset for the current control object
'             relative to another control object
'             The top value of the control is calculated
'             in such a way that no overlay of controls is allowed on the form and the
'             supplied offset value is calc form the outher edges of the relative control
'---------------------------------------------------------------------------------------
'
Private Function getTopOffsetWithoutOverlay( _
  ByVal obj As Object, _
  ByVal relObj As Object, _
  ByVal offset As Long _
) As Long
  
  If offset > 0 Then
    getTopOffsetWithoutOverlay = relObj.Top + relObj.Height + offset
  ElseIf ((offset < 0) And (relObj.Top + offset - obj.Height >= 0)) Then
    getTopOffsetWithoutOverlay = relObj.Top + offset - obj.Height
  Else
    If (relObj.Left >= obj.Left + obj.Width) _
    Or (relObj.Left + relObj.Width <= obj.Left) Then 'if no horizontal overlap occurs
      getTopOffsetWithoutOverlay = relObj.Top + offset
    Else
      getTopOffsetWithoutOverlay = relObj.Top + relObj.Height
    End If
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetTextBoxes
' Description : Returns a collection of all text boxes in provided form object
'---------------------------------------------------------------------------------------
Public Function GetTextBoxes( _
  ByRef Form As Object _
) As Collection

  Set GetTextBoxes = New Collection
  Dim ctl As Variant

  For Each ctl In Form.controls
    If TypeOf ctl Is MSforms.TextBox Then
      GetTextBoxes.Add ctl, ctl.Name
    End If
  Next ctl
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetLables
' Description : Returns a collection of all labels in provided form object
'---------------------------------------------------------------------------------------
Public Function GetLables( _
  ByRef Form As Object _
) As Collection

  Set GetLables = New Collection
  Dim ctl As Variant

  For Each ctl In Form.controls
    If TypeOf ctl Is MSforms.Label Then
      GetLables.Add ctl, ctl.Name
    End If
  Next ctl
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetButtons
' Description : Returns a collection of all command buttons in provided form object
'---------------------------------------------------------------------------------------
Public Function GetButtons( _
  ByRef Form As Object _
) As Collection

  Set GetButtons = New Collection
  Dim ctl As Variant

  For Each ctl In Form.controls
    If TypeOf ctl Is MSforms.CommandButton Then
      GetButtons.Add ctl, ctl.Name
    End If
  Next ctl
End Function
