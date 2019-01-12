VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim uw As UtilWindows
Dim ar As UtilArrays
Dim dt As UtilDateTime
Dim ex As UtilExcel
Dim fl As UtilFile
Dim OL As UtilOutlook
Dim prj As UtilProject
Dim reg As UtilRegistry
Dim str As UtilStrings
Dim uc As UtilUnits
Dim frm As UtilForm
Dim um As UtilModules
Dim rx As UtilRegEx
Dim rxp As UtilRegExPatterns
Dim pt As UtilTimer
Dim clr As UtilColor
Dim uh As UtilHash
Dim sc As UtilSmartCard
Dim cr As UtilCrypt
Dim acc As UtilAccess

Private Sub Class_Initialize()
  Set clr = New UtilColor
  Set uw = New UtilWindows
  Set ar = New UtilArrays
  Set dt = New UtilDateTime
  Set ex = New UtilExcel
  Set fl = New UtilFile
  Set OL = New UtilOutlook
  Set prj = New UtilProject
  Set reg = New UtilRegistry
  Set str = New UtilStrings
  Set uc = New UtilUnits
  Set frm = New UtilForm
  Set um = New UtilModules
  Set rx = New UtilRegEx
  Set rxp = New UtilRegExPatterns
  Set pt = New UtilTimer
  Set acc = New UtilAccess
End Sub

Private Sub Class_Terminate()
  Set clr = Nothing
  Set uw = Nothing
  Set ar = Nothing
  Set dt = Nothing
  Set ex = Nothing
  Set fl = Nothing
  Set OL = Nothing
  Set prj = Nothing
  Set reg = Nothing
  Set str = Nothing
  Set uc = Nothing
  Set frm = Nothing
  Set um = Nothing
  Set rx = Nothing
  Set rxp = Nothing
  Set pt = Nothing
  Set acc = Nothing
End Sub

Property Get Access() As UtilAccess
  If acc Is Nothing Then Set acc = New UtilAccess
  Set Access = acc
End Property

Property Get Hash() As UtilHash
  If uh Is Nothing Then Set uh = New UtilHash
  Set Hash = uh
End Property

Property Get SmartCard() As UtilSmartCard
  If sc Is Nothing Then Set sc = New UtilSmartCard
  Set SmartCard = sc
End Property

Property Get Crypt() As UtilCrypt
  If cr Is Nothing Then Set cr = New UtilCrypt
  Set SmartCard = cr
End Property

Property Get Timer() As UtilTimer
  Set Timer = pt
End Property

Property Get Windows() As UtilWindows
  Set Windows = uw
End Property

Property Get Arrays() As UtilArrays
  Set Arrays = ar
End Property

Property Get Color() As UtilColor
  Set Color = clr
End Property

Property Get DateTime() As UtilDateTime
  Set DateTime = dt
End Property

Property Get Excel() As UtilExcel
  Set Excel = ex
End Property

Property Get File() As UtilFile
  Set File = fl
End Property

Property Get Outlook() As UtilOutlook
  Set Outlook = OL
End Property

Property Get Project() As UtilProject
  Set Project = prj
End Property

Property Get Registry() As UtilRegistry
  Set Registry = reg
End Property

Property Get Strings() As UtilStrings
  Set Strings = str
End Property

Property Get Units() As UtilUnits
  Set Units = uc
End Property

Property Get Form() As UtilForm
  Set Form = frm
End Property

Property Get Modules() As UtilModules
  Set Modules = um
End Property

Property Get RegEx() As UtilRegEx
  Set RegEx = rx
End Property

Property Get RegExPatterns() As UtilRegExPatterns
  Set RegExPatterns = rxp
End Property

Public Sub Refresh()
  Call Class_Initialize
End Sub



