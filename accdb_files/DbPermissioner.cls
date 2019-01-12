VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbPermissioner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const CLEARANCE_LEVEL_SOURCE As String = "clearance_level"
Private Const USERS_CLEARANCE_LEVELS_SOURCE As String = "users_clearance_levels"
Private Const USERS_CLEARANCE_LEVELS_LEVELS_NAME_COLUMN_NAME As String = "level_name"
Private Const USERS_CLEARANCE_LEVELS_COLUMN_NAME As String = "level"
Private Const USERS_CLEARANCE_LEVELS_COND_COLUMN_NAME As String = "username"

Private Const FORMS_CLEARANCE_LEVEL_SOURCE As String = "forms_clearance_levels"
Private Const FORMS_CLEARANCE_LEVEL_COLUMN_NAME As String = "form_name"

Private Const CONTROLS_CLEARANCE_LEVEL_SOURCE As String = "controls_clearance_levels"
Private Const CONTROLS_CLEARANCE_LEVEL_COLUMN_NAME As String = "control_name"

'Private Const CONTROLS_FORM_NAMES_SOURCE As String = "controls_clearance_levels_distinct_form_names"

Public Function ExistsUser(ByVal username As String) As Boolean
  If getUserId(username) > 0 Then
    ExistsUser = True
  Else
    ExistsUser = False
  End If
End Function

Public Function getClearanceLevelByLabel( _
  ByVal Label As String _
) As Single
  getClearanceLevelByLabel = Nz(ELookup( _
    "[" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "]", _
    "[" & CLEARANCE_LEVEL_SOURCE & "]", _
    "[" & USERS_CLEARANCE_LEVELS_LEVELS_NAME_COLUMN_NAME & "]" & "=" & "'" & Label & "'" _
    ), _
    0)
End Function

Public Function getClearanceLevelLabel( _
  ByVal level As Single _
) As String
  getClearanceLevelLabel = Nz(ELookup( _
    "[" & USERS_CLEARANCE_LEVELS_LEVELS_NAME_COLUMN_NAME & "]", _
    "[" & CLEARANCE_LEVEL_SOURCE & "]", _
    "[" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "]" & "=" & level _
    ), _
    "")
End Function

Public Function getUserId(ByVal username As String) As Long
  getUserId = Nz(ELookup("user_id", "[user]", "username = '" & username & "'"), 0)
End Function

Public Function getClearanceLevelByUser(ByVal username As String) As Single
  getClearanceLevelByUser = CSng( _
    Nz(ELookup( _
      "MAX([" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "])", _
      "[" & USERS_CLEARANCE_LEVELS_SOURCE & "]", _
      "[" & USERS_CLEARANCE_LEVELS_COND_COLUMN_NAME & "]" & "=" & "'" & username & "'" _
      ), _
      getLowestClearanceLevel()))
End Function

Public Function getLowestClearanceLevel() As Single
  getLowestClearanceLevel = CSng( _
    Nz(ELookup( _
      "MIN([" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "])", _
      "[" & CLEARANCE_LEVEL_SOURCE & "]"), _
      0))
End Function

Public Function getHighestClearanceLevel() As Single
  getHighestClearanceLevel = CSng( _
    Nz(ELookup( _
      "MAX([" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "])", _
      "[" & CLEARANCE_LEVEL_SOURCE & "]"), _
      0))
End Function

Public Function getFormNamesRestrictedForClearanceLevel(ByVal lvl As Single) As Variant
  getFormNamesRestrictedForClearanceLevel = Array()
  
  Dim col As Collection
  Set col = EMultiLookup( _
    "[" & FORMS_CLEARANCE_LEVEL_COLUMN_NAME & "]", _
    "[" & FORMS_CLEARANCE_LEVEL_SOURCE & "]")
      
  Dim result As New ArrayList
  result.Append col
  getFormNamesRestrictedForClearanceLevel = result.Values
End Function

Public Function getFromNamesOfRestricedContolsForClearanceLevel(ByVal lvl As Single) As Variant
  getFromNamesOfRestricedContolsForClearanceLevel = Array()
  
  Dim col As Collection
  Set col = EMultiLookup( _
    "DISTINCT " & _
    "[" & FORMS_CLEARANCE_LEVEL_COLUMN_NAME & "]", _
    "[" & CONTROLS_CLEARANCE_LEVEL_SOURCE & "]", _
    "[" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "]" & ">" & Util.Units.SingleToString(lvl))
  
  Dim result As New ArrayList
  result.Append col
  getFromNamesOfRestricedContolsForClearanceLevel = result.Values
End Function

Public Function getControlNamesRestrictedForClearanceLevel( _
  ByVal lvl As Single, _
  Optional ByVal form_name As String = vbNullString _
) As Variant
  getControlNamesRestrictedForClearanceLevel = Array()
    
  Dim whereCond As String
  whereCond = "[" & USERS_CLEARANCE_LEVELS_COLUMN_NAME & "]" & ">" & Util.Units.SingleToString(lvl)
  whereCond = Util.Strings.Bracketize(whereCond)
  
  If Len(form_name) > 0 Then
    Dim addCond As String
    addCond = "[" & FORMS_CLEARANCE_LEVEL_COLUMN_NAME & "]" & "=" & "'" & form_name & "'"
    addCond = Util.Strings.Bracketize(addCond)
    whereCond = Util.Strings.Bracketize(whereCond & " AND " & addCond)
  End If
  
  Dim col As Collection
  Set col = EMultiLookup( _
    "[" & CONTROLS_CLEARANCE_LEVEL_COLUMN_NAME & "]", _
    "[" & CONTROLS_CLEARANCE_LEVEL_SOURCE & "]", _
    whereCond)
  
  Dim result As New ArrayList
  result.Append col
  getControlNamesRestrictedForClearanceLevel = result.Values
End Function

