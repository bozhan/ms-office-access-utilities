VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DateCode
  dd = 21
  mm = 20
  yy = 19
End Enum

Public Function GetDateStamp( _
  Optional dateFormat As String = "yyyy-mm-dd" _
) As String
  GetDateStamp = Format(VBA.DateTime.Date$(), dateFormat)
End Function

Public Function GetTimeStamp( _
  Optional timeFormat As String = "hhmmss" _
) As String
    GetTimeStamp = Format(VBA.DateTime.Time$(), timeFormat)
End Function

Public Function GetDateTimeStamp( _
  Optional dateFormat As String = "yyyy-mm-dd", _
  Optional timeFormat As String = "hhmmss", _
  Optional ByVal seperator As String = "_" _
) As String
    GetDateTimeStamp = _
      GetDateStamp(dateFormat) & _
      seperator & _
      Format(VBA.DateTime.Time$(), timeFormat)
End Function




