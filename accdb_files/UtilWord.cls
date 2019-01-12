VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : UtilWord
' Author    : Ivanov, Bozhan
' Purpose   : Provides ms word utilities to extend functionality and management and
'   generation of MS Word documents
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : GetHiddenColumnsIndeces
' Descr.    : Uses a saved word document tempalte as an OLE obejct in a worksheet
'   to generate and fill in a word document
'---------------------------------------------------------------------------------------
Public Function CreateWordDocument(ByVal strID As String, ByVal strRev As String, ByVal strTempPath As String) As String

Dim oleObject As Object
Dim wordDocument As Object

    Set oleObject = Sheet1.OLEObjects("Object 1")
    oleObject.Verb Verb:=xlPrimary
    ActiveSheet.Range("A1").Select
    
    Set wordDocument = oleObject.Object
    
    strTempPath = strTempPath & "\" & "Spec-" & strID & "_Rev-" & strRev & ".doc"
   
    wordDocument.SaveAs fileName:=strTempPath, FileFormat:=wdFormatDocument
  
    Set wordDocument = Nothing
    Set oleObject = Nothing

    CreateWordDocument = strTempPath
    
    Sheet2.Activate
    
End Function
