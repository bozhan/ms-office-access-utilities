VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : UtilFont
' Author      : Ivanov, Bozhan
' Description :
' Source      : https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/Font_Utilities
'---------------------------------------------------------------------------------------

Function GetTextPoints(sIn As String, sFontName As String, _
    nFontSize As Single, bFontBold As Boolean, _
    bFontItalic As Boolean) As Long
    'GetTextPoints returns points width of text.
    'When setting a control width, add two additional
    'space widths to these values to avoid end clipping.
    'Needs a user form called CountPoints. Form
    'is loaded and unloaded but never shown.
        
    'Monospace test: could be used here to identify monospaced fonts
    'If pair is same width then monospaced
    'IIIIIIIIII
    'HHHHHHHHHH
    
    'Kerning test pair used by printers: Wont work here since there is no kerning in userform controls.
    'If pair are different width then there is kerning.
    'AAAAATTTTT
    'ATATATATAT

    Dim oLbl As Control
    
    Load CountPoints
    Set oLbl = CountPoints.controls.Add("Forms.Label.1", "oLbl")

    'format the label with same fonts as sIn
    With oLbl
        .Width = 0
        .WordWrap = False
        .Visible = False
        .AutoSize = True
        .Caption = ""
        .Font.Size = nFontSize
        .Font.Name = sFontName
        .Font.bold = bFontBold
        .Font.Italic = bFontItalic
    End With

    'get points for sIn
    oLbl.Caption = sIn
    GetTextPoints = oLbl.Width

    Unload CountPoints

End Function

Sub ListAllExcelFonts()
    'Lists Excel fonts as monospaced or proportional
    'with a sample of text and its width in points
    'calls GetTextPoints to measure test strings
    'needs use of Sheet1 - clears all existing
    
    Dim FontList, sht As Worksheet, i As Long
    Dim sM1 As String, sM2 As String, sFN As String
    Dim sTest As String, nSize As Single
    Dim bBold As Boolean, bItalic As Boolean
    
    'monospaced test strings
    sM1 = "IIIIIIIIII"
    sM2 = "MMMMMMMMMM"
    
    'set a suitable test string here
    sTest = "The quick brown fox jumps over the lazy dog 1234567890"
    
    'set test parameters
    nSize = 10 'ten point for all tests
    bBold = False
    bItalic = False
    
    Set sht = ThisWorkbook.Worksheets("Sheet1")
    With sht
        .Activate
        .Range("A1:Z65536").ClearContents
        .Range("A1:Z65536").ClearFormats
    End With
    
    'get reference to the font list
    Set FontList = Application.CommandBars("Formatting").FindControl(ID:=1728)
    
    On Error Resume Next
    'work loop
    For i = 1 To FontList.ListCount
        
        sFN = FontList.List(i) 'font name
        
        'print general data to sheet
        With sht
            .Cells(i, 1) = sFN                                              'name
            .Cells(i, 3) = GetTextPoints(sTest, sFN, nSize, bBold, bItalic) 'test string pts width
        End With
        
        'set fonts for sample cell
        With sht.Cells(i, 4).Font
            .Name = sFN
            .Size = nSize
            .Italic = bItalic
            .bold = bBold
        End With
        
        'sample string to sheet
        sht.Cells(i, 4) = sTest
        
        'monospaced  test - true if both test strings equal in length
        If GetTextPoints(sM1, sFN, nSize, bBold, bItalic) = GetTextPoints(sM2, sFN, nSize, bBold, bItalic) Then
            'the test font is monospaced
            sht.Cells(i, 2) = "Monospaced"  'mono or prop
        Else
            sht.Cells(i, 2) = "Proportional"
        End If
    Next i
        
    With sht
        .Columns.AutoFit
        .Cells(1, 1).Select
    End With

End Sub

Private Sub testit()
    ' Find whether or not a font exists
    Dim sFontName As String
    
    sFontName = "Consolas"
    
    If FontExists(sFontName) Then
        MsgBox sFontName & " exists"
    Else
        MsgBox sFontName & " does not exist"
    End If

End Sub

Public Function FontExists(FontName As String) As Boolean
    ' Returns true in function name
    ' if parameter font name exists
    
    Dim oFont As New StdFont
    
    oFont.Name = FontName
    If StrComp(FontName, oFont.Name, vbTextCompare) = 0 Then
        FontExists = True
    End If
    
End Function

