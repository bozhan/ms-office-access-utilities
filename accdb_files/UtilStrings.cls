VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : modString
' Author    : Ivanov, Bozhan
' Descr.    : All string, char and text manipulation methods are defined here
'---------------------------------------------------------------------------------------
'
Option Explicit

Private Const ASCII_UPPER_A As Long = 65
Private Const ASCII_UPPER_Z As Long = 90
Private Const ASCII_LOWER_a As Long = 97
Private Const ASCII_LOWER_z As Long = 122
Private Const ASCII_NUM_0 As Long = 48
Private Const ASCII_NUM_9 As Long = 57
Private Const ASCII_SPECIAL_START As Long = 33
Private Const ASCII_SPECIAL_END As Long = 47

Public Enum CharacterType
  Numeric = 0
  Letter = 1
  Alphanumeric = 2
  NonAlphanumeric = 3
End Enum

Public Enum LetterCase
  upper = 0
  lower = 1
End Enum

'---------------------------------------------------------------------------------------
' Procedure   : GetBytes
' Description : returns input string into a byte array
'   Set a reference to mscorlib 4.0 64-bit
'---------------------------------------------------------------------------------------
Public Function GetBytes( _
  ByVal text As String _
) As Byte()
  Dim enc As Object, bytes() As Byte
  Set enc = CreateObject("System.Text.UTF8Encoding")
  bytes = enc.Getbytes_4(text)
  GetBytes = bytes
End Function

'Returns the number of subsring occurances in a provided string
Public Function getNumberOfOccurances( _
  ByVal text, _
  ByVal searchString, _
  Optional CompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
) As Long
  
  getNumberOfOccurances = 0
  Dim idx As Long
  idx = InStr(1, text, searchString, CompareMethod)
  Do While idx > 0
    getNumberOfOccurances = getNumberOfOccurances + 1
    idx = InStr(idx + 1, text, searchString, CompareMethod)
  Loop
End Function

Public Function GetDateFromString(ByVal Value As String) As String
  GetDateFromString = VBA.trim(Util.RegEx.Replace(Util.RegExPatterns.NotDatePattern, Value, vbNullString))
End Function

Public Function RemoveDateFromString(ByVal Value As String) As String
  RemoveDateFromString = Util.RegEx.Replace(Util.RegExPatterns.DatePattern, Value, vbNullString)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIndexOf
' Descr.    : Returns a zero-based index of the first occurrence of the specified
'             string in a provided substring
'---------------------------------------------------------------------------------------
'
Public Function GetIndexOf( _
  ByVal stringExpression As String, _
  ByVal soughtString As String, _
  Optional CompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
) As Long
    GetIndexOf = InStr(1, stringExpression, soughtString, CompareMethod)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLastIndexOf
' Descr.    : Returns a zero-based index of the last occurrence of the specified
'             string in a provided substring
'---------------------------------------------------------------------------------------
'
Public Function GetLastIndexOf( _
  ByVal Value As String, _
  ByVal searchString As String, _
  Optional CompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
) As Long
Dim c As Long
Dim searchResult As Long
  
  For c = 1 To Len(Value)
    searchResult = InStr(c, Value, searchString, CompareMethod)
    
    If searchResult > 0 Then ' string found in substring
      c = searchResult + Len(searchString) - 1
    Else
      Exit For
    End If
  Next c
  
  If c - Len(searchString) < 0 Then
    GetLastIndexOf = 0
  Else
    GetLastIndexOf = c - Len(searchString)
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFirstIndexOfCharType
' Descr.    : Returns a zero-based index of the first occurrence of the specified
'             type of character (enum CharacterType)
'---------------------------------------------------------------------------------------
'
Public Function GetFirstIndexOfCharType(ByVal Value As String, ByVal charType As CharacterType) As Long
Dim c As String
Dim i As Long

  GetFirstIndexOfCharType = 0
  
  If Value = vbNullString Then Exit Function
  
  For i = 1 To Len(Value)
    c = mid$(Value, i, 1)
    
    Select Case charType
      Case CharacterType.Numeric
        If IsNumeric(c) Then GetFirstIndexOfCharType = i: Exit Function
      Case CharacterType.Letter
        If IsLetter(c) Then GetFirstIndexOfCharType = i: Exit Function
      Case CharacterType.Alphanumeric
        If IsLetter(c) Or IsNumeric(c) Then GetFirstIndexOfCharType = i: Exit Function
      Case CharacterType.NonAlphanumeric
        If Not IsLetter(c) And Not IsNumeric(c) Then GetFirstIndexOfCharType = i: Exit Function
    End Select
  Next i
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLastIndexOfCharType
' Descr.    : Returns a zero-based index of the last occurrence of the specified
'             type of character (enum CharacterType)
'---------------------------------------------------------------------------------------
'TODO finish impl
Public Function GetLastIndexOfCharType(ByVal Value As String, ByVal charType As CharacterType) As Long
Dim c As String
Dim i As Long

  GetLastIndexOfCharType = 0
  
  If Value = vbNullString Then Exit Function
  
  For i = 1 To Len(Value)
    c = mid(Value, i, 1)
    
    Select Case charType
      Case CharacterType.Numeric
        If IsNumeric(c) Then
          
        End If
      Case CharacterType.Letter
        If IsLetter(c) Then
          
        End If
      Case CharacterType.Alphanumeric
        If IsLetter(c) Or IsNumeric(c) Then
          
        End If
      Case CharacterType.NonAlphanumeric
        If Not IsLetter(c) And Not IsNumeric(c) Then
          
        End If
    End Select
  Next i
'
'  For i = 1 To Len(value)
'    searchResult = InStr(i, value, searchString, compareMethod)
'
'    If searchResult > 0 Then ' string found in substring
'      i = searchResult + Len(searchString) - 1
'    Else
'      Exit For
'    End If
'  Next c
'
'  If i - Len(searchString) < 0 Then
'    GetLastIndexOfCharType = 0
'  Else
'    GetLastIndexOfCharType = i - Len(searchString)
'  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeftOfIndex
' Descr.    : returns a sub-string left of the specified string index
'---------------------------------------------------------------------------------------
'
Public Function GetLeftOfIndex(ByVal Value As String, ByVal Index As Long) As String
  GetLeftOfIndex = vbNullString
  If Index > 0 Then
    GetLeftOfIndex = Left$(Value, Index - 1)
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRightOfIndex
' Descr.    : returns a sub-string right of the specified string index
'---------------------------------------------------------------------------------------
'
Public Function GetRightOfIndex(ByVal Value As String, ByVal Index As Long) As String
  GetRightOfIndex = vbNullString
  If Index > 0 Then
    GetRightOfIndex = Right$(Value, Len(Value) - Index)
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : isLetter
' Descr.    : boolean function will return true if the string is only letters
'---------------------------------------------------------------------------------------
'TODO use regex to mach letters
Public Function IsLetter(ByVal Value As String) As Boolean
Dim lPos As Long

  For lPos = 1 To Len(Value)
    Select Case Asc(mid$(Value, lPos, 1))
      Case ASCII_UPPER_A To ASCII_UPPER_Z, ASCII_LOWER_a To ASCII_LOWER_z
        IsLetter = True
      Case Else
        IsLetter = False
        Exit For
    End Select
  Next
End Function

Public Function IsNumber(ByVal Value As String) As Boolean
Dim r As String
  r = Util.RegEx.Replace(Util.RegExPatterns.NumbersPattern, Value, "")
  If Len(r) = 0 Then
    IsNumber = True
  Else
    IsNumber = False
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : RemoveExcessSeperators
' Descr.    : Removes excess seprators from left and right side of the first string param
'   Also a double seperator string is replaced by a zero string
'---------------------------------------------------------------------------------------
Public Function RemoveExcessSeperators( _
  ByVal text As Variant, _
  ByVal seperator As String, _
  Optional ByVal duplicates As Boolean = True, _
  Optional ByVal leftSide As Boolean = True, _
  Optional ByVal rightSide As Boolean = True, _
  Optional ByVal trimText As Boolean = False _
) As String
Dim s As String
    
On Error GoTo RemoveExcessSeperators_Error

  If trimText Then
    s = trim(text)
  Else
    s = text
  End If
  
  If s = vbNullString Then GoTo RemoveExcessSeperators_Exit
  
  If duplicates Then
    While InStr(1, s, seperator & seperator, vbTextCompare) > 0
      s = Replace(s, seperator & seperator, "", 1, 1, vbTextCompare)
    Wend
  End If
  
  If leftSide Then
    If StrComp(Left$(s, Len(seperator)), seperator) = 0 Then
      s = Right$(s, Len(s) - Len(seperator))
    End If
  End If
  
  If rightSide Then
    If StrComp(Right$(s, Len(seperator)), seperator) = 0 Then
      s = Left$(s, Len(s) - Len(seperator))
    End If
  End If

RemoveExcessSeperators_Exit:
On Error Resume Next
  RemoveExcessSeperators = s
Exit Function

RemoveExcessSeperators_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ")" & _
    "in procedure RemoveExcessSeperators of module UtilStrings" & vbLf & _
    InfoErrMsg(), _
    vbExclamation, _
    "Error occurred..."
Resume RemoveExcessSeperators_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : StripChar
' Descr.    : Removes string from another string if found
'---------------------------------------------------------------------------------------
Public Function StripChar(ByVal text As String, ByVal char As String) As String
  StripChar = Replace(text, char, "")
End Function

'---------------------------------------------------------------------------------------
' Procedure : StripChars
' Descr.    : Removes any list of strings from another string if found
'---------------------------------------------------------------------------------------
Public Function StripChars(ByVal text As String, ParamArray chars()) As String
  StripChars = text
  If LBound(chars) < 0 Then Exit Function
  
  Dim result As String, i As Long
  
  result = text
  For i = LBound(chars) To UBound(chars)
    result = StripChar(result, chars(i))
  Next i
  StripChars = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIndexOfOccurance
' Descr.    : Returns the index of the searched string.
'   The occurance of the search string may be provided in order to find the index of
'   a particular occurance of that string. If the occurance parameter is larger than the
'   occurances of the search string then 0 will be returned
'---------------------------------------------------------------------------------------
Public Function GetIndexOfOccurance( _
  ByVal Value As String, _
  ByVal searchString As String, _
  Optional ByVal occurance As Long = 1, _
  Optional ByVal CompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare _
) As Long
Dim l As Long
Dim realIdx As Long
Dim tempIdx As Long

  GetIndexOfOccurance = 0
  realIdx = 0
  
  For l = 1 To occurance
    tempIdx = InStr(1, Value, searchString, CompareMethod)
    If tempIdx = 0 Then
      GetIndexOfOccurance = 0
      Exit Function
    End If
    Value = GetRightOfIndex(Value, tempIdx)
    realIdx = realIdx + tempIdx
  Next l
  
  GetIndexOfOccurance = realIdx
End Function

'TODO reserach locale number formats
'need to move function and need to make sure it does what it says
'Function used in producing a local independent date format
'to be used in Excel sheet cell number formatting -> custom

'    [$-409] - English, Untitled States
'    [$-1009] -English, Canada
'    [$-407] -German, Germany
'    [$-807] -German, Switzerland
'    [$-804] -Bengali, India
'    [$-804] -Chinese, China
'    [$-404] -Chinese, Taiwan

Public Function ConvertDateFormatBasedOnRegionalSettings( _
  ByVal dateFormat As String _
) As String
  dateFormat = Replace(dateFormat, "d", DateCode.dd, , , vbTextCompare)
  dateFormat = Replace(dateFormat, "m", DateCode.mm, , , vbTextCompare)
  dateFormat = Replace(dateFormat, "y", DateCode.yy, , , vbTextCompare)
  ConvertDateFormatBasedOnRegionalSettings = dateFormat
End Function

Public Function SplitOnDelimiter( _
  ByVal Value As String, _
  ByVal Delimiter As String _
) As Variant
  SplitOnDelimiter = Split(Value, Delimiter)
End Function

Public Function Bracketize( _
  ByVal text As String, _
  Optional ByVal exclusive As Boolean = True _
) As String
  If exclusive Then
    Bracketize = encloseWithCharsExclusive(text, "(", ")")
  Else
    Bracketize = encloseWithChars(text, "(", ")")
  End If
End Function

Public Function bracketizeSquare( _
  ByVal text As String, _
  Optional ByVal exclusive As Boolean = True _
) As String
  If exclusive Then
    bracketizeSquare = encloseWithCharsExclusive(text, "[", "]")
  Else
    bracketizeSquare = encloseWithChars(text, "[", "]")
  End If
End Function

Public Function bracketizeSquare_MatchInText( _
  ByVal text As String, _
  ByVal searchString As String _
) As String
  bracketizeSquare_MatchInText = encloseWithChars_MatchInText(text, searchString, "[", "]")
End Function

'---------------------------------------------------------------------------------------
' Procedure : encloseWithChars_MatchInText
' Purpose   : encloses all instances of the searchString parameter in exclusive manner
'---------------------------------------------------------------------------------------
Public Function encloseWithChars_MatchInText( _
  ByVal text As String, _
  ByVal searchString As String, _
  Optional openingChar As String = "(", _
  Optional closingChar As String = ")" _
) As String
  encloseWithChars_MatchInText = ""
  Dim idx As Long, res As String, l As String, r As String, m As String, tmp As String
  res = ""
  searchString = StripChars(searchString, openingChar, closingChar)
  tmp = text
  idx = InStr(1, tmp, searchString, vbTextCompare)
  
  Do While idx > 0
    l = Left(tmp, idx - 1)
    m = mid(tmp, idx, Len(searchString))
    res = res & l & encloseWithCharsExclusive(m, openingChar, closingChar)
    tmp = Right(tmp, Len(tmp) - (idx - 1 + Len(searchString)))
    idx = InStr(1, tmp, searchString, vbTextCompare)
  Loop
  
  'append last part of text after search string
  res = res & tmp
  
  'remove any duplicate square brackets occuring due to already bracketed strings in text
  res = Replace(res, openingChar & openingChar, openingChar)
  res = Replace(res, closingChar & closingChar, closingChar)
  encloseWithChars_MatchInText = res
End Function


Public Function encloseWithChars( _
  ByVal text As String, _
  ByVal leftChar As String, _
  Optional ByVal rightChar As String = vbNullString _
) As String
  encloseWithChars = text
  
  If rightChar = vbNullString Then
    encloseWithChars = leftChar & text & leftChar
  Else
    encloseWithChars = leftChar & text & rightChar
  End If
End Function

Public Function encloseWithCharsExclusive( _
  ByVal text As String, _
  ByVal leftChar As String, _
  Optional ByVal rightChar As String = vbNullString _
) As String
  encloseWithCharsExclusive = text
  
  If StrComp(Left(text, 1), leftChar, vbTextCompare) <> 0 Then
    encloseWithCharsExclusive = leftChar & encloseWithCharsExclusive
  End If
  
  If rightChar = vbNullString Then
    If StrComp(Right(text, 1), leftChar, vbTextCompare) <> 0 Then
      encloseWithCharsExclusive = encloseWithCharsExclusive & leftChar
    End If
  Else
    If StrComp(Right(text, 1), rightChar, vbTextCompare) <> 0 Then
      encloseWithCharsExclusive = encloseWithCharsExclusive & rightChar
    End If
  End If
End Function
 
Public Function getRandomLetter_RandomCase() As String
  If (Rnd(1) > 0.5) Then
    getRandomLetter_RandomCase = getRandomLetter_UpperCase()
  Else
    getRandomLetter_RandomCase = getRandomLetter_LowerCase()
  End If
End Function

'random char formula for range of numbers
'Int ((upperbound - lowerbound + 1) * Rnd + lowerbound)
Public Function getRandomLetter_UpperCase() As String
  getRandomLetter_UpperCase = Chr(Int((ASCII_UPPER_Z - ASCII_UPPER_A + 1) * Rnd + ASCII_UPPER_A))
End Function

Public Function getRandomLetter_LowerCase() As String
  getRandomLetter_LowerCase = Chr(Int((ASCII_LOWER_z - ASCII_LOWER_a + 1) * Rnd + ASCII_LOWER_a))
End Function

Public Function getRandomString_Alphabetic( _
  ByVal length As Integer _
) As String
Dim i As Integer
Dim result As String
  result = vbNullString
  For i = 1 To length
    result = result & getRandomLetter_RandomCase()
  Next i
  getRandomString_Alphabetic = result
End Function

Public Function getRandomString_AlphaNumeric( _
  ByVal length As Integer _
) As String
Dim i As Integer
Dim result As String
  result = vbNullString
  For i = 1 To length
    result = result & getRandomChar_AlphaNumeric()
  Next i
  getRandomString_AlphaNumeric = result
End Function

Public Function getRandomChar_AlphaNumeric() As String
  If (Rnd(1) > 0.5) Then
    getRandomChar_AlphaNumeric = getRandomLetter_RandomCase()
  Else
    getRandomChar_AlphaNumeric = getRandomNumber(0, 9, 1)
  End If
End Function

Public Function getRandomNumber( _
Optional lowerBound As Long = 0, _
Optional upperBound As Long = 9, _
Optional cycles As Long = 1 _
) As String
Dim i As Integer
Dim result As String
  result = vbNullString
  For i = 1 To cycles
    result = result & CStr(Int((upperBound - lowerBound + 1) * Rnd + lowerBound))
  Next i
  getRandomNumber = result
End Function


Public Function getSplitDSV( _
  ByVal content As String, _
  ByVal colDelimiter As String, _
  Optional ByVal rowDelimiter As String = vbCrLf _
) As Variant
Dim rows As Variant, r As Variant, cols As Variant, c As Variant
Dim result As Variant

  getSplitDSV = Array()
  If Len(content) <= 0 Then Exit Function
  
  rows = getSplitByRowsDSV(content, rowDelimiter)
  ReDim result(LBound(rows) To UBound(rows), 0 To 0)
  
  For r = LBound(rows) To UBound(rows)
    cols = Split(rows(r), colDelimiter)
    
    If UBound(result, 2) < UBound(cols) Then
      ReDim Preserve result(LBound(result, 1) To UBound(result, 1), LBound(result, 2) To UBound(cols))
    End If
    
    For c = LBound(cols) To UBound(cols)
      result(r, c) = cols(c)
    Next c
  Next r
  
  getSplitDSV = result
End Function

Public Function getSplitByRowsDSV( _
  ByVal content As String, _
  Optional ByVal Delimiter As String = vbCrLf _
) As Variant
Dim rows As Variant, r As Variant
Dim result As Variant

  getSplitByRowsDSV = Array()
  If Len(content) <= 0 Then Exit Function
  
  rows = Split(content, Delimiter)
  For Each r In rows
    If Len(r) > 0 Then
      Util.Arrays.IncreaseUpperBound result
      result(UBound(result)) = r
    End If
  Next r
  
  getSplitByRowsDSV = result
End Function

Public Function getSplitByColumnsDSVOnCondition( _
  ByVal content As String, _
  ByVal colDelimiter As String, _
  ByVal condition As Variant, _
  Optional ByVal rowDelimiter As String = vbCrLf _
) As Variant
Dim cols As Variant, c As Variant, rows As Variant, r As Long
Dim result As Variant
Dim matchPositions As Variant, colMatches As ArrayList
Dim i As Integer

  getSplitByColumnsDSVOnCondition = Array()
  If Len(content) <= 0 Then Exit Function

  result = Array()
  Set colMatches = New ArrayList
  
  matchPositions = getMatchPositionsInDSV(content, colDelimiter, condition, rowDelimiter)
  rows = getSplitByRowsDSV(content, rowDelimiter)
  
  ReDim result(UBound(rows), UBound(matchPositions))
  
  For r = LBound(rows) To UBound(rows)
    cols = Split(rows(r), colDelimiter)
    If Not Util.Arrays.IsEmpty(cols) And UBound(cols) >= matchPositions(UBound(matchPositions)) Then
      For i = LBound(matchPositions) To UBound(matchPositions)
        result(r, i) = cols(matchPositions(i))
      Next i
    End If
  Next r
  
  getSplitByColumnsDSVOnCondition = result
End Function

Public Function getMatchPositionsInDSV( _
  ByVal content As String, _
  ByVal Delimiter As String, _
  ByVal matchValues As Variant, _
  Optional ByVal rowDelimiter As String = vbCrLf _
) As Variant
Dim result As ArrayList
Dim m As Variant
Dim tmpPos As Long
  
  getMatchPositionsInDSV = Array()
  Set result = New ArrayList
  For m = LBound(matchValues) To UBound(matchValues)
    tmpPos = findMatchPositionInDSV(content, Delimiter, matchValues(m), rowDelimiter)
    If tmpPos > 0 Then
      result.Append tmpPos
    End If
  Next m
  getMatchPositionsInDSV = result.Values
End Function

Public Function findMatchPositionInDSV( _
  ByVal content As String, _
  ByVal Delimiter As String, _
  ByVal matchValue As Variant, _
  Optional ByVal rowDelimiter As String = vbCrLf _
) As Long
Dim splitByRows As Variant, splitByCols As Variant
Dim i As Long

  findMatchPositionInDSV = 0
  
  splitByRows = getSplitByRowsDSV(content, rowDelimiter)
  If Util.Arrays.IsEmpty(splitByRows) Then Exit Function
  
  For i = LBound(splitByRows) To UBound(splitByRows)
    splitByCols = Split(splitByRows(i), Delimiter)
    If Util.Arrays.Contains(splitByCols, matchValue, vbTextCompare) Then
      findMatchPositionInDSV = Util.Arrays.GetIndexOf(splitByCols, matchValue, vbTextCompare)
      Exit Function
    End If
  Next i
End Function

Function UTF8_Decode(ByVal sStr As String)
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
            iChar2 = Asc(mid(sStr, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(mid(sStr, l + 1, 1))
            iChar3 = Asc(mid(sStr, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function

Public Function getSpecialCharacters() As Variant
Dim chars As ArrayList
  Set chars = New ArrayList
  chars.Append ""
End Function

Public Function concat(ByVal stringsArray As Variant) As String
  concat = ""
  If LBound(stringsArray) < 0 Then Exit Function
  Dim i As Long
  For i = LBound(stringsArray) To UBound(stringsArray)
    concat = concat & stringsArray(i)
  Next i
End Function

'---------------------------------------------------------------------------------------
' Procedure : getStringWidthFromFontSize
' Author    : Ivanov, Bozhan
' Purpose   : returns a relative string width in Twips Per Inch
'---------------------------------------------------------------------------------------
Public Function getStringWidthFromFontSize( _
  ByVal fontSize As Integer, _
  ByVal stringLength As Long, _
  Optional ByVal adjDelta As Single = 0.75 _
) As Variant
  getStringWidthFromFontSize = _
    (((fontSize * adjDelta) * stringLength) / Util.Units.PixelsPerInch) * Util.Units.TwipsPerInch
End Function

Public Function SplitTrim( _
  ByVal Expression As String, _
  Optional ByVal Delimiter As Variant, _
  Optional ByVal limit As Long = -1, _
  Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
) As Variant
Dim result As Variant
Dim l As Long
  result = Split(Expression, Delimiter, limit, Compare)
  For l = LBound(result) To UBound(result)
    result(l) = trim$(result(l))
  Next l
  SplitTrim = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : CompareValues
' Author    : Ivanov, Bozhan
' Purpose   : returns the values 0,-1,1 according to output of the strcomp() function
'   Compares two elements V1 and V2
'   if both numeric or string with numeric,
'   they are converted to double and compared
'   If either one is String and does not contain numeric,
'   they are converted to string and compared with strcomp()
'   Case-sensitivity is controlled via the optional parameter compareMode
'---------------------------------------------------------------------------------------
Public Function CompareValues( _
  v1 As Variant, _
  v2 As Variant, _
  Optional CompareMethod As VbCompareMethod = VbCompareMethod.vbTextCompare _
) As Variant
  
  If Not IsSimpleDataType(v1) Or Not IsSimpleDataType(v2) Then
    CompareValues = Null
    Exit Function
  End If
  
  If IsNumeric(v1) And IsNumeric(v2) Then
    If CDbl(v1) = CDbl(v2) Then CompareValues = 0
    If CDbl(v1) < CDbl(v2) Then CompareValues = -1
    If CDbl(v1) > CDbl(v2) Then CompareValues = 1
    Exit Function
  End If
  
  If Not IsNumeric(v1) Or Not IsNumeric(v2) Then
    CompareValues = StrComp(CStr(v1), CStr(v2), CompareMethod)
  End If
End Function

Public Function IsSimpleDataType(v As Variant) As Boolean
On Error Resume Next

  If IsArray(v) Or IsObject(v) Then
    IsSimpleDataType = False
    Exit Function
  End If
  
  Select Case VarType(v)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType
      IsSimpleDataType = False
    Case Else
      IsSimpleDataType = True
  End Select
End Function

