VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module      : UtilHash
' Author      : Ivanov, Bozhan
' Sources     : https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
' Description : This class provides Hash functions with variable available algorithms.
'   generates the digests for the MD5, SHA1, SHA2-256, SHA2-384, and SHA2-512 hashes,
'   for strings, in either of the hex or base-64 output formats. These codings each make use of
'   MS Office's built-in functions, and provide consistent results. It has been noted that original
'   implementations elsewhere for the same digests can differ widely in their outputs.
'   Note the VBA references required for correct operation.
'   A reminder of these is given in some procedure headings.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure   : GetHash
' Description : Generic Hash function that takes in provider and any number of inputs
'   If input us of simple data type it is converted to string and hashed
'   if input is a byte array it will just be processed
'   if input does not match any conditions an empty string hash is returned
' References  :
'   mscorlib.dll - Microsoft.NET 4.0 64-bit,
'   msxml6.dll - Microsoft XML, v6.0
'---------------------------------------------------------------------------------------
Private Function GetHash( _
  vIn As Variant, _
  Provider As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  Dim bIn() As Byte

  If Util.Arrays.IsOfDataType(vIn, vbByte) Then
    'checks if array and if el. are Byte
    bIn = vIn
  ElseIf Util.Strings.IsSimpleDataType(vIn) Then
    If Util.File.FileExists(CStr(vIn)) Then
      bIn = Util.File.GetFileBytes(vIn)
    Else
      bIn = Util.Strings.GetBytes(CStr(vIn))
    End If
  End If

  If Util.Arrays.IsEmpty(bIn) Then
    err.raise Number:=ErrCode.InvalidProcedureCall, _
      Description:="Input cannot be empty!" & vbLf & _
      "in procedure GetHash of Class Module UtilSecurity" & vbLf & _
      InfoErrMsg()
  End If

  Dim hp As Object, bytes() As Byte
  Set hp = CreateObject(Provider)
  bytes = hp.ComputeHash_2((bIn))
  
  If ConvertToB64 Then
     GetHash = ConvToBase64String(bytes)
  Else
     GetHash = ConvToHexString(bytes)
  End If
 
  Set hp = Nothing
End Function

Public Function ConvToBase64String( _
  vIn() As Byte _
) As String
  Dim oD As Object
  Set oD = CreateObject("MSXML2.DOMDocument")
  oD.LoadXML "<root />"
  oD.DocumentElement.dataType = "bin.base64"
  oD.DocumentElement.nodeTypedValue = vIn
  ConvToBase64String = Replace(oD.DocumentElement.text, vbLf, "")
  Set oD = Nothing
End Function

Public Function ConvToHexString( _
  vIn() As Byte _
) As String
  Dim oD As Object
  Set oD = CreateObject("MSXML2.DOMDocument")
  oD.LoadXML "<root />"
  oD.DocumentElement.dataType = "bin.Hex"
  oD.DocumentElement.nodeTypedValue = vIn
  ConvToHexString = Replace(oD.DocumentElement.text, vbLf, "")
  Set oD = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure   : MD5
' Description : Returns MD5 hashed string from the input string
'   Test with empty string input:
'   Hex: d41d8cd98f00b204e9800998ecf8427e
'   Base-64: 1B2M2Y8AsgTpgAmY7PhCfg==
'---------------------------------------------------------------------------------------
Public Function MD5( _
  ByVal sIn As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  MD5 = GetHash( _
    sIn, _
    "System.Security.Cryptography.MD5CryptoServiceProvider", _
    ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : SHA1
' Description : Returns SHA1 hashed string from the input string
'    Test with empty string input:
'    40 Hex:   da39a3ee5e6...etc
'    28 Base-64:   2jmj7l5rSw0yVb...etc
'---------------------------------------------------------------------------------------
Public Function SHA1( _
  ByVal sIn As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  SHA1 = GetHash( _
    sIn, _
    "System.Security.Cryptography.SHA1Managed", _
    ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : SHA256
' Description : Returns SHA256 hashed string from the input string
'    Test with empty string input:
'    64 Hex:   e3b0c44298f...etc
'    44 Base-64:   47DEQpj8HBSa+/...etc
'---------------------------------------------------------------------------------------
Public Function SHA256( _
  ByVal sIn As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  SHA256 = GetHash( _
    sIn, _
    "System.Security.Cryptography.SHA256Managed", _
    ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : SHA384
' Description : Returns SHA384 hashed string from the input string
'   Test with empty string input:
'   96 Hex:   38b060a751ac...etc
'   64 Base-64:   OLBgp1GsljhM2T...etc
'---------------------------------------------------------------------------------------
Public Function SHA384( _
  ByVal sIn As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  SHA384 = GetHash( _
    sIn, _
    "System.Security.Cryptography.SHA384Managed", _
    ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : SHA512
' Description : Returns a SHA512 hashed string from the input string
'   Test with empty string input:
'   128 Hex:   cf83e1357eefb8bd...etc
'   88 Base-64:   z4PhNX7vuL3xVChQ...etc
'---------------------------------------------------------------------------------------
Public Function SHA512( _
  ByVal sIn As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  SHA512 = GetHash( _
    sIn, _
    "System.Security.Cryptography.SHA512Managed", _
    ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : FileToMD5
' Description : Returns a MD5 hashed string from file content
'---------------------------------------------------------------------------------------
Public Function FileToMD5( _
  FilePath As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  FileToMD5 = GetHash( _
  FilePath, _
  "System.Security.Cryptography.MD5CryptoServiceProvider", _
  ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : FileToSHA1
' Description : Returns a SHA1 hashed string from file content
'---------------------------------------------------------------------------------------
Public Function FileToSHA1( _
  FilePath As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  FileToSHA1 = GetHash( _
  FilePath, _
  "System.Security.Cryptography.SHA1CryptoServiceProvider", _
  ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : FileToSHA256
' Description : Returns a SHA256 hashed string from file content
'---------------------------------------------------------------------------------------
Public Function FileToSHA256( _
  FilePath As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  FileToSHA256 = GetHash( _
  FilePath, _
  "System.Security.Cryptography.SHA256Managed", _
  ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : FileToSHA384
' Description : Returns a SHA384 hashed string from file content
'---------------------------------------------------------------------------------------
Public Function FileToSHA384( _
  FilePath As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  FileToSHA384 = GetHash( _
  FilePath, _
  "System.Security.Cryptography.SHA384Managed", _
  ConvertToB64)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : FileToSHA512
' Description : Returns a SHA512 hashed string from file content
'---------------------------------------------------------------------------------------
Public Function FileToSHA512( _
  FilePath As String, _
  Optional ConvertToB64 As Boolean = False _
) As String
  FileToSHA512 = GetHash( _
  FilePath, _
  "System.Security.Cryptography.SHA512Managed", _
  ConvertToB64)
End Function


