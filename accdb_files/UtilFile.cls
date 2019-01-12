VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FileManager
' Author    : Ivanov, Bozhan
' Date      : 29.08.2014
' LastUpdate: 29.08.2014
' Purpose   : Provides a number of functions and precedures to help manage
'   file and folder paths and determine thir existance
' References:
'   FileSystemObject
'   https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/6tkce7xa%28v%3dvs.84%29
'---------------------------------------------------------------------------------------
Option Explicit

Const DEFAULT_ALLOWED_EXCEL_FILE_EXTENTIONS = "*.xls; *.xlsx; *.xlsm"
Const DEFAULT_ALLOWED_MS_PROJECT_FILE_EXTENTIONS = "*.mpp"
Const DEFAULT_ALLOWED_TEXT_FILE_EXTENTIONS = "*.txt"

'Required for function GetUNCPath
Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
  (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private fso As Scripting.FileSystemObject 'Object '

Private Property Get FileSys() As Scripting.FileSystemObject
On Error GoTo fso_Error

  If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
  Set FileSys = fso

Exit Property
fso_Error:
  Dim info As String
  If err.Number = 429 Then
    info = "Scripting.FileSystemObject could not be created in procedure FileSys of Class Module UtilFile"
  End If
  
  err.raise err.Number, err.source, err.Description & vbLf & info
End Property

Private Sub Class_Terminate()
  Set fso = Nothing
End Sub

Property Get DefaultExcelExt() As String
  DefaultExcelExt = DEFAULT_ALLOWED_EXCEL_FILE_EXTENTIONS
End Property

Property Get DefaultMSProjectExt() As String
  DefaultMSProjectExt = DEFAULT_ALLOWED_MS_PROJECT_FILE_EXTENTIONS
End Property

Property Get DefaultTextFileExt() As String
  DefaultTextFileExt = DEFAULT_ALLOWED_TEXT_FILE_EXTENTIONS
End Property

Property Get LastPath() As String
  If LAST_PATH <> vbNullString Then
    LastPath = GetFolderPath(LAST_PATH)
  End If
End Property

Property Let LastPath(ByVal Value As String)
  LAST_PATH = Value
End Property

Public Function CreateTextFile( _
  ByVal path As String, _
  Optional overwrite As Boolean = True, _
  Optional unicode As Boolean = False _
) As Object
  Set CreateTextFile = FileSys.CreateTextFile(path, overwrite, unicode)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFilePathViaDialogWindow
' Descr.    : Pop-up file dialog window returns the path of the selected file
'---------------------------------------------------------------------------------------
Public Function GetFilePathViaDialogWindow( _
  Optional ByVal InitialFilename As String = vbNullString, _
  Optional ByVal title As String = "Choose file...", _
  Optional ByVal buttonName As String = "Select...", _
  Optional ByVal InitialView As Long = Office.MsoFileDialogView.msoFileDialogViewDetails, _
  Optional ByVal FilterDesc As String = "All", _
  Optional ByVal FilterExt As String = "*.*", _
  Optional ByVal FilterPos As Long = 1 _
) As String
Dim fd As Office.FileDialog
 
  GetFilePathViaDialogWindow = vbNullString
  Set fd = Application.FileDialog(Office.MsoFileDialogType.msoFileDialogFilePicker)
  
  If InitialFilename <> vbNullString _
  And Not FileSys.FolderExists(InitialFilename) _
  And Not FileSys.FileExists(InitialFilename) Then
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Invalid initial path"
  End If
  
  With fd
    .AllowMultiSelect = False           ' Allow/Disable Multi file selection
    .InitialFilename = InitialFilename  ' Specify starting path
    If InitialFilename = vbNullString Then
      .InitialFilename = Environ("DESKTOP")
    End If
    .title = title                      ' Define dialog box title text
    .buttonName = buttonName            ' Define selection button text
    .InitialView = InitialView          ' Define Initial file/folder view style
    .Filters.Clear
    .Filters.Add FilterDesc, FilterExt, FilterPos
    
    
    'Use the Show method to display the File Picker dialog box and return the user's action.
    If .Show = -1 Then 'If the user presses the button...
      GetFilePathViaDialogWindow = Me.GetUNCPath(CStr(.SelectedItems(1)))
      LastPath = Me.GetUNCPath(CStr(.SelectedItems(1)))
    End If
  End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMultipleFilePathsViaDialogWindow
' Descr.    : Pop-up file dialog window returns an array of selected fiel paths
'---------------------------------------------------------------------------------------
Public Function GetMultipleFilePathsViaDialogWindow( _
  Optional ByVal InitialFilename As String = vbNullString, _
  Optional ByVal title As String = "Choose file(s)...", _
  Optional ByVal buttonName As String = "Select...", _
  Optional ByVal InitialView As Long = Office.MsoFileDialogView.msoFileDialogViewDetails, _
  Optional ByVal FilterDesc As String = "All", _
  Optional ByVal FilterExt As String = "*.*", _
  Optional ByVal FilterPos As Long = 1 _
) As Variant
Dim fd As Office.FileDialog
Dim result() As Variant
Dim item As Variant

  Set fd = Application.FileDialog(Office.MsoFileDialogType.msoFileDialogFilePicker)
  
  If InitialFilename <> vbNullString _
  And Not FileSys.FolderExists(InitialFilename) _
  And Not FileSys.FileExists(InitialFilename) Then
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Invalid initial path"
  End If
  
  With fd
    .AllowMultiSelect = True            ' Allow/Disable Multi file selection
    .InitialFilename = InitialFilename  ' Specify starting path
    If InitialFilename = vbNullString Then
      .InitialFilename = Environ("DESKTOP")
    End If
    .title = title                      ' Define dialog box title text
    .buttonName = buttonName            ' Define selection button text
    .InitialView = InitialView          ' Define Initial file/folder view style
    .Filters.Add FilterDesc, FilterExt, FilterPos
    
    'Use the Show method to display the File Picker dialog box and return the user's action.
    If .Show = -1 Then 'If the user presses the button...
      For Each item In .SelectedItems
        Util.Arrays.IncreaseUpperBound result
        result(UBound(result)) = GetUNCPath(CStr(item))
        LastPath = GetUNCPath(CStr(item))
      Next item
    End If
  End With
  
  GetMultipleFilePathsViaDialogWindow = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFolderPathViaDialogWindow
' Descr.    : Pop-up file dialog window returns the path of a selected folder
'---------------------------------------------------------------------------------------
Public Function GetFolderPathViaDialogWindow( _
  Optional ByVal InitialFilename As String = vbNullString, _
  Optional ByVal title As String = "Choose folder...", _
  Optional ByVal buttonName As String = "Select...", _
  Optional ByVal InitialView As Long = Office.MsoFileDialogView.msoFileDialogViewDetails _
) As String
Dim fd As Office.FileDialog

  GetFolderPathViaDialogWindow = vbNullString
  Set fd = Application.FileDialog(Office.MsoFileDialogType.msoFileDialogFolderPicker)
  
  If InitialFilename <> vbNullString _
  And Not FileSys.FolderExists(InitialFilename) _
  And Not FileSys.FileExists(InitialFilename) Then
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Invalid initial path"
  End If
  
  With fd
    .AllowMultiSelect = False           ' Allow/Disable Multi file selection
    .InitialFilename = InitialFilename  ' Specify starting path
    If InitialFilename = vbNullString Then
      .InitialFilename = Environ("DESKTOP")
    End If
    .title = title                      ' Define dialog box title text
    .buttonName = buttonName            ' Define selection button text
    .InitialView = InitialView          ' Define Initial file/folder view style
    
    'Use the Show method to display the folder Picker dialog box and return the user's action.
    If .Show = -1 Then 'If the user presses the button...
      Dim path As String
      path = GetUNCPath(CStr(.SelectedItems(1)))
      GetFolderPathViaDialogWindow = path
      LastPath = path
    End If
  End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveAsFilePathViaDialogWindow
' Descr.    : Pop-up file dialog window returns the save as path
'---------------------------------------------------------------------------------------
Public Function SaveAsFilePathViaDialogWindow( _
  Optional ByVal InitialFilename As String = vbNullString, _
  Optional ByVal title As String = "Choose save path...", _
  Optional ByVal buttonName As String = "Save", _
  Optional ByVal InitialView As Long = Office.MsoFileDialogView.msoFileDialogViewDetails _
) As String
Dim fd As Office.FileDialog

  SaveAsFilePathViaDialogWindow = vbNullString
  Set fd = Application.FileDialog(Office.MsoFileDialogType.msoFileDialogSaveAs)
   
  With fd
    .AllowMultiSelect = False           ' Allow/Disable Multi file selection
    .InitialFilename = InitialFilename  ' Specify starting path
    If InitialFilename = vbNullString Then
      .InitialFilename = Environ("DESKTOP")
    End If
    .title = title                      ' Define dialog box title text
    .buttonName = buttonName            ' Define selection button text
    .InitialView = InitialView          ' Define Initial file/folder view style
    
    'Use the Show method to display the File Picker dialog box and return the user's action.
    If .Show = -1 Then 'If the user presses the button...
      SaveAsFilePathViaDialogWindow = GetUNCPath(CStr(.SelectedItems(1)))
      LastPath = GetUNCPath(CStr(.SelectedItems(1)))
    End If
  End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : CopyFile
' Descr.    : Copy a file on disk from one location to another.
'---------------------------------------------------------------------------------------
Public Function CopyFile( _
  SourceFilePath As String, _
  DestFilePath As String, _
  Optional overwrite As Boolean = True _
) As Boolean
  
  CopyFile = False
  
  If FileSys.FileExists(SourceFilePath) _
  And Me.GetFolderPath(DestFilePath) <> vbNullString Then
    FileSys.CopyFile SourceFilePath, DestFilePath, overwrite
    
    If FileSys.FileExists(DestFilePath) Then
      CopyFile = True
    End If
  End If
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : CopyFileXCopy
' Descr.    : Copy a file on disk from one location to another using xcopy shell command
'---------------------------------------------------------------------------------------
Public Function CopyFileXCopy( _
  SourceFilePath As String, _
  DestFilePath As String, _
  Optional overwrite As Boolean = True _
) As Boolean
Dim shell As Object 'WshShell
Dim quoteMark As String

  CopyFileXCopy = False
  Set shell = CreateObject("WScript.Shell")
  quoteMark = Chr(38)
  
  If shell Is Nothing Then Exit Function

  If FileSys.FileExists(SourceFilePath) _
  And Me.GetFolderPath(DestFilePath) <> vbNullString Then
    Dim runStr As String
    runStr = "xcopy.exe " & _
      quoteMark & SourceFilePath & quoteMark & " " & _
      quoteMark & DestFilePath & "*" & quoteMark & " /R /H"
      
      If overwrite Then runStr = runStr & " /Y"
      
    shell.Run runStr, 0, True
    
    If FileSys.FileExists(DestFilePath) Then
      CopyFileXCopy = True
    End If
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : BackupFile
' Descr.    : makes a backup of a given file path and returns the new file path
'---------------------------------------------------------------------------------------
Public Function BackupFile( _
  ByVal path As String, _
  Optional ByVal dstFolder As String = vbNullString _
) As String
Dim bkpPath As String
Dim dstFolderPath As String

  BackupFile = vbNullString
  
  If FileSys.FileExists(path) Then
    If Me.GetFolderPath(dstFolder) <> vbNullString Then
      dstFolderPath = Me.GetFolderPath(dstFolder)
    Else
      dstFolderPath = Me.GetFolderPath(path)
    End If
    
    bkpPath = dstFolderPath & "\" & _
      "bkp_" & Format(VBA.DateTime.Date$(), "yyyy-mm-dd") & "_" & _
      Format(Time(), "hhmmss") & "_" & _
      Me.GetFileName(path)
      
    Me.CopyFile path, bkpPath
    
    If FileSys.FileExists(bkpPath) Then
      BackupFile = bkpPath
    End If
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFolderPath
' Descr.    : returns the deepest directory path without \ at the end
'   of the supplied path parameter without path parameter
'   Returns vbnullstring if path is not valid
'---------------------------------------------------------------------------------------
Public Function GetFolderPath( _
  ByVal path As String _
) As String
      
  'strip path down to the last existing folder
  Dim tmp As String
  tmp = path
  While Not (FileSys.FolderExists(tmp)) And Len(tmp) > 0
    tmp = Util.Strings.GetLeftOfIndex(tmp, Util.Strings.GetLastIndexOf(tmp, "\"))
  Wend
    
  If Len(tmp) > 0 Then
    If Right(tmp, 1) = ":" Then tmp = tmp & "\" 'drilled down to drive letter need to add "\"
    GetFolderPath = Me.GetUNCPath(GetFolder(tmp).path)
  Else
    GetFolderPath = vbNullString
  End If
End Function

Public Function GetFolder(ByVal path As String) As folder
  Set GetFolder = Nothing
  
  If Me.FolderExists(path) Then
    Set GetFolder = FileSys.GetFolder(path)
  Else
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Invalid folder path" & vbLf & _
        """" & path & """"
  End If
End Function

Public Function GetFile(ByVal path As String) As File
  Set GetFile = Nothing
  
  If Me.FileExists(path) Then
    Set GetFile = FileSys.GetFile(path)
  Else
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="invalid file path"
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileExtension
' Descr.    : Returns the file extension of a provided file path or file name
'   Returns vbnullstring if path is a directory
'   No check is done if file exitst on purpose, so procedure can be used to extract ext
'   from provided string paths
'---------------------------------------------------------------------------------------
Public Function GetFileExtension(path As String) As String
  If Not FileSys.FolderExists(path) Then
    GetFileExtension = FileSys.GetExtensionName(path)
  Else
    GetFileExtension = vbNullString
  End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
  FileExists = FileSys.FileExists(path)
End Function

Public Function FolderExists(ByVal path As String) As Boolean
  FolderExists = FileSys.FolderExists(path)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : GetFiles
' Description : returns a collection of Scripting.file objects in a folder
'   matching the ExtensionFilter, if no filter is provieded then all files are returned
'---------------------------------------------------------------------------------------
Public Function GetFiles( _
  ByVal FolderPath As String, _
  Optional ByVal Recurse As Boolean = False, _
  Optional ByVal ExtensionFilter As String = "" _
) As Collection

  Set GetFiles = New Collection
  If Not FileSys.FolderExists(FolderPath) Then Exit Function
  
  Dim al As New ArrayList
  Dim folder As Scripting.folder
  Dim f As Scripting.File
  Dim d As Scripting.folder
  
  Set folder = FileSys.GetFolder(FolderPath)
  
  For Each f In folder.Files
    If Len(ExtensionFilter) > 0 Then
      If InStr(1, ExtensionFilter, GetFileExtension(f.Name), vbTextCompare) _
      And Len(GetFileExtension(f.Name)) > 0 Then
        al.Append f
      End If
    Else
      al.Append f
    End If
  Next f
  
  If Recurse Then
    If folder.SubFolders.count > 0 Then
      For Each d In folder.SubFolders
        al.Append GetFiles(d.path, Recurse, ExtensionFilter)
      Next d
    End If
  End If
  
  Dim l As Long
  For l = al.lb To al.ub
    GetFiles.Add al(l)
  Next l
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileName
' Descr.    : Returns the file name with extension of a provided file path or file name
'   Returns vbnullstring if path is a directory
'---------------------------------------------------------------------------------------
Public Function GetFileName(ByRef path As String) As String
  If Not FileSys.FolderExists(path) Then
    GetFileName = FileSys.GetFileName(path)
  Else
    GetFileName = vbNullString
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetBaseName
' Descr.    : Returns the file name w/o extension of a provided file path or file name
'   Returns vbnullstring if path is a directory
'---------------------------------------------------------------------------------------
Public Function GetBaseName(path As String) As String
  If Not FileSys.FolderExists(path) Then
    Dim s As String
    s = FileSys.GetFileName(path)
    GetBaseName = Util.Strings.GetLeftOfIndex(s, Util.Strings.GetLastIndexOf(s, "."))
  Else
    GetBaseName = vbNullString
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetUNCPath
' Purpose   : Returns the Universal Naming Convention path, matching the assigned Letters
'---------------------------------------------------------------------------------------
Public Function GetUNCPath(ByVal localPath As String) As String
Dim uncPath As String, result As String, drive As String
Dim colonIndex As Long

  GetUNCPath = localPath
  colonIndex = Util.Strings.GetIndexOf(localPath, ":")
  
  If colonIndex = 0 Then Exit Function
  
  drive = Left$(localPath, colonIndex)
  uncPath = VBA.String(260, 0)
  
  If WNetGetConnection(drive, uncPath, Len(uncPath)) = 0 Then
    result = Left$(uncPath, InStr(uncPath, vbNullChar) - 1)
    If Len(result) > 0 Then GetUNCPath = result & mid$(localPath, colonIndex + 1)
  End If
End Function

Public Function GetFileNamesPathsFromFolder( _
  ByVal dirPath As String, _
  ByVal fileNames As Variant _
) As Variant
Dim names As ArrayList
Dim l As Long

  GetFileNamesPathsFromFolder = Array()
  
  If Not Util.Arrays.IsAllocated(fileNames) Or Util.Arrays.IsEmpty(fileNames) Then Exit Function
  If Not Me.FolderExists(dirPath) Then Exit Function
  
  Set names = New ArrayList
  names.Append fileNames
  dirPath = Me.GetFolderPath(dirPath)
  
  For l = names.lb To names.ub
    If Me.FileExists(dirPath & "\" & names(l)) Then
      names(l) = dirPath & "\" & names(l)
    End If
  Next l
  
  GetFileNamesPathsFromFolder = names.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsWorkbookPath
' Descr.    : returns TRUE if the given path is a valid path of an excel workbook
'---------------------------------------------------------------------------------------
Public Function IsWorkbookPath(ByVal path As String) As Boolean
  IsWorkbookPath = IsFileOfType(path, Me.DefaultExcelExt)
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsMSProjectPath
' Descr.    : returns TRUE if the given path is a valid path of an MS Project file
'---------------------------------------------------------------------------------------
Public Function IsMSProjectPath(ByVal path As String) As Boolean
  IsMSProjectPath = IsFileOfType(path, Me.DefaultMSProjectExt)
End Function

Private Function IsFileOfType( _
  ByVal path As String, _
  ByVal fileType As String _
) As Boolean
  
  IsFileOfType = False
  
  Dim ext As String
  ext = Me.GetFileExtension(path)
  
  If Len(ext) = 0 Then Exit Function
  If InStr(1, fileType, ext, vbTextCompare) > 0 Then IsFileOfType = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsTempFilePath
' Descr.    : returns TRUE if the given path is a temporary file
'---------------------------------------------------------------------------------------
Public Function IsTempFilePath(ByVal path As String) As Boolean
  IsTempFilePath = False
  If InStr(1, path, "~", vbTextCompare) > 0 Then IsTempFilePath = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetWorkbookPathsFromFolder
' Descr.    : Returns an array containing all excel workbook paths in a directory [and sub dirs]
'   If no start folder path is provided the current sheet directory is searched for WBs
'---------------------------------------------------------------------------------------
Public Function GetWorkbookPathsFromFolder( _
  Optional ByVal FolderPath As String = vbNullString, _
  Optional ByVal searchSubFolders As Boolean = False _
) As Variant
Dim wbPaths As ArrayList
Dim startFolder As Object 'Scripting.Folder
Dim subFolder As Object 'Scripting.Folder
Dim File As Object 'Scripting.File
On Error GoTo GetWorkbookPathsFromFolder_Error

  If FolderPath = vbNullString Then
    FolderPath = Me.GetFolderPath(ThisWorkbook.FullName)
  ElseIf Not Me.FolderExists(FolderPath) Then
    err.raise Number:=ErrCode.PathNotFound, _
      Description:="Path not found/Folder does not exist:" & FolderPath
  End If
  
  Set wbPaths = New ArrayList
  Set startFolder = Me.GetFolder(FolderPath)
  
  ' Cycle through all files in initial folder
  For Each File In startFolder.Files
    If Not Me.IsTempFilePath(File.path) _
    And Me.IsWorkbookPath(File.path) Then
      wbPaths.Append File.path
    End If
  Next File
  
  ' Search sub folders
  If searchSubFolders Then
    For Each subFolder In startFolder.SubFolders
      wbPaths.Append Me.GetWorkbookPathsFromFolder(subFolder.path, True)
    Next subFolder
  End If
 
  GetWorkbookPathsFromFolder = wbPaths.Values
 
GetWorkbookPathsFromFolder_Exit:
  On Error Resume Next
  Set startFolder = Nothing
  Set wbPaths = Nothing
Exit Function

GetWorkbookPathsFromFolder_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetWorkbookPathsFromFolder of module UtilExcel" & vbLf & _
    InfoErrMsg()
  Resume GetWorkbookPathsFromFolder_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileContent
' Purpose   : Returns content of a file as a string
'---------------------------------------------------------------------------------------
Public Function GetFileContent( _
  ByVal FilePath As String _
) As String

  GetFileContent = vbNullString
  
  If Not FileExists(FilePath) Then
    err.raise Number:=ErrCode.FileNotFound, _
      Description:="No file with the path: """ & FilePath & """ was found!"
  End If
  
  Dim fileNumber As Long, length As Long, res As String
  fileNumber = VBA.FreeFile()
  length = FileLen(FilePath)
  res = VBA.Space(length)
  
  Open FilePath For Binary Access Read As #fileNumber
  Get #fileNumber, , res
  Close fileNumber
  
  GetFileContent = res
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileBytes
' Purpose   : Returns content of a file as a byte array
'---------------------------------------------------------------------------------------
Public Function GetFileBytes( _
  ByVal FilePath As String _
) As Byte()
  Dim res() As Byte
  GetFileBytes = res()
  
  If Not FileExists(FilePath) Then
    err.raise Number:=ErrCode.FileNotFound, _
      Description:="No file with the path: """ & FilePath & """ was found!"
  End If
  
  Dim fileNumber As Integer, length As Long
  fileNumber = VBA.FreeFile()
  length = FileLen(FilePath)
  
  Open FilePath For Binary Access Read As #fileNumber
  
  'check if file is empty -> returns default empty byte array
  If LOF(fileNumber) = 0 Then Exit Function
  
  ReDim res(0 To LOF(fileNumber) - 1&) As Byte
  Get #fileNumber, , res
  Close fileNumber
  
  GetFileBytes = res
  Erase res
End Function


Public Function CreateFolder(ByVal path As String) As folder
  Dim dirPath As String
  Set CreateFolder = Nothing
  dirPath = GetFolderPath(path)
  
  'if path ends with '\' the diff will be 1 else <1
  If Len(path) - Len(dirPath) <= 1 Then 'path already exists
    Exit Function
  End If
  
  If Len(dirPath) = 0 Then
    err.raise Number:=ErrCode.PathNotFound, Description:="Invalid Dir path/Access denied"
  End If
  
  Dim nonExistantPaths As Variant
  nonExistantPaths = Split(Right(path, Len(path) - Len(dirPath)), "\")
  
  Dim i As Integer
  Dim f As folder
  For i = LBound(nonExistantPaths) To UBound(nonExistantPaths)
    If Len(nonExistantPaths(i)) > 0 Then
      Set f = FileSys.CreateFolder(dirPath & "\" & nonExistantPaths(i))
      
      If Not f Is Nothing Then
        dirPath = dirPath & "\" & nonExistantPaths(i)
      Else
        err.raise Number:=ErrCode.PathNotFound, _
          Description:="The path """ & dirPath & "\" & nonExistantPaths(i) & """ is invalid"
      End If
    End If
  Next i
  
  If StrComp(GetFolderPath(path), GetFolderPath(dirPath), vbTextCompare) = 0 Then
    Set CreateFolder = f
  End If
  
End Function

Public Function HasWritePermissions(ByVal path As String) As Boolean
Dim stream As Object 'Scripting.TextStream
Dim testPath As String

  HasWritePermissions = False
  Set fso = CreateObject("Scripting.FileSystemObject")
  testPath = Me.GetFolderPath(path) & "\" & Format(Now(), "yyyy-mm-dd_hh-nn-ss") & ".txt"

  On Error Resume Next
  Set stream = FileSys.CreateTextFile(testPath, True)
  
  If err.Number = 70 Then
   HasWritePermissions = False
   Exit Function
  End If
  
  On Error GoTo 0
  If Not stream Is Nothing Then
    stream.Close
    FileSys.DeleteFile testPath, True
    HasWritePermissions = True
  End If
End Function


