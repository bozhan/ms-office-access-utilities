VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ref:
'http://www.techrepublic.com/blog/10-things/10-plus-of-my-favorite-windows-api-functions-to-use-in-office-applications/

#If VBA7 And Win64 Then
  Private Declare PtrSafe Function FindWindow _
    Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, lpdwprocessid As Long) As Long
  Private Declare PtrSafe Function GetWindow Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal wCmd As Long) As Long
  Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare PtrSafe Function GetParent Lib "user32" _
    (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function Get_User_Name Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare PtrSafe Function Get_ComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
  Private Declare PtrSafe Function GetActiveWindow Lib "user32" _
    () As Long
  Private Declare PtrSafe Function Get_TempPath Lib "kernel32" Alias "GetTempPathA" _
   (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
  Private Declare PtrSafe Function Get_TempFileName Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, ByVal lpPrefixString As String, _
    ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
  Private Declare PtrSafe Function DesktopWindow Lib "user32" _
    () As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal nCmdSHow As Long) As Long
  Private Declare PtrSafe Function BringWindowToTop Lib "user32" _
    (ByVal lngHWnd As LongPtr) As Long
  Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" _
    () As Long
  Private Declare PtrSafe Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)
  Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
  Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
    (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
  Private Declare PtrSafe Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
#Else
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwprocessid As Long) As Long
  Private Declare Function GetWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal wCmd As Long) As Long
  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare Function GetParent Lib "user32" _
    (ByVal hwnd As Long) As Long
  Private Declare Function Get_User_Name Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare Function Get_ComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
  Private Declare Function GetActiveWindow Lib "user32" _
    () As Long
  Private Declare Function Get_TempPath Lib "kernel32" Alias "GetTempPathA" _
   (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
  Private Declare Function Get_TempFileName Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, ByVal lpPrefixString As String, _
    ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
  Private Declare Function DesktopWindow Lib "user32" _
    () As Long
  Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdSHow As Long) As Long
  Private Declare Function BringWindowToTop Lib "user32" _
    (ByVal lngHWnd As Long) As Long
  Private Declare Function GetCurrentProcessId Lib "kernel32" _
    () As Long
  Private Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)
  Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
  Private Declare Function OpenProcess Lib "kernel32.dll" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
  Private Declare Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
  Private Declare Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
#End If

Private Const GW_HWNDNEXT = 2
Private Const SYNCHRONIZE = &H100000
Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const STATUS_WAIT_0 As Long = &H0
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_INFINITE = -1&

Public Enum ShellAndWaitResult
    Success = 0
    Failure = 1
    TimeOut = 2
    InvalidParameter = 3
    SysWaitAbandoned = 4
    UserWaitAbandoned = 5
    UserBreak = 6
End Enum

Public Enum ActionOnBreak
    IgnoreBreak = 0
    AbandonWait = 1
    PromptUser = 2
End Enum

Public Enum WindowShowState
  SW_FORCEMINIMIZE = 11 'Minimizes a window.
  SW_HIDE = 0 'Hides the window and activates another window.
  SW_MAXIMIZE = 3 'Maximizes a window.
  SW_MINIMIZE = 6 'Minimizes the specified window and activates the next top-level window.
  SW_RESTORE = 9 'Activates and displays the window.
  SW_SHOW = 5 'Activates the window.
  SW_SHOWMAXIMIZED = 3 'Activates the window and displays it as a maximized window.
  SW_SHOWMINIMIZED = 2 'Activates the window and displays it as a minimized window.
  SW_SHOWMINNOACTIVE = 7 'Displays the window as a minimized window (without activating the window).
  SW_SHOWNA = 8 'Displays the window in its current size and position (without activating the window).
  SW_SHOWNOACTIVATE = 4 'Displays a window in its most recent size and position (without activating the window).
  SW_SHOWNORMAL = 1 'Activates and displays a window.
End Enum

Public Property Get FrameClassNames() As Object
Dim fc As Object
  With fc
    .Access = "OMain"
    .Excel = "XLMAIN"
    .FrontPage = "FrontPageExplorerWindow40"
    .Outlook = "rctrl_renwnd32"
    .PowerPoint_2000 = "PP9FrameClass"
    .PowerPoint_2010 = "PPTFrameClass"
    .PowerPoint_95 = "PP7FrameClass"
    .PowerPoint_97 = "PP97FrameClass"
    .PowerPoint_XP = "PP10FrameClass"
    .Project = "JWinproj-WhimperMainClass"
    .UserForm_2000 = "ThunderDFrame"
    .UserForm_97 = "ThunderXFrame"
    .Word = "OpusApp"
    .VBE = "wndclass_desked_gsk"
  End With
  FrameClassNames = fc
End Property

Public Sub CopyMemoryValueAdrFromReferenceAdr( _
  ByVal hpvDest As Object, _
  ByRef hpvSource As Object, _
  ByVal cbCopy As Long _
)
  CpyMemValAdrFromRefAdr hpvDest, hpvSource, cbCopy
End Sub

Public Sub CopyMemoryReferenceAdrFromValueAdr( _
  ByRef hpvDest As Object, _
  ByVal hpvSource As Object, _
  ByVal cbCopy As Long _
)
  CpyMemRefAdrFromValAdr hpvDest, hpvSource, cbCopy
End Sub


'The window name (lpWindowName) in findWindow is usually the window's caption;
'see the window's title bar. Passing the object in the following form might also work: Object.caption
'FindWindow(FrameClasses.Class, frm.Caption)
Public Function WindowHandle( _
  ByVal ClassName As String, _
  Optional WindowsName As String = vbNullString _
) As Long
  WindowHandle = FindWindow(ClassName, WindowsName)
End Function

Public Function ProcIdFromWinClass( _
  ByVal ClassName As String, _
  Optional WindowsName As String = vbNullString _
) As Long
Dim idProc As Long
   GetWindowThreadProcessId WindowHandle(ClassName, WindowsName), idProc ' Get PID for this HWnd
   ProcIdFromWinClass = idProc
End Function

Public Function ProcIdFromWinHandle( _
  ByVal hwnd As Long _
) As Long
Dim idProc As Long
  GetWindowThreadProcessId hwnd, idProc ' Get PID for this HWnd
  ProcIdFromWinHandle = idProc
End Function

Public Function CurrentProcessId() As Long
  CurrentProcessId = GetCurrentProcessId()
End Function

Public Function WinHandleFromProcId( _
  hInstance As Long _
) As Long
Dim tempHwnd As Long
  ' Grab the first window handle that Windows finds:
  tempHwnd = FindWindow(vbNullString, vbNullString)

  ' Loop until you find a match or there are no more window handles:
  Do Until tempHwnd = 0
    If GetParent(tempHwnd) = 0 Then ' Check if no parent for this window
      If hInstance = ProcIdFromWinHandle(tempHwnd) Then ' Check for PID match
        WinHandleFromProcId = tempHwnd ' Return found handle
        Exit Do
      End If
    End If
    tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT) ' Get the next window handle
  Loop
End Function

Public Function WinCaptionFromProcId( _
  ByVal hInstance As Long _
) As String
Dim hWndApp As Long           ' Window handle from WinHandleFromProcId.
Dim buffer As String          ' Holds caption of Window.
Dim numChars As Integer       ' Count of bytes returned.
  hWndApp = WinHandleFromProcId(hInstance)
  buffer = Space$(128)
  numChars = GetWindowText(hWndApp, buffer, Len(buffer))
  WinCaptionFromProcId = Left$(buffer, numChars)
End Function

Public Function GetWinCaptionFromWinHandle( _
ByVal winHndl As Long _
) As String
Dim buffer As String          ' Holds caption of Window.
Dim numChars As Integer       ' Count of bytes returned.
  buffer = Space$(128)
  numChars = GetWindowText(winHndl, buffer, Len(buffer))
  GetWinCaptionFromWinHandle = Left$(buffer, numChars)
End Function

Public Function username() As String
Dim lpBuff As String * 25
  Get_User_Name lpBuff, 25
  username = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function
          
Public Function ComputerName() As String
Dim lResponse As Long
Dim sUserName As String * 32
  lResponse = Get_ComputerName(sUserName, 32)
  ComputerName = Left(sUserName, InStr(sUserName, Chr$(0)) - 1)
End Function

'Resturns exectuable for passed data file. Returns null string if no application found
Public Function FindExe( _
  ByVal FilePath As String, _
  Optional ByVal dirPath As String = "" _
) As String
Dim appId As Long
Dim appName As String

  appName = Space(260)
  appId = FindExecutable(FilePath, dirPath, appName)
  
  If appId > 32 Then
    FindExe = appName
  Else
    FindExe = vbNullString
  End If
End Function

'Retrieves the window handle to the active window attached to the calling thread's message queue.
'You can quickly learn whether a specific form is still the active window
'using the following expression if the host application supports a handle property (Access does):
'GetActiveWindow() = | <> form.hWND
Public Function ActiveWindowHandle()
  ActiveWindowHandle = GetActiveWindow()
End Function

'Returns path to system's temporary folder.
Public Function TempPath() As String
Dim path As String * 512
Dim handle As Long
  handle = Get_TempPath(512, path)
  TempPath = Left(path, InStr(1, path, vbNullChar))
End Function

'Creates a name for a temporary file.
'If a unique file name is generated, an empty file is created and the handle to it is released;
'otherwise, only a file name is generated.
Public Function TempFileName( _
  Optional ByVal prefix As String = "" _
) As String
Dim path As String * 512
Dim fileName As String * 576
Dim handle As Long

  handle = Get_TempPath(512, path)

  If (handle > 0 And handle < 512) Then
     handle = Get_TempFileName(path, prefix, 0, fileName)
     If handle <> 0 Then
        TempFileName = Left$(fileName, InStr(fileName, vbNullChar))
     End If
  End If
End Function

'Retrieves a handle to the desktop window. The desktop window covers the entire screen.
'The desktop window is the area on top of which other windows are painted.
Public Function DesktopWindowHandle() As Long
  DesktopWindowHandle = DesktopWindow()
End Function

'Sets the specified window's show state.
'Returns True if operation was successfully executed
Function SetWindowState( _
  ByVal ClassName As String, _
  ByVal WindowsState As WindowShowState, _
  Optional ByVal WindowName As String = vbNullString _
) As Boolean
  Dim handle As Long
  handle = FindWindow(ClassName, WindowName)
  SetWindowState = CBool(ShowWindow(handle, WindowsState))
End Function

Public Sub SleepApplication( _
  ByVal miliseconds As Long _
)
  Sleep miliseconds
End Sub

Public Function FocusWindow( _
  ByVal winHandle As Long _
) As Boolean
  FocusWindow = BringWindowToTop(winHandle)
End Function

'TODO Revise
'---------------------------------------------------------------------------------------
' Procedure : ShellAndWait
' Purpose   : By Chip Pearson, chip@cpearson.com, www.cpearson.com
' This page on the web site: www.cpearson.com/Excel/ShellAndWait.aspx
'---------------------------------------------------------------------------------------
Public Function ShellAndWait(shellCommand As String, _
                    TimeOutMs As Long, _
                    ShellWindowState As VbAppWinStyle, _
                    BreakKey As ActionOnBreak) As ShellAndWaitResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShellAndWait
'
' This function calls Shell and passes to it the command text in ShellCommand. The function
' then waits for TimeOutMs (in milliseconds) to expire.
'
'   Parameters:
'       ShellCommand
'           is the command text to pass to the Shell function.
'
'       TimeOutMs
'           is the number of milliseconds to wait for the shell'd program to wait. If the
'           shell'd program terminates before TimeOutMs has expired, the function returns
'           ShellAndWaitResult.Success = 0. If TimeOutMs expires before the shell'd program
'           terminates, the return value is ShellAndWaitResult.TimeOut = 2.
'
'       ShellWindowState
'           is an item in VbAppWinStyle specifying the window state for the shell'd program.
'
'       BreakKey
'           is an item in ActionOnBreak indicating how to handle the application's cancel key
'           (Ctrl Break). If BreakKey is ActionOnBreak.AbandonWait and the user cancels, the
'           wait is abandoned and the result is ShellAndWaitResult.UserWaitAbandoned = 5.
'           If BreakKey is ActionOnBreak.IgnoreBreak, the cancel key is ignored. If
'           BreakKey is ActionOnBreak.PromptUser, the user is given a ?Continue? message. If the
'           user selects "do not continue", the function returns ShellAndWaitResult.UserBreak = 6.
'           If the user selects "continue", the wait is continued.
'
'   Return values:
'            ShellAndWaitResult.Success = 0
'               indicates the the process completed successfully.
'            ShellAndWaitResult.Failure = 1
'               indicates that the Wait operation failed due to a Windows error.
'            ShellAndWaitResult.TimeOut = 2
'               indicates that the TimeOutMs interval timed out the Wait.
'            ShellAndWaitResult.InvalidParameter = 3
'               indicates that an invalid value was passed to the procedure.
'            ShellAndWaitResult.SysWaitAbandoned = 4
'               indicates that the system abandoned the wait.
'            ShellAndWaitResult.UserWaitAbandoned = 5
'               indicates that the user abandoned the wait via the cancel key (Ctrl+Break).
'               This happens only if BreakKey is set to ActionOnBreak.AbandonWait.
'            ShellAndWaitResult.UserBreak = 6
'               indicates that the user broke out of the wait after being prompted with
'               a ?Continue message. This happens only if BreakKey is set to
'               ActionOnBreak.PromptUser.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TaskID As Long
Dim ProcHandle As Long
Dim WaitRes As Long
Dim Ms As Long
Dim MsgRes As VbMsgBoxResult
Dim SaveCancelKey As XlEnableCancelKey
Dim ElapsedTime As Long
Dim Quit As Boolean
Const ERR_BREAK_KEY = 18
Const DEFAULT_POLL_INTERVAL = 500

  If trim(shellCommand) = vbNullString Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
  End If
  
  If TimeOutMs < 0 Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
  ElseIf TimeOutMs = 0 Then
    Ms = WAIT_INFINITE
  Else
    Ms = TimeOutMs
  End If

  Select Case BreakKey
    Case AbandonWait, IgnoreBreak, PromptUser
      ' valid
    Case Else
      ShellAndWait = ShellAndWaitResult.InvalidParameter
      Exit Function
  End Select
  
  Select Case ShellWindowState
    Case vbHide, vbMaximizedFocus, vbMinimizedFocus, vbMinimizedNoFocus, vbNormalFocus, vbNormalNoFocus
      ' valid
    Case Else
      ShellAndWait = ShellAndWaitResult.InvalidParameter
      Exit Function
  End Select
  
  On Error Resume Next
  err.Clear
  TaskID = shell(shellCommand, ShellWindowState)
  If (err.Number <> 0) Or (TaskID = 0) Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
  End If
  
  ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskID)
  If ProcHandle = 0 Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
  End If
  
  On Error GoTo errh:
  'Application.EnableCancelKey available only in excel
  'SaveCancelKey = Application.EnableCancelKey
  'Application.EnableCancelKey = xlErrorHandler
  WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
  Do Until WaitRes = WAIT_OBJECT_0
    DoEvents
    Select Case WaitRes
      Case WAIT_ABANDONED
        ' Windows abandoned the wait
        ShellAndWait = ShellAndWaitResult.SysWaitAbandoned
        Exit Do
      Case WAIT_OBJECT_0
        ' Successful completion
        ShellAndWait = ShellAndWaitResult.Success
        Exit Do
      Case WAIT_FAILED
        ' attach failed
        ShellAndWait = ShellAndWaitResult.Success
        Exit Do
      Case WAIT_TIMEOUT
        ' Wait timed out. Here, this time out is on DEFAULT_POLL_INTERVAL.
        ' See if ElapsedTime is greater than the user specified wait
        ' time out. If we have exceed that, get out with a TimeOut status.
        ' Otherwise, reissue as wait and continue.
        ElapsedTime = ElapsedTime + DEFAULT_POLL_INTERVAL
        If Ms > 0 Then
            ' user specified timeout
            If ElapsedTime > Ms Then
                ShellAndWait = ShellAndWaitResult.TimeOut
                Exit Do
            Else
                ' user defined timeout has not expired.
            End If
        Else
            ' infinite wait -- do nothing
        End If
        ' reissue the Wait on ProcHandle
        WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
      Case Else
        ' unknown result, assume failure
        ShellAndWait = ShellAndWaitResult.Failure
        Quit = True
    End Select
  Loop
  
  CloseHandle ProcHandle
  'Application.EnableCancelKey = SaveCancelKey
  Exit Function
  
errh:
  'Debug.Print "ErrH: Cancel: " & Application.EnableCancelKey
  If err.Number = ERR_BREAK_KEY Then
    If BreakKey = ActionOnBreak.AbandonWait Then
      CloseHandle ProcHandle
      ShellAndWait = ShellAndWaitResult.UserWaitAbandoned
      'Application.EnableCancelKey = SaveCancelKey
      Exit Function
    ElseIf BreakKey = ActionOnBreak.IgnoreBreak Then
      err.Clear
      Resume
    ElseIf BreakKey = ActionOnBreak.PromptUser Then
      MsgRes = MsgBox("User Process Break." & vbCrLf & _
         "Continue to wait?", vbYesNo)
      If MsgRes = vbNo Then
        CloseHandle ProcHandle
        ShellAndWait = ShellAndWaitResult.UserBreak
        'Application.EnableCancelKey = SaveCancelKey
      Else
        err.Clear
        Resume Next
      End If
    Else
      'Debug.Print "Unknown value of 'BreakKey': " & CStr(BreakKey)
      CloseHandle ProcHandle
      'Application.EnableCancelKey = SaveCancelKey
      ShellAndWait = ShellAndWaitResult.Failure
    End If
  Else
    ' some other error. assume failure
    CloseHandle ProcHandle
    ShellAndWait = ShellAndWaitResult.Failure
  End If
  
  'Application.EnableCancelKey = SaveCancelKey
End Function

Public Function DotNetVersion() As String
  Dim keys As New ArrayList
  DotNetVersion = ""
  
  keys.Append "HKLM\Software\Microsoft\Active Setup\Installed Components\{78705f0d-e8db-4b2d-8193-982bdda15ecd}\Version"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.0\Version"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v3.5\Version"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Client\Version"
  keys.Append "HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full\Version"
  
  Dim i As Integer
  Dim s As String
  For i = keys.lb To keys.ub
    s = Util.Registry.ReadKey(keys(i))
    If s <> "" Then DotNetVersion = keys(i) & ": " & s
  Next
End Function

Public Function ShellRun( _
  ExecutionText As String _
) As Variant
  Dim shell As Object
  Set shell = CreateObject("WScript.Shell")
  
  Dim exec As Object
  Set exec = shell.exec(ExecutionText)
  
  Dim out As Object
  Set out = exec.StdOut
  
  Dim result As New ArrayList
  Do Until out.AtEndOfStream
    result.Append out.ReadLine
  Loop
  
  ShellRun = result.Values
End Function

Private Sub ActiveDirectoryTest()
'Active Directroy all attributes
'https://msdn.microsoft.com/en-us/library/ms675090(v=vs.85).aspx

'List of LDAP Attributes Supported by ADMP
'https://www.manageengine.com/products/ad-manager/help/csv-import-management/active-directory-ldap-attributes.html

Set rootDSE = GetObject("LDAP://RootDSE")
base = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
'filter on user objects with the given account name
fltr = "(&(objectClass=user)(objectCategory=Person)" & _
        "(sAMAccountName=" & username() & "))"
'add other attributes according to your requirements
attr = "displayName,mail,distinguishedName,sAMAccountName"
scope = "SubTree"

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

Set cmd = CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandText = base & ";" & fltr & ";" & attr & ";" & scope

Set rs = cmd.Execute
Do Until rs.EOF
  Debug.Print rs.fields("displayName").Value
  rs.MoveNext
Loop
rs.Close

conn.Close
End Sub
