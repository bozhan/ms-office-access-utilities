Attribute VB_Name = "ErrorHandlerExample"
Option Compare Database
Option Explicit


'AppTitle               Property  Get the current value of AppTitle
'CurrentOperation       Property  Get the current value of CurrentOperation. Use this property to create logical checkpoints in your procedures. Before crucial sections of code, assign a 'unique value to the CurrentOperation property. If the error handler is triggered after assigning this value, you will have a better idea of where the error occurred. This is especially 'helpful if you do not use line numbers in your code.
'Destination            Property  Get the current value of the Destination property. If you assigned an error log path name string to this property, a string is returned. If you assigned a Recordset 'object to this property, an object pointer to that recordset is returned.
'ShowErrorFile          Property  Determine if the user is shown the error log after it's created
'OverwriteLog           Property  Determine if the error log is overwritten each time or new info appended to the existing file
'DisplayMsgOnError      Property  Get the current value of DisplayMsgOnError.
'ErrorDescription       Property  Get the description of the error that caused the error handler to be triggered (undefined until an error is handled by the HandleError method). Refer to this 'property in the AppSpecificErrorHandler procedure, or in the code triggered in response to the AfterHandlerCalled event.
'ErrorLine              Property  Get the line number of the procedure that caused the error when the error handler was triggered (undefined until an error is handled by the HandleError method). Refer 'to this property in the AppSpecificErrorHandler procedure, or in the code triggered in response to the AfterHandlerCalled event.
'ErrorNumber            Property  Get the error code that caused the error when the error handler is triggered (undefined until an error is handled by the HandleError method). Refer to this property 'in the AppSpecificErrorHandler procedure, or in the code triggered in response to the AfterHandlerCalled event.
'IncludeExpandedInfo    Property  Determine whether additional information about the user's machine environment is included in the error log file
'ProcName               Property  Get the name of the procedure containing the error that triggered the error handler (undefined until an error is handled by the HandleError method). Refer to this 'property in the AppSpecificErrorHandler procedure, or in the code triggered in response to the AfterHandlerCalled event.
'TraceExecution         Property  Determine whether procedure execution is traced (in addition to handling errors). If this property is true, a log file is maintained which tracks each procedure 'which is logged by the Push method. Even if no error occurs in the procedure, an entry is written. This information can be useful to trace the order in which procedures are called in your 'application. By default the information is stored in a file named "trace.log" in your application's directory.
'Class_Initialize       Initialize  Set initial values to defaults which may be overridden with property settings
'Class_Terminate        Terminate   Close trace log if opened, and check to be sure that all Push calls are balanced by Pop calls. If you raise an assertion at this point, double-check your code 'to see that you are calling the error handler correctly.
'ClearLog               Method  Clear the error log (or table depending on the setting of the Destination property). If Destination is a string property indicating a DOS text file, this function 'simply deletes the file. If Destination is a Recordset object, this method deletes all rows from the recordset.
'HandleError            Method  Logs errors and responds to run-time errors.
'                         This is the main routine which handles a run-time error. It is responsible for displaying an error message, writing log 'information to a file, and raising events back into the caller program. This method is called if a run-time error branches to a section of code which calls this method. This method should 'only be called AFTER an error has occurred. First the current state of the error object is saved (error message, line, description). The BeforeHandlerCalled event is raised in the program 'which instantiated the error handler. The Cancel argument to BeforeHandlerCalled is tested. If the user does not cancel the error handler, by setting Cancel to True, the error handler 'proceeds to save error state information to the log file. After the error information has been saved, the AfterHandlerCalled event is raised to give the caller a chance to provide a 'custom response to the error, and the "AppSpecificHandler" private subroutine is called.
'                         If your program is not receiving the BeforeHandlerCalled or AfterHandlerCalled events, you can use 'this procedure to provide custom error handling, for example to display a custom form, or to provide the user a chance to exit the program
'Pop                    Method  Pops the current procedure name off the error handling stack. Call this method when your code successfully exits a procedure with no errors. It must be balanced by a call to 'the Push method when the procedure is first called.
'Push                   Method  Pushes the supplied procedure name onto the error handling stack. Call this method at the beginning of a procedure. It must be balanced by a call to the Pop method which is 'called when the procedure exits normally.
'AppSpecificErrHandler  Private   Custom error handling stub subroutine (perform a custom action here). The recommended way to use the CErrorHandlerVBA class is to have it raise events 'into your program. You can receive the BeforeHandlerCalled and AfterHandlerCalled events. There you can provide specific functionality for your program, such as displaying a custom error 'message form, or closing the program cleanly. The advantage to using the events is that the CErrorHandlerVBA object can be completely generic. All custom app-specific logic is contained 'in your program, not in the class itself. If you cannot use the events however, because you are declaring the error handler object variable in a standard module, which cannot receive 'events, you can customize this procedure to provide app-specific error-handling logic. All app-specific logic should be confined to this procedure. That way when you incorporate the 'CErrorHandlerVBA class into your application, this is the only part of the code that will need
'to be modified.
'GetLastErr             Private   This method is called to get the last error that occurred. We increment the pointer as one of the last things done in Push() in order to add the next procedure to 'the next item available in the stack array. However, if an error occurred before we could add the next procedure to the stack, we need to go back to the previous item in the array to get 'the error that occurred.
'CreateErrorLog         Private   Generate the error log text
'AppendTextFile         Private   Create a file with the error log or append to the file if it already exists.
'DeleteFile             Private   Delete the named file, handling errors if the file does not exist
'LogErrorToTable        Private   Logs the most recent error to a table. The table is specified by setting the Destination property of the error handler object to a recordset object which you 'create in your application
'ShowFile               Private   Open the error file (assumes a default program will open it based on its extension)

' Example of CErrorHandlerVBA
'
' To try this example, do the following:
' 1. Create a new form. This example uses the Access forms object model.
' 2. Add a command button named 'cmdTest'
' 3. Add a check box named 'chkResponseMode'.
'       When this box is unchecked the default error handling message will be displayed.
'       When it is checked, a custom error handling action will be taken.
' 4. Paste all the code from this example to the new form's module.

'Private WithEvents mErrHandler As CErrorHandlerVBA

' This example assumes that the sample files are located in the folder named by the following constant.
'Private Const mcstrSamplePath As String = "C:\TVSBSamp"
'
'Private Sub Form_Load()
'  On Error GoTo PROC_ERR
'
'  ' Create the error handler object and set some of its optional properties
'  Set mErrHandler = New CErrorHandlerVBA
'
'  With mErrHandler
'    .AppTitle = "My Application"
'    .Destination = mcstrSamplePath & "\err.txt"
'    .DisplayMsgOnError = True
'    .IncludeExpandedInfo = True
'    .CurrentOperation = "Initializing"
'
'    ' Set to not overwrite the error log so you can see new errors appeneded to the end of the file
'    .OverwriteLog = False
'
'    ' Shows the error log to the user
'    .ShowErrorFile = True
'
'    .Push "Form_Load"
'  End With
'
'PROC_EXIT:
'  mErrHandler.Pop
'  Exit Sub
'
'PROC_ERR:
'  mErrHandler.HandleError
'  GoTo PROC_EXIT
'End Sub
'
'Private Sub cmdTest_Click()
'  ' Comments: Standard error handling style used for all non-trivial procedures.
'  '           The procedure starts with a Push and the name of the procedure to the error handler stack.
'  '           It ends with a Pop to take it off the stack.
'  '           For this technique to be effective, procedures should always end at the bottom without an EXIT in the middle.
'
'  On Error GoTo PROC_ERR
'  mErrHandler.Push "cmdTest_Click"
'
'  ' Call the next procedure
'  Call Proc1
'
'PROC_EXIT:
'  mErrHandler.Pop
'  Exit Sub
'
'PROC_ERR:
'  mErrHandler.HandleError
'  GoTo PROC_EXIT
'End Sub
'
'Private Sub Proc1()
'  ' Comments: This procedure is called from the cmdTest_Click procedure and adds its name to the procedure call stack with the Push command.
'  '           That lets the error handler track the chain of procedures that are called.
'
'  On Error GoTo PROC_ERR
'  mErrHandler.Push "Proc1"
'
'  Call Proc2
'
'PROC_EXIT:
'  mErrHandler.Pop
'  Exit Sub
'
'PROC_ERR:
'  mErrHandler.HandleError
'  GoTo PROC_EXIT
'End Sub
'
'Private Sub Proc2()
'  ' Comments: This procedure is called from the Proc1 procedure and adds its name to the procedure call stack with the Push command.
'  '           That lets the error handler track this third procedure that's called.
'
'  On Error GoTo PROC_ERR
'  mErrHandler.Push "Proc2"
'
'  ' Specify a checkpoint to narrow down the location of the error if you are not using line numbers
'  mErrHandler.CurrentOperation = "I'm about to die"
'
'  ' Simulate an error
'  MsgBox (1 / 0)
'
'  ' This code is bypassed by the error handler
'  mErrHandler.CurrentOperation = "I'll never make it"
'
'PROC_EXIT:
'  mErrHandler.Pop
'  Exit Sub
'
'PROC_ERR:
'  mErrHandler.HandleError
'  GoTo PROC_EXIT
'End Sub
'
'Private Sub chkResponseMode_Click()
'  ' Comments: When unchecked, let the error handler display a standard error message.
'  '           When checked, no message will be displayed, so display a custom message in the BeforeHandlerCalled event of the error object.
'
'  On Error GoTo PROC_ERR
'  mErrHandler.Push "chkResponseMode_Click"
'
'  mErrHandler.DisplayMsgOnError = Not (chkResponseMode.value)
'
'PROC_EXIT:
'  mErrHandler.Pop
'  Exit Sub
'
'PROC_ERR:
'  mErrHandler.HandleError
'  GoTo PROC_EXIT
'End Sub
'
'Private Sub mErrHandler_AfterHandlerCalled()
'  ' Comments: This event is raised after the error is triggrered and the .HandleError method is called.
'  '           The BeforeHandlerCalled event can be used to prevent this error from being raised
'
'  ' If you have the error handler already displaying the error message, nothing happens in this procedure
'  If Not mErrHandler.DisplayMsgOnError Then
'    ' This section lets you handle situations where the error handler doesn't show anything to the user and you want to control it here
'
'    MsgBox "Closing application"
'
'    ' If you are writing the error to a text file this will open the file for viewing
'    shell "notepad.exe " & mErrHandler.Destination, vbNormalFocus
'    Quit
'  End If
'
'End Sub
'
'Private Sub mErrHandler_BeforeHandlerCalled(Cancel As Boolean)
'  ' Comments: This event is raised when the .HandleError method is called on the errorhandler object.
'  '           This is your opportunity to handle the error differently.
'  '           This procedure is not necessary if you want the error handler to run without special intervention.
'  ' Set     : Cancel          If this is set to True, then the AfterHandlerCalled event is not raised.
'
'  Dim strCustomMessage As String
'
'  ' By default, the error handler class will
'  Cancel = False
'
'  With mErrHandler
'    If Not .DisplayMsgOnError Then
'      strCustomMessage = "Custom Error Handler" & vbCrLf & _
'        "Something bad happened while doing: " & .CurrentOperation & vbCrLf & _
'        "The current procedure is: " & .ProcName & vbCrLf & _
'        "The error message is: " & .ErrorDescription & vbCrLf & _
'        "The error number is: " & .ErrorNumber & vbCrLf & _
'        "The error log will now be opened."
'
'      MsgBox strCustomMessage, , .AppTitle
'
'      If MsgBox("Close the application?", vbQuestion + vbYesNo, .AppTitle) = vbNo Then
'        Cancel = True
'      End If
'    End If
'  End With
'
'End Sub
'
