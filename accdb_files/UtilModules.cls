VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TODO revise whole module improve on procedures add import single moduel from path

Public Enum vbext_ComponentType
  vbext_ct_ActiveXDesigner = 11
  vbext_ct_ClassModule = 2
  vbext_ct_Document = 100
  vbext_ct_MSForm = 3
  vbext_ct_StdModule = 1
End Enum

Public Sub ExportModules(ByRef wb As Workbook)
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As Variant

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles(wb) = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
        Kill FolderWithVBAProjectFiles(wb) & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = wb.Name
    Set wkbSource = wb.Application.Workbooks(szSourceWorkbook)

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles(wb) & "\"

    For Each cmpComponent In wkbSource.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ComponentType.vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ComponentType.vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ComponentType.vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ComponentType.vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select

        If bExport Then
            ''' Export the component to a text file.
           cmpComponent.Export szExportPath & szFileName
          'Debug.Print szFileName
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent

        End If

    Next cmpComponent

    MsgBox "Export is ready"
End Sub

Public Sub ImportModules(ByRef wb As Workbook)
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As Object

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles(wb) = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = wb.Name
    Set wkbTarget = wb.Application.Workbooks(szTargetWorkbook)

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles(wb) & "\"

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files

        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If

    Next objFile

    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles(ByRef obj As Object) As String
    Dim fso As Object
    Dim SpecialPath As String

    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = Left(obj.FullName, Len(obj.FullName) - Len(obj.Name))

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If

    If fso.FolderExists(SpecialPath & obj.Name & "_Files") = False Then
        On Error Resume Next
        MkDir SpecialPath & obj.Name & "_Files"
        On Error GoTo 0
    End If

    If fso.FolderExists(SpecialPath & obj.Name & "_Files") = True Then
        FolderWithVBAProjectFiles = SpecialPath & obj.Name & "_Files"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If

End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As Object
        Dim VBComp As Variant

        Set VBProj = ActiveWorkbook.VBProject

        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

'Export modules procedure for ms access
Public Sub ExportAccessModules(ByRef obj As Object)
    Dim bExport As Boolean
    Dim szFileName As String
    Dim cmpComponent As Variant
    
    If FolderWithVBAProjectFiles(obj) = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    On Error Resume Next
        Kill FolderWithVBAProjectFiles(obj) & "\*.*"
    On Error GoTo 0
    
    Dim exportPath  As String
    exportPath = FolderWithVBAProjectFiles(obj) & "\"

    For Each cmpComponent In obj.Parent.Application.VBE.VBProjects.item(1).VBComponents
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ComponentType.vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ComponentType.vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ComponentType.vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ComponentType.vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select

        If bExport Then
           cmpComponent.Export exportPath & szFileName
        End If

    Next cmpComponent

    MsgBox "Export is ready"
End Sub

