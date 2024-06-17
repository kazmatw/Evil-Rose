Attribute VB_Name = "ModExport"
Sub ExportModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim file As String

    exportPath = ThisWorkbook.Path & "\ExportedModules\"
    
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    file = Dir(exportPath & "*.*")
    Do While file <> ""
        Kill exportPath & file
        file = Dir()
    Loop

    Set vbProj = ThisWorkbook.VBProject

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case vbext_ct_Document
            Case Else
        End Select
    Next vbComp
    
    MsgBox "Export complete!"
End Sub


Sub ImportModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim fileName As String
    Dim importPath As String
    Dim fileType As String
    Dim compName As String

    importPath = ThisWorkbook.Path & "\ExportedModules\"
    Set vbProj = ThisWorkbook.VBProject

    ' Loop through all files in the export directory
    fileName = Dir(importPath & "*.*")
    Do While fileName <> ""
        ' Determine the file type based on the extension
        fileType = LCase(Right(fileName, 4))
        compName = Left(fileName, Len(fileName) - 4)
        
        ' Delete the existing component if it exists
        On Error Resume Next
        Set vbComp = vbProj.VBComponents(compName)
        If Not vbComp Is Nothing Then
            vbProj.VBComponents.Remove vbComp
        End If
        On Error GoTo 0
        
        ' Import the new component
        Select Case fileType
            Case ".bas", ".cls", ".frm"
                vbProj.VBComponents.Import importPath & fileName
            Case Else
                ' Skip other file types
        End Select
        
        fileName = Dir
    Loop

    MsgBox "Import complete!"
End Sub


