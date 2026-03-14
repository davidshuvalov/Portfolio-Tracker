Attribute VB_Name = "Module1"
Sub ExportVBAModules()

    Dim comp As Object
    Dim exportPath As String
    
    exportPath = ThisWorkbook.Path & "\VBA_Export\"
    
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        Select Case comp.Type
        
            Case 1 'Standard Module
                comp.Export exportPath & comp.name & ".bas"
                
            Case 2 'Class Module
                comp.Export exportPath & comp.name & ".cls"
                
            Case 3 'UserForm
                comp.Export exportPath & comp.name & ".frm"
                
        End Select
        
    Next comp
    
    MsgBox "Export complete"

End Sub
