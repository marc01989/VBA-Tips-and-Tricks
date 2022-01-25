'----------------------------------------------------------------
'must have reference "Microsoft Office 2016 Object Library" selected to use FileDialog object
    'go to tools -> references -> Microsoft Office 2016 Object Library
'----------------------------------------------------------------

Private Sub btnImportFile_Click()
    Dim fDialog As FileDialog
    Dim filePath As String, fileDefault As String
    
On Error GoTo err1:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        
    'default folder location - where filepicker will default to
    fileDefault = "C:\Users\Public\Desktop"

    'open file dialog window
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select a file to import"
        .InitialFileName = fileDefault
        .Filters.Clear
            If .Show = 0 Then
                End
            Else
                filePath = .SelectedItems(1)
            End If
    End With
        
    'prompt user if they want to import the selected file
    If MsgBox("Do you want to import the selected file?", vbQuestion + vbYesNo, "Import?") = vbYes Then        
        DoCmd.SetWarnings False
        'import file from excel into table
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "File Import", filePath, True   
    Else
        End
    End If

    MsgBox ("Import Complete")

err1:
    Select Case Err.Number
        Case 0
        Case Else
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

