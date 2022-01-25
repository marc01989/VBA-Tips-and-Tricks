'----------------------------------------------------------------
'must have reference "Microsoft Office 2016 Object Library" selected to use FileDialog object
    'go to tools -> references -> Microsoft Office 2016 Object Library
'----------------------------------------------------------------

Private Sub btnExport_Click()
    
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim exportDate As String
    exportDate = Format(Date(), "yyyy_mm_dd")

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fileLocation = "C:\Users\Public\Desktop"
    
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select a location to export files"
        .InitialFileName = fileLocation
        .Filters.Clear
        If .Show = 0 Then
            End
        Else
            filePath = .SelectedItems(1)
        End If
    End With

    filePath = filePath & "\"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, expQry, filePath & exportDate, True
    
    If MsgBox("Export complete - go to file folder?", vbYesNo) = vbYes Then
        Shell "explorer.exe " & filePath, vbNormalFocus
    End If
    
End Sub
