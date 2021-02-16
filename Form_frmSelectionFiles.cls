VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSelectionFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    optFileLOB.SetFocus
End Sub

Private Sub btnReturnToMainMenu_Click()
    DoCmd.Close acForm, "frmSelectionFiles", acSaveNo
    DoCmd.OpenForm "frmLandingDRG", acNormal, , , , acWindowNormal
End Sub

Private Sub btnViewSelectionFileData_Click()
    If Application.CurrentData.AllQueries("qrySelectionFileReport").IsLoaded = True Then DoCmd.Close acQuery, "qrySelectionFileReport", acSaveNo
    DoCmd.OpenQuery "qrySelectionFileReport", acViewNormal, acReadOnly
End Sub
Private Sub btnImportSelectionFile_Click()
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String, fileName As String, fileType As String
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    Dim lobid As Long
    Dim subDate As String: subDate = Now()
    
'On Error GoTo err1:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'default folder location of selection file
    fileLocation = "\\acct.upmchs.net\hp\Quality Audit\Quality Audit(Custom)\HDI\Files From FTP"
    
    'test location
    'fileLocation = "X:\QA - Tech Team\QA - Tech Team(DeptUsers)\Focused_Review\HMS File Drop\"

    'open file dialog window
    With fDialog
        .AllowMultiSelect = False
        .Title = "Pick a Selection File to import"
        .InitialFileName = fileLocation
        .Filters.Clear
            If .Show = 0 Then
                End
            Else
                filePath = .SelectedItems(1)
                fileName = Right$(.SelectedItems(1), Len(.SelectedItems(1)) - InStrRev(.SelectedItems(1), "\"))
                If InStr(fileName, ".") > 0 Then: fileName = Left$(fileName, InStr(fileName, ".") - 1)
                
                If InStr(fileName, "Medicare") > 0 Then
                    fileType = "Medicare"
                    lobid = 1
                ElseIf InStr(fileName, "Medicaid") > 0 Then
                    fileType = "Medicaid"
                    lobid = 3
                ElseIf InStr(fileName, "Commercial") > 0 Then
                    fileType = "Commercial"
                    lobid = 2
                End If
            End If
    End With
    
    'validate if good file
    If ValidateForm(1, fileName) = True Then
        If MsgBox("Error importing file. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
    'prompt user if they'd like it import file - check if file already imported
    If MsgBox("Do you want to import the selected file?", vbQuestion + vbYesNo, "Import HMS Selection File") = vbYes Then
    Else
        Exit Sub
    End If
    
    DoCmd.SetWarnings False
    'import file from excel into table
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tblSelection", filePath, True

    'set file name to each new record in table
    DoCmd.RunSQL ("UPDATE [tblSelection] SET file_name = '" & fileName & "', lob_id = " & lobid & " WHERE file_name IS NULL;")

    'check - paid_dt > 365
    DoCmd.RunSQL ("UPDATE [tblSelection] " & _
        " SET selection_status = IIF([paid_dt]<(Date()-365), 'D','A'), " & _
        " selection_status_desc = IIF([paid_dt]<(Date()-365), 'Paid Date greater than 365 days','') " & _
        " WHERE file_name = '" & fileName & "';")
    
    'check - provider exclusion Non DRG
    'matching on provider name
    DoCmd.RunSQL ("UPDATE tblSelection " & _
        " INNER JOIN tblExclusionProvidersNonDRG ON tblExclusionProvidersNonDRG.provider_name = tblSelection.provider_name " & _
        " SET selection_status = 'D', " & _
        " selection_status_desc = 'Provider is on exclusion list' " & _
        " WHERE file_name = '" & fileName & "' AND (tblExclusionProvidersNonDRG.lob_id = " & lobid & " OR tblExclusionProvidersNonDRG.lob_id = 4);")
        
    'check - provider exclusion CCBH
    'match on provider_id/billingprovider
    DoCmd.RunSQL ("UPDATE tblSelection " & _
        " INNER JOIN tblExclusionProvidersCCBH ON tblExclusionProvidersCCBH.vendor_id = tblSelection.billingprovider " & _
        " SET selection_status = 'D', " & _
        " selection_status_desc = 'Provider is on exclusion list' " & _
        " WHERE file_name = '" & fileName & "';")
    
    'check - member exclusion list
    DoCmd.RunSQL ("UPDATE tblSelection " & _
        " INNER JOIN tblExclusionMembers ON tblExclusionMembers.member_id = tblSelection.Patient_ID " & _
        " SET selection_status = 'D', " & _
        " selection_status_desc = 'Member is on exclusion list' " & _
        " WHERE file_name = '" & fileName & "';")
    
    
    'update file log
    DoCmd.RunSQL ("INSERT INTO tblFileLogSelection (user_id, import_date, file_name, lob_id) " & _
        "VALUES ('" & userID & "', '" & subDate & "', '" & fileName & "', " & lobid & "); ")
        
    'export to excel file
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef

    Set db = CurrentDb
    Set qd = db.QueryDefs("qryExportSelectionFile")

    DoCmd.RunSQL ("DELETE * FROM UPMCComplexAcceptfileOut;")
    qd.Parameters("[File Name]").Value = fileName
    qd.Execute
    qd.Close: Set qd = Nothing
    db.Close: Set db = Nothing

    Call UpdateSubmissionLog(userID, "xxxxxxxxxxx", fileName, 4, lobid)
    Call ExportSelectionFile(fileName)
        
'err1:
'    Select Case Err.Number
'        Case 0
'        Case Else
'            Call LogError(Err.Number & " " & Err.Description, "frmSelectionFiles; btnImportSelectionFile_Click()")
'            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
'            Exit Sub
'    End Select

End Sub


Public Sub ExportSelectionFile(fileName As String)
        
On Error GoTo err1:
    If MsgBox("Process complete - please choose a location to save Selection Return File.", vbOKOnly, "Complete") = vbOK Then:

        Dim fDialog As FileDialog
        Dim filePath As String, fileLocation As String
        fileLocation = "X:\QA - Overpayment Recovery\QA - Overpayment Recovery(DeptUsers)\HMS\"
        
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
        With fDialog
            .AllowMultiSelect = False
            .Title = "Select a location to save the return file"
            .InitialFileName = fileLocation
            .Filters.Clear
            If .Show = 0 Then
                End
            Else
                filePath = .SelectedItems(1)
            End If
        End With
        'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryComplexCommercialReturnFile", filePath & "\" & fileName & "ReturnClaimFile", True
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "UPMCComplexAcceptfileOut", filePath & "\" & fileName & "Return_Complex", True
        
        If MsgBox("Export complete - go to file folder?", vbYesNo) = vbYes Then
            Shell "explorer.exe " & filePath, vbNormalFocus
        End If
        
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmSelectionFiles; ExportSelectionFile()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
        
End Sub

Public Function ValidateForm(submitType As Integer, fileName As String) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1 'Import file validation
            'StrComp() = 0 means exact match
            'InStr() = 0 means not found in sring
            intComp = InStr(fileName, "Select")
            If intComp = 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is not a <b>Selection File</b>.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
            intComp = InStr(fileName, "DRG")
            If intComp = 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is not <b>DRG</b>.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
            
            intComp = InStr(fileName, "Return")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is a <b>Return File</b>.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
            intComp = InStr(fileName, "Adjust")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
            intComp = InStr(fileName, "Eligible")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
            
            If DCount("file_name", "tblFileLogSelection", "file_name = '" & fileName & "'") > 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected <b>Selection File</b> has already been imported.</li>"
                ctlName = "btnImportSelectionFile,"
            End If
                
        End Select
    
    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(255, 255, 255)
        ValidateForm = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "<b>Submission Errors</b>"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        varFocus = Split(ctlName, ",")
        Me.Controls(varFocus(0)).SetFocus
        ValidateForm = True
    End If

End Function
