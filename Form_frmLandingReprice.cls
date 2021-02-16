VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLandingReprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    cboFileName.RowSource = ""
End Sub
Private Sub Form_Activate()
    Call ClearForm
    Call RefreshData
End Sub

Private Sub optFileStatus_AfterUpdate()
    Call ClearForm
    Call UpdateRowSource
End Sub
Private Sub optLOB_AfterUpdate()
    Call ClearForm
    Call UpdateRowSource
End Sub
Public Sub UpdateRowSource()
    If optLOB.Value <> "" And optFileStatus.Value <> "" Then
        cboFileName.RowSource = "SELECT DISTINCT file_name FROM tblFileLogReprice WHERE lob_id = " & optLOB.Value & " AND return_file_sent = " & optFileStatus.Value & " ORDER BY file_name;"
        cboFileName.Value = Null
        cboFileName.Requery
    End If
    Call RefreshData
End Sub
Private Sub cboFileName_Change()
    Call ClearForm
    Call RefreshData
End Sub

Private Sub btnImportRepriceFile_Click()
    'import reprice file
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String, fileName As String, fileType As String
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    Dim lobid As Long
    Dim subDate As String: subDate = Now()
    Dim intComp As Integer
    
On Error GoTo err1:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    'default folder location of selection file
    fileLocation = "\\acct.upmchs.net\hp\Quality Audit\Quality Audit(Custom)\HDI\Files From FTP"
    'fileLocation = "X:\QA - Tech Team\QA - Tech Team(DeptUsers)\Focused_Review\HMS File Drop\"

    'open file dialog window
    With fDialog
        .AllowMultiSelect = False
        .Title = "Pick a Reprice File to import"
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
    
    'prompt user if they want to import selected file
    If MsgBox("Do you want to import the selected Reprice file?", vbQuestion + vbYesNo, "Import HMS Reprice File") = vbYes Then
    Else
        Exit Sub
    End If

    DoCmd.SetWarnings False
    
    'import file from excel into table
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tblReprice", filePath, True
    
    'UPDATED
    'set [File Name] to name of complex file and [audit status] to 1
    DoCmd.RunSQL ("UPDATE [tblReprice] SET [File Name] = '" & fileName & "', [Audit Status] = 1, [lob_id] = " & lobid & " WHERE [File Name] IS NULL;")

    'check against all selection files
    DoCmd.RunSQL ("UPDATE tblReprice " & _
        " LEFT JOIN tblSelection ON tblReprice.Claim_number = tblSelection.claim_no " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 1, " & _
        " [New Allowed Amt] = 0, " & _
        " [Overpayment_Response] = 'N'," & _
        " [Overpayment_Exception_Reason] = 99," & _
        " [Overpayment_Exception_Comments] = 'Claim was not approved for review on the selection files'," & _
        " [Calculated_Savings] = 0 " & _
        " WHERE ([selection_status] = 'D' OR [selection_status] is null) " & _
        " AND [File Name] = '" & fileName & "' ;")
    
    'check if [paid_dt] is greater than 365
    DoCmd.RunSQL ("UPDATE tblReprice " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments]= 2, " & _
        " [New Allowed Amt] = 0, " & _
        " [Overpayment_Response] = 'N'," & _
        " [Overpayment_Exception_Reason] = 99," & _
        " [Overpayment_Exception_Comments] = 'Claim paid date is greater than 365 days'," & _
        " [Calculated_Savings] = 0" & _
        " WHERE DATEDIFF('d', DATESERIAL(LEFT([paid_date], 4), MID([paid_date], 5,2), RIGHT([paid_date], 2)), FORMAT(NOW(), 'mm/dd/yyyy')) > 365 " & _
        " AND [File Name] = '" & fileName & "';")

    'check for provider exclusions - CCBH
    DoCmd.RunSQL ("UPDATE tblReprice " & _
        " INNER JOIN tblExclusionProvidersCCBH ON tblReprice.[Servicing_Provider_ID] = tblExclusionProvidersCCBH.vendor_id " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 3, " & _
        " [New Allowed Amt] = 0, " & _
        " [Overpayment_Response] = 'N'," & _
        " [Overpayment_Exception_Reason] = 99," & _
        " [Overpayment_Exception_Comments] = 'Provider is on exclusion list'," & _
        " [Calculated_Savings] = 0 " & _
        " WHERE [File Name] = '" & fileName & "' AND (tblExclusionProvidersCCBH.lob_id = " & lobid & " or tblExclusionProvidersCCBH.lob_id = 4);")


    'check for provider exclusions - NonDRG
    DoCmd.RunSQL ("UPDATE tblReprice " & _
        " INNER JOIN tblExclusionProvidersNonDRG ON tblReprice.[Provider_Name] = tblExclusionProvidersNonDRG.provider_name " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 3, " & _
        " [New Allowed Amt] = 0, " & _
        " [Overpayment_Response] = 'N'," & _
        " [Overpayment_Exception_Reason] = 99," & _
        " [Overpayment_Exception_Comments] = 'Provider is on exclusion list'," & _
        " [Calculated_Savings] = 0 " & _
        " WHERE [File Name] = '" & fileName & "' AND (tblExclusionProvidersNonDRG.lob_id = " & lobid & " or tblExclusionProvidersNonDRG.lob_id = 4);")

        
    'check for member exclusion
    DoCmd.RunSQL ("UPDATE tblReprice " & _
        " INNER JOIN tblExclusionMembers ON tblReprice.patient_id = tblExclusionMembers.member_id " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 4, " & _
        " [New Allowed Amt] = 0, " & _
        " [Overpayment_Response] = 'N'," & _
        " [Overpayment_Exception_Reason] = 99," & _
        " [Overpayment_Exception_Comments] = 'Member is on exclusion list', " & _
        " [Calculated_Savings] = 0 " & _
        " WHERE [File Name] = '" & fileName & "';")
        
    'update file log
    DoCmd.RunSQL ("INSERT INTO tblFileLogReprice (user_id, import_date, file_name, lob_id) " & _
        "VALUES ('" & userID & "', '" & subDate & "', '" & fileName & "', " & lobid & "); ")

    Call UpdateSubmissionLog(userID, "xxxxxxxxxxx", fileName, 5, lobid)

    DoCmd.SetWarnings True
    Call RefreshData
    
    MsgBox ("Import Complete.")
 

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingReprice; btnImportRepriceFile_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub btnOpenReprice_Click()
    Dim args As String
    If cboFileName.Value <> "" Then
        args = cboFileName.Value
        args = args & "," & optLOB.Value & "," & optFileStatus.Value
    End If
    DoCmd.OpenForm "frmWorksheetReprice", acNormal, , , , acFormPropertySettings, args
End Sub

Private Sub btnRepriceReturnFile_Click()
    'reprice return file
    'naming convention:
    '202007080700_UPMC_53_Commercial_HMSReturnReprice_Complex
    
    If ValidateForm(2, "") = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

'On Error GoTo err1:
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim fileName As String: fileName = cboFileName.Value
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    Dim lobid As Long: lobid = optLOB.Value
    Dim exportName As String
    fileLocation = "X:\QA - Overpayment Recovery\QA - Overpayment Recovery(DeptUsers)\HMS"
    
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select a location to export file"
        .InitialFileName = fileLocation
        .Filters.Clear
        If .Show = 0 Then
            End
        Else
            filePath = .SelectedItems(1)
        End If
    End With

    
    DoCmd.SetWarnings False
    
    exportName = Replace(fileName, "HMSReprice_Complex", "HMSReturnReprice_Complex")
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryRepriceReturnFile", filePath & "\" & exportName, True
    DoCmd.RunSQL ("UPDATE tblFileLogReprice SET return_file_sent = True, return_file_sent_date = '" & Now() & "' WHERE file_name = '" & fileName & "';")
    Call UpdateSubmissionLog(userID, "xxxxxxxxxxx", cboFileName.Value, 7, lobid)
    DoCmd.SetWarnings True
    
    Call RefreshData
    MsgBox ("Export complete")
    
    
'err1:
'    Select Case Err.Number
'        Case 0
'        Case Else
'            Call LogError(Err.Number & " " & Err.Description, "frmLandingReprice; btnRepriceReturnFile_Click()")
'            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
'            Exit Sub
'    End Select
End Sub
Private Sub btnReturnToMainMenu_Click()
    DoCmd.Close acForm, "frmLandingReprice", acSaveNo
    DoCmd.OpenForm "frmLandingDRG", acNormal, , , , acWindowNormal
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
            intComp = InStr(fileName, "Reprice")
            If intComp = 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is not a <b>Repricing File</b>.</li>"
                ctlName = "btnImportRepriceFile,"
            End If
'            intComp = InStr(fileName, "DRG")
'            If intComp = 0 Then
'                msgStr = msgStr & "<li>Cannot import: the selected file is not <b>DRG</b>.</li>"
'                ctlName = "btnImportRepriceFile,"
'            End If
            
            intComp = InStr(fileName, "Return")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is a <b>Return File</b>.</li>"
                ctlName = "btnImportRepriceFile,"
            End If
            intComp = InStr(fileName, "Adjust")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportRepriceFile,"
            End If
            intComp = InStr(fileName, "Eligible")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportRepriceFile,"
            End If
            If DCount("file_name", "tblFileLogReprice", "file_name = '" & fileName & "'") > 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected <b>Reprice File</b> has already been imported.</li>"
                ctlName = "btnImportRepriceFile,"
            End If
            
        Case 2 'Return file report
            If cboFileName.Value = "" Or IsNull(cboFileName.Value) = True Then
                msgStr = msgStr & "<li>You must select a <b>Reprice File</b> before exporting a report. </li>"
                ctlName = "cboFileName,"
            End If
            If DCount("Claim_Number", "tblReprice", "[Audit Status] = 1 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                msgStr = msgStr & "<li>You must review all <b>New Records</b> on the selected Reprice File before generating a Return File.</li>"
                ctlName = "cboFileName,"
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

Public Sub RefreshData()
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varStatus As Variant
    Dim total As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT import_date, file_name, return_file_sent, return_file_sent_date " & _
        " FROM tblFileLogReprice " & _
        " WHERE ([file_name] = '" & cboFileName.Value & "');", dbOpenSnapshot)
        
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            txtImportDate.Value = ![import_date]
            txtReturnFileSent.Value = ![return_file_sent]
            txtReturnFileSentDate.Value = ![return_file_sent_date]
        End If
    End With
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT [Status Desc] AS Status, SUM(IIF([Audit Status] IS NULL, 0, 1)) AS [Nbr of Claims] " & _
        " FROM REFAuditStatus LEFT JOIN (SELECT tblReprice.[Audit Status] " & _
        " FROM tblReprice " & _
        " WHERE [File Name] = '" & cboFileName.Value & "')  AS sub ON REFAuditStatus.ID = sub.[Audit Status] " & _
        " WHERE REFAuditStatus.[ID] IN (1, 2, 8)  " & _
        " GROUP BY [status desc], [Process Order] " & _
        " ORDER BY [Process Order];", dbOpenSnapshot)
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            varStatus = .GetRows(.RecordCount)
            txtNewRecord.Value = varStatus(1, 0)
            txtNotEligible.Value = varStatus(1, 1)
            txtRepricingCompleted.Value = varStatus(1, 2)
            For i = 0 To UBound(varStatus, 2)
                total = total + varStatus(1, i)
            Next i
            txtTotal.Value = total
        End If

    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing

    DoCmd.Requery ("cboFileName")

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingReprice; RefreshData()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Public Sub ClearForm()
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(255, 255, 255)
    txtImportDate.Value = ""
    txtReturnFileSent.Value = ""
    txtReturnFileSentDate.Value = ""
    lstStatusCount.RowSource = ""
End Sub
