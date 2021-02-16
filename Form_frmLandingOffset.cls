VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLandingOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database





Private Sub btnImportOffsetFile_Click()
'must have reference "Microsoft Office 2016 Object Library" selected to use FileDialog object
    'go to tools -> references -> Microsoft Office 2016 Object Library

    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String, fileName As String, fileType As String
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    Dim subDate As String: subDate = Now()
    Dim intComp As Integer
    Dim lobid As Long
    
On Error GoTo err1:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    'default folder location of complex file
    fileLocation = "\\acct.upmchs.net\hp\Quality Audit\Quality Audit(Custom)\HDI\Files From FTP"
    'fileLocation = "X:\QA - Tech Team\QA - Tech Team(DeptUsers)\Focused_Review\HMS File Drop\"

    'open file dialog window
    With fDialog
        .AllowMultiSelect = False
        .Title = "Pick an Offset File to import"
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
    If ValidateForm(3, fileName) = True Then
        If MsgBox("Error importing file. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

    'prompt user if they'd like it import file - check if file already imported
    If MsgBox("Do you want to import the selected Offset file?", vbQuestion + vbYesNo, "Import HMS Offset File") = vbYes Then
    Else
        Exit Sub
    End If
    
    DoCmd.SetWarnings False

    'import file from excel into table
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tblOffset", filePath, True
    
    'set [File Name] to name of complex file and [audit status] to 1
    DoCmd.RunSQL ("UPDATE [tblOffset] SET [File Name] = '" & fileName & "', [Audit Status] = 1, [lob_id] = " & lobid & " WHERE [File Name] IS NULL;")
    
    'check against all selection files
    DoCmd.RunSQL ("UPDATE tblOffset " & _
        " LEFT JOIN tblSelection ON tblOffset.claim_no = tblSelection.claim_no " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 1, " & _
        " [New Allowed Amt] = 0, " & _
        " [Adjusted_Status] = 'N', " & _
        " [Adjusted Amount] = 0, " & _
        " [Adjustment_Exception_Reason] = 99," & _
        " [Adjustment or Exception Comments] = 'Claim was not approved for review on the selection files'" & _
        " WHERE ([selection_status] = 'D' OR [selection_status] is null)" & _
        " AND [File Name] = '" & fileName & "';")
    
    'check if [paid_dt] is greater than 365
    DoCmd.RunSQL ("UPDATE tblOffset " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments]= 2, " & _
        " [New Allowed Amt] = 0, " & _
        " [Adjusted_Status] = 'N'," & _
        " [Adjusted Amount] = 0, " & _
        " [Adjustment_Exception_Reason] = 99," & _
        " [Adjustment or Exception Comments] = 'Claim paid date is greater than 365 days'" & _
        " WHERE DATEDIFF('d', DATESERIAL(LEFT([paid_dt], 4), MID([paid_dt], 5,2), RIGHT([paid_dt], 2)), FORMAT(NOW(), 'mm/dd/yyyy')) > 365 " & _
        " AND [File Name] = '" & fileName & "';")

    'check for provider exclusions - CCBH
    DoCmd.RunSQL ("UPDATE tblOffset " & _
        " INNER JOIN tblExclusionProvidersCCBH ON tblOffset.[Servicing Provider ID] = tblExclusionProvidersCCBH.vendor_id " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 3, " & _
        " [New Allowed Amt] = 0, " & _
        " [Adjusted_Status] = 'N'," & _
        " [Adjusted Amount] = 0, " & _
        " [Adjustment_Exception_Reason] = 99," & _
        " [Adjustment or Exception Comments] = 'Provider is on exclusion list'" & _
        "WHERE [File Name] = '" & fileName & "';")

    'check for provider exclusions - NonDRG
    DoCmd.RunSQL ("UPDATE tblOffset " & _
        " INNER JOIN tblExclusionProvidersNonDRG ON tblOffset.provider_name = tblExclusionProvidersNonDRG.provider_name " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 3, " & _
        " [New Allowed Amt] = 0, " & _
        " [Adjusted_Status] = 'N'," & _
        " [Adjusted Amount] = 0, " & _
        " [Adjustment_Exception_Reason] = 99," & _
        " [Adjustment or Exception Comments] = 'Provider is on exclusion list'" & _
        "WHERE [File Name] = '" & fileName & "' AND (tblExclusionProvidersNonDRG.lob_id = " & lobid & " OR tblExclusionProvidersNonDRG.lob_id = 4 );")

    'check for member exclusion
    DoCmd.RunSQL ("UPDATE tblOffset " & _
        " INNER JOIN tblExclusionMembers ON tblOffset.patient_id = tblExclusionMembers.member_id " & _
        " SET [Audit Status] = 2, " & _
        " [QA Comments] = 4, " & _
        " [New Allowed Amt] = 0, " & _
        " [Adjusted_Status] = 'N'," & _
        " [Adjusted Amount] = 0, " & _
        " [Adjustment_Exception_Reason] = 99," & _
        " [Adjustment or Exception Comments] = 'Member is on exclusion list'" & _
        " WHERE [File Name] = '" & fileName & "';")
        
    'check against repricing table for data such as [LOB], [Original Paid Amt], [New Allowed Amt], [Internal Notes], [UPMC_New_DRG]
    DoCmd.RunSQL ("Update tblOffset " & _
        " INNER JOIN tblReprice ON tblOffset.claim_no = tblReprice.Claim_Number " & _
        " OR tblOffset.x_ref_claim_no = tblReprice.x_ref_claim_no " & _
        " SET tblOffset.[LOB] = tblReprice.[LOB], " & _
        " tblOffset.[New Allowed Amt] = tblReprice.[New Allowed Amt], " & _
        " tblOffset.[Original Paid Amt] = tblReprice.[Original Paid Amt], " & _
        " tblOffset.[internal_notes] = tblReprice.[internal_notes], " & _
        " tblOffset.[UPMC_New_DRG] = tblReprice.[UPMC_New_DRG] " & _
        " WHERE tblOffset.[File Name] = '" & fileName & "';")
        
    Select Case lobid
    Case 2
    'check against ARCHIVE repricing table for data such as LOB, original paid amt and new allowed amt
    'this should eventually be removed once all repricing done in database
        DoCmd.RunSQL ("Update tblOffset " & _
             " INNER JOIN tblRepriceCommercialArchive ON tblOffset.claim_no = tblRepriceCommercialArchive.[Claim Number] " & _
             " OR tblOffset.x_ref_claim_no = tblRepriceCommercialArchive.[Claim Number] " & _
             " SET tblOffset.[LOB] = tblRepriceCommercialArchive.[LOB], " & _
             " tblOffset.[New Allowed Amt] = tblRepriceCommercialArchive.[New Allowed Amount], " & _
             " tblOffset.[Original Paid Amt] = tblRepriceCommercialArchive.[Orig Paid Amnt], " & _
             " tblOffset.[internal_notes] = tblRepriceCommercialArchive.[QA Comments] " & _
             " WHERE tblOffset.[File Name] = '" & fileName & "';")
    Case 3
        'check against ARCHIVE repricing table for data such as [LOB], [Original Paid Amt], [New Allowed Amt]
        DoCmd.RunSQL ("UPDATE tblOffset INNER JOIN tblRepriceMedicaidArchive ON tblOffset.claim_no = tblRepriceMedicaidArchive.[Claim Number] " & _
            " OR tblOffset.x_ref_claim_no = tblRepriceMedicaidArchive.[Claim Number] " & _
            " SET tblOffset.LOB = tblRepriceMedicaidArchive.[LOB], " & _
            " tblOffset.[New Allowed Amt] = tblRepriceMedicaidArchive.[New Allowed Amount], " & _
            " tblOffset.[Original Paid Amt] = tblRepriceMedicaidArchive.[Orig Paid Amnt], " & _
            " tblOffset.[internal_notes] = tblRepriceMedicaidArchive.[QA Comments] " & _
            " WHERE tblOffset.[File Name] = '" & fileName & "';")
    End Select
        
        
    'update file log
    DoCmd.RunSQL ("INSERT INTO tblFileLogOffset (user_id, import_date, file_name, lob_id) " & _
        "VALUES ('" & userID & "', '" & subDate & "', '" & fileName & "', " & lobid & "); ")

    'update sub log
    Call UpdateSubmissionLog(userID, "xxxxxxxxxxx", fileName, 6, lobid)
    DoCmd.SetWarnings True
    Call RefreshData
    
    MsgBox ("Import Complete.")
        
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingOffset; btnImportOffsetFile_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

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
        cboFileName.RowSource = "SELECT DISTINCT file_name FROM tblFileLogOffset WHERE lob_id = " & optLOB.Value & " AND return_file_sent = " & optFileStatus.Value & " ORDER BY file_name;"
        cboFileName.Value = Null
        cboFileName.Requery
    End If
    Call RefreshData
End Sub

Private Sub btnComplexReturnFile_Click()
    'naming convention
    '5 - 202006100500_UPMC_53_Medicare_HMSClaims_ComplexReturnClaimFile
    
    If ValidateForm(2, "") = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

On Error GoTo err1:
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim fileName As String: fileName = cboFileName.Value
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    fileLocation = "X:\QA - Overpayment Recovery\QA - Overpayment Recovery(DeptUsers)\HMS\"
    Dim lobid As Long: lobid = optLOB.Value
    
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
    DoCmd.RunSQL ("UPDATE tblOffset SET [Adjusted Date] = IIF([Adjusted_status] = 'Y', DATE(), '') WHERE [File Name] = '" & fileName & "';")
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryComplexReturnFile", filePath & "\" & fileName & "ReturnClaimFile", True
    DoCmd.RunSQL ("UPDATE tblFileLogOffset SET return_file_sent = True, return_file_sent_date = '" & Now() & "' WHERE file_name = '" & fileName & "';")
    Call UpdateSubmissionLog(userID, "xxxxxxxxxxx", fileName, 8, lobid)
    DoCmd.SetWarnings True
    
    Call RefreshData
    MsgBox ("Export complete")
    
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingOffset; btnComplexReturnFile_Click")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
  
End Sub

Private Sub btnOpenPricingForm_Click()
    Dim args As String
    If cboFileName.Value <> "" Then
        args = cboFileName.Value
        args = args & "," & optLOB.Value & "," & optFileStatus.Value
    End If
    DoCmd.OpenForm "frmWorksheetOffset", acNormal, , , , acFormPropertySettings, args
End Sub
Private Sub cboFileName_Change()
    Call ClearForm
    Call RefreshData
End Sub

Private Sub btnAdjustmentRequestReport_Click()

    If ValidateForm(1, "") = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

On Error GoTo err1:
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim fileName As String: fileName = cboFileName.Value
    fileLocation = "X:\QA - Overpayment Recovery\QA - Overpayment Recovery(DeptUsers)\HMS\"
    
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
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryAdjustmentRequestReport", filePath & "\" & fileName & "_Adjustments", True
    DoCmd.RunSQL ("UPDATE tblOffset SET [Adjustment Requested] = True, " & _
        "[Audit Status] = 5 " & _
        " WHERE [Audit Status] = 3 AND [File Name] = '" & fileName & "';")
    DoCmd.RunSQL ("UPDATE tblFileLogOffset SET adjustment_sent = True, adjustment_sent_date = '" & Now() & "' WHERE file_name = '" & fileName & "';")
    DoCmd.SetWarnings True
    
    Call RefreshData
    MsgBox ("Export complete")
    
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingOffset; btnAdjustmentRequestReport_Click")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub btnAdjustmentAddReviews_Click()
    If ValidateForm(1, "") = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

On Error GoTo err1:
    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String
    Dim fileName As String: fileName = cboFileName.Value
    fileLocation = "X:\QA - Overpayment Recovery\QA - Overpayment Recovery(DeptUsers)\HMS\"
    
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
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryAdjustmentErrorsReport", filePath & "\" & fileName & "_AdjustmentErrors", True
    DoCmd.SetWarnings True
    
    Call RefreshData
    MsgBox ("Export complete")
    
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingOffset; btnAdjustmentAddReviews_Click")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub


Private Sub btnReturnToMainMenu_Click()
    DoCmd.Close acForm, "frmLandingOffset", acSaveNo
    DoCmd.OpenForm "frmLandingDRG", acNormal, , , , acWindowNormal
End Sub
Public Sub ClearForm()
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(255, 255, 255)
    txtImportDate.Value = ""
    txtAdjustmentSent.Value = ""
    txtAdjustmentSentDate.Value = ""
    txtReturnFileSent.Value = ""
    txtReturnFileSentDate.Value = ""
    'lstStatusCount.RowSource = ""
End Sub
Public Sub RefreshData()
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varStatus As Variant
    Dim total As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT import_date, file_name, adjustment_sent, adjustment_sent_date, return_file_sent, return_file_sent_date " & _
        " FROM tblFileLogOffset " & _
        " WHERE ([file_name] = '" & cboFileName.Value & "');", dbOpenSnapshot)
        
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            txtImportDate.Value = ![import_date]
            txtAdjustmentSent.Value = ![adjustment_sent]
            txtAdjustmentSentDate.Value = ![adjustment_sent_date]
            txtReturnFileSent.Value = ![return_file_sent]
            txtReturnFileSentDate.Value = ![return_file_sent_date]
        End If
    End With
    
    rs.Close

    Set rs = db.OpenRecordset("SELECT [Status Desc] AS Status, SUM(IIF([Audit Status] IS NULL, 0, 1)) AS [Nbr of Claims] " & _
        " FROM REFAuditStatus LEFT JOIN (SELECT tblOffset.[Audit Status] " & _
        " FROM tblOffset " & _
        " WHERE [File Name] = '" & cboFileName.Value & "')  AS sub ON REFAuditStatus.ID = sub.[Audit Status] " & _
        " WHERE REFAuditStatus.[ID] <> 8  " & _
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
            txtEligibleNoAdjustmentNeeded.Value = varStatus(1, 2)
            txtAdjustmentNeeded.Value = varStatus(1, 3)
            txtPendingAdjustmentVerification.Value = varStatus(1, 4)
            txtAdjustmentCompletedWithErrors.Value = varStatus(1, 5)
            txtAdjustmentCompleted.Value = varStatus(1, 6)
            txtAppealedOverturned = varStatus(1, 7)
            
            For i = 0 To UBound(varStatus, 2)
                total = total + varStatus(1, i)
                'Debug.Print varStatus(1, i)
            Next i
            txtTotal.Value = total
'            Do While Not .EOF
'                lstStatusCount.AddItem .Fields(0).Value & "," & .Fields(1).Value
'                lstStatusCount.GridlineColor = vbBlack
'                lstStatusCount.GridlineStyleLeft = Solid
'                .MoveNext
'            Loop
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

Public Function ValidateForm(submitType As Integer, fileName As String) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1 'Adjustment report
            If cboFileName.Value = "" Or IsNull(cboFileName.Value) = True Then
                msgStr = msgStr & "<li>You must select an <b>Offset File</b> before exporting a report. </li>"
                ctlName = "cboFileName,"
            Else
                If DCount("claim_no", "tblOffset", "[Audit Status] = 1 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                    msgStr = msgStr & "<li>You must review all <b>New Records</b> on the Offset File before requesting adjustment.</li>"
                    ctlName = "cboFileName,"
                End If
            End If
        Case 2 'Return file report
            If cboFileName.Value = "" Or IsNull(cboFileName.Value) = True Then
                msgStr = msgStr & "<li>You must select a <b>Offset File</b> before exporting a report. </li>"
                ctlName = "cboFileName,"
            End If
            If DCount("claim_no", "tblOffset", "[Audit Status] = 1 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                msgStr = msgStr & "<li>You must review all <b>New Records</b> on the selected Offset File before generating a Return File.</li>"
                ctlName = "cboFileName,"
            End If
            If DCount("claim_no", "tblOffset", "[Audit Status] = 3 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                msgStr = msgStr & "<li>You must review all <b>Adjustment Needed</b> records on the selected Offset File before generating a Return File.</li>"
                ctlName = "cboFileName,"
            End If
            If DCount("claim_no", "tblOffset", "[Audit Status] = 5 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                msgStr = msgStr & "<li>You must review all <b>Pending Adjustment Verification</b> records on the selected Offset File before generating a Return File.</li>"
                ctlName = "cboFileName,"
            End If
            If DCount("claim_no", "tblOffset", "[Audit Status] = 7 AND [File Name] = '" & cboFileName.Value & "'") > 0 Then
                msgStr = msgStr & "<li>You must review all <b>Adjustment Completed with Errors</b> records on the selected Offset File before generating a Return File.</li>"
                ctlName = "cboFileName,"
            End If
        Case 3 'Import file validation
            'StrComp() = 0 means exact match
            'InStr() = 0 means not found in sring
            intComp = InStr(fileName, "Complex")
            If intComp = 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is not a <b>Complex File</b>.</li>"
                ctlName = "btnImportOffsetFile,"
            End If
'            intComp = InStr(fileName, "DRG")
'            If intComp = 0 Then
'                msgStr = msgStr & "<li>Cannot import: the selected file is not <b>DRG</b>.</li>"
'                ctlName = "btnImportOffsetFile,"
'            End If
            
            intComp = InStr(fileName, "Return")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is a <b>Return File</b>.</li>"
                ctlName = "btnImportOffsetFile,"
            End If
            intComp = InStr(fileName, "Adjust")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportOffsetFile,"
            End If
            intComp = InStr(fileName, "Eligible")
            If intComp <> 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected file is an <b>Adjustment File</b>.</li>"
                ctlName = "btnImportOffsetFile,"
            End If
            If DCount("file_name", "tblFileLogOffset", "file_name = '" & fileName & "'") > 0 Then
                msgStr = msgStr & "<li>Cannot import: the selected <b>Offset File</b> has already been imported.</li>"
                ctlName = "btnImportOffsetFile,"
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

