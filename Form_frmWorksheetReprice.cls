VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmWorksheetReprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varReprice As Variant
Dim recordMax As Long, recordCounter As Long
Dim userID As Integer
Dim bolChanged As Boolean



Private Sub Form_Load()
    Dim splitArgs As Variant
    Call ClearForm
    Call SetListeners
    
    If Not IsNull(Forms![frmWorksheetReprice].OpenArgs) Then
        args = Forms![frmWorksheetReprice].OpenArgs
        splitArgs = Split(args, ",")
        cboFileName.Value = splitArgs(0)
        optLOB.Value = splitArgs(1)
        optFileStatus.Value = splitArgs(2)
        cboFileName.RowSource = "SELECT DISTINCT file_name FROM tblFileLogReprice WHERE lob_id = " & optLOB.Value & " AND return_file_sent = " & optFileStatus.Value & " ORDER BY file_name;"
        Call btnSearch_Click
    Else
        cboFileName.RowSource = ""
    End If
End Sub
Private Sub btnClear_Click()
    Call ClearForm
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
End Sub
Private Sub btnSearch_Click()
    If ValidateForm(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Call ClearForm
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblReprice WHERE [File Name] = '" & cboFileName.Value & "' ORDER BY [Provider_Name];", dbOpenSnapshot)
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            recordMax = .RecordCount - 1
            recordCounter = 0
            varReprice = .GetRows(.RecordCount)
            
            Call PopulateBoxes(recordCounter)
            Call RefreshData
        End If
    End With

    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetReprice; btnSearch_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Public Sub PopulateBoxes(i As Long)

On Error GoTo err1:
    txtPatientID.Value = varReprice(0, i)
    txtClaimNumber.Value = varReprice(1, i)
    txtServicingProviderID.Value = varReprice(3, i)
    txtProviderName.Value = varReprice(4, i)
    txtBillFromDate.Value = varReprice(5, i)
    txtBillThruDate.Value = varReprice(6, i)
    txtPaidDate.Value = varReprice(7, i)
    txtOverpaymentReasonDescription.Value = varReprice(8, i)
    txtNewDischrgStatus.Value = varReprice(9, i)
    txtOldDRG.Value = varReprice(10, i)
    txtHMSNewDRG.Value = varReprice(11, i)
    txtNewPrimDiag.Value = varReprice(12, i)
    txtNewDiagCode2.Value = varReprice(14, i)
    txtNewDiagCode3.Value = varReprice(15, i)
    txtNewDiagCode4.Value = varReprice(16, i)
    txtNewDiagCode5.Value = varReprice(17, i)
    txtNewDiagCode6.Value = varReprice(18, i)
    txtNewDiagCode7.Value = varReprice(19, i)
    txtNewDiagCode8.Value = varReprice(20, i)
    txtNewDiagCode9.Value = varReprice(21, i)
    txtNewDiagCode10.Value = varReprice(22, i)
    txtNewProcCd1.Value = varReprice(55, i)
    txtNewProcCd2.Value = varReprice(56, i)
    txtNewProcCd3.Value = varReprice(57, i)
    txtNewProcCd4.Value = varReprice(58, i)
    txtNewProcCd5.Value = varReprice(59, i)
    txtAltClaimNbr.Value = varReprice(69, i)
    txtOverpaymentResponse.Value = varReprice(71, i)
    txtOverpaymentExceptionReason.Value = varReprice(72, i)
    txtOverpaymentExceptionComments.Value = varReprice(73, i)
    txtCalculatedSavings.Value = varReprice(74, i)
    txtUPMCNewDRG.Value = varReprice(75, i)
    cboLOB.Value = varReprice(77, i)
    cboQAComments.Value = varReprice(78, i)
    txtNewAllowedAmt.Value = varReprice(79, i)
    txtOriginalPaidAmt.Value = varReprice(80, i)
    cboAuditStatus.Value = varReprice(81, i)
    If varReprice(81, i) = 2 Then
        Activate cboQAComments
    Else
        Deactivate cboQAComments
    End If
    txtLastUpdate.Value = varReprice(82, i)
    If Not IsNull(varReprice(83, i)) Then
        txtLastUpdateUser.Value = DLookup("user_name", "tblUsers", "user_id = " & varReprice(83, i))
    Else
        txtLastUpdateUser.Value = Null
    End If
    If Not IsNull(varReprice(84, i)) Then: txtLOBID.Value = DLookup("lob", "REFLob", "lob_id = " & varReprice(84, i))
    'UPDATE LOB COMBOBOX ACCORDING TO LOB
    If optLOB.Value = 2 Then
        cboLOB.RowSource = "ASO;FI;FI Individual"
    ElseIf optLOB.Value = 3 Then
        cboLOB.RowSource = "Medicaid;CHC Southwest;CHC Southeast;CHC Northwest;CHC Northeast;CHC Lehigh Cap"
    Else
        cboLOB.RowSource = ""
    End If
    cboLOB.Requery
    txtInternalNotes.Value = varReprice(85, i)
    txtID.Value = varReprice(86, i)
    
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetReprice; PopulateBoxes()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub btnNextRecord_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        Call PopulateBoxes(recordCounter)
    End If
End Sub
Private Sub btnPrevRecord_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        Call PopulateBoxes(recordCounter)
    End If
End Sub
Private Sub btnClose_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    DoCmd.Close acForm, "frmWorksheetReprice", acSaveNo
End Sub
Private Sub cboFileName_Change()
    Call ClearForm
    DoCmd.Requery ("cboFileName")
End Sub


Private Sub btnSave_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Cancel = True
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    
    Dim subDate As String: subDate = Now()
    Dim networkId As String: networkId = Environ$("username")
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & networkId & "'")
    Dim lobid As Long: lobid = optLOB.Value
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblReprice", dbOpenDynaset)
    
    With rs
        .FindFirst ("[ID] = " & txtID.Value)
        If Not .NoMatch Then
            'UPDATE TABLE
            .Edit
                ![x_ref_claim_no] = txtAltClaimNbr
                ![Overpayment_Response] = txtOverpaymentResponse
                ![Overpayment_Exception_Reason] = txtOverpaymentExceptionReason
                ![Overpayment_Exception_Comments] = txtOverpaymentExceptionComments
                ![Calculated_Savings] = txtCalculatedSavings
                ![UPMC_New_DRG] = txtUPMCNewDRG
                ![LOB] = cboLOB
                ![QA Comments] = cboQAComments
                ![New Allowed Amt] = txtNewAllowedAmt
                ![Original Paid Amt] = txtOriginalPaidAmt
                ![Audit Status] = cboAuditStatus
                ![Last Update] = subDate
                ![Last Update User] = userID
                ![internal_notes] = txtInternalNotes
            .Update
            .Close
            
            'UPDATE LOCAL ARRAY
            varReprice(69, recordCounter) = txtAltClaimNbr
            varReprice(71, recordCounter) = txtOverpaymentResponse
            varReprice(72, recordCounter) = txtOverpaymentExceptionReason
            varReprice(73, recordCounter) = txtOverpaymentExceptionComments
            varReprice(74, recordCounter) = txtCalculatedSavings
            varReprice(75, recordCounter) = txtUPMCNewDRG
            varReprice(77, recordCounter) = cboLOB
            varReprice(78, recordCounter) = cboQAComments
            varReprice(79, recordCounter) = txtNewAllowedAmt
            varReprice(80, recordCounter) = txtOriginalPaidAmt
            varReprice(81, recordCounter) = cboAuditStatus
            varReprice(82, recordCounter) = subDate
            varReprice(83, recordCounter) = userID
            varReprice(85, recordCounter) = txtInternalNotes
            
            'UPDATE FORM
            txtLastUpdate.Value = subDate
            txtLastUpdateUser.Value = DLookup("user_name", "tblUsers", "user_id = " & userID)
            
            bolChanged = False
            Call UpdateSubmissionLog(userID, txtClaimNumber.Value, cboFileName.Value, 2, lobid)
            Call RefreshData
            MsgBox ("Item Saved")
        Else
            MsgBox ("Cannot submit")
        End If
    End With
    
    Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetReprice; btnSave_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub cboAuditStatus_Click()
    
    If cboAuditStatus.Value = 1 Then
        Deactivate cboQAComments
        
        txtOverpaymentResponse.Value = Null
        txtOverpaymentExceptionReason.Value = Null
        txtOverpaymentExceptionComments.Value = Null
    ElseIf cboAuditStatus.Value = 2 Then
        Activate cboQAComments

        txtOverpaymentResponse.Value = "N"
        cboQAComments.Value = Null

    ElseIf cboAuditStatus.Value = 8 Then
        Deactivate cboQAComments
        
        txtOverpaymentResponse.Value = "Y"
        txtOverpaymentExceptionReason.Value = Null
        txtOverpaymentExceptionComments.Value = Null

        Call CalculateSavings
        
    End If
End Sub

Private Sub txtNewAllowedAmt_Exit(Cancel As Integer)
    'New Allowed Amt
    Call CalculateSavings
End Sub
Private Sub txtOriginalPaidAmt_Exit(Cancel As Integer)
    'Original Paid Amt
    Call CalculateSavings
End Sub
Private Sub cboQAComments_Change()
    txtOverpaymentResponse.Value = "N"
    txtCalculatedSavings.Value = Null
    txtOverpaymentExceptionReason.Value = cboQAComments.Column(2)
    txtOverpaymentExceptionComments.Value = cboQAComments.Column(1)
    txtNewAllowedAmt.Value = 0
    txtCalculatedSavings.Value = 0
End Sub


Public Sub CalculateSavings()
    If txtOriginalPaidAmt.Value > txtNewAllowedAmt.Value Then
        txtCalculatedSavings.Value = txtOriginalPaidAmt.Value - txtNewAllowedAmt.Value
    Else
        txtCalculatedSavings.Value = 0
    End If
End Sub





Private Sub SetListeners()
    cboAuditStatus.AfterUpdate = "=MarkChanged()"
    cboQAComments.AfterUpdate = "=MarkChanged()"
    txtOriginalPaidAmt.AfterUpdate = "=MarkChanged()"
    txtNewAllowedAmt.AfterUpdate = "=MarkChanged()"
    txtCalculatedSavings.AfterUpdate = "=MarkChanged()"
    cboLOB.AfterUpdate = "=MarkChanged()"
    txtAltClaimNbr.AfterUpdate = "=MarkChanged()"
    txtUPMCNewDRG.AfterUpdate = "=MarkChanged()"
    txtInternalNotes.AfterUpdate = "=MarkChanged()"
End Sub
Private Function CheckUnsavedChanges() As Boolean
    If bolChanged = True Then
        If MsgBox("You have unsaved changes." & vbCrLf & vbCrLf & "Are you sure you want to leave this record?", vbYesNo + vbQuestion, "Unsaved Changes") = vbYes Then
            CheckUnsavedChanges = True
            bolChanged = False
        Else
            CheckUnsavedChanges = False
        End If
    Else
        CheckUnsavedChanges = True
    End If
End Function
Private Function MarkChanged()
    bolChanged = True
End Function
Public Sub ClearForm()
    If Not IsEmpty(varReprice) Then: Set varReprice = Nothing
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(217, 217, 217)
    Deactivate cboQAComments
    cboLOB.RowSource = ""
    recordMax = 0
    bolChanged = False
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "cboFileName" Then
                ctrl.Value = Null
                ctrl.Enabled = True
            End If
        Next
    End With
End Sub
Sub Activate(ctrl As Control)
    With ctrl
        .Enabled = True
        '.Locked = False
        .Visible = True
        '.Value = Null
    End With
End Sub
Sub Deactivate(ctrl As Control)
    With ctrl
        .Enabled = False
        '.Locked = True
        .Visible = False
        '.Value = Null
    End With
End Sub
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant

    headerStr = "<ul>"
    footerStr = "</ul>"

    Select Case submitType
    
        Case 1 'submit
            If IsNull(txtClaimNumber) Or txtClaimNumber.Value = "" Then
                msgStr = msgStr & "<li>You must select a claim record before submitting information.</li>"
                ctlName = ctlName & "cboFileName,"
            End If
            
            Select Case cboAuditStatus.Value
                Case 1 'new record
                    If cboAuditStatus.Value = 1 Then
                        msgStr = msgStr & "<li><b>Audit Status</b> must be updated.</li>"
                        ctlName = ctlName & "cboAuditStatus,"
                    End If
                    
                Case 2 'Not eligible for recovery
                    If txtOriginalPaidAmt.Value = "" Or IsNull(txtOriginalPaidAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>Original Paid Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtOriginalPaidAmt,"
                    End If
                    If txtNewAllowedAmt.Value <> 0 Or IsNull(txtNewAllowedAmt) = True Then
                        msgStr = msgStr & "<li><b>New Allowed Amt</b> must be set to $0.00.</li>"
                        ctlName = ctlName & "txtNewAllowedAmt,"
                    End If
                        If txtCalculatedSavings.Value <> 0 Or IsNull(txtCalculatedSavings) = True Then
                        msgStr = msgStr & "<li><b>Calculated Savings</b> must be set to $0.00.</li>"
                        ctlName = ctlName & "txtCalculatedSavings,"
                    End If
                    If cboLOB.Value = "" Or IsNull(cboLOB.Value) = True Then
                        msgStr = msgStr & "<li><b>LOB</b> cannot be blank.</li>"
                        ctlName = ctlName & "cboLOB,"
                    End If
                    If IsNull(cboQAComments.Value) = True Or cboQAComments.Value = "" Then
                        msgStr = msgStr & "<li><b>Not Eligible Reason</b> cannot be blank.</li>"
                        ctlName = ctlName & "cboQAComments,"
                    End If
                
                Case 8 'Repricing Completed
                    If txtOriginalPaidAmt.Value = "" Or IsNull(txtOriginalPaidAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>Original Paid Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtOriginalPaidAmt,"
                    End If
                    If txtNewAllowedAmt.Value = "" Or IsNull(txtNewAllowedAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>New Allowed Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtNewAllowedAmt,"
                    End If
                    If cboLOB.Value = "" Or IsNull(cboLOB.Value) = True Then
                        msgStr = msgStr & "<li><b>LOB</b> cannot be blank.</li>"
                        ctlName = ctlName & "cboLOB,"
                    End If
            End Select
            
        Case 2 'search
            If optLOB.Value = "" Or IsNull(optLOB.Value) = True Then
                msgStr = msgStr & "<li>You must select a <b>Product</b>. </li>"
                ctlName = "optLOB,"
            End If
            If optFileStatus.Value = "" Or IsNull(optFileStatus.Value) = True Then
                msgStr = msgStr & "<li>You must select a <b>File Status</b>. </li>"
                ctlName = "optFileStatus,"
            End If
            If cboFileName.Value = "" Or IsNull(cboFileName.Value) = True Then
                msgStr = msgStr & "<li><b>File Name</b> cannot be blank.</li>"
                ctlName = ctlName & "cboFileName,"
            End If
        
    End Select

    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(217, 217, 217)
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

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmLandingReprice; RefreshData()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub


