VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmWorksheetOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varOffset As Variant
Dim recordMax As Long, recordCounter As Long
Dim userID As Integer
Dim bolChanged As Boolean

Private Sub btnImportMC400Financials_Click()
    'module requires reference to Microsoft ActiveX Data Objects (ADO)
    'Tools -> References -> Microsoft ActiveX Data Objects 6.1 Library
    'Tools -> References -> Microsoft ActiveX Data Objects REcordset 6.0 Library
    'use ODBC to connect to DW tables using Tech Team credentials
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cn As New ADODB.Connection
    Dim rs2 As New ADODB.Recordset
    Dim sql As String, sql2 As String, criteria As String, strCon As String
    Dim fileName As String

    If ValidateForm(3) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

On Error GoTo err1:
    DoCmd.SetWarnings False
    fileName = cboFileName.Value
    
    Select Case optImportFinancials.Value
        Case 1 'batch
            sql = "SELECT IIF((x_ref_claim_no IS NOT NULL AND x_ref_claim_no <> ''), x_ref_claim_no, [claim_no]) AS claim_nbr FROM tblOffset WHERE [File Name] = '" & fileName & "';"
            
            Set db = CurrentDb
            Set rs = db.OpenRecordset(sql)
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do Until rs.EOF = True
                        criteria = criteria & "'" & ![claim_nbr] & "',"
                        'Debug.Print ![claim_nbr]
                        .MoveNext
                    Loop
                Else
                    Exit Sub
                End If
            End With
            rs.Close: Set rs = Nothing
            db.Close: Set db = Nothing
            criteria = Left(criteria, Len(criteria) - 1)
            
            'delete existing financial info, if any
            DoCmd.RunSQL ("DELETE * FROM tblClaimFinancials WHERE FileName = '" & fileName & "';")
            
        Case 2 'curr record
            If Not IsNull(txtAltClaimNbr) And txtAltClaimNbr.Value <> "" Then
                criteria = "'" & txtAltClaimNbr.Value & "'"
            Else
                criteria = "'" & txtClaimNo.Value & "'"
            End If

            'delete existing financial info, if any
            DoCmd.RunSQL ("DELETE * FROM tblClaimFinancials WHERE FileName = '" & fileName & "' AND (ClaimNbr = " & criteria & " OR AltClaimNbr = " & criteria & ");")
    End Select
    
    'prep db objects for local table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblClaimFinancials", dbOpenDynaset)
    
    'prep connection to Data Warehouse tables
    sql2 = "SELECT Sum(AllowedAmt) AS SumOfAllowedAmt, Sum(NetAmt) AS SumOfNetAmt, GetDate() AS CurrDate, ClaimNbr FROM ClaimsDetail WHERE ClaimNbr IN (" & criteria & ") GROUP BY ClaimNbr ORDER BY ClaimNbr;"
    Set cn = New ADODB.Connection
    strCon = "Driver={ODBC};" & _
        "Provider=SQLOLEDB;" & _
        "Data Source=HEALTHPLANDW2;" & _
        "Database=dwOAO;" & _
        "Uid=QATechTeam;Pwd=qattpw_1;"
    cn.Open (strCon)
    rs2.Open sql2, cn, adOpenKeyset
    
    'update local tblClaimsFinancial table with results of Data Warehouse query
    With rs2
        .MoveFirst
        .MoveLast
        .MoveFirst
        Do While Not .EOF
            rs.AddNew
                rs.Fields(3) = .Fields(0)
                rs.Fields(4) = .Fields(1)
                rs.Fields(5) = .Fields(2)
                rs.Fields(6) = .Fields(3)
            rs.Update
            Results = Results & .Fields(0) & "," & .Fields(1) & "," & .Fields(2) & "," & .RecordCount
        .MoveNext
        Loop
    End With
    rs.Close: Set rs = Nothing
    rs2.Close: Set rs2 = Nothing
    cn.Close: Set cn = Nothing
    
    'pull in additional info from tblOffset
    DoCmd.RunSQL ("UPDATE tblClaimFinancials AS a" & _
        " INNER JOIN tblOffset AS b ON a.ClaimNbrDW = b.claim_no OR a.ClaimNbrDW = b.x_ref_claim_no " & _
        " SET a.FileName = b.[File Name], " & _
        " a.ClaimNbr = b.claim_no, " & _
        " a.AltClaimNbr = b.x_ref_claim_no " & _
        " WHERE b.[File Name] = '" & fileName & "';")

    DoCmd.SetWarnings True
    DoCmd.Requery "qryFinancialsMedicare subform"
    
    MsgBox ("Financial Import Complete")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetOffset; btnImportMC400Financials_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub


Private Sub Form_Load()
    Dim splitArgs As Variant
    Call ClearForm
    Call SetListeners
    If Not IsNull(Forms![frmWorksheetOffset].OpenArgs) Then
        args = Forms![frmWorksheetOffset].OpenArgs
        splitArgs = Split(args, ",")
        cboFileName.Value = splitArgs(0)
        optLOB.Value = splitArgs(1)
        optFileStatus.Value = splitArgs(2)
        cboFileName.RowSource = "SELECT DISTINCT file_name FROM tblFileLogOffset WHERE lob_id = " & optLOB.Value & " AND return_file_sent = " & optFileStatus.Value & " ORDER BY file_name;"
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
        cboFileName.RowSource = "SELECT DISTINCT file_name FROM tblFileLogOffset WHERE lob_id = " & optLOB.Value & " AND return_file_sent = " & optFileStatus.Value & " ORDER BY file_name;"
        cboFileName.Value = Null
        cboFileName.Requery
    End If
End Sub
Private Sub cboFileName_Change()
    'cboFileName name change
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
    'Set rs = db.OpenRecordset("SELECT * FROM tblOffset WHERE [File Name] = '" & cboFileName.Value & "' ORDER BY [Provider_Name];", dbOpenSnapshot)
    Set rs = db.OpenRecordset("SELECT tblOffset.*, adjustment_sent_date FROM tblOffset " & _
        " LEFT JOIN tblFileLogOffset ON tblOffset.[File Name] = tblFileLogOffset.file_name " & _
        " WHERE [File Name] = '" & cboFileName.Value & "' ORDER BY [Provider_Name];", dbOpenSnapshot)

    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            recordMax = .RecordCount - 1
            recordCounter = 0
            varOffset = .GetRows(.RecordCount)
            
            'populate LOB dropdown depending on selected product type
            If optLOB.Value = 1 Then
                cboLOB.RowSource = "Medicare;SNP"
            ElseIf optLOB.Value = 2 Then
                cboLOB.RowSource = "ASO;FI;FI Individual"
            ElseIf optLOB.Value = 3 Then
                cboLOB.RowSource = "Medicaid;CHC Southwest;CHC Southeast;CHC Northwest;CHC Northeast;CHC Lehigh Cap"
            Else
                cboLOB.RowSource = ""
            End If
            
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
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetOffset; btnSearch_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Public Sub PopulateBoxes(i As Long)

On Error GoTo err1:
    txtPatientID.Value = varOffset(0, i)
    txtClaimNo.Value = varOffset(1, i)
    txtServicingProviderID.Value = varOffset(2, i)
    txtProviderName.Value = varOffset(3, i)
    txtStmtCovFromDt.Value = varOffset(4, i)
    txtStmtCovToDt.Value = varOffset(5, i)
    txtPaidDt.Value = varOffset(6, i)
    txtOverpaymentReasonDescription.Value = varOffset(7, i)
    txtNewDischrgStatus.Value = varOffset(8, i)
    txtOldDrgNo.Value = varOffset(9, i)
    txtNewDrgNo.Value = varOffset(10, i)
    txtNewDiagCode1.Value = varOffset(13, i)
    txtNewDiagCode2.Value = varOffset(14, i)
    txtNewDiagCode3.Value = varOffset(15, i)
    txtNewDiagCode4.Value = varOffset(16, i)
    txtNewDiagCode5.Value = varOffset(17, i)
    txtNewDiagCode6.Value = varOffset(18, i)
    txtNewDiagCode7.Value = varOffset(19, i)
    txtNewDiagCode8.Value = varOffset(20, i)
    txtNewDiagCode9.Value = varOffset(21, i)
    txtNewDiagCode10.Value = varOffset(22, i)
    txtNewProcCd1.Value = varOffset(55, i)
    txtNewProcCd2.Value = varOffset(56, i)
    txtNewProcCd3.Value = varOffset(57, i)
    txtNewProcCd4.Value = varOffset(58, i)
    txtNewProcCd5.Value = varOffset(59, i)
    txtPotentialSavings.Value = varOffset(70, i)
    txtAltClaimNbr.Value = varOffset(71, i)
    cboAdjustedStatus.Value = varOffset(72, i)
    txtAdjustedAmount.Value = varOffset(74, i)
    cboAdjustedExceptionReason.Value = varOffset(75, i)
    
    txtAdjustmentExceptionReason.Value = varOffset(76, i)
    cboLOB.Value = varOffset(80, i)
    cboQAComments.Value = varOffset(81, i)
    txtNewAllowedAmt.Value = varOffset(82, i)
    txtOriginalPaidAmt.Value = varOffset(83, i)
    txtNewPaidAmt.Value = varOffset(84, i)
    txtAdjustmentRequestSent.Value = varOffset(85, i)
    cboAuditStatus.Value = varOffset(87, i)
    txtLastUpdate.Value = varOffset(88, i)
    If Not IsNull(varOffset(89, i)) Then
        txtLastUpdateUser.Value = DLookup("user_name", "tblUsers", "user_id = " & varOffset(89, i))
    Else
        txtLastUpdateUser.Value = Null
    End If
    txtInternalNotes.Value = varOffset(91, i)
    txtUPMCNewDRG.Value = varOffset(92, i)
    optOverturned.Value = varOffset(93, i)
    txtOverturnedAmt.Value = varOffset(94, i)
    txtOverturnedDate.Value = varOffset(95, i)
    txtID.Value = varOffset(96, i)
    If varOffset(85, i) = True Then
        txtAdjustmentSentDate.Value = varOffset(97, i)
    Else
        txtAdjustmentSentDate.Value = Null
    End If
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1
    
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(217, 217, 217)
    Call cboAuditStatus_BeforeUpdate(0)

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmWorksheetReprice; PopulateBoxes()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub cboAuditStatus_BeforeUpdate(Cancel As Integer)

    If cboAuditStatus.Value = 1 Then
        Deactivate cboQAComments
        Deactivate txtNewPaidAmt
        txtNewPaidAmt.Value = Null
        txtAdjustedAmount.Visible = False
        Call GetReturnInfo(1)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 2 Then
        Activate cboQAComments
        Deactivate txtNewPaidAmt
        txtAdjustedAmount.Visible = False
        Call GetReturnInfo(2)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 3 Then
        Deactivate cboQAComments
        Deactivate txtNewPaidAmt
        txtNewPaidAmt.Value = Null
        txtAdjustedAmount.Visible = False
        Call GetReturnInfo(1)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 4 Then
        Deactivate cboQAComments
        Activate txtNewPaidAmt
        txtAdjustedAmount.Visible = True
        Call GetReturnInfo(3)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 5 Then
        Deactivate cboQAComments
        Deactivate txtNewPaidAmt
        txtAdjustedAmount.Visible = False
        Call GetReturnInfo(1)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 6 Then
        Deactivate cboQAComments
        Activate txtNewPaidAmt
        txtAdjustedAmount.Visible = True
        Call GetReturnInfo(3)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 7 Then
        Deactivate cboQAComments
        Activate txtNewPaidAmt
        txtAdjustedAmount.Visible = False
        Call GetReturnInfo(1)
        ShowOverturnedInfo (False)
        
    ElseIf cboAuditStatus.Value = 9 Then
        Deactivate cboQAComments
        Activate txtNewPaidAmt
        txtAdjustedAmount.Visible = True
        ShowOverturnedInfo (True)
        'Call GetReturnInfo(1)
    End If
    
End Sub
Public Sub GetReturnInfo(returnOpt As Integer)
    Select Case returnOpt
    
        Case 1 'set null
            cboAdjustedStatus.Value = Null
            txtAdjustedAmount.Value = Null
            cboAdjustedExceptionReason.Value = Null
            txtAdjustmentExceptionReason.Value = Null
            
        Case 2 'not eligible
            cboAdjustedStatus.Value = "N"
            txtAdjustedAmount.Value = 0
            
        Case 3 'calculate recovery
            cboAdjustedStatus.Value = "Y"
            If txtOriginalPaidAmt.Value > txtNewPaidAmt.Value Then
                txtAdjustedAmount.Value = txtOriginalPaidAmt.Value - txtNewPaidAmt.Value
            Else
                txtAdjustedAmount.Value = 0
            End If
            cboAdjustedExceptionReason.Value = Null
            txtAdjustmentExceptionReason.Value = Null
   
    End Select
End Sub
Public Sub ShowOverturnedInfo(bolShow As Boolean)
    If bolShow = True Then
        lblOverturned.Visible = True
        Activate optOverturned
        Activate txtOverturnedAmt
        Activate txtOverturnedDate
    Else
        lblOverturned.Visible = False
        Deactivate optOverturned
        Deactivate txtOverturnedAmt
        Deactivate txtOverturnedDate
    End If
End Sub
Private Sub cboQAComments_Change()
    cboAdjustedExceptionReason.Value = cboQAComments.Column(2)
    txtAdjustmentExceptionReason.Value = cboQAComments.Column(1)
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
    Dim processId As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOffset", dbOpenDynaset)
    
    With rs
        .FindFirst ("[ID] = " & txtID.Value)
        If Not .NoMatch Then
            'UPDATE TABLE
            .Edit
                ![x_ref_claim_no] = txtAltClaimNbr
                ![Adjusted_status] = cboAdjustedStatus
                ![Adjusted Amount] = txtAdjustedAmount
                ![Adjustment_Exception_Reason] = cboAdjustedExceptionReason
                ![Adjustment or Exception Comments] = txtAdjustmentExceptionReason
                ![LOB] = cboLOB
                ![QA Comments] = cboQAComments
                ![New Allowed Amt] = txtNewAllowedAmt
                ![Original Paid Amt] = txtOriginalPaidAmt
                ![New Paid Amount] = txtNewPaidAmt
                '![Recovery] = txtRecovery
                ![Audit Status] = cboAuditStatus
                ![Last Update] = subDate
                ![Last Update User] = userID
                ![internal_notes] = txtInternalNotes
                ![UPMC_New_DRG] = txtUPMCNewDRG
                ![overturned] = optOverturned
                ![overturned_amt] = txtOverturnedAmt
                ![overturned_date] = txtOverturnedDate

            .Update
            .Close
            
            'UPDATE LOCAL ARRAY
            varOffset(71, recordCounter) = txtAltClaimNbr
            varOffset(72, recordCounter) = cboAdjustedStatus
            varOffset(74, recordCounter) = txtAdjustedAmount
            varOffset(75, recordCounter) = cboAdjustedExceptionReason
            varOffset(76, recordCounter) = txtAdjustmentExceptionReason
            varOffset(80, recordCounter) = cboLOB
            varOffset(81, recordCounter) = cboQAComments
            varOffset(82, recordCounter) = txtNewAllowedAmt
            varOffset(83, recordCounter) = txtOriginalPaidAmt
            varOffset(84, recordCounter) = txtNewPaidAmt
            'varOffset(86, recordCounter) = txtRecovery
            varOffset(87, recordCounter) = cboAuditStatus
            varOffset(88, recordCounter) = subDate
            varOffset(89, recordCounter) = userID
            varOffset(91, recordCounter) = txtInternalNotes
            varOffset(92, recordCounter) = txtUPMCNewDRG
            varOffset(93, recordCounter) = optOverturned
            varOffset(94, recordCounter) = txtOverturnedAmt
            varOffset(95, recordCounter) = txtOverturnedDate
            
            'UPDATE FORM
            txtLastUpdate.Value = subDate
            txtLastUpdateUser.Value = DLookup("user_name", "tblUsers", "user_id = " & userID)
            
            bolChanged = False
            
            Select Case lobid
                Case 1 'medicare
                    If cboAuditStatus.Value >= 2 And cboAuditStatus.Value <= 4 Then
                        processId = 1
                    Else
                        processId = 9
                    End If
                Case 2, 3 'comm/medicaid
                    If cboAuditStatus.Value >= 2 And cboAuditStatus.Value <= 4 Then
                        processId = 3
                    Else
                        processId = 9
                    End If
            End Select
            Call UpdateSubmissionLog(userID, txtClaimNo.Value, cboFileName.Value, processId, lobid)
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
            Call LogError(Err.Number & " " & Err.Description, "btnSave_Click(); btnSave_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub txtOriginalPaidAmt_Exit(Cancel As Integer)
    If cboAuditStatus.Value = 4 Or cboAuditStatus.Value = 6 Then
        Call GetReturnInfo(3)
    End If
End Sub
Private Sub txtNewPaidAmt_Exit(Cancel As Integer)
    If cboAuditStatus.Value = 4 Or cboAuditStatus.Value = 6 Then
        Call GetReturnInfo(3)
    End If
End Sub

Private Sub btnClose_Click()
    If CheckUnsavedChanges = False Then Exit Sub
    DoCmd.Close acForm, "frmWorksheetOffset", acSaveNo
End Sub
Private Sub SetListeners()
    cboAuditStatus.AfterUpdate = "=MarkChanged()"
    cboQAComments.AfterUpdate = "=MarkChanged()"
    txtOriginalPaidAmt.AfterUpdate = "=MarkChanged()"
    txtNewAllowedAmt.AfterUpdate = "=MarkChanged()"
    txtNewPaidAmt.AfterUpdate = "=MarkChanged()"
    'txtCalculatedSavings.AfterUpdate = "=MarkChanged()"
    'txtRecovery.AfterUpdate = "=MarkChanged()"
    cboLOB.AfterUpdate = "=MarkChanged()"
    txtAltClaimNbr.AfterUpdate = "=MarkChanged()"
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

Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant

    headerStr = "<ul>"
    footerStr = "</ul>"

    Select Case submitType
    
        Case 1 'submit
            If IsNull(txtClaimNo) Or txtClaimNo.Value = "" Then
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
                    If IsNull(cboQAComments.Value) = True Or cboQAComments.Value = "" Then
                        msgStr = msgStr & "<li><b>Not Eligible Reason</b> cannot be blank.</li>"
                        ctlName = ctlName & "cboQAComments,"
                    End If
                    If txtOriginalPaidAmt.Value = "" Or IsNull(txtOriginalPaidAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>Original Paid Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtOriginalPaidAmt,"
                    End If
                     If txtNewAllowedAmt.Value <> 0 Or IsNull(txtNewAllowedAmt) = True Then
                        msgStr = msgStr & "<li><b>New Allowed Amt</b> must be set to $0.00.</li>"
                        ctlName = ctlName & "txtNewAllowedAmt,"
                    End If
                    If cboLOB.Value = "" Or IsNull(cboLOB.Value) = True Then
                        msgStr = msgStr & "<li><b>LOB</b> cannot be blank.</li>"
                        ctlName = ctlName & "cboLOB,"
                    End If
                    
                Case 3 'Adjustment needed
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
                    
                Case 4 'Eligible for Recovery - No Adjustment Needed
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
                    If txtNewPaidAmt.Value = "" Or IsNull(txtNewPaidAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>New Paid Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtnewpaidamt,"
                    End If
                    If txtAdjustedAmount.Value = "" Or IsNull(txtAdjustedAmount.Value) = True Then
                        msgStr = msgStr & "<li><b>Recovery</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtadjustedamount,"
                    End If
                    
                Case 5 'Pending adjustment verification
                    If cboAuditStatus.Value = 5 Then
                        msgStr = msgStr & "<li><b>Audit Status</b> must be updated.</li>"
                        ctlName = ctlName & "cboAuditStatus,"
                    End If
                    
                Case 6 'Adjustment Completed
                    If txtAdjustmentRequestSent.Value = 0 Then
                        msgStr = msgStr & "<li>You must export the <b>Adjustment Request Report</b> before changing status to Adjustment Completed.</li>"
                        ctlName = ctlName & "cboAuditStatus,"
                    Else
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
                        If txtNewPaidAmt.Value = "" Or IsNull(txtNewPaidAmt.Value) = True Then
                            msgStr = msgStr & "<li><b>New Paid Amt</b> cannot be blank.</li>"
                            ctlName = ctlName & "txtnewpaidamt,"
                        End If
                        If txtAdjustedAmount.Value = "" Or IsNull(txtAdjustedAmount.Value) = True Then
                            msgStr = msgStr & "<li><b>Recovery</b> cannot be blank.</li>"
                            ctlName = ctlName & "txtadjustedamount,"
                        End If
                    End If
                    
                Case 7 'Adjustment Completed with Errors
                    If txtAdjustmentRequestSent.Value = 0 Then
                        msgStr = msgStr & "<li>You must export the <b>Adjustment Request Report</b> before changing status to Adjustment Completed with Errors.</li>"
                        ctlName = ctlName & "cboAuditStatus,"
                    Else
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
'                        If txtNewPaidAmt.Value = "" Or IsNull(txtNewPaidAmt.Value) = True Then
'                            msgStr = msgStr & "<li><b>New Paid Amt</b> cannot be blank.</li>"
'                            ctlName = ctlName & "txtnewpaidamt,"
'                        End If
'                        If txtAdjustedAmount.Value = "" Or IsNull(txtAdjustedAmount.Value) = True Then
'                            msgStr = msgStr & "<li><b>Recovery</b> cannot be blank.</li>"
'                            ctlName = ctlName & "txtadjustedamount,"
'                        End If
                    End If
                    
                Case 9 'Appealed/overturned
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
                    If txtNewPaidAmt.Value = "" Or IsNull(txtNewPaidAmt.Value) = True Then
                        msgStr = msgStr & "<li><b>New Paid Amt</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtnewpaidamt,"
                    End If
                    If txtAdjustedAmount.Value = "" Or IsNull(txtAdjustedAmount.Value) = True Then
                        msgStr = msgStr & "<li><b>Recovery</b> cannot be blank.</li>"
                        ctlName = ctlName & "txtadjustedamount,"
                    End If
                    If optOverturned.Value = "" Or IsNull(optOverturned.Value) = True Then
                        msgStr = msgStr & "<li><b>Overturned?</b> cannot be blank.</li>"
                        ctlName = ctlName & "optOverturned,"
                    ElseIf optOverturned.Value = True Then
                        If txtOverturnedAmt.Value = "" Or IsNull(txtOverturnedAmt.Value) = True Then
                            msgStr = msgStr & "<li><b>Overturned Amt</b> cannot be blank.</li>"
                            ctlName = ctlName & "txtOverturnedAmt,"
                        End If
                        If txtOverturnedDate.Value = "" Or IsNull(txtOverturnedDate.Value) = True Then
                            msgStr = msgStr & "<li><b>Overturned Date</b> cannot be blank.</li>"
                            ctlName = ctlName & "txtOverturnedDate,"
                        End If
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
            
        Case 3 'import financials
            If cboFileName.Value = "" Or IsNull(cboFileName.Value) = True Then
                msgStr = msgStr & "<li><b>File Name</b> cannot be blank.</li>"
                ctlName = ctlName & "cboFileName,"
            End If
            If optImportFinancials.Value = "" Or IsNull(optImportFinancials.Value) = True Then
                msgStr = msgStr & "<li>You must select a type from the <b>Import Financials</b> section. </li>"
                ctlName = "optImportFinancials,"
            End If
            If IsNull(txtClaimNo) Or txtClaimNo.Value = "" Then
                msgStr = msgStr & "<li>You must select a claim record before submitting information.</li>"
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

Public Sub ClearForm()
    If Not IsEmpty(varOffset) Then: Set varOffset = Nothing
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(217, 217, 217)
    Deactivate cboQAComments
    Deactivate txtNewPaidAmt
    cboLOB.RowSource = ""
    optImportFinancials.Value = Null
    txtAdjustedAmount.Visible = False
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
Public Sub RefreshData()
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varStatus As Variant
    Dim total As Integer
    
    Set db = CurrentDb
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
            txtAppealedOverturned.Value = varStatus(1, 7)
            
            For i = 0 To UBound(varStatus, 2)
                total = total + varStatus(1, i)
                'Debug.Print varStatus(1, i)
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
Sub Activate(ctrl As Control)
    With ctrl
        .Visible = True
        .Enabled = True
        '.Locked = False
        '.Value = Null
    End With
End Sub
Sub Deactivate(ctrl As Control)
    With ctrl
        .Visible = False
        .Enabled = False
        '.Locked = True
        .Value = Null
    End With
End Sub

