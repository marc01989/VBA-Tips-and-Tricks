VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Utilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub btnCloseDatabase_Click()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
          DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
    DoCmd.OpenForm "Main Menu", acNormal
End Sub

Private Sub btnSubmitProvider_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblExclusionProviders", dbOpenDynaset)
    With rs
        .AddNew
        ![vendor_id] = txtVendorId.Value
        ![provider_name] = txtProviderName.Value
        ![lob_id] = cboLOB.Value
        .Update
        .Close
    End With
    
    MsgBox ("Item successfully submitted.")
    
    Set rs = Nothing
    db.Close: Set db = Nothing
    
End Sub

Private Sub btnSubmitMember_Click()

    If ValidateForm(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblExclusionMembers", dbOpenDynaset)
    With rs
        .AddNew
        ![member_id] = txtMemberId.Value
        ![member_name] = txtMemberName.Value
        .Update
        .Close
    End With
    
    MsgBox ("Item successfully submitted.")
    
    Set rs = Nothing
    db.Close: Set db = Nothing
End Sub
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant

    headerStr = "<ul>"
    footerStr = "</ul>"

    Select Case submitType
    
        Case 1 'update provider exclusions
            If IsNull(txtVendorId) Or txtVendorId.Value = "" Then
                msgStr = msgStr & "<li><b>Vendor ID</b> cannot be blank.</li>"
                ctlName = ctlName & "txtVendorId,"
            End If
            If IsNull(txtProviderName) Or txtProviderName.Value = "" Then
                msgStr = msgStr & "<li><b>Provider Name</b> cannot be blank.</li>"
                ctlName = ctlName & "txtProviderName,"
            End If
            If IsNull(cboLOB) Or cboLOB.Value = "" Then
                msgStr = msgStr & "<li><b>Line of Business</b> cannot be blank.</li>"
                ctlName = ctlName & "cboLOB,"
            End If
            
        Case 2
            If IsNull(txtMemberId) Or txtMemberId.Value = "" Then
                msgStr = msgStr & "<li><b>Member ID</b> cannot be blank.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
            If IsNull(txtMemberName) Or txtMemberName.Value = "" Then
                msgStr = msgStr & "<li><b>Member Name</b> cannot be blank.</li>"
                ctlName = ctlName & "txtMemberName,"
            End If
            
        Case 3
            If IsNull(txtFullName) Or txtFullName.Value = "" Then
                msgStr = msgStr & "<li><b>Full Name</b> cannot be blank.</li>"
                ctlName = ctlName & "txtFullName,"
            End If
            If IsNull(txtNetworkId) Or txtNetworkId.Value = "" Then
                msgStr = msgStr & "<li><b>Network ID</b> cannot be blank.</li>"
                ctlName = ctlName & "txtNetworkId,"
            End If
            If IsNull(txtEmail) Or txtEmail.Value = "" Then
                msgStr = msgStr & "<li><b>Email</b> cannot be blank.</li>"
                ctlName = ctlName & "txtEmail,"
            End If
            If IsNull(cboManagerAccess) Or cboManagerAccess.Value = "" Then
                msgStr = msgStr & "<li><b>Manager Access</b> cannot be blank.</li>"
                ctlName = ctlName & "cboManagerAccess,"
            End If
            If IsNull(cboIsUserActive) Or cboIsUserActive.Value = "" Then
                msgStr = msgStr & "<li><b>Is User Active?</b> cannot be blank.</li>"
                ctlName = ctlName & "cboIsUserActive,"
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

Private Sub cboUsers_AfterUpdate()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblUsers WHERE user_id = " & cboUsers.Value & ";", dbReadOnly)
    
    With rs
        If rs.RecordCount > 0 Then
            txtFullName.Value = ![user_name]
            txtNetworkId.Value = ![network_id]
            txtEmail.Value = ![Email]
            cboManagerAccess.Value = ![is_manager]
            cboIsUserActive.Value = ![is_active]
        End If
    End With
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
End Sub
Private Sub btnSaveUser_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    If ValidateForm(3) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblUsers", dbOpenDynaset)

    With rs
    
        If optUpdateType.Value = 1 Then
            .AddNew
                ![user_name] = txtFullName.Value
                ![network_id] = txtNetworkId.Value
                ![Email] = txtEmail.Value
                ![is_manager] = cboManagerAccess.Value
                ![is_active] = cboIsUserActive.Value
            .Update
            MsgBox ("User successfully added")
        Else
            .FindFirst "user_id = " & cboUsers.Value
            If Not .NoMatch Then
                .Edit
                    ![user_name] = txtFullName.Value
                    ![network_id] = txtNetworkId.Value
                    ![Email] = txtEmail.Value
                    ![is_manager] = cboManagerAccess.Value
                    ![is_active] = cboIsUserActive.Value
                .Update
                MsgBox ("User successfully updated")
            Else
                MsgBox ("Cannot find user")
            End If
        End If

    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Utilities; btnSaveUser_Click()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Private Sub Form_Load()
    Deactivate cboUsers
End Sub

Private Sub optUpdateType_AfterUpdate()
    Select Case optUpdateType
        Case 1
            Deactivate cboUsers
        Case 2
            Activate cboUsers
    End Select
    
    cboUsers.Value = ""
    txtFullName.Value = ""
    txtNetworkId.Value = ""
    txtEmail.Value = ""
    cboManagerAccess.Value = ""
    cboIsUserActive.Value = ""
    cboUsers.Requery
    
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
