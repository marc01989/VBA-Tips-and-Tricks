VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnResetPassword_Click()
    DoCmd.OpenForm "Password Reset", acNormal, , , acFormPropertySettings
    DoCmd.Close acForm, "Login", acSaveNo
End Sub

Private Sub btnSubmit_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:

    '----SUBMIT INFO HERE----
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim userName As String, passwordInput As String, passwordSaved As String
    userName = txtUserName.Value
    passwordInput = txtPassword.Value
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblAuditors", dbOpenDynaset)
        With rs
            .FindFirst ("network_id = '" & userName & "'")
                If .NoMatch Then '--user not on tblAuditors
                    If ValidateForm(2) = True Then
                        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
                        End If
                    End If
                Else '--user is on tblAuditors

                    If rs![is_active] = 0 Then '--user not active
                        If ValidateForm(3) = True Then
                            If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
                            End If
                        End If
                    Else '--user is active
                        Dim intComp As Integer
                        passwordSaved = rs![password]
                        intComp = StrComp(passwordInput, passwordSaved, vbBinaryCompare)
                        
                        If intComp = 0 Then '--match
                            Dim userId As Integer
                            userId = rs![auditor_id]
                            DoCmd.OpenForm "Home", acNormal, , , , acWindowNormal, userId
                            DoCmd.Close acForm, "Login", acSaveNo
                        Else '--no match
                            If ValidateForm(2) = True Then
                                If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
                                End If
                            End If
                        End If
                    End If
                End If
        End With
        
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Login; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select


End Sub
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1
            If txtUserName.Value = "" Or IsNull(txtUserName.Value) = True Then
                msgStr = msgStr & "<li><b>User Name</b> cannot be blank.</li>"
                ctlName = "txtUserName,"
            End If
            If txtPassword.Value = "" Or IsNull(txtPassword.Value) = True Then
                msgStr = msgStr & "<li><b>Password</b> cannot be blank.</li>"
                ctlName = ctlName & "txtPassword,"
            End If
        Case 2
            msgStr = msgStr & "<li>Username and password combination was not found. Please try again.</li>"
            ctlName = ctlName & "txtUserName,"
        Case 3
            msgStr = msgStr & "<li>You no longer have access to use this database. Please see admin for assistance.</li>"
            ctlName = ctlName & "txtUserName,"
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


Private Sub txtUserName_Exit(Cancel As Integer)
    If txtUserName.Value = "" Or IsNull(txtUserName) Then
        txtUserName.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    txtUserName.ForeColor = vbBlack
End Sub
Private Sub txtUserName_GotFocus()
    If IsNull(txtUserName) Then
        txtUserName.Format = ""
    End If
End Sub
Private Sub txtUserName_LostFocus()
    If IsNull(txtUserName) Then
        txtUserName.Format = "@;required"
    End If
End Sub
Private Sub txtPassword_Exit(Cancel As Integer)
    If txtPassword.Value = "" Or IsNull(txtPassword) Then
        txtPassword.ForeColor = RGB(166, 166, 166)
        txtPassword.Format = "@;required"
    Else
        txtPassword.Format = ""
        txtPassword.InputMask = "Password"
    End If
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    txtPassword.ForeColor = vbBlack
End Sub
Private Sub txtPassword_GotFocus()
    If IsNull(txtPassword) Then
        txtPassword.Format = ""
    End If
End Sub
Private Sub txtPassword_LostFocus()
    If IsNull(txtPassword) Then
        txtPassword.Format = "@;required"
    End If
End Sub
