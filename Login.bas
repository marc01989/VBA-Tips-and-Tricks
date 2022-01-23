
'------------------------------------------------------------------
'this module is a simple example of a login form for a MS Access database
'assumes you have a table storing user names, pswds, user id's, active status, etc (tblUsers)
'entry form is very simple, contains 2 textboxes (txtUserName and txtPassword)

'------------------------------------------------------------------


Private Sub btnSubmit_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim userName As String, passwordInput As String, passwordSaved As String
    userName = txtUserName.Value
    passwordInput = txtPassword.Value
    
    Set db = CurrentDb
      Set rs = db.OpenRecordset("tblUsers", dbOpenDynaset)
        With rs
            .FindFirst ("network_id = '" & userName & "'")
                If .NoMatch Then '--user not on tblUsers
                    If ValidateForm(2) = True Then
                        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
                        End If
                    End If
            
                Else '--user is on tblUsers

                    If rs![is_active] = 0 Then '--user is on table but no longer active
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


