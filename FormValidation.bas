
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'ValidateForm() function accepts int as param used in case statement to accomodate different validation types
'function checks the state of different controls on form, concatenates an error message and sends to textboxes on form named txtErrorBox and txtErrorBar
'txtErrorBar changes color to alert user of submission error
'focus is set to the control that needs attention 
'example below assumes you have form with textboxes named txtName and txtPassword, and button named btnSubmit
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btnSubmit_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
  
    'CODE TO SUBMIT RECORD GOES HERE, ONLY EXECUTE IF IT PASSES VALIDATION FUNCTION ABOVE
  
End Sub

                                            
                                            
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant

    headerStr = "<ul>"
    footerStr = "</ul>"

    Select Case submitType
    
        Case 1 'submit
            If IsNull(txtName) Or txtName.Value = "" Then
                msgStr = msgStr & "<li><b>Name</b> cannot be blank.</li>"
                ctlName = ctlName & "txtName,"
            End If
            If IsNull(txtPassword) Or txtPassword.Value = "" Then
                msgStr = msgStr & "<li><b>Password</b> cannot be blank.</li>"
                ctlName = ctlName & "txtPassword,"
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



