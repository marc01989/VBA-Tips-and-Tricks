'---------------------------------------------------------
'Ways to display helpful placeholder text in textboxes
'Placehoder text in this example reads "required" in lightgrey, indicating that use must enter username
'the below events are all for a single textbox control named txtUserName
'Note: set default forecolor in property sheet for control to #A6A6A6 (light grey). Modules below will change text color as needed
'Same events are useful for password entry boxes - just change input mask to "password" in property sheet to obsure password
'---------------------------------------------------------

'by default, textbox will display the text "required" in light grey until the textbox has focus or until text is keyed in
Private Sub txtUserName_Exit(Cancel As Integer)
    If txtUserName.Value = "" Or IsNull(txtUserName) Then
        txtUserName.ForeColor = RGB(166, 166, 166)
    End If
End Sub

'when user types values into textbox, change fore color to black
Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    txtUserName.ForeColor = vbBlack
End Sub

'when textbox gets focus, placeholder text will disappear 
Private Sub txtUserName_GotFocus()
    If IsNull(txtUserName) Then
        txtUserName.Format = ""
    End If
End Sub

'when focus leaves textbox, if textbox is null, reapply placeholder
Private Sub txtUserName_LostFocus()
    If IsNull(txtUserName) Then
        txtUserName.Format = "@;required"
    End If
End Sub
