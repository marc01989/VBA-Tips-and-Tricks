Public Sub ClearForm()

'call this sub from other routines to clear the controls on a form, empty a global array, change control formatting, etc

    'empty global variant varOffset
    If Not IsEmpty(varOffset) Then: Set varOffset = Nothing
    
    'change color
    txtErrorBar.BackColor = RGB(217, 217, 217)
    
    'change global bool to false
    bolChanged = False
    
    'iterate all controls on form. If textbox, clear contents and enable. If combobox and NOT named cboFileName, do the same
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
