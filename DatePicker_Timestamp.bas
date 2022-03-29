'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'For some reason, the Date Picker property of a textbox doesn't not append the time after selecting a date from calendar
'This sub is a workaround to insert the time component as well, w/o changing the functionality of date picker
'Example assumes you have textbox named Date_Review_Completed
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub Date_Review_Completed_AfterUpdate()
    Dim inputValue As String

    If Not IsNull(Date_Review_Completed) And Date_Review_Completed.Value <> "" Then
        inputValue = Date_Review_Completed.Value
        inputValue = DateValue(inputValue)
        
        If IsDate(inputValue) Then
            Date_Review_Completed.Value = Null
            Date_Review_Completed.Value = inputValue & " " & Time()
        End If
    End If
End Sub
