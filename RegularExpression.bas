Public Sub RegularExpression(searchStr As String)

'----------------------------------------------------------------------
  'This sub is a basic example of how to use Regular Expression searches in VBA
  'Sub takes string as paramater, uses regex to find all lines of text that contain it
  'Form has basic controls, which allow user to enter date range and search string
  'Records on Access table are searched - regex matches are displayed to user via Text Box with Text Format property set to "Rich Text", which allows HTML formatting 
  'Reference required to use regex object:
      'Microsoft VBScript Regular Expressions 5.5
'----------------------------------------------------------------------  
  
    Dim regex As RegExp
    Dim colRegMatch As MatchCollection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim resultStr As String, dateStart As String, dateEnd As String
    Dim counter As Integer

    dateStart = Format(dtPickerStart.Value, "yyyymmdd")
    dateEnd = Format(dtPickerEnd.Value, "yyyymmdd")
    resultStr = ""
    counter = 0

    'Instantiate regex object, set properties
    'Set Pattern property to searchStr param - this is the value you are/aren't looking for
    'If you want to find string where this expression is not present, wrap searchStr in this:  ^((?! PHRASE HERE  ).)*$
    Set regex = New RegExp
    With regex
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = searchStr
    End With


    'Open recordset to iterate through - this can be any Access table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT transcript, [Chat Start Time], [Chat End Time], [Chat Time in Mins] " & _
        " FROM chat_data " & _
        " WHERE [Chat Start Time] >= '" & dateStart & "' AND [Chat Start Time] <= '" & dateEnd & "'" & _
        " ORDER BY [Chat Start Time];")

    'With recordset, test each record against the pattern in regex
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Dim tempStr As String
              
    'If pattern is found, save the record to string with HTML formatting to highlight the found searchStr in bold/italics
                    If regex.test(.Fields(0)) = True Then
                        tempStr = .Fields(0)
                        tempStr = regex.Replace(tempStr, "[<i><b>$&</b></i>]")

    'Concat the rest of the fields into resultStr - this displays the entire record formatted together into a readable paragraph using HTML tags
                        counter = counter + 1
                        resultStr = resultStr & "<b>" & counter & ".) </b><br>" & _
                        "<b>Date: </b>" & Left(.Fields(1), 10) & "<br>" & _
                        "<b>Time: </b>" & Right(.Fields(1), 8) & " to " & Right(.Fields(2), 8) & "<br>" & _
                        "<B>Total time: </b>" & .Fields(3) & " minutes <br><br>" & _
                        tempStr & "<br>"
                
                    End If
                .MoveNext
            Loop
        End If
    End With

    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing

    'Display results to user on form
    txtResults.Value = Null
    txtResults.Value = resultStr
    txtChatCount.Value = Null
    txtChatCount.Value = counter

End Sub
