Public Sub RegularExpression(searchStr As String)

'----------------------------------------------------------------------
  'This sub is a basic example of how to use Regular Expression searches in VBA
  'Sub takes string as paramater, uses regex to find all lines of text that contain it
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

    'instantiate regex object, set properties
    'set pattern to sub param - this is what you are/aren't looking for
    'find string where this expression is not present:  ^((?! PHRASE HERE  ).)*$
    Set regex = New RegExp
    With regex
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = searchStr
    End With

    resultStr = ""
    counter = 0

    'open recordset to iterate through
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT chat_transcripts.transcript, [Chat Start Time], [Chat End Time], [Chat Time in Mins] " & _
    " FROM chat_data INNER JOIN chat_transcripts ON chat_data.Connid = chat_transcripts.Connid " & _
    " WHERE FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') >= '" & dateStart & "' AND FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') <= '" & dateEnd & "'" & _
    " ORDER BY FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') ;")

    'with recordset, test each record against the pattern in regex
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Dim tempStr As String
              
    'if pattern is found, save the record to string with HTML formatting to highlight the found searchStr
                    If regex.test(.Fields(0)) = True Then
                        tempStr = .Fields(0)
                        tempStr = regex.Replace(tempStr, "[<i><b>$&</b></i>]")
                        'Debug.Print tempStr

                        counter = counter + 1
                        'Set colRegMatch = regex.Execute(.Fields(0))
                        'Debug.Print .Fields(0) & "<br><br>"
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

    MsgBox "no errors"

    txtResults.Value = Null
    txtResults.Value = resultStr
    txtChatCount.Value = Null
    txtChatCount.Value = counter

End Sub
