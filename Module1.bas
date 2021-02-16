Attribute VB_Name = "Module1"
Option Compare Database

Public Sub LogError(strError, modName As String)
    'call this sub from other modules in VBA project to save the error to txt file
    'list of errors: https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx

    Dim strPath As String, userName As String
    Dim fs As Object
    Dim a As Object

    'GET USERNAME OF LOGGED ON USER
    userName = Environ$("username")
    'PATH TO SAVE TEXT FILE
    strPath = "C:/Users/Public/Desktop"

    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strPath & "\ErrorLog.txt") = True Then
            Set a = fs.Opentextfile(strPath & "\ErrorLog.txt", 8)
        Else
            Set a = fs.createtextfile(strPath & "\ErrorLog.txt")
        End If
    
        a.writeline "--------------------------------------------------------------------------"
        a.writeline "DATE: " & Date + Time
        a.writeline "ERROR: " & strError
        a.writeline "USER: " & userName
        a.writeline "MODULE: " & modName
        a.writeline "VERSION: 1.7"
        a.writeline "--------------------------------------------------------------------------"
        a.Close
    Set fs = Nothing
End Sub

Public Sub UpdateSubmissionLog(userID As Long, claimNbr As String, fileName As String, processId As Long, lobid As Long)

    DoCmd.SetWarnings False
        
On Error GoTo err1:
        
    DoCmd.RunSQL ("INSERT INTO tblSubmissionLog (user_id, claim_nbr, file_name, process_id, lob_id, sub_date) " & _
    " VALUES ( " & userID & ", '" & claimNbr & "', '" & fileName & "', " & processId & ", " & lobid & ", '" & Now() & "');")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Module1; UpdateSubmissionLog()")
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub


