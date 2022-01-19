Attribute VB_Name = "ErrorLogger"
Option Compare Database

Public Sub LogError(strError, modName As String)
    
    '--------------------------------------------------------------------------------------------
    'call this sub from other modules in VBA project to save the error to txt file
    'this sub accepts the captured error and the name of module where error was thrown
    'list of Ms Access errors: https://msdn.microsoft.com/en-us/library/bb221208(v=office.12).aspx
    '--------------------------------------------------------------------------------------------
    
    Dim strPath As String, userName As String
    Dim fs As Object
    Dim a As Object

    'GET USERNAME OF LOGGED ON USER
    userName = Environ$("username")
    'PATH TO SAVE TEXT FILE
    strPath = "C:\Users\Public\Desktop"

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
        a.writeline "VERSION: xx.xx.xx"
        a.writeline "--------------------------------------------------------------------------"
        a.Close
    Set fs = Nothing
End Sub

                
'THE BELOW SUB IS AN EXAMPLE OF HOW THE ErrorLogger() FUNCTION IS CALLED FROM DIFFEREND SUB
                
Private Sub btnSubmit_Click()        
    On Error GoTo err1:

        'CODE THAT COULD THROW ERROR GOES HERE

    err1:
        Select Case Err.Number
            Case 0
            Case Else
                Call LogError(Err.Number & " " & Err.Description, "MainForm; btnSubmit_Click()")
                If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
                Exit Sub
        End Select
End Sub


