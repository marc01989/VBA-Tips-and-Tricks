Private Sub btnImportOffsetFile_Click()
'must have reference "Microsoft Office 2016 Object Library" selected to use FileDialog object
    'go to tools -> references -> Microsoft Office 2016 Object Library


    Dim fDialog As FileDialog
    Dim filePath As String, fileLocation As String, fileName As String, fileType As String
    Dim userID As Long: userID = DLookup("user_id", "tblUsers", "network_id = '" & Environ("USERNAME") & "'")
    Dim subDate As String: subDate = Now()
    Dim intComp As Integer
    Dim lobid As Long
    
On Error GoTo err1:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
