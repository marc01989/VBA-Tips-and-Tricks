Private Sub btnQuerySQLServer_Click()
  
    'module requires reference to Microsoft ActiveX Data Objects (ADO)
    'Tools -> References -> Microsoft ActiveX Data Objects 6.1 Library
    'Tools -> References -> Microsoft ActiveX Data Objects REcordset 6.0 Library
    'use ODBC to connect to SQL tables using credentials stored in code
    'In this example, we're going to get claim data form a SQL Server table (Claims) and append it to local Access table (tblClaimFinancials)
      'This involves opening 2 recordsets simultaneously and appending records from one into the other
    'NOTE: since you're storing account name/pswd in code, save the Access app as compile only (.accde) to lock down the code
  
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cn As New ADODB.Connection
    Dim rs2 As New ADODB.Recordset
    Dim sql As String, criteria As String, strCon As String
    Dim fileName As String

On Error GoTo err1:
    DoCmd.SetWarnings False
             
    'populate list of claim numbers as our query criteria
    criteria = "'c239452', 'c463623', 'c444603'"
  
    'prep db objects for local table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblClaimFinancials", dbOpenDynaset)
    
    'prep connection to SQL Server tables
    sql = "SELECT Sum(BilledAmt) AS SumOfBilledAmt, Sum(PaidAmt) AS SumOfPaidAmt, GetDate() AS CurrDate, ClaimNbr FROM Claims WHERE ClaimNbr IN (" & criteria & ") GROUP BY ClaimNbr;"
    Set cn = New ADODB.Connection
    strCon = "Driver={ODBC};" & _
        "Provider=SQLOLEDB;" & _
        "Data Source=SERVERNAME;" & _
        "Database=dbName;" & _
        "Uid=userID;" & _
        "Pwd=fakePswd;"
    cn.Open (strCon)
    rs2.Open sql, cn, adOpenKeyset
    
    'update local tblClaimsFinancial table with results of SQL Server query
    With rs2
        .MoveFirst
        .MoveLast
        .MoveFirst
        Do While Not .EOF
            'this adds new record to local table tblClaimsFinancial, populates with data from SQL Server qry
            rs.AddNew
                rs.Fields(3) = .Fields(0)
                rs.Fields(4) = .Fields(1)
                rs.Fields(5) = .Fields(2)
                rs.Fields(6) = .Fields(3)
            rs.Update
        .MoveNext
        Loop
    End With
    
    'clean up
    rs.Close: Set rs = Nothing
    rs2.Close: Set rs2 = Nothing
    cn.Close: Set cn = Nothing
      
    MsgBox ("Financial Import Complete")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            If MsgBox("Error performing operation. See database admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
