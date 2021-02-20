VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USER CHECK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUserCheck_Click()

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'query schema info to show which users are accessing records in Access db backend
'module assumes you have a form with listbox to display results and table that houses user names/computer names
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Long
    Dim strCon As String
    Dim compStr As String
    
On Error GoTo err1:
    lstResults.RowSource = ""
    
    'UPDATE strCon WITH PATH TO DATABASE BACKEND, ADD USERID/PSWD AS NEEDED
    Set cn = New ADODB.Connection
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\Public\Desktop\Database1.accdb;" & _
        "User Id=admin;Password="
    cn.Open (strCon)

    'THE USER ROSTER IS EXPOSED AS A PROVIDER-SPECIFIC SCHEMA ROWSET IN THE JET 4.0 OLE DB PROVIDER.  YOU HAVE TO USE A GUID TO
    'REFERENCE THE SCHEMA, AS PROVIDER-SPECIFIC SCHEMAS ARE NOT LISTED IN ADO'S TYPE LIBRARY FOR SCHEMA ROWSETS
    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'OUTPUT HEADERS TO LISTBOX
    lstResults.AddItem rs.Fields(0).Name & "," & rs.Fields(1).Name & "," & rs.Fields(2).Name & "," _
    & rs.Fields(3).Name & "," & "USER_NAME"
    
    
    While Not rs.EOF
        'GET EMPLOYEE NAME BASED ON RETURNED COMPUTER NAME 
        compStr = DLookup("employee_name", "tblEmployeeNames", "[Computer Name] = '" & Clean(rs.Fields(0).Value) & "'") & vbNullString
        
        'OUTPUT RESULTS TO LISTBOX
        lstResults.AddItem Clean(rs.Fields(0).Value) & "," & Clean(rs.Fields(1).Value) & "," & rs.Fields(2).Value & "," & rs.Fields(3).Value & "," & compStr
        rs.MoveNext
    Wend
    
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
    
     'OPTIONAL ERROR HANDLING
err1:
    Select Case Err.Number
        Case 0
        Case Else
            If MsgBox("Error performing operation. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
              
              

