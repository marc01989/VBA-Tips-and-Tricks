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
'query schema info to show what users are accessing records in HICS_2.0_be database
'users must not be accessing backend data in order for admin to perform weekly backup
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Long
    Dim strCon As String
    Dim compStr As String
    
On Error GoTo err1:
    lstResults.Value = Null
    lstResults.RowSource = ""
    
    Set cn = New ADODB.Connection
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Database\HICS 2.0 BE\HICS_2.0_BE.accdb;" & _
        "User Id=admin;Password="
    cn.Open (strCon)

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4.0 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the database.
    'Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, _
    "", rs.Fields(2).Name, rs.Fields(3).Name
    
    lstResults.AddItem rs.Fields(0).Name & "," & rs.Fields(1).Name & "," & rs.Fields(2).Name & "," _
    & rs.Fields(3).Name & "," & "USER_NAME"
    
    
    While Not rs.EOF

        compStr = DLookup("specialist_name", "tblEmployeeNames", "[Computer Name] = '" & Clean(rs.Fields(0).Value) & "'") & vbNullString
        lstResults.AddItem Clean(rs.Fields(0).Value) & "," & Clean(rs.Fields(1).Value) & "," & rs.Fields(2).Value & "," & rs.Fields(3).Value & "," & compStr

        rs.MoveNext
    Wend
    
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "USER CHECK; btnUserCheck_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
              
              

