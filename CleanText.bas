Attribute VB_Name = "CleanText"
Option Compare Database

Function Clean(txtClean As String)

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Function can be called from other modules of Access project
'Accepts string as param and removes invisible chars (spaces, tabs, etc)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

      txtClean = Replace(txtClean, vbLf, "")
      txtClean = Replace(txtClean, vbTab, "")
      txtClean = Replace(txtClean, vbCr, "")
      txtClean = Replace(txtClean, vbCrLf, "")
      txtClean = Replace(txtClean, vbNewLine, "")
      txtClean = Replace(txtClean, Chr(160), "")
      txtClean = Replace(txtClean, Chr(146), "")
      txtClean = Replace(txtClean, Chr(39), "")
      txtClean = Trim(txtClean)
      Clean = txtClean
                
End Function
