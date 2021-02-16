VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmDRGReporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnReturnToMainMenu_Click()
    DoCmd.Close acForm, "frmDRGReporting", acSaveNo
    DoCmd.OpenForm "frmLandingDRG", acNormal, , , , acWindowNormal
End Sub

Private Sub btnRunReport_Click()
    
    'run reports
On Error GoTo err1:
    
    Select Case optReports.Value
        Case 1 'good
            If Application.CurrentProject.AllReports("Auditor Production Report").IsLoaded = True Then DoCmd.Close acReport, "Auditor Production Report", acSaveNo
            DoCmd.OpenReport "Auditor Production Report", acViewReport
            
        Case 2
            If Application.CurrentProject.AllReports("Inventory Report").IsLoaded = True Then DoCmd.Close acReport, "Auditor Production Report", acSaveNo
            DoCmd.OpenReport "Inventory Report", acViewReport

        Case 3 'good
            If Application.CurrentData.AllQueries("qryFileStatusOffset").IsLoaded = True Then DoCmd.Close acQuery, "qryFileStatusOffset", acSaveNo
            DoCmd.OpenQuery "qryFileStatusOffset", acViewNormal, acReadOnly
            
        Case 4
            If Application.CurrentProject.AllReports("Recovery Detail Report").IsLoaded = True Then DoCmd.Close acReport, "Auditor Production Report", acSaveNo
            DoCmd.OpenReport "Recovery Detail Report", acViewReport

        Case 5
            If Application.CurrentProject.AllReports("Reprice Outstanding Claims").IsLoaded = True Then DoCmd.Close acReport, "Reprice Outstanding Claims", acSaveNo
            DoCmd.OpenReport "Reprice Outstanding Claims", acViewReport
             
        Case 6
            If Application.CurrentData.AllQueries("qryNotEligibleForRecovery").IsLoaded = True Then DoCmd.Close acQuery, "qryNotEligibleForRecovery", acSaveNo
            DoCmd.OpenQuery "qryNotEligibleForRecovery", acViewNormal, acReadOnly
            
        Case 7 'good
            If Application.CurrentProject.AllReports("Recovery Report").IsLoaded = True Then DoCmd.Close acReport, "Auditor Production Report", acSaveNo
            DoCmd.OpenReport "Recovery Report", acViewReport
        
        Case 8
            If Application.CurrentData.AllQueries("qryDRGTrends").IsLoaded = True Then DoCmd.Close acQuery, "qryDRGTrends", acSaveNo
            DoCmd.OpenQuery "qryDRGTrends", acViewNormal, acReadOnly

        Case 9
            If Application.CurrentData.AllQueries("qryRawDataReprice").IsLoaded = True Then DoCmd.Close acQuery, "qryRawDataReprice", acSaveNo
            DoCmd.OpenQuery "qryRawDataReprice", acViewNormal, acReadOnly
        
        Case 10
            If Application.CurrentData.AllQueries("qryRawDataOffset").IsLoaded = True Then DoCmd.Close acQuery, "qryRawDataOffset", acSaveNo
            DoCmd.OpenQuery "qryRawDataOffset", acViewNormal, acReadOnly
        
        Case 11 'good
            If Application.CurrentData.AllQueries("qryFileStatusReprice").IsLoaded = True Then DoCmd.Close acQuery, "qryFileStatusReprice", acSaveNo
            DoCmd.OpenQuery "qryFileStatusReprice", acViewNormal, acReadOnly
        
        Case 12 'good
            If Application.CurrentData.AllQueries("qryFileStatusSelection").IsLoaded = True Then DoCmd.Close acQuery, "qryFileStatusSelection", acSaveNo
            DoCmd.OpenQuery "qryFileStatusSelection", acViewNormal, acReadOnly
        
        Case 13 'good
            If Application.CurrentData.AllQueries("qryRawDataOffsetOverturned").IsLoaded = True Then DoCmd.Close acQuery, "qryRawDataOffsetOverturned", acSaveNo
            DoCmd.OpenQuery "qryRawDataOffsetOverturned", acViewNormal, acReadOnly

        
    End Select
    
err1:
    Select Case Err.Number
        Case 0
        Case 2501
        Case 2001
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Reporting; btnRunReport_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

