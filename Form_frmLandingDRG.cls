VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLandingDRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCloseDatabase_Click()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
          DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
    DoCmd.OpenForm "Main Menu", acNormal
End Sub

Private Sub btnInvoices_Click()
    DoCmd.OpenForm "Invoices", acNormal
End Sub

Private Sub btnMedicareAutomated_Click()
    MsgBox ("Coming soon")
End Sub

Private Sub btnOpenCommercialReprice_Click()
    DoCmd.OpenForm "frmCommercialLandingReprice", acNormal
End Sub

Private Sub btnOpenMedicaidReprice_Click()
    DoCmd.OpenForm "frmMedicaidLandingReprice", acNormal
End Sub

Private Sub btnOpenOffset_Click()
    DoCmd.OpenForm "frmLandingOffset", acNormal
End Sub

Private Sub btnOpenReprice_Click()
    DoCmd.OpenForm "frmLandingReprice", acNormal
End Sub

Private Sub btnReports_Click()
    DoCmd.OpenForm "frmDRGReporting", acNormal
End Sub

Private Sub Command1_Click()
    DoCmd.OpenForm "frmCommercialLandingComplex", acNormal
End Sub

Private Sub Command2_Click()
    DoCmd.OpenForm "frmMedicaidLandingComplex", acNormal
End Sub

Private Sub Command3_Click()
    DoCmd.OpenForm "frmSelectionFiles", acNormal
End Sub

Private Sub MedicareMainMenu_Click()
    DoCmd.OpenForm "Medicare Main Menu", acNormal
End Sub
