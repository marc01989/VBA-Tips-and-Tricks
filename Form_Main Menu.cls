VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCloseDatabase_Click()
    DoCmd.Quit
End Sub

Private Sub btnDRG_Click()
    DoCmd.Close acForm, "Main Menu", acSaveNo
    DoCmd.OpenForm "frmLandingDRG", acNormal, , , , acWindowNormal
End Sub

Private Sub btnUtilities_Click()
    DoCmd.Close acForm, "Main Menu", acSaveNo
    DoCmd.OpenForm "Utilities", acNormal, , , , acWindowNormal
End Sub
