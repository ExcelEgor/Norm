VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPolirovaniye 
   Caption         =   "Полирование"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5670
   OleObjectBlob   =   "frmPolirovaniye.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPolirovaniye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMaterial_Change(): Call RaschotVremeni: End Sub

Private Sub cboRa_Change(): Call RaschotVremeni: End Sub

Private Sub optSlozhnost1_Click(): Call RaschotVremeni: End Sub

Private Sub optSlozhnost2_Click(): Call RaschotVremeni: End Sub

Private Sub optSlozhnost3_Click(): Call RaschotVremeni: End Sub

Private Sub txtDlina_Change(): Call RaschotVremeni: End Sub

Private Sub txtShirina_Change(): Call RaschotVremeni: End Sub

Private Sub UserForm_Initialize()

    DobavitMaterial "Пластмасса", PLASTMASSA
    DobavitMaterial "Оргстекло", ORGSTEKLO
    DobavitMaterial "Алюминиевые сплавы", ALUMINIYEVYYE_SPLAVY
    DobavitMaterial "Медные сплавы", MEDNYYE_SPLAVY
    DobavitMaterial "Сталь", STAL_UGLERODISTAYA
    cboMaterial.ListIndex = 1
    
    With cboRa
        .AddItem "2,5"
        .AddItem "1,25...0,63"
        .AddItem "0,32...0,16"
        .ListIndex = 1
    End With
    
End Sub

Private Sub DobavitMaterial(Material, id)
    With cboMaterial
        .AddItem
        .List(.ListCount - 1, 0) = Material
        .List(.ListCount - 1, 1) = id
    End With
End Sub

Private Sub RaschotVremeni()

    txtTpz_Kontrol.text = Empty
    txtTsht_Kontrol.text = Empty
    txtTpz_Polirov.text = Empty
    txtTsht_Polirov.text = Empty

    If IsNumeric(txtDlina) And IsNumeric(txtShirina) And cboMaterial.ListIndex <> -1 And cboRa.ListIndex <> -1 Then
    
        Dim Slozhnost As Integer
        Select Case True
            Case optSlozhnost1.Value: Slozhnost = 1
            Case optSlozhnost2.Value: Slozhnost = 2
            Case optSlozhnost3.Value: Slozhnost = 3
        End Select
        
        Dim Ra As Double
        Select Case cboRa.ListIndex
            Case 0: Ra = 2.5
            Case 1: Ra = 1.25
            Case 2: Ra = 0.32
        End Select
        
        txtTpz_Polirov.text = 5
        txtTsht_Polirov.text = CStr(OkruglenieTsht(1.3 * Polirovanie(txtDlina, txtShirina, MaterialFromCboMaterial(cboMaterial), Slozhnost, Ra)))
        
        txtTpz_Kontrol.text = 0
        txtTsht_Kontrol.text = CStr(OkruglenieTsht(1.3 * VisualnyyKontrol(txtShirina, txtDlina, Slozhnost, False, True)))
        
    End If
    
End Sub




