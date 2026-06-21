VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShlifonaiye 
   Caption         =   "Шлифование"
   ClientHeight    =   8595.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   OleObjectBlob   =   "frmShlifonaiye.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShlifonaiye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    Call ZagruzkaFormy_Shlifovaniye

End Sub

Private Sub cboDopusk_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub cboDopuskPloskParallel_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Ploskoe
End Sub

Private Sub cboDupusk_Naruzhnoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub cboGaltel_Naruzhnoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub cboIT_Naruzhnoye_Change()
    Call SpisokSherohovatostey(cboIT_Naruzhnoye, cboRa_Naruzhnoye, 1)
End Sub

Private Sub cboIT_Ploskoye_Change()
    
    With cboRa_Ploskoye
        .Clear
        If cboIT_Ploskoye.ListIndex <> -1 Then
            Select Case cboIT_Ploskoye.ListIndex
                Case 0
                    .AddItem "1,25"
                Case 1
                    .AddItem "1,25"
                    .AddItem "0,63"
                    .AddItem "0,32"
                Case 2
                    .AddItem "0,63"
                    .AddItem "0,32"
            End Select
            .ListIndex = .ListCount - 1
        End If
    End With
    
    Call RaschetVremeni_Ploskoe
    
End Sub

Private Sub cboIT_Vnutrenneye_Change()
    Call SpisokSherohovatostey(cboIT_Vnutrenneye, cboRa_Vnutrenneye, 2)
End Sub

Private Sub cboMaterial_Change()
    Call IzmeneniyeMateriala
End Sub

Private Sub cboPripusk_Naruzhnoye_Change()
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub cboPripusk_Ploskoye_Change()
    Call RaschetVremeni_Ploskoe
End Sub

Private Sub cboPripusk_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub cboRa_Naruzhnoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub cboRa_Ploskoye_Change()
    Call RaschetVremeni_Ploskoe
End Sub

Private Sub cboRa_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub cboShirinaKruga_Naruzhnoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub cboShirinaKruga_Ploskoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Ploskoe
End Sub

Private Sub cboShirinaKruga_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub cboTipProizvodstva_Change()
    Call IzmeneniyeTipaProizvodstva
End Sub

Private Sub cboTolschinaDetali_Ploskoye_Change()
    If Not tglIdotZagruzkaFormy Then Call RaschetVremeni_Ploskoe
End Sub



Private Sub cboVozrast_Change()
    Call RaschetGlavnogoKoeffitsienta
End Sub

Private Sub chkNezhestkaya_Click()
    Call RaschetGlavnogoKoeffitsienta
End Sub

Private Sub chkShlifovaniyeDvukhPoverhnostey_Click()
    Call RaschetGlavnogoKoeffitsienta
End Sub

Private Sub cmdAdd_Ploskoye_Click()

    Select Case mltShlifonaiye.SelectedItem.Name
        Case "pNaruzhnoye"
            Call DobavitVTablitsu_Naruzhnoye
        Case "pPloskoye"
            Call DobavitVTablitsu_Ploskoye
        Case "pVnutrenneye"
            Call DobavitVTablitsu_Vnutrenneye
    End Select
    
End Sub

Private Sub cmdDelete_Click()
    Call mShlifovaniye_RabotaSFormoy.UdaleniyeIzTablitsy
End Sub

Private Sub cmdEdit_Click()
    Call mShlifovaniye_RabotaSFormoy.RedaktirovaniyePerekhoda
End Sub

Private Sub cmdOtmenitRedaktirovaniye_Click()
    Call IzmenitRezhimRedaktirovaniya(False)
End Sub


Private Sub optCilindricheskaya_Naruzhnoye_Click()
    optKonicheskaya_Naruzhnoye.Value = Not optCilindricheskaya_Naruzhnoye
    Call RaschetVremeni_Naruzhnoye
End Sub


Private Sub optCilindricheskaya_Vnutrenneye_Click()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub optKonicheskaya_Naruzhnoye_Click()
    optKonicheskaya_Naruzhnoye.Value = Not optCilindricheskaya_Naruzhnoye
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub optKonicheskaya_Vnutrenneye_Click()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub optPereferiyaKruga_Click()
    Call Visible_ShirinaKruga
End Sub

Private Sub optProdolnaya_Naruzh_Click()
    cboShirinaKruga_Naruzhnoye.Enabled = optProdolnaya_Naruzh.Value
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub optRadialnaya_Naruzh_Click()
    cboShirinaKruga_Naruzhnoye.Enabled = optProdolnaya_Naruzh.Value
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub optToretsKruga_Click()
    Call Visible_ShirinaKruga
End Sub

Private Sub txtDiametr_Naruzhnoye_Change()
    Call SpisokPripuskov_Naruzhnoye
End Sub

Private Sub txtDiametr_Vnutrenneye_Change()
    Call SpisokPripuskov_Vnutrenneyee
End Sub

Private Sub txtDlina_Naruzhnoye_Change()
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub txtDlina_Ploskoye_Change()
    Call RaschetVremeni_Ploskoe
End Sub

Private Sub txtDlina_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub txtKolVo_Naruzhnoye_Change()
    Call RaschetVremeni_Naruzhnoye
End Sub

Private Sub txtKolVo_Ploskoye_Change()
    Call RaschetVremeni_Ploskoe
End Sub

Private Sub txtKolVo_Vnutrenneye_Change()
    Call RaschetVremeni_Vnutrenneye
End Sub

Private Sub txtShirina_Ploskoye_Change()
    Call RaschetVremeni_Ploskoe
End Sub


Private Sub Visible_ShirinaKruga()
    cboShirinaKruga_Ploskoye.Visible = optPereferiyaKruga.Value
    lblShirinaKruga_Ploskoye.Visible = optPereferiyaKruga.Value
End Sub

Private Sub SpisokSherohovatostey(cboIT As MSForms.ComboBox, cboRa As MSForms.ComboBox, ValOtverstiye As Integer)

    Dim IT As Integer
    IT = cboIT.List(cboIT.ListIndex, 1)
    
    With cboRa
        .Clear
        .ListWidth = .Width & " pt"
        Select Case IT
            Case Is <= IIf(ValOtverstiye = 1, 5, 6)
                .List = Split("1,25 0,63 0,32")
            Case Is <= IIf(ValOtverstiye = 1, 6, 7)
                .List = Split("1,25 0,63 0,32")
            Case Is <= 9
                .List = Split("2,5 1,25 0,63")
            Case Else
                .List = Split("2,5 1,25")
        End Select
        .ListIndex = 1
    End With

End Sub












