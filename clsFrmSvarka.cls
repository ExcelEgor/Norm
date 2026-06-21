VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrmSvarka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Private pKolVoStrokSvarka As Integer

Private Type KoeffitsientyNaTipProizvodtstva
    kSvarka As Double
    kZachistka As Double
    kPravka As Double
    kKontrol As Double
    kNaneseniyeMela As Double
    kKerosin As Double
End Type

Private Type ResultatyRaschota
    tZachistkaShvov As Double
    tZachistkaPodSvarku As Double
    tVisualKontrol As Double
    tNaneseniyaMela As Double
    tKerosin As Double
    tSvarkaItogo As Double
    tPravkaPosleSvarki As Double
    tKomplektovaniye As Double
    tPeskostruy As Double
End Type

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Let KolVoStrokSvarka(Value As Integer): pKolVoStrokSvarka = Value: End Property

Private Sub txt_Change(): Call GlavnyyRaschot: End Sub

Private Sub cbo_Change(): Call GlavnyyRaschot: End Sub

Sub GlavnyyRaschot()

    Dim Material As EnumMaterialy
    Material = CInt(frmSvarka.cboMaterial.List(frmSvarka.cboMaterial.ListIndex, 1))

    Dim kTipProizvodtstva As KoeffitsientyNaTipProizvodtstva
    kTipProizvodtstva = RaschotKoeffitsientovNaTipProizvodtstva

    Dim Massa As Double
    Massa = DblFromCtrl(frmSvarka.txtMassa)

    Dim KolVoDetaley As Double
    KolVoDetaley = DblFromCtrl(frmSvarka.txtKolVoDetaley)
    
    Call RaschotGalvaniki(KolVoDetaley, kTipProizvodtstva, Material)
    
    Call RaschotKomplektovaniyaSborkiPodSvarku(KolVoDetaley, Massa)
    
    Call RaschotIzTablitsySvarki(Material, kTipProizvodtstva)
    
End Sub

Private Function RaschotKoeffitsientovNaTipProizvodtstva() As KoeffitsientyNaTipProizvodtstva

    With RaschotKoeffitsientovNaTipProizvodtstva
        
        .kSvarka = 1
        .kKontrol = 1
        .kNaneseniyeMela = 1
        .kKerosin = 1
        .kZachistka = 1
        .kPravka = 1

        If frmSvarka.cboTipProizvodstva.ListIndex <> -1 Then

            Select Case frmSvarka.cboTipProizvodstva.ListIndex
                Case 0
                    .kSvarka = 1.5
                    .kZachistka = 1.3
                    .kKontrol = 1.3
                    .kNaneseniyeMela = 1.3
                    .kKerosin = 1.1
                    .kPravka = 1.1
                Case 2
                    .kSvarka = 0.7
                    .kZachistka = 0.8
                    .kKontrol = 0.7
                    .kNaneseniyeMela = 0.7
                    .kKerosin = 0.9
                    .kPravka = 0.9
            End Select
        
        End If
        
    End With

End Function

Private Sub RaschotGalvaniki(KolVoDetaley As Double, kTipProizvodtstva As KoeffitsientyNaTipProizvodtstva, Material As EnumMaterialy)
    
    frmSvarka.txtVremyaTravleniya.text = "-"
    frmSvarka.txtVremyaVisualKontrolTravleniya.text = "-"
    
    If KolVoDetaley <= 0 Then Exit Sub
    
    Dim Ploschad As Double
    Ploschad = DblFromCtrl(frmSvarka.txtPloschad) / KolVoDetaley
    
    If Ploschad <= 0 Then Exit Sub
    
    Dim Operatsiya As String
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        Operatsiya = "Травление"
    Else
        Operatsiya = "Обезжиривание"
    End If
    
    Dim tTravleniye As Double, tKontrol As Double
    tTravleniye = 2 * Galvanika(Operatsiya, Ploschad) * KolVoDetaley
    
    tKontrol = kTipProizvodtstva.kKontrol * VisualnyyKontrolPoPloshadi(Ploschad * 1000000) * KolVoDetaley
    
    frmSvarka.txtVremyaTravleniya.text = OkruglenieTsht(tTravleniye)
    frmSvarka.txtVremyaVisualKontrolTravleniya.text = OkruglenieTsht(tKontrol)

    
End Sub

Private Sub RaschotKomplektovaniyaSborkiPodSvarku(KolVoDetaley As Double, Massa As Double)

    Dim tSborka As Double
    tSborka = SborkaMetallokonstruktsiyIzListovogoMetalla(Massa, KolVoDetaley)
    
    Dim tKomplektovaniye As Double
    tKomplektovaniye = Komplektovaniye(Massa, KolVoDetaley)
    
    Call ZapisatVTextBox(0.8 * tSborka, frmSvarka.txtSborkaPodSvarku)
    Call ZapisatVTextBox(0.2 * tSborka, frmSvarka.txtZachistkaPrikhvatok)
    Call ZapisatVTextBox(tKomplektovaniye, frmSvarka.txtKomplektovaniye)

    
End Sub

Private Sub RaschotIzTablitsySvarki(Material As EnumMaterialy, kTipProizvodtstva As KoeffitsientyNaTipProizvodtstva)
    
    Dim Resultaty As ResultatyRaschota
    
    Call ZapisatResultaty(Resultaty)
    
    Dim Tolshchina As Double, Dlina As Double, KolVo As Double
    
    Dim txtVremya As MSForms.TextBox
    Dim Stroka As Integer
    Dim tSvarka As Double
    Dim MaksimalnayaTolshchina As Double
    
    For Stroka = 1 To pKolVoStrokSvarka
    
        Tolshchina = DblFromCtrl(frmSvarka.Controls("txtTolshchina_" & Stroka))
        If Tolshchina > MaksimalnayaTolshchina Then MaksimalnayaTolshchina = Tolshchina
            
        Dlina = DblFromCtrl(frmSvarka.Controls("txtDlina_" & Stroka))
        KolVo = KolVoPoverhnosteiProhodov(frmSvarka.Controls("txtKolVo_" & Stroka))
        
        Set txtVremya = frmSvarka.Controls("txtVremya_" & Stroka)
        
        tSvarka = kTipProizvodtstva.kSvarka * KolVo * SvarkaArgonoDugovaya(Material, Tolshchina, Dlina, True, VTavr, True, True)
        If tSvarka > 0 Then
            txtVremya.text = CStr(WorksheetFunction.Round(tSvarka, 2))
            
        Else
            txtVremya.text = "-"
        End If
        
        If Dlina > 0 Then
            With Resultaty
                .tSvarkaItogo = .tSvarkaItogo + tSvarka
                .tZachistkaShvov = .tZachistkaShvov + kTipProizvodtstva.kZachistka * KolVo * ZachistkaShvovShlifovalnoyMashinkoy(Material, Dlina)
                .tZachistkaPodSvarku = .tZachistkaPodSvarku + kTipProizvodtstva.kZachistka * KolVo * ZachistkaPodSvarkuNapilnikom(Material, Dlina)
                .tVisualKontrol = .tVisualKontrol + kTipProizvodtstva.kKontrol * KolVo * KontrolSvarnykhShvov(Dlina)
                .tNaneseniyaMela = .tNaneseniyaMela + kTipProizvodtstva.kNaneseniyeMela * KolVo * Lakirovanie((Tolshchina * Dlina) / 10000, 2, 1)
                .tKerosin = .tKerosin + kTipProizvodtstva.kKerosin * KolVo * ProverGermetSvarShvaKerosinom(Dlina)
            End With
        End If
        
    Next

    Dim DlinaIzdeliya As Double, ShirinaIzdeliya As Double
        
    DlinaIzdeliya = DblFromCtrl(frmSvarka.txtDlina)
    ShirinaIzdeliya = DblFromCtrl(frmSvarka.txtShirina)
    
    Resultaty.tPravkaPosleSvarki = PravkaSvarnykhUzlovVruchnuyuNaPlite(Material, DlinaIzdeliya, ShirinaIzdeliya, MaksimalnayaTolshchina)
    Resultaty.tPeskostruy = kTipProizvodtstva.kZachistka * PeskostruynayaOchistkaSvarnukhIzdeleiy(DlinaIzdeliya, ShirinaIzdeliya, DblFromCtrl(frmSvarka.txtVysota), DblFromCtrl(frmSvarka.txtKolVoDetaley))

    Call ZapisatResultaty(Resultaty)
    
End Sub

Private Sub ZapisatResultaty(Resultaty As ResultatyRaschota)

    With frmSvarka
    
        Call ZapisatVTextBox(Resultaty.tZachistkaShvov, .txtItogoVremyaZachistki)
        Call ZapisatVTextBox(Resultaty.tZachistkaPodSvarku, .txtZachistkaPodSvarku)
        Call ZapisatVTextBox(Resultaty.tVisualKontrol, .txtKontrolVisualnyy)
        Call ZapisatVTextBox(Resultaty.tNaneseniyaMela, .txtNaneseniyeMela)
        Call ZapisatVTextBox(Resultaty.tKerosin, .txtKerosin)
        Call ZapisatVTextBox(Resultaty.tSvarkaItogo, .txtItogoVremyaSvarki)
        Call ZapisatVTextBox(Resultaty.tPravkaPosleSvarki, .txtVremyaPravki)
        Call ZapisatVTextBox(Resultaty.tPeskostruy, .txtPeskostruy)
        
    End With

End Sub

Private Sub ZapisatVTextBox(Vremya As Double, TextBox As MSForms.TextBox)
    
    If Vremya > 0 Then
        TextBox.text = WorksheetFunction.RoundUp(Vremya, 0)
    Else
        TextBox.text = "-"
    End If
    
End Sub




