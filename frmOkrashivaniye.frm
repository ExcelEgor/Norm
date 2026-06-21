VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOkrashivaniye 
   Caption         =   "ŒÍ‡¯Ë‚‡ÌËÂ"
   ClientHeight    =   11475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   OleObjectBlob   =   "frmOkrashivaniye.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmOkrashivaniye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LBL_TOP As Integer = 9

Private FormaZagruzhaetsya As Boolean

Const KolVoStrokIzolOtv As Integer = 7

Private Sub cboKolVoDet_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboMetodLitya_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboSposobIzolOtv_1_Change(): Call GlavnyyRaschot: End Sub

Private Sub cmdNachalnyyeZnacheniya_Click(): Call UstanovitParametryPoUmolchaniyu: End Sub

Private Sub lblNastroyki_Click(): mltOkrashivaniye.Value = 1: End Sub

Private Sub lblRaschot_Click(): mltOkrashivaniye.Value = 0: End Sub

Private Sub mltOkrashivaniye_Change(): Call IzmenitStilVkladok: End Sub

Private Sub optSloznost_1_Click(): Call GlavnyyRaschot: End Sub

Private Sub optSloznost_2_Click(): Call GlavnyyRaschot: End Sub

Private Sub optSloznost_3_Click(): Call GlavnyyRaschot: End Sub

Private Sub txtKolVoOtv_1_Change(): Call GlavnyyRaschot: End Sub

Private Sub txtPloschadIzolCilindrPoverh_Change(): Call GlavnyyRaschot: End Sub

Private Sub txtPloschadIzolPloskPoverh_Change(): Call GlavnyyRaschot: End Sub

Private Sub txtPloschadPokrytiya_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboKolVoSloyovEmal_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboKolVoSloyovGruntovka_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboSposobIzolPloskPoverh_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboTipPoverhnosti_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboEdinitsaIzmereniya_Change(): Call GlavnyyRaschot: End Sub

Private Sub cboEmal_Change(): Call ZagruzkaKolVaSloyov(cboEmal, "tbOkrashSloi", cboKolVoSloyovEmal): End Sub

Private Sub cboGruntovka_Change(): Call ZagruzkaKolVaSloyov(cboGruntovka, "tbOkrashGruntSloi", cboKolVoSloyovGruntovka): End Sub

Private Sub UserForm_Initialize()

    FormaZagruzhaetsya = True
    
    mltOkrashivaniye.Style = fmTabStyleNone
    mltOkrashivaniye.Value = 0
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Call ZagruzkaSpiskov
    
    Call UstanovitParametryPoUmolchaniyu
    
    Call VyravnitElementy
    
    Call IzmenitStilVkladok
    
    FormaZagruzhaetsya = False
    
    txtPloschadPokrytiya.SetFocus

End Sub

Private Sub ZagruzkaSpiskov()

    cboEmal.List = wsLakokras_Spiski.ListObjects("tbOkrashPokritie").DataBodyRange.Value
    
    cboEdinitsaIzmereniya.List = Split("ÏÏ2 ÒÏ2 ‰Ï2 Ï2")
    
    cboSposobPokrytiya.List = wsLakokras_Spiski.ListObjects("tbOkrashSposob").DataBodyRange.Value
    
    cboKolVoSloyovGruntovka.ListIndex = cboKolVoSloyovGruntovka.ListCount - 1
    
    cboTipPoverhnosti.List = wsLakokras_Spiski.ListObjects("tbOkrashPoverh").DataBodyRange.Value
    
    cboKolVoDet.List = wsLakokras_Spiski.ListObjects("tbOkrash_KoefKolVoDet").DataBodyRange.Value

    cboMetodLitya.List = wsLakokras_Spiski.ListObjects("tbOkrash_KoefMetodLitya").DataBodyRange.Value
    
    cboSposobIzolOtv_1.List = wsLakokras_Spiski.ListObjects("tbIzolOtv").DataBodyRange.Value

    cboSposobIzolPloskPoverh.List = wsLakokras_Spiski.ListObjects("tbIzolCilindr").DataBodyRange.Value
    
    
End Sub

Private Sub UstanovitParametryPoUmolchaniyu()
    
    optSloznost_3.Value = True
    cboEmal.Value = "œ‘-218√—"
    cboEdinitsaIzmereniya.ListIndex = 3
    cboSposobPokrytiya.ListIndex = 1
    cboGruntovka.Value = "¬À-02"
    cboKolVoSloyovEmal.ListIndex = 0
    cboTipPoverhnosti.ListIndex = 1
    cboKolVoDet.ListIndex = 0
    cboMetodLitya.ListIndex = 0
    cboSposobIzolOtv_1.ListIndex = 0
    cboSposobIzolPloskPoverh.ListIndex = cboSposobIzolPloskPoverh.ListCount - 1
    
End Sub

Private Sub VyravnitElementy()
    
    lblRaschot.Top = LBL_TOP
    lblNastroyki.Top = LBL_TOP
    lblLine.Top = lblNastroyki.Top + lblNastroyki.Height + 2
    
    mltOkrashivaniye.Top = lblLine.Top + 1
    mltOkrashivaniye.Left = 3
    mltOkrashivaniye.Height = fraNormy.Top + fraNormy.Height + 9
    mltOkrashivaniye.Width = fraNormy.Left + fraNormy.Width + 9
     
    lblRaschot.Left = mltOkrashivaniye.Left
    lblNastroyki.Left = lblRaschot.Left + lblRaschot.Width
    
    lblLine.Left = lblRaschot.Left
    lblLine.Width = mltOkrashivaniye.Width
    
    Me.Width = mltOkrashivaniye.Width + mltOkrashivaniye.Left + PLUS_SHIRINA
    
    cmdNovyyRaschet.Left = mltOkrashivaniye.Left
    cmdNovyyRaschet.Top = mltOkrashivaniye.Top + mltOkrashivaniye.Height + 2
    
    fraNastroyki.Top = 3
    fraNastroyki.Left = 3
    fraNastroyki.Width = mltOkrashivaniye.Width - mltOkrashivaniye.Left - 6
    
    Me.Height = cmdNovyyRaschet.Top + cmdNovyyRaschet.Height + PLUS_VYSOTA
    
End Sub

Private Sub GlavnyyRaschot()
    
    txtTsht_Okrashivaniye.text = Empty
    txtTpz_Okrashivaniye.text = Empty
    txtTsht_Kontrol.text = Empty
    
    Call ProstavitEdinitsyIzmereniya
    
    If FormaZagruzhaetsya Then Exit Sub
    
    Dim Sloznost As Integer
    If optSloznost_1.Value Then
        Sloznost = 1
    ElseIf optSloznost_2.Value Then
        Sloznost = 2
    ElseIf optSloznost_3.Value Then
        Sloznost = 3
    Else
        Exit Sub
    End If

    Dim tPokrytiya As Double
    tPokrytiya = RaschotPokrytiya(Sloznost)
    
    Dim tIzolOtv As Double
    If txtKolVoOtv_1.text <> "" Then tIzolOtv = KolVoPoverhnosteiProhodov(txtKolVoOtv_1) * IzolirovaniyeOtverstiy(cboSposobIzolOtv_1.ListIndex + 1, 10, Sloznost)
    
    Dim tIzolPoverkh As Double
    tIzolPoverkh = RaschotIzolirovaniyePoverhnostey(Sloznost)

    Dim kKolVoDet As Double
    kKolVoDet = cboKolVoDet.List(cboKolVoDet.ListIndex, 1)
    
    Dim Tsht As Double
    Tsht = (kKolVoDet * tPokrytiya + tIzolOtv + tIzolPoverkh)

    If Tsht > 0 Then
        txtTsht_Okrashivaniye = CStr(OkruglenieTsht(Tsht))
        
        If tIzolOtv + tIzolPoverkh And tPokrytiya > 0 Then
            txtTpz_Okrashivaniye = 10
        Else
            txtTpz_Okrashivaniye = 5
        End If
        
        Dim Ploschad_mm2 As Double
        Ploschad_mm2 = KonvertPloschad_mm2(cboEdinitsaIzmereniya, txtPloschadIzolPloskPoverh) + KonvertPloschad_mm2(cboEdinitsaIzmereniya, txtPloschadPokrytiya)
        
        Dim tVisualKontrol As Double
        tVisualKontrol = VisualnyyKontrolPoPloshadi(Ploschad_mm2) * 1.3 + 0.01 * KolVoPoverhnosteiProhodov(txtKolVoOtv_1)
        
        txtTsht_Kontrol = CStr(OkruglenieTsht(tVisualKontrol))
        
    End If

End Sub

Private Sub ProstavitEdinitsyIzmereniya()
    
    Dim EdIzmer As String
    Select Case cboEdinitsaIzmereniya.text
        Case "ÏÏ2": EdIzmer = lbl_mm2
        Case "ÒÏ2": EdIzmer = lbl_sm2
        Case "‰Ï2": EdIzmer = lbl_dm2
        Case "Ï2": EdIzmer = lbl_m2
    End Select
    
    lbl_SIzolPlosk.Caption = EdIzmer
    lbl_SIzolTsilindr.Caption = EdIzmer
    lbl_Spokrytiya.Caption = EdIzmer
    
End Sub

Private Function RaschotPokrytiya(Sloznost As Integer) As Double
    
    Dim tGruntovki As Double, tOkrashivaniye As Double, tLakirovaniye As Double, Ploschad_dm2 As Double
    
    If IsNumeric(txtPloschadPokrytiya) Then
    
        Select Case cboEdinitsaIzmereniya.ListIndex
            Case 0: Ploschad_dm2 = txtPloschadPokrytiya * 0.0001
            Case 1: Ploschad_dm2 = txtPloschadPokrytiya * 0.01
            Case 2: Ploschad_dm2 = txtPloschadPokrytiya * 1
            Case 3: Ploschad_dm2 = txtPloschadPokrytiya * 100
        End Select
        
        If chkGruntovaniye Then
            If IsNumeric(cboKolVoSloyovGruntovka) Then
                tGruntovki = Gruntovaniye(Ploschad_dm2, cboGruntovka, cboKolVoSloyovGruntovka, cboSposobPokrytiya.ListIndex + 1, Sloznost, cboTipPoverhnosti.ListIndex + 1)
            End If
        End If
        
        If chkOkrashivaniye Then
            If IsNumeric(cboKolVoSloyovEmal) Then
                tOkrashivaniye = Okrashivaniye(Ploschad_dm2, cboEmal, Sloznost, cboKolVoSloyovEmal, cboSposobPokrytiya.ListIndex + 1)
            End If
        End If
        
        If chkLakirovaniye Then
            tLakirovaniye = Lakirovanie(Ploschad_dm2, cboSposobPokrytiya.ListIndex + 1, Sloznost)
        End If
        
        Dim kMetodLitya As Double
        kMetodLitya = cboMetodLitya.List(cboMetodLitya.ListIndex, 1)
        
        RaschotPokrytiya = kMetodLitya * (tGruntovki + tOkrashivaniye + tLakirovaniye)
        
    End If
    
End Function

Private Function RaschotIzolirovaniyePoverhnostey(Sloznost As Integer) As Double
    
    Dim tIzolPlosk As Double
    If IsNumeric(txtPloschadIzolPloskPoverh) Then
        tIzolPlosk = fIzolPloskPov(cboSposobIzolPloskPoverh.ListIndex + 1, KonvertPloschad(cboEdinitsaIzmereniya, txtPloschadIzolPloskPoverh), Sloznost)
    End If
    
    Dim tIzolCilindr As Double
    If IsNumeric(txtPloschadIzolCilindrPoverh) Then
        tIzolCilindr = fIzolCilindrPov(cboSposobIzolPloskPoverh.ListIndex + 1, KonvertPloschad(cboEdinitsaIzmereniya, txtPloschadIzolCilindrPoverh), Sloznost)
    End If
    
    RaschotIzolirovaniyePoverhnostey = tIzolPlosk + tIzolCilindr
    
End Function

Private Sub ZagruzkaKolVaSloyov(cboPokritie As MSForms.ComboBox, tbSloiName As String, cboSloi As MSForms.ComboBox)
    
    cboSloi.Clear
    
    If cboPokritie.ListCount > 0 Then
        Dim i As Integer
        Dim tbSloi As ListObject:    Set tbSloi = wsLakokras_Spiski.ListObjects(tbSloiName)
        
        With tbSloi
            For i = 1 To .ListRows.Count
                If .DataBodyRange(i, 1) = cboPokritie.Value Then
                    cboSloi.AddItem .DataBodyRange(i, 2)
                    If .DataBodyRange(i, 1) <> .DataBodyRange(i + 1, 1) Then Exit For
                End If
            Next
        End With
        
        cboSloi.ListIndex = 0
        
    End If
    
End Sub

Private Sub cboSposobPokrytiya_Change()
    
    Dim i As Integer
    Dim tbOkrashGruntovka As ListObject:    Set tbOkrashGruntovka = wsLakokras_Spiski.ListObjects("tbOkrashGruntovka")
    
    With cboGruntovka
        .Clear
        For i = 1 To tbOkrashGruntovka.ListRows.Count
            If tbOkrashGruntovka.DataBodyRange(i, 1) = "√‘-020" Or tbOkrashGruntovka.DataBodyRange(i, 1) = "√‘-021" Then
                If cboSposobPokrytiya.ListIndex <> 1 Then .AddItem tbOkrashGruntovka.DataBodyRange(i, 1)
            Else
                .AddItem tbOkrashGruntovka.DataBodyRange(i, 1)
            End If
        Next
        .ListIndex = 0
    End With
    
    Call GlavnyyRaschot
    
End Sub

Private Sub chkGruntovaniye_Click()
        
    cboGruntovka.Enabled = chkGruntovaniye.Value
    cboKolVoSloyovGruntovka.Enabled = chkGruntovaniye.Value
    Call GlavnyyRaschot
    
End Sub

Private Sub chkOkrashivaniye_Click()

    cboEmal.Enabled = chkOkrashivaniye.Value
    cboKolVoSloyovEmal.Enabled = chkOkrashivaniye.Value
    Call GlavnyyRaschot
    
End Sub

Private Sub chkLakirovaniye_Click()

    chkGruntovaniye.Value = Not chkLakirovaniye.Value
    chkOkrashivaniye.Value = Not chkLakirovaniye.Value
    chkGruntovaniye.Enabled = Not chkLakirovaniye.Value
    chkOkrashivaniye.Enabled = Not chkLakirovaniye.Value
    Call GlavnyyRaschot
    
End Sub


Private Sub cmdNovyyRaschet_Click()

    Dim tp As Double, lf As Double
    tp = Me.Top
    lf = Me.Left

    Unload frmOkrashivaniye
    
    With frmOkrashivaniye
        .Top = tp
        .Left = lf
        .Show
    End With
    
End Sub

Private Function KonvertPloschad(cboEdinitsaIzmereniya As MSForms.ComboBox, txtPloschad As MSForms.TextBox) As Double

    Dim Ploschad As Double
    Ploschad = DblFromCtrl(txtPloschad)
    Select Case cboEdinitsaIzmereniya.ListIndex
        Case 0: KonvertPloschad = Ploschad * 0.0001
        Case 1: KonvertPloschad = Ploschad * 0.01
        Case 2: KonvertPloschad = Ploschad * 1
        Case 3: KonvertPloschad = Ploschad * 100
    End Select
    
End Function
Private Function KonvertPloschad_mm2(cboEdinitsaIzmereniya As MSForms.ComboBox, txtPloschad As MSForms.TextBox) As Double

    Dim Ploschad As Double
    Ploschad = DblFromCtrl(txtPloschad)
    Select Case cboEdinitsaIzmereniya.ListIndex
        Case 0: KonvertPloschad_mm2 = Ploschad / 1
        Case 1: KonvertPloschad_mm2 = Ploschad / 0.01
        Case 2: KonvertPloschad_mm2 = Ploschad / 0.0001
        Case 3: KonvertPloschad_mm2 = Ploschad / 0.000001
    End Select
    
End Function

Private Sub IzmenitStilVkladok()

    Dim lblAktiv As Control, lblNeaktiv As Control

    If mltOkrashivaniye.Value = 0 Then
        Set lblAktiv = lblRaschot
        Set lblNeaktiv = lblNastroyki
    Else
        Set lblAktiv = lblNastroyki
        Set lblNeaktiv = lblRaschot
    End If
    
    With lblAktiv
        .BackColor = &HFFFFFF
        .ForeColor = &H8000000D
        .Font.Name = "Segoe UI Semibold"
        .Top = LBL_TOP
        .Height = 21
    End With
    
    With lblNeaktiv
        .BackColor = &H8000000F
        .ForeColor = &H80000011
        .Font.Name = "Segoe UI"
        .Top = LBL_TOP + 3
        .Height = 18
    End With
        

    
End Sub






