VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGidroabraziv 
   Caption         =   "ГИДРОАБРАЗИВНАЯ РЕЗКА"
   ClientHeight    =   12045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "frmGidroabraziv.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmGidroabraziv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const KOLVO_STROK As Integer = 5

Dim ArrTxtGidroabraziv() As New clsmZapuskRaschetov
Dim ArrCboGidroabraziv() As New clsmZapuskRaschetov

Private Enum TipKontura
    ePerimetr = 1
    eDiametr = 2
    eProizvolnyy = 3
End Enum

Private Type GabarityGidroabraziv
    Dlina As Double
    Shirina As Double
    Diametr As Double
    Tolshchina As Double
End Type

Private Type ParametryVyrezki
    a As Double
    b As Double
    d As Double
    DlinaProizvolnogoKontura As Double
    KolVo As Double
End Type

Private Type ResultatyVyrezki
    tSlesar As Double
    tKontrol As Double
    DlinaRezaniya As Double
    tVisual As Double
    tUpakovka As Double
    tPeskostruy As Double
End Type

Private Type ItogovyyeResultaty
    tGidroabraziv As Double
    tKontrol As Double
    tUpakovyvaniye As Double
    tVisual As Double
    tSlesar As Double
    tPeskostruy As Double
End Type


Private Sub mltTipZagotovki_Change(): Call GlavnyyRaschot: End Sub

Private Sub chkPravka_Click(): Call GlavnyyRaschot: End Sub

Private Sub optRz20_Click(): Call GlavnyyRaschot: End Sub

Private Sub optRz40_Click(): Call GlavnyyRaschot: End Sub

Private Sub cmdFrezerovatGabarity_Click(): Call FrezerovatGabarity(RaschotGabaritovOsnovnoyVyrezki(OpredelitTipZagotovki)): End Sub

Private Sub cboMaterial_Change(): Call GlavnyyRaschot: End Sub

Private Sub cmdClear_Click(): Call OchistitRaschot: End Sub

Private Sub UserForm_Initialize()

    Call SozdatTablitsuRaschotaDopontilenoyVyrezki


    Dim MATERIALY
    MATERIALY = Array(EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.STAL_UGLERODISTAYA, _
        EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.TITANOVYYE_SPLAVY, EnumMaterialy.POLIAMID)
    
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    

    Dim ctrl As Control
    Dim iCbo As Integer, iTxt As Integer
    
    For Each ctrl In Me.Controls
        
        Select Case TypeName(ctrl)
            Case "ComboBox"
                If ctrl.Name Like "cboTip*" Then
                    ctrl.List = Split("Периметр Контур Отверстие")
                End If
                ReDim Preserve ArrCboGidroabraziv(iCbo)
                Set ArrCboGidroabraziv(iCbo).cboGidroabraziv = ctrl
                iCbo = iCbo + 1
            Case "TextBox"
                If Not ctrl.Name Like "txtVremya*" And ctrl.Name <> "txtMassa" And ctrl.Name <> "txtPodacha" Then
                    ReDim Preserve ArrTxtGidroabraziv(iTxt)
                    Set ArrTxtGidroabraziv(iTxt).txtGidroabraziv = ctrl
                    iTxt = iTxt + 1
                End If
        End Select
        
        If ctrl.Name Like "*_*" And (TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "TextBox") Then Call VyravnivanieElementa(ctrl)
        
    Next
    
    Call VyravnitElementy
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub

Private Sub VyravnitElementy()

    fraZagotov.Top = 3
    fraZagotov.Left = 3
    
    fraNaruzh.Left = 3
    fraNaruzh.Top = fraZagotov.Top + fraZagotov.Height + 6
    
    mltTipZagotovki.Width = fraNaruzh.Width - 9 - mltTipZagotovki.Left
    
    fraVnutr.Left = 3
    fraVnutr.Top = fraNaruzh.Top + fraNaruzh.Height + 6
    
    Dim ctrl As Control
    Set ctrl = fraGidro.Controls("txtA_" & KOLVO_STROK)
    
    fraGidro.Left = 3
    fraGidro.Width = fraVnutr.Width - fraGidro.Left - 6
    fraGidro.Height = ctrl.Top + ctrl.Height + 9
    
    fraVnutr.Height = fraGidro.Height + fraGidro.Top + 15
    
    fraRaschot.Top = fraZagotov.Top
    fraRaschot.Left = fraZagotov.Left + fraZagotov.Width + 6
    fraRaschot.Height = fraVnutr.Top + fraVnutr.Height - fraRaschot.Top
    
    Me.Width = fraRaschot.Left + fraRaschot.Width + 9
    Me.Height = fraVnutr.Top + fraVnutr.Height + 30
    
    VyravnitPoLevomuKrayuNazvaniyaVkladok mltTipZagotovki
    
End Sub

Private Sub SozdatTablitsuRaschotaDopontilenoyVyrezki()

    Dim cboTip As MSForms.ComboBox
    Dim txtA As MSForms.TextBox, txtB As MSForms.TextBox, txtKontur As MSForms.TextBox, txtOtv As MSForms.TextBox, txtKolVo As MSForms.TextBox
    
    Dim i As Integer
    For i = 2 To KOLVO_STROK
    
        Set cboTip = fraGidro.Controls.Add("Forms.ComboBox.1", "cboTip_" & i)
        cboTip.Style = fmStyleDropDownList
        
        Set txtA = fraGidro.Controls.Add("Forms.TextBox.1", "txtA_" & i)
        Set txtB = fraGidro.Controls.Add("Forms.TextBox.1", "txtB_" & i)
        Set txtKontur = fraGidro.Controls.Add("Forms.TextBox.1", "txtKontur_" & i)
        Set txtOtv = fraGidro.Controls.Add("Forms.TextBox.1", "txtOtv_" & i)
        Set txtKolVo = fraGidro.Controls.Add("Forms.TextBox.1", "txtKolvo_" & i)
        
    Next
    
End Sub

Private Sub VyravnivanieElementa(ctrl As Control)
    
    Dim num As Integer
    num = NumControl(ctrl)
    
    If num > 1 Then
    
        Dim prevCtrl As Control
        
        With ctrl
            
            Set prevCtrl = Me.Controls(Left(.Name, Len(.Name) - Len(CStr(num))) & 1)
            
            .ForeColor = &H404040
            .SpecialEffect = fmSpecialEffectEtched
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = &HC0C0C0
            .Width = prevCtrl.Width
            .Left = prevCtrl.Left
            .Top = (prevCtrl.Top + (prevCtrl.Height * (num - 1))) + 3 * (num - 1)
            .SelectionMargin = False
            If Not (.Name Like "txtKolvo*" Or .Name Like "txtVremya*" Or .Name Like "cboTip*") Then .Enabled = False
            
        End With
        
    End If
    
    Me.Repaint
    
End Sub

Sub GlavnyyRaschot()

    lstOper.Clear
    txtMassa.text = Empty
    IzmenitTsvetRamkiElemetovDopolnitelnoyVyrezki
    
    Dim Material As EnumMaterialy
    Material = MaterialFromCboMaterial(cboMaterial)
    Call IzmenitTsvetFormy(Material)
    
    Dim TipZagotovki As TipKontura
    TipZagotovki = OpredelitTipZagotovki
    
    Dim Gabarity As GabarityGidroabraziv
    Gabarity = RaschotGabaritovOsnovnoyVyrezki(TipZagotovki)

    If Gabarity.Tolshchina <= 0 Then Exit Sub
    
    Dim Massa As Double
    Massa = RaschotMassy(Material, Gabarity.Tolshchina, TipZagotovki)

    If Massa > 0 Then txtMassa = Format(Massa, "0.000")
    
    Call ZapisatMinutnuyuPodachu(Material, Gabarity.Tolshchina)
    
    Dim tOsnovnoyVyrezki As ResultatyVyrezki
    tOsnovnoyVyrezki = RaschotOsnovnoyVyrezki(Gabarity.Tolshchina, Gabarity, Material, Massa, TipZagotovki)
    
    Dim tDopolnitelnoyVyrezky As ResultatyVyrezki
    tDopolnitelnoyVyrezky = RaschotDopolnitelnoyVyrezki(Material, Gabarity.Tolshchina, Massa)
    
    Dim DlinaRezaniya_Summarnaya As Double
    DlinaRezaniya_Summarnaya = tOsnovnoyVyrezki.DlinaRezaniya + tDopolnitelnoyVyrezky.DlinaRezaniya
    
    If DlinaRezaniya_Summarnaya <= 0 Then Exit Sub
    
    Dim tUst As Double
    tUst = UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(Massa) + UstanovkaSlesar_NaStoleBezKrepleniya_Povorot(Massa)
    
    Dim Resultaty As ItogovyyeResultaty
    With Resultaty
        .tGidroabraziv = GidrAbraziv_Vremya(Material, Gabarity.Tolshchina, DlinaRezaniya_Summarnaya, KolVoPoverhnosteiProhodov(txtKolVoUglov), optRz20.Value = True)
        .tKontrol = 1.3 * (tOsnovnoyVyrezki.tKontrol + tDopolnitelnoyVyrezky.tKontrol)
        .tPeskostruy = 1.3 * tOsnovnoyVyrezki.tPeskostruy
        .tSlesar = 1.3 * (tUst + tOsnovnoyVyrezki.tSlesar + tDopolnitelnoyVyrezky.tSlesar)
        .tUpakovyvaniye = 1.3 * tOsnovnoyVyrezki.tUpakovka
        .tVisual = 1.3 * tOsnovnoyVyrezki.tVisual
    End With

    Call ZapolnitTablitsuRashcotnymyZnacheniyami(Resultaty)

End Sub

Private Function OpredelitTipZagotovki() As TipKontura

    Select Case mltTipZagotovki.SelectedItem.Name
        Case "pDiametr"
            OpredelitTipZagotovki = eDiametr
        Case "pPerimetr"
            OpredelitTipZagotovki = ePerimetr
        Case "pKontur"
            OpredelitTipZagotovki = eProizvolnyy
    End Select
    
End Function

Private Sub ZapisatMinutnuyuPodachu(Material As EnumMaterialy, Tolshchina As Double)
    
    Dim Smm_min As Double
    
    txtPodacha.text = Empty
    If Tolshchina > 0 Then
        Smm_min = GidrAbrazivSmm_min(Material, Tolshchina, optRz20.Value = True)
        txtPodacha = Format(Smm_min, "0.00")
    End If
    
End Sub

Private Function RaschotMassy(Material As EnumMaterialy, Tolshchina As Double, TipZagotovki As TipKontura)

    
    Dim MassaGabaritov As Double
    If TipZagotovki = eDiametr Then
        MassaGabaritov = MassaKruga(DblFromCtrl(txtD_Diametr), Tolshchina, CInt(Material))
    ElseIf TipZagotovki = ePerimetr Then
        MassaGabaritov = MassaLista(DblFromCtrl(txtA_Perimetr), DblFromCtrl(txtB_Perimetr), Tolshchina, CInt(Material))
    ElseIf TipZagotovki = eProizvolnyy Then
        MassaGabaritov = MassaLista(DblFromCtrl(txtA_Kontur), DblFromCtrl(txtB_Kontur), Tolshchina, CInt(Material))
    End If
    
    Dim VychetaemayaMassa(1 To KOLVO_STROK)
    Dim Kontur As TipKontura
    Dim Parametry As ParametryVyrezki

    Dim i As Integer
    For i = 1 To KOLVO_STROK

        Kontur = OpredelitTipKontura(Me.Controls("cboTip_" & i))
         
        Parametry = PoluchitParametryVyrezki(i)
        
        With Parametry
        
            If Kontur = ePerimetr Then
                If .a > 0 And .b > 0 Then
                    VychetaemayaMassa(i) = .KolVo * MassaLista(.a, .b, Tolshchina, CInt(Material))
                End If
                
            ElseIf Kontur = eProizvolnyy Then
                If .DlinaProizvolnogoKontura > 0 Then
                    VychetaemayaMassa(i) = .KolVo * 0.5 * MassaLista(Kontur / 4, Kontur / 4, Tolshchina, CInt(Material))
                End If
                
            ElseIf Kontur = eDiametr Then
                If .d > 0 Then
                    VychetaemayaMassa(i) = .KolVo * MassaKruga(.d, Tolshchina, CInt(Material))
                End If
                
            Else
                VychetaemayaMassa(i) = 0
            End If
        
        End With
            
    Next
    
    RaschotMassy = MassaGabaritov - WorksheetFunction.Sum(VychetaemayaMassa)

End Function
 
Private Function RaschotGabaritovOsnovnoyVyrezki(TipZagotovki As TipKontura) As GabarityGidroabraziv
    
    With RaschotGabaritovOsnovnoyVyrezki
    
        If TipZagotovki = eDiametr Then
            .Diametr = DblFromCtrl(txtD_Diametr)
            .Shirina = .Diametr
        ElseIf TipZagotovki = ePerimetr Then
            .Dlina = DblFromCtrl(txtA_Perimetr)
            .Shirina = DblFromCtrl(txtB_Perimetr)
        ElseIf TipZagotovki = eProizvolnyy Then
            .Dlina = DblFromCtrl(txtA_Kontur)
            .Shirina = DblFromCtrl(txtB_Kontur)
        End If
        
        .Tolshchina = DblFromCtrl(txtTolshchina)
    
    End With
    
End Function

Private Function RaschotOsnovnoyVyrezki(Tolshchina As Double, Gabarity As GabarityGidroabraziv, Material As EnumMaterialy, Massa As Double, TipZagotovki As TipKontura) As ResultatyVyrezki
    
    Dim tZausentsy As Double, tPravka As Double
    
    With RaschotOsnovnoyVyrezki

        If TipZagotovki = eDiametr Then
            .DlinaRezaniya = PI * Gabarity.Diametr
            .tKontrol = IzmerShtangenCircul(Gabarity.Diametr, Tolshchina)
            
            .tUpakovka = UpakovyvaniyeKruglogoPrutka(Gabarity.Diametr, Tolshchina)
            .tVisual = VisualnyyKontrol(Gabarity.Diametr, Tolshchina, 1)
            tZausentsy = ZachistkaZausencev_PoKonturu_Napilnikom(CInt(Material), 2 * .DlinaRezaniya, False, Massa > 20)
            
        Else
        
            .tUpakovka = UpakovyvaniyePramougolnika(Gabarity.Dlina, Gabarity.Shirina, Tolshchina)
            .tVisual = VisualnyyKontrol(Gabarity.Dlina, Gabarity.Shirina, 1)
        
            If TipZagotovki = ePerimetr Then
                .DlinaRezaniya = 2 * (Gabarity.Dlina + Gabarity.Shirina)
                .tKontrol = IzmerShtangenCircul_CPL(Gabarity.Dlina, Gabarity.Shirina)
                tZausentsy = SnyatiyeZausentsevPoKonturu(4 * (Gabarity.Dlina + Gabarity.Shirina) + 4 * Tolshchina, Material, False, NAPILNIK_SHABER, False)
                
            ElseIf TipZagotovki = eProizvolnyy Then
                .DlinaRezaniya = DblFromCtrl(txtL_Kontur)
                
                If DblFromCtrl(txtKolVoRazmerov) > 2 Then
                    .tKontrol = IzmerShtangenCircul(Gabarity.Dlina) + IzmerShtangenCircul(Gabarity.Shirina) * (DblFromCtrl(txtKolVoRazmerov) - 1)
                Else
                    .tKontrol = IzmerShtangenCircul_CPL(Gabarity.Dlina, Gabarity.Shirina)
                End If
                
                tZausentsy = ZachistkaZausencev_PoKonturu_Napilnikom(CInt(Material), 2 * .DlinaRezaniya, False, Massa > 20)
                
            End If
            
        End If
        
        Const kPravka As Double = 2.7 'Опытно-статистический коэффициент на правку
    
        If chkPravka Then
            If TipZagotovki = eDiametr Then
                tPravka = kPravka * PravkaZagIzListMatVRuch(Gabarity.Diametr, Gabarity.Diametr, Tolshchina, CInt(Material))
            Else
                tPravka = kPravka * PravkaPloskikhDetaley(Tolshchina, Gabarity.Dlina, Gabarity.Shirina, CInt(Material))
            End If
        End If
        
        .tSlesar = tZausentsy + tPravka
        
        .tPeskostruy = PeskostruynayaOchistka(Gabarity.Shirina, Gabarity.Dlina, 2, 2, Massa, Material)

    End With
    
End Function


Private Function RaschotDopolnitelnoyVyrezki(Material As EnumMaterialy, Tolshchina As Double, Massa As Double) As ResultatyVyrezki
        
    Dim tSlesar(1 To KOLVO_STROK) As Double, tKontrol(1 To KOLVO_STROK) As Double, DlinaRezaniya(1 To KOLVO_STROK) As Double

    If Tolshchina <= 0 Then Exit Function

    Dim Kontur As TipKontura
    Dim Parametry As ParametryVyrezki
    
    Dim i As Integer
    For i = 1 To KOLVO_STROK

        Kontur = OpredelitTipKontura(Me.Controls("cboTip_" & i))
         
        Parametry = PoluchitParametryVyrezki(i)
        
        With Parametry
        
            If Kontur = ePerimetr Then
                If .a > 0 And .b > 0 Then
                    DlinaRezaniya(i) = .KolVo * 2 * (.a + .b)
                    tSlesar(i) = .KolVo * ZachistkaZausencev_PoKonturu_Napilnikom(CInt(Material), 4 * (.a + .b), False, Massa > 20)
                    tKontrol(i) = .KolVo * IzmerShtangenCircul_CPL(.a, .b)
                End If
                
            ElseIf Kontur = eProizvolnyy Then
                If .DlinaProizvolnogoKontura > 0 Then
                    DlinaRezaniya(i) = .KolVo * .DlinaProizvolnogoKontura
                    tSlesar(i) = .KolVo * ZachistkaZausencev_PoKonturu_Napilnikom(CInt(Material), 2 * .DlinaProizvolnogoKontura, False, Massa > 20)
                    tKontrol(i) = .KolVo * IzmerShtangenCircul_CPL(.a / 4, .b / 4)
                End If
                
            ElseIf Kontur = eDiametr Then
                If .d > 0 Then
                    DlinaRezaniya(i) = .KolVo * .d * PI
                    tSlesar(i) = .KolVo * ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), .d, 2, 2, Massa > 20)
                    If .d > DMAX_IZMERENIE_PROBKOI Then
                        tKontrol(i) = .KolVo * IzmerShtangenCircul(.d, Tolshchina)
                    Else
                        tKontrol(i) = .KolVo * IzmerenieOtverstiiProbkoi(.d, Tolshchina)
                    End If
                End If
                
            Else
                DlinaRezaniya(i) = 0
                tSlesar(i) = 0
                tKontrol(i) = 0
            End If
        
        End With
            
    Next
    
    With RaschotDopolnitelnoyVyrezki
        .DlinaRezaniya = WorksheetFunction.Sum(DlinaRezaniya)
        .tKontrol = WorksheetFunction.Sum(tKontrol)
        .tSlesar = WorksheetFunction.Sum(tSlesar)
    End With

End Function

Private Sub ZapolnitTablitsuRashcotnymyZnacheniyami(Resultaty As ItogovyyeResultaty)

    Dim Operations As Variant
    
    With Resultaty
        Operations = Array( _
            Array("Подготов. для гидроабразивной резки", 10, 0), _
            Array("Гидроабразивная", 20, .tGidroabraziv), _
            Array("Контрольная", 0, .tKontrol), _
            Array("Маркирование (этикеткой)", 0, 3), _
            Array("Упаковывание", 0, .tUpakovyvaniye), _
            Array("Контрольная", 0, .tVisual), _
            Array("Пескоструйная", 5, .tPeskostruy), _
            Array("Контрольная", 0, .tVisual), _
            Array("Слесарная", 5, .tSlesar))
    End With

    Call ZapolnitListBoksIzMassiva(Operations, lstOper)
    
End Sub

Private Sub imgStenka_Click()
    With fraStenka
        If .Left = 129 Then
            .Left = 210
            .Height = "17,25"
            .Width = "17,25"
        Else
            .Height = "233,95"
            .Width = "461,95"
            .Left = 129
        End If
    End With
End Sub

Private Sub FrezerovatGabarity(Gabarity As GabarityGidroabraziv)

    Dim Tolshchina As Double
    Tolshchina = DblFromCtrl(txtTolshchina)

    Dim i As Integer
    Dim ArrGabarity(1 To 3) As Double
    Dim a As Double, b As Double, H As Double
    
    If Not Tolshchina > 0 Then Exit Sub
    
    If mltTipZagotovki.SelectedItem.Name = "pDiametr" Then
        If Not Gabarity.Diametr > 0 Then Exit Sub
        
        a = Gabarity.Diametr
        b = Gabarity.Diametr
    Else
        If Not Gabarity.Dlina > 0 Or Not Gabarity.Shirina > 0 Then Exit Sub
        
        a = Gabarity.Dlina
        b = Gabarity.Shirina
    End If
    
    For i = 1 To 3
        ArrGabarity(i) = WorksheetFunction.Large(Array(a, b, Tolshchina), i)
    Next
    
    a = ArrGabarity(1)
    b = ArrGabarity(2)
    H = ArrGabarity(3)
        
    If a * b * H > 0 Then

        With frmFrezGabarit
    
            Dim ListIndexMaterial As Integer
            For ListIndexMaterial = 0 To .cboMaterial.ListCount - 1
                If .cboMaterial.List(ListIndexMaterial, 1) = cboMaterial.List(cboMaterial.ListIndex, 1) Then
                    .cboMaterial.ListIndex = ListIndexMaterial
                    Exit For
                End If
            Next
        
            .txtA_zag = CStr(ArrGabarity(2))
            .txtB_zag = CStr(ArrGabarity(1))
            .txtH_zag = CStr(ArrGabarity(3))
            .txtA_det = CStr(ArrGabarity(2) - 5)
            .txtB_det = CStr(ArrGabarity(1) - 5)
            .txtH_det = CStr(ArrGabarity(3) - 5)
            .Show
        End With
        
    End If

End Sub

Private Function PoluchitParametryVyrezki(i As Integer) As ParametryVyrezki
    
    Dim Parametry As ParametryVyrezki
    With Parametry
        .a = DblFromCtrl(Me.Controls("txtA_" & i))
        .b = DblFromCtrl(Me.Controls("txtB_" & i))
        .DlinaProizvolnogoKontura = DblFromCtrl(Me.Controls("txtKontur_" & i))
        .d = DblFromCtrl(Me.Controls("txtOtv_" & i))
        .KolVo = KolVoPoverhnosteiProhodov(Me.Controls("txtKolvo_" & i))
    End With
         
    PoluchitParametryVyrezki = Parametry
    
End Function

Private Function OpredelitTipKontura(cboTip As MSForms.ComboBox) As TipKontura
    Select Case cboTip.text
        Case "Периметр"
            OpredelitTipKontura = ePerimetr
        Case "Контур"
            OpredelitTipKontura = eProizvolnyy
        Case "Отверстие"
            OpredelitTipKontura = eDiametr
    End Select
End Function


Private Sub IzmenitTsvetFormy(Material As EnumMaterialy)

    If cboMaterial.ListIndex = 0 Then
        optRz20.Value = True
        optRz40.Enabled = False
    Else
        optRz40.Value = True
        optRz40.Enabled = True
    End If

    Dim MaterialColor As Double
    MaterialColor = TsvetMateriala(Material)

    lblTipObrabotki.BackColor = MaterialColor
    lblA.BackColor = MaterialColor
    lblB.BackColor = MaterialColor
    lblL.BackColor = MaterialColor
    lblD.BackColor = MaterialColor
    lblKolVo.BackColor = MaterialColor
    lblOperatsiya.BackColor = MaterialColor
    lblTpz.BackColor = MaterialColor
    lblTsht.BackColor = MaterialColor
    cmdFrezerovatGabarity.BackColor = MaterialColor

    
End Sub

Private Sub OchistitRaschot()

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "ComboBox" Then
            If ctrl.Name = cboMaterial.Name Then
                ctrl.ListIndex = 0
            Else
                ctrl.ListIndex = -1
            End If
        ElseIf TypeName(ctrl) = "TextBox" Then
            If ctrl.Name = txtKolVoUglov.Name Then
                ctrl.text = 4
            Else
                ctrl.text = Empty
            End If
        End If
    Next
    
    optRz40.Value = True
    
    Call IzmenitTsvetRamkiElemetovDopolnitelnoyVyrezki

End Sub

Private Sub IzmenitTsvetRamkiElemetovDopolnitelnoyVyrezki()
    
    Dim i As Integer
    For i = 1 To KOLVO_STROK
        
        Dim Kontur As TipKontura
        Kontur = OpredelitTipKontura(Me.Controls("cboTip_" & i))

        If Kontur = ePerimetr Then
            Call ActiveBorderColor(Controls("txtA_" & i))
            Call ActiveBorderColor(Controls("txtB_" & i))
            Call InactiveBorderColor(Controls("txtKontur_" & i))
            Call InactiveBorderColor(Controls("txtOtv_" & i))
            
        ElseIf Kontur = eProizvolnyy Then
            Call InactiveBorderColor(Controls("txtA_" & i))
            Call InactiveBorderColor(Controls("txtB_" & i))
            Call ActiveBorderColor(Controls("txtKontur_" & i))
            Call InactiveBorderColor(Controls("txtOtv_" & i))
            
        ElseIf Kontur = eDiametr Then
            Call InactiveBorderColor(Controls("txtA_" & i))
            Call InactiveBorderColor(Controls("txtB_" & i))
            Call InactiveBorderColor(Controls("txtKontur_" & i))
            Call ActiveBorderColor(Controls("txtOtv_" & i))
            
        Else
            Call InactiveBorderColor(Controls("txtA_" & i))
            Call InactiveBorderColor(Controls("txtB_" & i))
            Call InactiveBorderColor(Controls("txtKontur_" & i))
            Call InactiveBorderColor(Controls("txtOtv_" & i))
        End If
   
    Next

End Sub

Private Sub InactiveBorderColor(txt As MSForms.TextBox)
    With txt
        .BorderColor = &HC0C0C0
        .Enabled = False
        .text = Empty
    End With
End Sub

Private Sub ActiveBorderColor(txt As MSForms.TextBox)
    With txt
        .BorderColor = &H8000000D
        .Enabled = True
    End With
End Sub













