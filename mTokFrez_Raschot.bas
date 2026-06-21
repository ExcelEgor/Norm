Attribute VB_Name = "mTokFrez_Raschot"
Option Explicit

Public Type RezultatyTokFrez
    Normy(1 To 3) As Double
    OpisaniyePerekhoda As String
End Type

Private Parametry As clsTokFrez
Private Rezultaty As RezultatyTokFrez
Private NekorrektnyyeZnacheniya As Boolean
Private frm As frmTokFrez
Private mlt As MSForms.MultiPage

Sub GlavnyyRaschot_TokFrez()

    If frmTokFrez.ZagruzkaFormy Then Exit Sub
    
    Call OchistitVremyaIPerekhod

    With frmTokFrez
        .Material = CInt(.cboMaterial.List(.cboMaterial.ListIndex, 1))
        .PoRezhimam = .chkPoRezhimam.Value
        .Tonkostennaya = .chkTonkosten
        .TipStanka = OpredelitTipsStanka
        .CHPU = .chkCHPU
    End With

    Set frm = frmTokFrez
    NekorrektnyyeZnacheniya = False
    Set Parametry = New clsTokFrez
    
    If frm.mltMain.SelectedItem.Name = "pRastFrez" Then
        Set mlt = frm.mltNormativRastFrez

    ElseIf frm.mltMain.SelectedItem.Name = "pTok" Then
        Set mlt = frm.mltNormativTok
    End If
    
    PrisvoitOsnovnyyeParametry
    
    Call RaschotRezultatov
    
    Rezultaty = Parametry.RaschotNormTokFrezIzParametrov

    ZapisatRezultaty
    

End Sub

Private Function OpredelitTipsStanka() As Integer

    Dim TipStanka As Integer
    Select Case frmTokFrez.tabForMlt.Value
    Case 0: TipStanka = 0
    Case 1: TipStanka = 2
    Case 2: TipStanka = 1
    End Select
    
    OpredelitTipsStanka = TipStanka
    
End Function

Private Sub OchistitVremyaIPerekhod()
    
    With frmTokFrez
        .tSles = 0
        .tKontrol = 0
        .tTokFrez = 0
        .txtTime = ""
        .txtPerekhod.text = Empty
    End With
        
End Sub

Private Function RaFromCbo(cbo As MSForms.ComboBox) As Double
    If cbo.ListIndex <> -1 And cbo.ColumnCount = 4 Then
        RaFromCbo = CDbl(Split(cbo.Column(1, cbo.ListIndex), "...")(0))
    End If
End Function

Private Sub PrisvoitOsnovnyyeParametry()

    Dim ctrl As Control
    
    With Parametry
    
        .id = OpredelitIdPerekhoda
        .TipStanka = frm.TipStanka
        .UchestKontrol = frm.chk_tKontr
        .UchestSlesar = frm.chk_tSlec
        .Material = frm.Material
        .PoRezhimam = frm.PoRezhimam
        .Tonkostennaya = frm.Tonkostennaya
        .CHPU = frm.CHPU

        For Each ctrl In mlt.SelectedItem.Controls
            If ctrl.Name Like "txtL_*" Then
                .L = DblFromCtrl(ctrl)
            ElseIf ctrl.Name Like "txtD1_*" Or ctrl.Name Like "txtD_*" Then
                .D1 = DblFromCtrl(ctrl)
            ElseIf ctrl.Name Like "txtd2_*" Then
                .d2 = DblFromCtrl(ctrl)
            ElseIf LCase(ctrl.Name) Like "txtkolvo_*" Then
                .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(ctrl)
            ElseIf LCase(ctrl.Name) Like "txtkolvoprohodov_*" Then
                .KolVoProkhodov = KolVoPoverhnosteiProhodov(ctrl)
            ElseIf ctrl.Name Like "chkUdar_*" Then
                .Udar = ctrl.Value
            ElseIf ctrl.Name Like "chkKorka_*" Then
                .Korka = ctrl.Value
            ElseIf ctrl.Name Like "optVUpor_*" Or ctrl.Name Like "chkVUpor_*" Then
                .VUpor = ctrl.Value
            ElseIf ctrl.Name Like "cboRaRz_*" Then
                .Ra = RaFromCbo(ctrl)
            End If
        Next
        
    End With
    
End Sub

Private Function OpredelitIdPerekhoda() As String

    Dim id As String
    id = mlt.SelectedItem.Name
    
    Select Case id
    Case "pObrabotkaOtverstiy"
        Select Case frm.cboPerehod_Sverl.ListIndex
        Case 0:     id = "pSverl_Tok"
        Case 1:     id = "pRassverl_Tok"
        Case 2:     id = "pZenker_Tok"
        Case 3, 4:  id = "pRazvert_Tok"
        End Select
    Case "pMetrRezb"
        With frm
            If .chkAvtoVybor Then
                id = "RezbaAvtoVybor"
            Else
                If .optPlashkoyMetchikom.Value = True Then
                    If .optNaruzhnaya_MetrRezb.Value = True Then
                        id = "pPlashka"
                    Else
                        id = "pMetchik"
                    End If
                End If
            End If
        End With
    End Select
    
    OpredelitIdPerekhoda = id
    
End Function

Private Sub ProveritNaMaksimalnoyeZnacheniye(txt As Control, ByVal Znacheniye As Double, ByVal MaksZnacheniye As Double, Optional lblDO As Control = Nothing)
    
    txt.BackColor = vbWhite
    
    If Not lblDO Is Nothing Then
        lblDO.Caption = "-"
        If MaksZnacheniye > 0 Then lblDO.Caption = "до " & MaksZnacheniye & " мм"
    End If
    
    If Znacheniye > 0 Then
        If Znacheniye > MaksZnacheniye And MaksZnacheniye <> 0 Then
            txt.BackColor = vbRed
            NekorrektnyyeZnacheniya = True
        End If
    End If
    
End Sub

Private Sub RaschotRezultatov()

    Select Case mlt.SelectedItem.Name
    
    Case "pUstRastFrez"
        PrisvoitParametry_UstRastFrez
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtMass_RastFrez, Parametry.Massa, MaksMassaUstanovki_RastFrez(Parametry.Sposob), frm.lblMaksMassa_RastFrez)
                
    Case "pFrezPloskKonc"
        PrisvoitParametry_FrezPloskKonc
            
    Case "pFrezPloskTorc"
        PrisvoitParametry_FrezPloskTorc
        
    Case "pFrezUstup"
        PrisvoitParametry_FrezUstup
                
    Case "pFrezPazov"
        PrisvoitParametry_FrezPazov
        Proverit_FrezPazov
            
    Case "pFrezShlic"
        PrisvoitParametry_FrezShlic
            
    Case "pFrezOtverstii"
        PrisvoitParametry_FrezOtverstii
        NastroitOtobrazheniye_FrezOtverstii
        Proverit_FrezOtverstii
           
    Case "pObrabotkaRezbovihOtv"
        PrisvoitParametry_ObrabotkaRezbovihOtv
        NastroitOtobrazheniye_ObrabotkaRezbovihOtv
        Proverit_ObrabotkaRezbovihOtv
            
    Case "pFrezOkon"
        PrisvoitParametry_FrezOkon
            
    Case "pFrezKnopki"
        PrisvoitParametry_FrezKnopki
            
    Case "pShestigrannik"
        PrisvoitParametry_Shestigrannik
                
        'Токарный
    Case "pUst"
        PrisvoitParametry_Ust
            
    Case "pOtrezka"
        PrisvoitParametry_Otrezka
            
    Case "pPoperechnoye"
        PrisvoitParametry_Poperechnoye
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Pop, Parametry.D1, DMAX_TOKARNYY)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtd2_Pop, Parametry.d2, Parametry.D1)
            
    Case "pProdolnoye"
        PrisvoitParametry_Prodolnoye
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Prod, Parametry.D1, IIf(Parametry.PoRezhimam, 0, DMAX_TOKARNYY), frm.lblD1_Prod_Dmax)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtd2_Prod, Parametry.d2, Parametry.D1)
            
    Case "pRastachivaniye"
        PrisvoitParametry_Rastachivaniye
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Rast, Parametry.D1, IIf(Parametry.PoRezhimam, 0, DMAX_TOKARNYY), frm.lblDmax_Rastachivaniye)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtd2_Rast, Parametry.d2, Parametry.D1)
            
    Case "pKonus"
        PrisvoitParametry_Konus
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Konus, Parametry.D1, DMAX_KONUS)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtL_Konus, Parametry.L, Lmax_Konus(Parametry.D1), frm.lblLmax_Konus)
            
    Case "pFaski"
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Faski, Parametry.D1, DMAX_TOKARNYY)
            
    Case "pGaltel"
        PrisvoitParametry_Galtel
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Galtel, Parametry.D1, DMAX_TOKARNYY)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtL_Galtel, Parametry.L, Lmax_Galtel(Parametry.Material, Parametry.D1), frm.lblRmax_Galtel)
            
    Case "pObrabotkaOtverstiy"
        PrisvoitParametry_ObrabotkaOtverstiy
            
    Case "pObrabOtvSNulya"
        PrisvoitParametry_ObrabOtvSNulya
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_ObrabOtvSNulya, Parametry.D1, IIf(Parametry.PoRezhimam, 0, DMAX_TOKARNYY), frm.lblDmax_ObrabOtvSNulya)
            
    Case "pNakativaniyeRifleniy"
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_NakatRifl, Parametry.D1, DMAX_NAKATYVANIYE)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtL_NakatRifl, Parametry.L, Lmax_NakativaniyeRifleniy(Parametry.D1), frm.lblLmax_NakatRifl)
        
    Case "pKanavka"
        PrisvoitParametry_Kanavka
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Kanav, Parametry.D1, DMAX_TOKARNYY)

    Case "pMetrRezb"
        PrisvoitParametry_MetrRezb
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_MetrRezb, Parametry.D1, DMAX_METRREZBAREZTSOM)
        OpredelitSposobNarezki_MetrRezb
            
    Case "pDovodka"
        PrisvoitParametry_Dovodka
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Dovod, Parametry.D1, DMAX_DOVODKA)
        
    Case "pTsentrovaniye"
        PrisvoitParametry_Tsentrovaniye
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD_Tsentrovaniye, Parametry.D1, DMAX_TOKARNYY)
            
    Case "pPritupleniye"
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_Pritupleniye, Parametry.D1, DMAX_TOKARNYY)
            
    Case "pPolirovaniyeTok"
        PrisvoitParametry_PolirovaniyeTok
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD_Polirov, Parametry.D1, DMAX_KRAMATORSK)
                
    End Select

End Sub

Private Sub PrisvoitParametry_UstRastFrez()

    With frm
        If .lstSposobUst_RastFrez.ListIndex = -1 Or .cboTochnostVyverki_RastFrez.ListIndex = -1 Or .cboKharakterVyverki_RastFrez.ListIndex = -1 Then Exit Sub
        Parametry.Sposob = .lstSposobUst_RastFrez.List(.lstSposobUst_RastFrez.ListIndex, .lstSposobUst_RastFrez.ColumnCount - 1)
        Parametry.Massa = DblFromCtrl(.txtMass_RastFrez)
        Parametry.KharkterVyverki = .cboKharakterVyverki_RastFrez.List(.cboKharakterVyverki_RastFrez.ListIndex, .cboKharakterVyverki_RastFrez.ColumnCount - 1)
        Parametry.TochnostVyverki = DblFromCtrl(.cboTochnostVyverki_RastFrez)
    End With

End Sub

Private Sub PrisvoitParametry_FrezPloskKonc()

    With frm
        Parametry.Pripusk = DblFromCtrl(.txtPripusk_FrezPloskKonc)
        Parametry.IT = CInt(.cboIT_FrezPloskKonc.text)
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_FrezPloskKonc)
        Parametry.AvtoRezhimy = .chkAvtoRezhimy_FrezPloskKonc.Value
        Parametry.Dfrezy = DblFromCtrl(.txtDfrezy_Konts)

    End With
    
    ZapisatRezhimy_FrezKonts
    
    Dim Rezhimy As New clsMillingRegimes
    With Rezhimy
            
        .ChernKork.Pokazateli.S = DblFromCtrl(frm.txtSchernKork_FrezKonts)
        .ChernKork.iB = DblFromCtrl(frm.txtIBChernKork_FrezKonts)
        .ChernKork.iH = DblFromCtrl(frm.txtIHChernKork_FrezKonts)
        .ChernKork.Bfrez = DblFromCtrl(frm.txtBChernKork_FrezKonts)
        .ChernKork.GlubRez = DblFromCtrl(frm.txtTChernKork_FrezKonts)
        .ChernKork.Pokazateli.N = DblFromCtrl(frm.txtNKork_FrezKonts)
        .ChernKork.Pokazateli.V = DblFromCtrl(frm.txtVKork_FrezKonts)

        .Chern.Pokazateli.S = DblFromCtrl(frm.txtSchern_FrezKonts)
        .Chern.iB = DblFromCtrl(frm.txtIBChern_FrezKonts)
        .Chern.iH = DblFromCtrl(frm.txtIHChern_FrezKonts)
        .Chern.Bfrez = DblFromCtrl(frm.txtBChern_FrezKonts)
        .Chern.GlubRez = DblFromCtrl(frm.txtTChern_FrezKonts)
        .Chern.Pokazateli.N = DblFromCtrl(frm.txtNchern_FrezKonts)
        .Chern.Pokazateli.V = DblFromCtrl(frm.txtVchern_FrezKonts)

        .Chist.Pokazateli.S = DblFromCtrl(frm.txtSchist_FrezKonts)
        .Chist.iB = DblFromCtrl(frm.txtIBChist_FrezKonts)
        .Chist.iH = DblFromCtrl(frm.txtIHChist_FrezKonts)
        .Chist.Bfrez = DblFromCtrl(frm.txtBChist_FrezKonts)
        .Chist.GlubRez = DblFromCtrl(frm.txtTChist_FrezKonts)
        .Chist.Pokazateli.N = DblFromCtrl(frm.txtNChist_FrezKonts)
        .Chist.Pokazateli.V = DblFromCtrl(frm.txtVChist_FrezKonts)
    End With
        
    Set Parametry.Rezhimy = Rezhimy

End Sub

Private Sub ZapisatRezhimy_FrezKonts()

    If frm.chkAvtoRezhimy_FrezPloskKonc.Value = True Then OchistitRezhimy Else Exit Sub

    frm.StopEvents = True

    Dim PloskKonts As New clsRezhimy_FrezKonts

    Dim Rezhimy As New clsMillingRegimes
    
    With Parametry
        Set Rezhimy = PloskKonts.RaschotRezhimov(.Material, .Pripusk, .Dfrezy, .ShirinaPoverhnosti, .IT, .Ra, .Korka)
        With frm
            
            .txtDfrezy_Konts.text = PloskKonts.OpredelitDfrezy_FrezPloskKonts(Parametry.ShirinaPoverhnosti)
            
            .txtSchernKork_FrezKonts = Rezhimy.ChernKork.Pokazateli.S
            .txtIBChernKork_FrezKonts = Rezhimy.ChernKork.iB
            .txtIHChernKork_FrezKonts = Rezhimy.ChernKork.iH
            .txtBChernKork_FrezKonts = Rezhimy.ChernKork.Bfrez
            .txtTChernKork_FrezKonts = Rezhimy.ChernKork.GlubRez
            .txtNKork_FrezKonts = Rezhimy.ChernKork.Pokazateli.N
            .txtVKork_FrezKonts = Rezhimy.ChernKork.Pokazateli.V
            
            .txtSchern_FrezKonts = Rezhimy.Chern.Pokazateli.S
            .txtIBChern_FrezKonts = Rezhimy.Chern.iB
            .txtIHChern_FrezKonts = Rezhimy.Chern.iH
            .txtBChern_FrezKonts = Rezhimy.Chern.Bfrez
            .txtTChern_FrezKonts = Rezhimy.Chern.GlubRez
            .txtNchern_FrezKonts = Rezhimy.Chern.Pokazateli.N
            .txtVchern_FrezKonts = Rezhimy.Chern.Pokazateli.V
            
            .txtSchist_FrezKonts = Rezhimy.Chist.Pokazateli.S
            .txtIBChist_FrezKonts = Rezhimy.Chist.iB
            .txtIHChist_FrezKonts = Rezhimy.Chist.iH
            .txtBChist_FrezKonts = Rezhimy.Chist.Bfrez
            .txtTChist_FrezKonts = Rezhimy.Chist.GlubRez
            .txtNChist_FrezKonts = Rezhimy.Chist.Pokazateli.N
            .txtVChist_FrezKonts = Rezhimy.Chist.Pokazateli.V
            
        End With
        
        Dim ctrl As Control
        For Each ctrl In frm.fraRezhimy_FrezKonts.Controls
            If TypeOf ctrl Is MSForms.TextBox Then
                If ctrl.text = 0 Then
                    ctrl.text = "-"
                End If
            End If
        Next
        
    End With
    
    frm.StopEvents = False
                
End Sub

Private Sub OchistitRezhimy()
    
    frm.StopEvents = True
    
    Dim ctrl As Control
    For Each ctrl In frm.fraRezhimy_FrezKonts.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.text = "-"
        End If
    Next
    
    frm.StopEvents = False
    
End Sub

Private Sub PrisvoitParametry_FrezPloskTorc()

    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_FrezPloskTorc)
        Parametry.Pripusk = DblFromCtrl(.txtPripusk_FrezPloskTorc)
        Parametry.IT = CInt(.cboIT_FrezPloskTorc.text)
    End With

End Sub

Private Sub PrisvoitParametry_FrezUstup()

    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_FrezUstup)
        Parametry.Pripusk = DblFromCtrl(.txtPripusk_FrezUstup)
        Parametry.IT = CInt(.cboIT_FrezUstup.text)
    End With
        
End Sub

Private Sub PrisvoitParametry_FrezPazov()

    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_FrezPazov)
        Parametry.GlubinaPaza = DblFromCtrl(.txtGlubina_FrezPazov)
        Parametry.ZachistkaBokovykhStoron = .chkZachistkaBokovykhStoron.Value
        Parametry.IT = CInt(.cboIT_FrezPazov.text)
    End With

End Sub

Private Sub Proverit_FrezPazov()

    With frm
        
        .lblMaksShirinaPaza.Visible = Not .chkZachistkaBokovykhStoron.Value
        .txtB_FrezPazov.BackColor = vbWhite
        
        If Parametry.ZachistkaBokovykhStoron = False Then
            If Parametry.ShirinaPoverhnosti > 0 And Parametry.ShirinaPoverhnosti > MAX_SHIRINA_KONTS_USTUP_PAZ Then
                .txtB_FrezPazov.BackColor = vbRed
                NekorrektnyyeZnacheniya = True
            End If
        End If
        
    End With
    
End Sub

Private Sub PrisvoitParametry_FrezShlic()
    
    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.cboB_FrezShlic)
        Parametry.GlubinaRezaniya = DblFromCtrl(.cboGlubinaRez_FrezShlic)
    End With
    
End Sub

Private Sub PrisvoitParametry_FrezOtverstii()

    With frm
        Parametry.DBObnizki = DblFromCtrl(.txtDBObniz_FrezOtv)
        Parametry.GlubinaObnizki = DblFromCtrl(.txtGlubinaObniz_FrezOtv)
        Parametry.IT = CInt(.cboIT_FrezOtv)
        Parametry.FrezPosleSverl = .chkFrezPosleSverl.Value
        Parametry.UchestFasku = .chkFaska_FrezOtv.Value
    End With

End Sub

Private Sub NastroitOtobrazheniye_FrezOtverstii()

    With frm
    
        .chkFrezPosleSverl.Enabled = True
        .cboRaRz_FrezOtv.Enabled = True
        .cboIT_FrezOtv.Enabled = True
        
        .lblLObnMax_ObrOtv.Caption = "-"
        .lblLmax_ObrOtv.Caption = "-"
        .txtL_FrezOtv.BackColor = vbWhite
        .txtGlubinaObniz_FrezOtv.BackColor = vbWhite
        .txtDBObniz_FrezOtv.BackColor = vbWhite

        If Parametry.D1 > MAKS_D_SVERL_FREZ_STANOK Then
            .chkFrezPosleSverl.Value = True
            .chkFrezPosleSverl.Enabled = False
        End If
        
        '        If Parametry.D1 <= MIN_D_FREZ Then
        '            .chkFrezPosleSverl.Value = False
        '            .chkFrezPosleSverl.Enabled = False
        '        End If
        
        If .chkFrezPosleSverl.Value = False Then
            .cboRaRz_FrezOtv.ListIndex = 0
            .cboIT_FrezOtv.ListIndex = .cboIT_FrezOtv.ListCount - 1
            .cboRaRz_FrezOtv.Enabled = False
            .cboIT_FrezOtv.Enabled = False
        End If
        
    End With
    
End Sub

Private Sub Proverit_FrezOtverstii()

    With frm
        
        Dim Lmax As Integer
        Lmax = Lmax_Sverleniye_Frezernyy(Parametry.D1, CInt(Parametry.TipStanka))

        Call ProveritNaMaksimalnoyeZnacheniye(.txtL_FrezOtv, Parametry.L, Lmax, .lblLmax_ObrOtv)
        
        If NekorrektnyyeZnacheniya = False And Parametry.L >= 0 Then

            If Parametry.DBObnizki > 0 Then
            
                If Parametry.DBObnizki <= Parametry.D1 Then
                    NekorrektnyyeZnacheniya = True
                    .txtDBObniz_FrezOtv.BackColor = vbRed
                End If

                Call ProveritNaMaksimalnoyeZnacheniye(.txtGlubinaObniz_FrezOtv, Parametry.GlubinaObnizki, Parametry.L, .lblLObnMax_ObrOtv)
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub NastroitOtobrazheniye_ObrabotkaRezbovihOtv()
    
    With frm

        .txtD1_FrezRezb.BackColor = vbWhite
        .lblDmax_FrezRezb.Caption = "до " & MAKS_D_REZBY_FREZERNYY & " мм"
        
        .txtL_FrezRezb.BackColor = vbWhite
        .lblLotvMax_FrezRezb.Caption = "-"
        
        .txtLrez_FrezRezb.BackColor = vbWhite
        .lblLrezMax_FrezRezb.Caption = "-"
        
        .txtDBObniz_FrezRezb.BackColor = vbWhite
        .txtGlubinaObniz_FrezRezb.BackColor = vbWhite
        
    End With

End Sub

Private Sub Proverit_ObrabotkaRezbovihOtv()

    If Parametry.D1 > 0 And Parametry.D1 <= MAKS_D_REZBY_FREZERNYY Then

        Dim LmaxOtv As Integer
        LmaxOtv = Lmax_Sverleniye_Frezernyy(Parametry.D1, CInt(frm.TipStanka))
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtL_FrezRezb, Parametry.L, LmaxOtv, frm.lblLotvMax_FrezRezb)
        
        Dim LmaxRezb As Double
        LmaxRezb = WorksheetFunction.Min(Lmax_Metchik(Parametry.D1), Parametry.L)
        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtLrez_FrezRezb, Parametry.Lrezby, LmaxRezb, frm.lblLrezMax_FrezRezb)

    Else

        Call ProveritNaMaksimalnoyeZnacheniye(frm.txtD1_FrezRezb, Parametry.D1, MAKS_D_REZBY_FREZERNYY)
        
    End If

    Proverit_ObrabotkaRezbovihOtv_Obnizka
   
End Sub

Private Sub Proverit_ObrabotkaRezbovihOtv_Obnizka()

    With frm
        If NekorrektnyyeZnacheniya = False And Parametry.L > 0 And Parametry.Lrezby > 0 And Parametry.DBObnizki > 0 Then
            If Parametry.DBObnizki <= Parametry.D1 Then
                NekorrektnyyeZnacheniya = True
                .txtDBObniz_FrezRezb.BackColor = vbRed
            Else
                Call ProveritNaMaksimalnoyeZnacheniye(.txtGlubinaObniz_FrezRezb, Parametry.GlubinaObnizki, Parametry.L)
            End If
        End If
    End With
    
End Sub

Private Sub PrisvoitParametry_ObrabotkaRezbovihOtv()

    With frm
        Parametry.Lrezby = DblFromCtrl(.txtLrez_FrezRezb)
        Parametry.DBObnizki = DblFromCtrl(.txtDBObniz_FrezRezb)
        Parametry.GlubinaObnizki = DblFromCtrl(.txtGlubinaObniz_FrezRezb)
        Parametry.IT = IIf(.tabIT_FrezRezb.Value = 0, 6, 7)
        Parametry.ShagRezby = DblFromCtrl(.cboS_FrezRezb)
        Parametry.KalibrRezbSlesar = .chkKalibr_FrezRezb.Value
        Parametry.FrezPosleSverl = .chkFrezPosleSverl_Rezb.Value
    End With
    
End Sub

Private Sub PrisvoitParametry_FrezOkon()

    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_FrezOkon)
        Parametry.Pripusk = DblFromCtrl(.txtPripusk_FrezOkon)
        If .chkObniz_FrezOkon Then
            Parametry.DlinaObnizki = DblFromCtrl(.txtLObniz_FrezOkon)
            Parametry.DBObnizki = DblFromCtrl(.txtDBObniz_FrezOkon)
            Parametry.GlubinaObnizki = DblFromCtrl(.txtGlubinaObniz_FrezOkon)
        End If
        Parametry.IT = CInt(.cboIT_FrezOkon.text)
        Parametry.TipStanka = frm.TipStanka
    End With
  
End Sub

Private Sub PrisvoitParametry_FrezKnopki()

    With frm
        Parametry.L = DblFromCtrl(.txtRazmer_Knopki)
        Parametry.Pripusk = DblFromCtrl(.txtTolshcina_Knopki)
    End With

End Sub

Private Sub PrisvoitParametry_Shestigrannik()

    With frm
        Parametry.ShagRezby = DblFromCtrl(.txtS_Shestigrannik)
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_Shestigrannik)
        Parametry.Massa = DblFromCtrl(.txtMassa_Shestigrannik)
        Parametry.IT = CInt(.cboIT_Shestigrannik)
    End With
    
End Sub

Private Sub PrisvoitParametry_Ust()

    With Parametry
        .Massa = DblFromCtrl(frm.txtMass_Tok)
        .Sposob = frm.lstSposobUst_Tok.ListIndex + 1
        .KharkterVyverki = frm.cboKharakterVyverki_Tok.ListIndex + 1
        .TochnostVyverki = DblFromCtrl(frm.cboTochnostVyverki_Tok)
    End With

End Sub

Private Sub PrisvoitParametry_Otrezka()

    With frm
        Parametry.L = 0.5 * (DblFromCtrl(.txtD1_Otrezka) - DblFromCtrl(.txtd2_Otrezka))
    End With
    
End Sub

Private Sub PrisvoitParametry_Poperechnoye()

    With frm
        Parametry.Pripusk = DblFromCtrl(.txtPripusk_Pop)
        Parametry.IT = CInt(.cboIT_Pop)
        Parametry.GlubinaRezaniya = DblFromCtrl(.cboGlubinaRez_Pop)
        Parametry.Vconst = .chkVconst_Poperech.Value
        Parametry.TocheniyeDiametra = .chkTocheniyeDiametra.Value
    End With
    
End Sub

Private Sub PrisvoitParametry_Prodolnoye()
    
    With frm
        Parametry.IT = CInt(.cboIT_Prod)
        Parametry.GlubinaRezaniya = DblFromCtrl(.cboGlubinaRez_Prod)
        Parametry.PodrezkaTortsa = .chkPodrezka_Prod.Value
    End With

End Sub

Private Sub PrisvoitParametry_Rastachivaniye()

    With frm
        Parametry.IT = CInt(.cboIT_Rast)
        Parametry.GlubinaRezaniya = DblFromCtrl(.cboGlubinaRez_Rast)
        Parametry.PodrezkaTortsa = .chkPodrezka_Rast.Value
    End With

End Sub

Private Sub PrisvoitParametry_Konus()
    
    With frm
        Parametry.TipPoverkhnosti = IIf(.optNaruzh_Konus, 1, 2)
        If .cboRaRz_Konus.ListIndex = 0 Then
            Parametry.Ra = 5
        Else
            Parametry.Ra = DblFromCtrl(.cboRaRz_Konus)
        End If
        Parametry.IT = CInt(.cboIT_Konus)
    End With

End Sub

Private Sub PrisvoitParametry_Galtel()

    With frm
        Select Case .cboRaRz_Galtel.ListIndex
        Case 0: Parametry.Ra = 10
        Case 1: Parametry.Ra = 1.25
        Case 2: Parametry.Ra = 0.63
        End Select
    End With

End Sub

Private Sub PrisvoitParametry_ObrabotkaOtverstiy()

    With frm
        If .cboPerehod_Sverl = "Развертывание H7" Then
            Parametry.Ra = 1.25
        Else
            Parametry.Ra = 2.5
        End If
    End With

End Sub

Private Sub PrisvoitParametry_ObrabOtvSNulya()

    With frm
        Parametry.IT = CInt(.cboIT_ObrabOtvSNulya.text)
        Parametry.GlubinaRezaniya = DblFromCtrl(.cboGlubinaRez_ObrabOtvSNulya)
    End With

End Sub

Private Sub PrisvoitParametry_Kanavka()

    With frm
        Parametry.ShirinaPoverhnosti = DblFromCtrl(.txtB_Kanav)
        Parametry.Ra = CDbl(.cboRaRz_Kanav.List(.cboRaRz_Kanav.ListIndex, 1))
        Parametry.TipPoverkhnosti = IIf(.optNaruzhnaya_Kanavka, 1, 2)
        Parametry.IT = CInt(.cboIT_Kanav)
    End With

End Sub

Private Sub PrisvoitParametry_MetrRezb()
    
    With frm
        Parametry.TipPoverkhnosti = IIf(.optNaruzhnaya_MetrRezb, 1, 2)
        Parametry.IT = IIf(.tabIT_MetrRezb.Value = 0, 6, 7)
        Parametry.ShagRezby = DblFromCtrl(.cboS_MetrRezb)
    End With
    
End Sub

Private Sub OpredelitSposobNarezki_MetrRezb()

    frm.optReztsom.Enabled = Not frm.chkAvtoVybor.Value
    frm.optPlashkoyMetchikom.Enabled = Not frm.chkAvtoVybor.Value

    Dim SposobNarezki As EnumSposobNarezkiRezby
       
    If frm.chkAvtoVybor.Value = True Then
        SposobNarezki = OpredelitSposobNarezkiRezbyNaTokarnomStanke(IIf(frm.optNaruzhnaya_MetrRezb, 1, 2), Parametry.D1)
        
        If SposobNarezki = Reztsom Then
            frm.optReztsom.Value = True
        
        ElseIf SposobNarezki = PlashkoyIliMetchikom Then
            frm.optPlashkoyMetchikom.Value = True
            
        End If
        
    End If
    
    If (Parametry.TipPoverkhnosti = 1 And Parametry.D1 > DMAX_METCHIK) Or (Parametry.TipPoverkhnosti = 2 And Parametry.D1 > DMAX_PLASHKA) Then
        frm.chkAvtoVybor.Value = False
        frm.chkAvtoVybor.Enabled = False
        frm.optReztsom.Value = True
        frm.optPlashkoyMetchikom.Enabled = False
    Else
        frm.chkAvtoVybor.Enabled = True
    End If

End Sub

Private Sub PrisvoitParametry_Dovodka()
    
    With frm
        Parametry.TipPoverkhnosti = .cboTipPov_Dovod.ListIndex + 1
        Parametry.Ra = DblFromCtrl(.cboRaRz_Dovod)
        Parametry.Otkloneniye = DblFromCtrl(.cboOtkl_Dovod)
    End With
    
End Sub

Private Sub PrisvoitParametry_Tsentrovaniye()
    
    With frm
        If .optTsentrovaniye_1 Then
            Parametry.TipTsentrovaniya = 1
        ElseIf .optTsentrovaniye_2 Then
            Parametry.TipTsentrovaniya = 2
        ElseIf .optTsentrovaniye_3 Then
            Parametry.TipTsentrovaniya = 3
        End If
        Parametry.Ra = DblFromCtrl(.cboRa_Tsentrovaniye)
    End With

End Sub

Private Sub PrisvoitParametry_PolirovaniyeTok()

    With frm
        Parametry.Ra = DblFromCtrl(.cboRaRz_Polirov)
        Select Case True
        Case .optNaruzh_Polirov: Parametry.TipPoverkhnosti = 1
        Case .optVnutr_Polirov: Parametry.TipPoverkhnosti = 2
        Case .optTortsov_Polirov: Parametry.TipPoverkhnosti = 3
        End Select
    End With
    
End Sub

Private Sub ZapisatRezultaty()

    If Rezultaty.Normy(1) = 0 Or NekorrektnyyeZnacheniya = True Then Exit Sub

    With frm
        
        .tTokFrez = Rezultaty.Normy(1)
        .tSles = Rezultaty.Normy(2)
        .tKontrol = Rezultaty.Normy(3)
        
        .txtPerekhod.text = Rezultaty.OpisaniyePerekhoda
        
        .txtTime = Format(.tTokFrez, "0.00")
        
        If .tKontrol > 0 And Parametry.UchestKontrol = True Then
            .txt_tKontr.text = Format(.tKontrol, "0.00")
        Else
            .txt_tKontr.text = "-"
        End If
        
        If .tSles > 0 And Parametry.UchestSlesar = True Then
            .txt_tSles.text = Format(.tSles, "0.00")
        Else
            .txt_tSles.text = "-"
        End If
        
    End With
    
    Set frm.Parametry_TokarnoFrezer = Parametry
    
End Sub

