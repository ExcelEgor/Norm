Attribute VB_Name = "mTokFrez_Redaktirovaniye"
Option Explicit
Dim Parametry As clsTokFrez

Sub Redaktirovaniye_TokFrez()
    
    If frmTokFrez.lstTokarnoFrezer.ListIndex = -1 Then Exit Sub

    frmTokFrez.RezhimRedaktirovaniya = True
    
    Dim StrokaRedaktirovaniya As Integer
    StrokaRedaktirovaniya = frmTokFrez.lstTokarnoFrezer.ListIndex + 1
    
    Set Parametry = frmTokFrez.Raschot.ZapolnitParametryIzStrokiTablitsy(StrokaRedaktirovaniya)

    Call OtkrytNuzhnyyNormativ
            
    Dim pageName As String
    pageName = OpredelitImyaVkladki()
    
    Call OtkrytNuzhnuyuVkladku(frmTokFrez.mltNormativ, pageName)

    Call ZapolnitElemtyUpravleniya(pageName)

    Call frmTokFrez.IzmenitRezhimFormy

    Call GlavnyyRaschot_TokFrez
    
    frmTokFrez.lstTokarnoFrezer.ListIndex = StrokaRedaktirovaniya - 1
    
    frmTokFrez.txtPerekhod.text = Parametry.OpisaniyePerekhoda

End Sub

Private Sub OtkrytNuzhnyyNormativ()

    Dim TipStanka As Integer
    TipStanka = Parametry.TipStanka

    With frmTokFrez
    
        Select Case TipStanka
            Case 1, 2
            
                .mltMain.Value = .mltMain.Pages("pRastFrez").Index
                Set .mltNormativ = .mltNormativRastFrez
                
                If TipStanka = 1 Then
                    .tabForMlt.Value = 2
                Else
                    .tabForMlt.Value = 1
                End If
                
            Case Else
            
                .mltMain.Value = .mltMain.Pages("pTok").Index
                Set .mltNormativ = .mltNormativTok
                
        End Select

    End With

End Sub

Private Sub OtkrytNuzhnuyuVkladku(mltNormativ As MSForms.MultiPage, pageName As String)
    
    Select Case pageName
        Case "RezbaAvtoVybor", "pPlashka", "pMetchik"
            mltNormativ.Value = mltNormativ.Pages("pMetrRezb").Index
        Case Else
            Dim i As Integer
            For i = 0 To mltNormativ.Pages.Count - 1
                If mltNormativ.Pages(i).Name = pageName Then
                    mltNormativ.Value = mltNormativ.Pages(pageName).Index
                    Exit For
                End If
            Next
    End Select
    
End Sub

Private Function OpredelitImyaVkladki() As String
    
    Dim pageName As String
    If Parametry.id Like "pUstTok_*" Then
        pageName = "pUst"
    Else
        Select Case Parametry.id
            Case "pSverl_Tok", "pRassverl_Tok", "pZenker_Tok", "pRazvert_Tok"
                pageName = "pObrabotkaOtverstiy"
            Case Else
                pageName = Parametry.id
        End Select
    End If
    
    OpredelitImyaVkladki = pageName
    
End Function

Private Sub ZapolnitElemtyUpravleniya(pageName As String)
    
    Select Case pageName
        Case "pUst"
            Call ZapolnitUstanovkuTokar
        
        Case "pUstRastFrez"
            Call ZapolnitUstanovkuRastFrez
        
        Case "pObrabotkaOtverstiy"
            Call ZapolnitObrabotkuOtverstiyTokar
        
        Case "pTsentrovaniye"
            Call ZapolnitTsentrovaniye
            
        Case "pMetrRezb"
            Call ZapolnitNarezaniyeRezby
            Call ZapolnitOstalnoye

        Case Else
            Call ZapolnitOstalnoye
                
    End Select
    
    frmTokFrez.chk_tKontr = Parametry.UchestKontrol
    frmTokFrez.chk_tSlec = Parametry.UchestSlesar
    
End Sub

Private Sub ZapolnitUstanovkuTokar()

    With frmTokFrez
        .txtMass_Tok.text = CStr(Parametry.Massa)
        .lstSposobUst_Tok.ListIndex = Parametry.Sposob - 1
        .cboKharakterVyverki_Tok.ListIndex = Parametry.KharkterVyverki - 1
        .cboTochnostVyverki_Tok.text = IIf(Parametry.TochnostVyverki = 0, "-", Parametry.TochnostVyverki)
        .txtKolVoProhodov_Ust = Parametry.KolVoProkhodov
    End With
    
End Sub

Private Sub ZapolnitTsentrovaniye()

    With frmTokFrez
    
        Select Case Parametry.TipTsentrovaniya
            Case 1: .optTsentrovaniye_1.Value = True
            Case 2: .optTsentrovaniye_2.Value = True
            Case 3: .optTsentrovaniye_3.Value = True
        End Select
        
        .txtD_Tsentrovaniye = CStr(Parametry.D1)
        .txtKolVo_Tsentrovaniye = CStr(Parametry.KolVoPoverkhnostey)
        .cboRa_Tsentrovaniye.ListIndex = ZapolnitRa(.cboRa_Tsentrovaniye, Parametry.Ra)
        
    End With
    
End Sub

Private Sub ZapolnitUstanovkuRastFrez()

    Dim i As Integer
    
    With frmTokFrez
        
        .txtMass_RastFrez.text = CStr(Parametry.Massa)
        
        With .lstSposobUst_RastFrez
            For i = 0 To .ListCount - 1
                If .List(i, .ColumnCount - 1) = CStr(Parametry.Sposob) Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End With
        
        With .cboKharakterVyverki_RastFrez
            For i = 0 To .ListCount - 1
                If .List(i, .ColumnCount - 1) = CStr(Parametry.KharkterVyverki) Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End With

        .cboTochnostVyverki_RastFrez.text = IIf(Parametry.TochnostVyverki = 0, "-", Parametry.TochnostVyverki)
        .txtKolVoProhodov_RastFrez = CStr(Parametry.KolVoProkhodov)
        
    End With
    
End Sub

Private Sub ZapolnitObrabotkuOtverstiyTokar()

    With frmTokFrez
    
        Select Case Parametry.id
            Case "pSverl_Tok"
                .cboPerehod_Sverl.ListIndex = 0
            Case "pRassverl_Tok"
                .cboPerehod_Sverl.ListIndex = 1
            Case "pZenker_Tok"
                .cboPerehod_Sverl.ListIndex = 2
            Case "pRazvert_Tok"
                .cboPerehod_Sverl.ListIndex = IIf(Parametry.Ra > 1.25, 3, 4)
        End Select
        
        .txtD_Sverl.text = CStr(Parametry.D1)
        .txtL_Sverl.text = CStr(Parametry.L)
        .txtKolVo_Sverl = CStr(Parametry.KolVoPoverkhnostey)
        .optVUpor_ObrOtv = Parametry.VUpor
        
    End With

End Sub

Private Sub ZapolnitGlubinuRezaniyaTokar(cbo As MSForms.ComboBox, GlubinaRezaniya As Double)
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If CDbl(cbo.List(i)) >= GlubinaRezaniya Then
            cbo.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub ZapolnitOtkloneniye(cbo As MSForms.ComboBox, Otkloneniye As Double)
    Dim i As Integer
    cbo.ListIndex = 0
    For i = cbo.ListCount - 1 To 0 Step -1
        If CDbl(cbo.List(i)) <= Otkloneniye Then
            cbo.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub ZapolnitShagRezby(cbo As MSForms.ComboBox, Shag As Double)
    Dim i As Integer
    cbo.ListIndex = 0
    For i = cbo.ListCount - 1 To 0 Step -1
        If CDbl(cbo.List(i)) >= Shag Then
            cbo.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub ZapolnitOstalnoye()
    
    ZapolnitNarezaniyeRezby

    Dim ctrl As Control
    
    For Each ctrl In frmTokFrez.mltNormativ.SelectedItem.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                Call ZapolnitTextBox(ctrl)
            Case "CheckBox"
                Call ZapolnitCheckBox(ctrl)
            Case "ComboBox"
                Call ZapolnitComboBox(ctrl)
            Case "OptionButton"
                Call ZapolnitOptionButton(ctrl)
            Case "TabStrip"
                Call ZapolnitTabStrip(ctrl)
        End Select
    Next
End Sub

Private Sub ZapolnitTextBox(txt As MSForms.TextBox)

    If txt.Name Like "txtD1_*" Or txt.Name Like "txtD_*" Then
        txt.text = Parametry.D1
    ElseIf txt.Name Like "txtd2_*" Then
        txt.text = CStr(Parametry.d2)
    ElseIf txt.Name Like "txtL_*" Then
        txt.text = Parametry.L
    ElseIf txt.Name Like "txtLObniz_*" Then
        txt.text = CStr(Parametry.DlinaObnizki)
    ElseIf txt.Name Like "txtLrez_*" Then
        txt.text = CStr(Parametry.Lrezby)
    ElseIf txt.Name Like "txtB_*" Then
        txt.text = CStr(Parametry.ShirinaPoverhnosti)
    ElseIf LCase(txt.Name) Like "txtkolvo_*" Then
        txt.text = CStr(Parametry.KolVoPoverkhnostey)
    ElseIf txt.Name Like "txtKolvoProhodov_*" Then
        txt.text = CStr(Parametry.KolVoProkhodov)
    ElseIf txt.Name Like "txtPripusk_*" Then
        txt.text = CStr(Parametry.Pripusk)
    ElseIf txt.Name Like "txtDBObniz_*" Then
        txt.text = CStr(Parametry.DBObnizki)
    ElseIf txt.Name Like "txtGlubinaObniz_*" Then
        txt.text = CStr(Parametry.GlubinaObnizki)
    ElseIf txt.Name = "txtGlubina_*" Then
        txt.text = CStr(Parametry.GlubinaPaza)
    ElseIf txt.Name Like "txtMassa_*" Then
        txt.text = CStr(Parametry.Massa)
    End If
    
    txt.text = Replace(txt.text, ".", ",")
    
End Sub

Private Sub ZapolnitCheckBox(chk As MSForms.CheckBox)

    If chk.Name Like "chkGluhoe_*" Or chk.Name Like "chkVUpor_*" Then
        chk = Parametry.VUpor
    ElseIf chk.Name Like "chkUdar_*" Then
        chk = Parametry.Udar
    ElseIf chk.Name Like "chkKorka_*" Then
        chk = Parametry.Korka
    ElseIf chk.Name Like "chkVconst_*" Then
        chk = Parametry.Vconst
    ElseIf chk.Name Like "chkPodrezka_*" Then
        chk = Parametry.PodrezkaTortsa
    ElseIf chk.Name = "chkTocheniyeDiametra" Then
        chk = Parametry.TocheniyeDiametra
    ElseIf chk.Name Like "chkKalibr_*" Then
        chk = Parametry.KalibrRezbSlesar
    ElseIf chk.Name = "chkObniz_FrezOkon" Then
        chk = Parametry.DBObnizki > 0
    End If

End Sub
                
Private Sub ZapolnitComboBox(cbo As MSForms.ComboBox)

    If cbo.Name Like "cboGlubinaRez_*" Then
        Call ZapolnitGlubinuRezaniyaTokar(cbo, Parametry.GlubinaRezaniya)
        
    ElseIf cbo.Name Like "cboIT_*" Then
        Call ZapolnitKvalitet(cbo, Parametry.IT)
        
    ElseIf cbo.Name Like "cboRaRz_*" Then
        cbo.ListIndex = ZapolnitRa(cbo, Parametry.Ra)
            
    ElseIf cbo.Name Like "cboB_*" Then
        Call ZapolnitShagRezby(cbo, Parametry.ShirinaPoverhnosti)
    
    ElseIf cbo.Name Like "cboS_*" Then
        Call ZapolnitShagRezby(cbo, Parametry.ShagRezby)
    
    ElseIf cbo.Name Like "cboOtkl_*" Then
        Call ZapolnitOtkloneniye(cbo, Parametry.Otkloneniye)
    
    ElseIf cbo.Name Like "cboTipPov_*" Then
        cbo.ListIndex = Parametry.TipPoverkhnosti - 1
    
    End If
    
End Sub

Private Sub ZapolnitNarezaniyeRezby()

    With frmTokFrez
        .chkAvtoVybor.Value = Parametry.id = "RezbaAvtoVybor"
        
        Select Case Parametry.id
            Case "pPlashka", "pMetchik"
                .optPlashkoyMetchikom.Value = True
            Case "pMetrRezb"
                .optReztsom.Value = True
        End Select
        
    End With
    
End Sub
           
Private Sub ZapolnitOptionButton(opt As MSForms.OptionButton)
    If opt.Name Like "optNaruzh*" Then
        opt = Parametry.TipPoverkhnosti = 1
    ElseIf opt.Name Like "optVnutr*" Then
        opt = Parametry.TipPoverkhnosti = 2
    ElseIf opt.Name Like "optVUpor_*" Then
        opt = Parametry.VUpor
    ElseIf opt.Name Like "optSkvoznoye_*" Then
        opt = Not Parametry.VUpor
    End If
End Sub
               
Private Sub ZapolnitTabStrip(tbs As MSForms.TabStrip)
    If tbs.Name Like "tabIT_*" Then
        tbs.Value = IIf(Parametry.IT > 6, 1, 0)
    End If
End Sub
               
Private Sub ZapolnitKvalitet(cbo As MSForms.ComboBox, IT As Integer)
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = CStr(IT) Then
            cbo.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Function ZapolnitRa(cbo As MSForms.ComboBox, Ra As Double) As Integer
    
    Dim i As Integer
    ZapolnitRa = cbo.ListCount - 1
    For i = 0 To cbo.ListCount - 1
        If Ra >= CDbl(Split(cbo.Column(IIf(cbo.ColumnCount > 1, 1, 0), i), "...")(0)) Then
            ZapolnitRa = i
            Exit For
        End If
    Next
End Function

