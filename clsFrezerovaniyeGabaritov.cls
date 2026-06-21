VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrezerovaniyeGabaritov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents tbs As MSForms.TabStrip
Attribute tbs.VB_VarHelpID = -1

Private RazmeryZapolnenyKorrektno As Boolean

Private Material As EnumMaterialy

Public DlinaZagotovki As Double, ShirinaZagotovki As Double, TolshchinaZagotovki As Double
Public Dlina As Double, Shirina As Double, Tolshchina As Double

Private Massa As Double

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Set TabStripControl(vControl As MSForms.TabStrip): Set tbs = vControl: End Property

Public Property Set CheckBoxControl(vControl As MSForms.CheckBox): Set chk = vControl: End Property

Private Sub cbo_Change():   Call RaschotVremeni: End Sub

Private Sub chk_Click():    Call RaschotVremeni: End Sub

Private Sub tbs_Change():   Call RaschotVremeni: End Sub

Private Sub txt_Change():   Call RaschotVremeni: End Sub

Private Sub RaschotVremeni()

    With frmFrezGabarit
        
        .lstObrazec.Clear
        .lstFrezTime.Clear
        .lstOper.Clear
        .lstTsht.Clear
        .txtSum.text = Empty
        
        If .cboRa_a.ListIndex = -1 Or .cboRa_b.ListIndex = -1 Or .cboRa_h.ListIndex = -1 Then Exit Sub
        If .chkA.Value = False And .chkB.Value = False And .chkH = False Then Exit Sub

    End With
    
    Call ProveritUkazannyyeRazmery
    If RazmeryZapolnenyKorrektno = False Then Exit Sub
    
    Call PrisvoitRazmery
    
    Call RaschotFrezernoyObrabotki
    
    Call RaschotDrugikhOperatsii
    
End Sub

Private Sub ProveritUkazannyyeRazmery()

    RazmeryZapolnenyKorrektno = True
    
    With frmFrezGabarit
        Call ProvertiParametryStorony(.txtA_det, .txtA_zag, .chkA)
        Call ProvertiParametryStorony(.txtB_det, .txtB_zag, .chkB)
        Call ProvertiParametryStorony(.txtH_det, .txtH_zag, .chkH)
        .Repaint
    End With

End Sub

Private Sub PrisvoitRazmery()

    With frmFrezGabarit
        Dlina = DblFromCtrl(IIf(.chkB, .txtB_det, .txtB_zag))
        Shirina = DblFromCtrl(IIf(.chkA, .txtA_det, .txtA_zag))
        Tolshchina = DblFromCtrl(IIf(.chkH, .txtH_det, .txtH_zag))
    End With
    
End Sub

Private Sub ProvertiParametryStorony(det As MSForms.TextBox, Zag As MSForms.TextBox, chk As MSForms.CheckBox)
    
    det.BorderColor = &H8000000A
    Zag.BorderColor = &H8000000A
    If chk.Value = True Then
        If IsNumeric(det) And IsNumeric(Zag) Then
            If CDbl(det.text) > CDbl(Zag.text) Then
                det.BorderColor = vbRed
                Zag.BorderColor = vbRed
                RazmeryZapolnenyKorrektno = False
            End If
        Else
            RazmeryZapolnenyKorrektno = False
        End If
    Else
        If Not IsNumeric(Zag) Then RazmeryZapolnenyKorrektno = False
    End If

End Sub

Private Sub RaschotFrezernoyObrabotki()

    With frmFrezGabarit

        Material = .cboMaterial.List(.cboMaterial.ListIndex, 1)
        
        Massa = MassaLista(Dlina, Shirina, Tolshchina, CInt(Material))
        .txtMassa = Format(Massa, "0.000")
        
        Dim tUst As Double
        If Massa <= 20 Then
            tUst = Ustanov_Vtiskah(Massa)
        Else
            tUst = Ustanov_NaStoleKreplBoltPlank(Massa)
        End If
        
        Dim KolUst As Integer
        KolUst = (Abs(CInt(.chkA)) * (.tab_a.SelectedItem.Index + 1)) + _
            (Abs(CInt(.chkB)) * (.tab_b.SelectedItem.Index + 1)) + _
            (Abs(CInt(.chkH)) * (.tab_h.SelectedItem.Index + 1))
        tUst = tUst * KolUst * 1.3
        
        With frmFrezGabarit.lstObrazec
            .AddItem
            .List(.ListCount - 1, 0) = "Установить (" & KolUst & " раз" & IIf(KolUst > 4, "", "а") & ")"
        End With
        
        With .lstFrezTime
            .AddItem
            .List(.ListCount - 1, 0) = Format(tUst, "#,##0.00")
        End With

        Dim time_h As Double, time_a As Double, time_b As Double
        time_h = FrezRazmer(.txtA_zag, .txtB_zag, .txtH_zag, .txtH_det, .cboFrez_h, Material, .cboRa_h, .chkH, .tab_h)
        time_a = FrezRazmer(.txtH_zag, .txtB_zag, .txtA_zag, .txtA_det, .cboFrez_a, Material, .cboRa_a, .chkA, .tab_a)
        time_b = FrezRazmer(.txtH_zag, .txtA_zag, .txtB_zag, .txtB_det, .cboFrez_b, Material, .cboRa_b, .chkB, .tab_b)

        .txtSum = CStr(OkruglenieTsht(tUst + time_h + time_a + time_b))
            
    End With
    
End Sub

Private Function FrezRazmer(txtBsurf As MSForms.TextBox, txtL As MSForms.TextBox, txtZag As MSForms.TextBox, txtDet As MSForms.TextBox, _
    cboFrez As MSForms.ComboBox, Material As EnumMaterialy, cboRa As MSForms.ComboBox, chkFrez As MSForms.CheckBox, tabKol As MSForms.TabStrip)
    
    If chkFrez.Value = True Then
        
        Dim L As Double:            L = WorksheetFunction.Max(CDbl(txtL.text), CDbl(txtBsurf.text))
        Dim Ra As Double:           Ra = CDbl(cboRa)
        Dim bSurf As Double:        bSurf = WorksheetFunction.Min(CDbl(txtL.text), CDbl(txtBsurf.text))
        Dim Kol As Integer:         Kol = tabKol.SelectedItem.Index + 1
        Dim Pripusk As Double:      Pripusk = (CDbl(txtZag.text) - CDbl(txtDet.text)) / Kol
        
        If cboFrez.ListIndex = 0 Then

            If Pripusk > 0 Then
                Dim ctx As New clsMillingContext
                
                With ctx
                    .TipStanka = VertikalnoGorizontalnoFrezernyy
                    .Material = Material
                    .DlinaPoverkhnosti = L
                    .ShirinaPoverkhnosti = bSurf
                    
                    .Pripusk = Pripusk
                    .IT = 14
                    .Ra = Ra
                    
                    .NaUdar = False
                    .Korka = False
                    .TonkostennayaKonfiguraciya = False
                    
                    .CHPU = False
                    .PoRezhimam = True
                    .AvtoRezhimy = True
                    
                End With
                
                FrezRazmer = FrezKontsevymiFrezami_Kompleks(ctx)(1)
            End If
        Else
            FrezRazmer = FrezTortsovymyFrezami_Kompleks(2, CInt(Material), Pripusk, bSurf, L, , Ra)(1)
        End If
        FrezRazmer = FrezRazmer * 1.3 * Kol
        
        With frmFrezGabarit.lstObrazec
            .AddItem "Фрезеровать размер " & txtDet.text
        End With
        
        With frmFrezGabarit.lstFrezTime
            .AddItem Format(FrezRazmer, "#,##0.00")
        End With
        
    End If

    
End Function

Private Sub RaschotDrugikhOperatsii()

    With frmFrezGabarit

        If Dlina > 0 And Shirina > 0 And Tolshchina > 0 Then

            With .lstOper
                .AddItem "Слесарная"
                .AddItem "Промывка"
                .AddItem "Контрольная"
            End With
        
            With .lstTsht
                .AddItem CStr(OkruglenieTsht(1.3 * tSlesar))
                .AddItem CStr(OkruglenieTsht(1.3 * tPromyvki))
                .AddItem CStr(OkruglenieTsht(1.3 * tKontrol))
            End With
        
        End If
        
    End With
    
End Sub

Private Property Get tSlesar() As Double

    Dim tUst As Double
    tUst = UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(Massa) + UstanovkaSlesar_NaStoleBezKrepleniya_Povorot(Massa)

    Dim KolVoStoron As Integer
    KolVoStoron = CInt(frmFrezGabarit.tab_h.SelectedItem.Caption)
    
    Dim DlinaZachistki As Double
    DlinaZachistki = RaschotDlinyZachistki
        
    Dim tZaus As Double
    tZaus = ZachistkaZausencev_PoKonturu_Napilnikom(Material, DlinaZachistki, False, Massa > 20)
    
    tSlesar = tUst + tZaus
    
End Property

Private Function tPromyvki() As Double

    tPromyvki = Promivka_m64(1, 3, Shirina, Dlina, 2) + Obduvka_m65(1, Shirina, Dlina, 2)
    
End Function

Private Function tKontrol() As Double

    With frmFrezGabarit

        If .chkH Then tKontrol = IzmerShtangenCircul(Tolshchina, Dlina)
        
        If .chkA Then tKontrol = tKontrol + IzmerShtangenCircul(Shirina, Dlina)
        
        If .chkB Then tKontrol = tKontrol + IzmerShtangenCircul(Dlina, Shirina)
        
    End With
    
End Function

Private Function RaschotDlinyZachistki() As Double

    Dim Rebra(1 To 12) As Double
    Rebra(1) = Dlina:       Rebra(2) = Dlina:       Rebra(3) = Dlina:       Rebra(4) = Dlina
    Rebra(5) = Shirina:     Rebra(6) = Shirina:     Rebra(7) = Shirina:     Rebra(8) = Shirina
    Rebra(9) = Tolshchina:  Rebra(10) = Tolshchina: Rebra(11) = Tolshchina: Rebra(12) = Tolshchina
    
    Dim RebraDlina(), RebraShirina(), RebraTolshchina()
    Call PrisvoitRebra(RebraDlina, RebraShirina, RebraTolshchina)
    
    Dim RebraDlyaZachistki As Collection
    Set RebraDlyaZachistki = SozdatKollektsiyuReber(RebraDlina, RebraShirina, RebraTolshchina)

    Dim Rebro
    For Each Rebro In RebraDlyaZachistki
        RaschotDlinyZachistki = RaschotDlinyZachistki + CDbl(Rebra(Rebro))
    Next
    
End Function

Private Sub PrisvoitRebra(ByRef RebraDlina, ByRef RebraShirina, ByRef RebraTolshchina)

    With frmFrezGabarit

        If .chkA = True Then
            If .tab_a.SelectedItem.Caption = "1" Then
                RebraDlina = Array(1, 2, 10, 12)
            Else
                RebraDlina = Array(1, 2, 10, 12, 3, 4, 9, 11)
            End If
        End If

        If .chkB = True Then
            If .tab_a.SelectedItem.Caption = "1" Then
                RebraShirina = Array(5, 6, 9, 10)
            Else
                RebraShirina = Array(5, 6, 9, 10, 7, 8, 11, 12)
            End If
        End If
        
        If .chkH = True Then
            If .tab_h.SelectedItem.Caption = "1" Then
                RebraTolshchina = Array(1, 3, 5, 7)
            Else
                RebraTolshchina = Array(1, 3, 5, 7, 2, 4, 6, 8)
            End If
        End If
        
    End With
    
End Sub

Private Function SozdatKollektsiyuReber(RebraDlina, RebraShirina, RebraTolshchina) As Collection

    Dim Rebra As New Collection

    Dim i As Integer
    
    On Error Resume Next
    For i = LBound(RebraDlina) To UBound(RebraDlina)
        Rebra.Add RebraDlina(i), CStr(RebraDlina(i))
    Next
    For i = LBound(RebraShirina) To UBound(RebraShirina)
        Rebra.Add RebraShirina(i), CStr(RebraShirina(i))
    Next
    For i = LBound(RebraTolshchina) To UBound(RebraTolshchina)
        Rebra.Add RebraTolshchina(i), CStr(RebraTolshchina(i))
    Next
    On Error GoTo 0
    
    Set SozdatKollektsiyuReber = Rebra
    
End Function

