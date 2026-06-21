VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSlesarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mlt As MSForms.MultiPage
Attribute mlt.VB_VarHelpID = -1
Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents cboPerekhod As MSForms.ComboBox
Attribute cboPerekhod.VB_VarHelpID = -1
Private WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents cmdClear As MSForms.CommandButton
Attribute cmdClear.VB_VarHelpID = -1

Dim pKolVoStrok As Integer

Private Material As EnumMaterialy

Private OstanovitVypolneniye As Boolean


Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set MultiPageControl(vControl As MSForms.MultiPage): Set mlt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Set ComboBoxPerekhodControl(vControl As MSForms.ComboBox): Set cboPerekhod = vControl: End Property

Public Property Set OptionButtonControl(vControl As MSForms.OptionButton): Set opt = vControl: End Property

Public Property Set CheckBoxControl(vControl As MSForms.CheckBox): Set chk = vControl: End Property

Public Property Set CmdClearControl(vControl As MSForms.CommandButton): Set cmdClear = vControl: End Property

Private Sub cboPerekhod_Change(): Call ZapolnitKvalitet: End Sub

Private Sub chk_Click(): Call RaschotTsht(chk): End Sub

Private Sub cmdClear_Click(): Call OchistitStrokuObrabotkiOtverstiy: End Sub

Private Sub mlt_Change(): Call RaschotTsht: End Sub

Private Sub opt_Click(): Call RaschotTsht: End Sub

Private Sub txt_Change(): Call RaschotTsht(txt): End Sub

Private Sub cbo_Change(): Call RaschotTsht(cbo): End Sub

Public Sub Init(KolVoStrok As Integer)
    pKolVoStrok = KolVoStrok
End Sub

Public Sub RaschotTsht(Optional ctrl As Control)
    
    If OstanovitVypolneniye Then Exit Sub
    If ctrl Is Nothing Then Exit Sub

    Dim Tsht As Double

    frmSlesar.txt_tSlesar.text = "-"

    Material = CInt(frmSlesar.cboMaterial.List(frmSlesar.cboMaterial.ListIndex, 1))
    
    Select Case VybrannayaKarta.SelectedItem.Name
        Case "Zausentsy_List"
            Tsht = RaschotTsht_Zausentsy_List
        Case "Sverleniye"

        Case "NarezaniyeRezby"
            Call RaschotTsht_ObrabotkaOtverstiy(ctrl)
        Case "Pravka_List"
            Tsht = RaschotTsht_Pravka_List
        Case "pZachistka"
            Tsht = RaschotTsht_Zachistka
        Case "pPritirka_Plosk"
            Tsht = RaschotTsht_Pritirka
    End Select
    
    If Tsht > 0 Then
        frmSlesar.txt_tSlesar.text = Format(Tsht, "0.0")
    End If
    
End Sub

Public Property Get VybrannayaKarta() As MSForms.MultiPage
    Dim ctrl As Control
    For Each ctrl In frmSlesar.mltSlesar.SelectedItem.Controls
        If TypeName(ctrl) = "MultiPage" Then
            Set VybrannayaKarta = ctrl
            Exit For
        End If
    Next
    If VybrannayaKarta Is Nothing Then
        Set VybrannayaKarta = frmSlesar.mltSlesar
    End If
End Property

Private Function RaschotTsht_Zausentsy_List()

    Dim Dlina As Double, Shirina As Double, Tolshchina As Double
    With frmSlesar
        Dlina = DblFromCtrl(.txtDlina_Zausentsy_List)
        Shirina = DblFromCtrl(.txtShirina_Zausentsy_List)
        Tolshchina = DblFromCtrl(.txtTolshchina_Zausentsy_List)
    End With

    RaschotTsht_Zausentsy_List = SnyatiyeZausentsevList(Dlina, Shirina, Tolshchina, Material)

End Function

Private Sub RaschotTsht_ObrabotkaOtverstiy(ctrl As Control)
    
    Dim txtDotv As MSForms.TextBox
    Dim txtLotv As MSForms.TextBox
    Dim txtLrezb As MSForms.TextBox
    Dim txtKolVoOtv As MSForms.TextBox
    Dim cboITotv As MSForms.ComboBox
    Dim chkGlukhoyeOtv As MSForms.CheckBox
    Dim txtVremya As MSForms.TextBox
    
    Dim TipObrabotki As TipObrabotkiOtverstiya, d As Double, KolVo As Double, Lotv As Double, Lnarez As Double, Glukhoye As Boolean, IT As Integer
    
    Dim Stroka As Integer
    Stroka = CInt(Split(ctrl.Name, "_")(1))
    
    Dim tOtv As Double
  
    Set cboPerekhod = frmSlesar.Controls("cboPerekhod_" & Stroka)
    Set txtDotv = frmSlesar.Controls("txtDotv_" & Stroka)
    Set txtLotv = frmSlesar.Controls("txtLotv_" & Stroka)
    Set txtLrezb = frmSlesar.Controls("txtLrezb_" & Stroka)
    Set txtKolVoOtv = frmSlesar.Controls("txtKolVoOtv_" & Stroka)
    Set cboITotv = frmSlesar.Controls("cboITotv_" & Stroka)
    Set chkGlukhoyeOtv = frmSlesar.Controls("chkGlukhoyeOtv_" & Stroka)
    Set txtVremya = frmSlesar.Controls("txtVremya_" & Stroka)
    
    txtLotv.BackColor = vbWhite
    txtVremya.text = Empty

    If cboPerekhod.ListIndex <> -1 Then
    
        TipObrabotki = CInt(cboPerekhod.List(cboPerekhod.ListIndex, 1))

        d = DblFromCtrl(txtDotv)

        Lotv = DblFromCtrl(txtLotv)
        
        If Lotv > LmaxObrabotkaOtverstiy(TipObrabotki, d) Then
            txtLotv.BackColor = vbRed
            Exit Sub
        End If
        
        Lnarez = DblFromCtrl(txtLrezb)
        KolVo = KolVoPoverhnosteiProhodov(txtKolVoOtv)
        IT = IIf(cboITotv.text = "H9", 9, 7)
        Glukhoye = chkGlukhoyeOtv.Value
        
        If d > 0 Then
            tOtv = ObrabotkaOtverstiySleasar(TipObrabotki, CInt(Material), d, KolVo, Lotv, Lnarez, Glukhoye, IT).tOtv
            txtVremya.text = CStr(WorksheetFunction.Round(tOtv, 2))
        End If
        
    End If
        

End Sub

Private Function LmaxObrabotkaOtverstiy(TipObrabotki As TipObrabotkiOtverstiya, d As Double) As Integer
    
    Dim Lmax As Integer
    If TipObrabotki = TipObrabotkiOtverstiya.Sverleniye Then
        Lmax = LmaxSverleniyeSlesar(CInt(Material), d)
    ElseIf TipObrabotki = NarezaniyeRezby Then
        Lmax = LmaxNarezaniyeRezbyMetchikomSlesar(CInt(Material), d)
    ElseIf TipObrabotki = Razvertyvaniye Then
        Lmax = LmaxRazvertyvaniyeSlesar(CInt(Material), d)
    End If
    
    LmaxObrabotkaOtverstiy = Lmax
    
End Function

Private Sub ZapolnitKvalitet()

    Dim Stroka As Integer
    Stroka = CInt(Split(cboPerekhod.Name, "_")(1))
    
    Dim cboITotv As MSForms.ComboBox
    Set cboITotv = frmSlesar.Controls("cboITotv_" & Stroka)
    
    cboITotv.Clear
    
    Dim TipObrabotki As TipObrabotkiOtverstiya
    
    If cboPerekhod.ListIndex <> -1 Then
    
        TipObrabotki = CInt(cboPerekhod.List(cboPerekhod.ListIndex, 1))
        
        If TipObrabotki = NarezaniyeRezby Or TipObrabotki = Razvertyvaniye Or TipObrabotki = SverleniyeRazvertyvaniye Or TipObrabotki = SverleniyeRezba Then
            With cboITotv
                .AddItem "H9"
                .AddItem "H7"
                .ListIndex = 1
                .Enabled = True
            End With
        
        Else
            With cboITotv
                .AddItem "-"
                .ListIndex = 0
                .Enabled = False
            End With
            
        End If
        
    End If
    
    Call RaschotTsht
    
End Sub

Private Function RaschotTsht_Pravka_List()
    
    Dim Tolshchina As Double, Dlina As Double, Shirina As Double, Tochnost As Double
    
    With frmSlesar
        Tolshchina = DblFromCtrl(.txtTolshchina_PravkaList)
        Dlina = DblFromCtrl(.txtDlina_PravkaList)
        Shirina = DblFromCtrl(.txtShirina_PravkaList)
        Tochnost = DblFromCtrl(.cboTochnost_PravkaList)
    End With
    
    If Tolshchina > 0 And Dlina > 0 And Shirina > 0 And Tochnost > 0 Then
        RaschotTsht_Pravka_List = PravkaPloskikhDetaley(Tolshchina, Dlina, Shirina, Material, Tochnost)
    End If
    
End Function

Private Function RaschotTsht_Zachistka()
    
    Dim Zachistka As New clsZachistkaPoverkhnosti
    
    With frmSlesar
        .optPryamoyShtrikh_Zachistka.Enabled = True
        .optPutannyyShtrikh_Zachistka.Enabled = True
        
        If .optShlifMashinka_Zachistka Then
            .optPryamoyShtrikh_Zachistka.Enabled = False
            .optPutannyyShtrikh_Zachistka.Value = True
        End If
    End With

    With Zachistka
    
        .Dlina = DblFromCtrl(frmSlesar.txtDlina_Zachistka)
        .Shirina = DblFromCtrl(frmSlesar.txtShirina_Zachistka)
        
        If .Dlina <= 0 Or .Shirina <= 0 Then Exit Function
        
        .GlubokieRiski = frmSlesar.chkGlubokiyeRiski_Zachistka.Value
        .Material = Material
        
        .Ra = IIf(frmSlesar.cboRa_Zachistka.ListIndex = 0, 5, 1.25)

        If frmSlesar.optShkurka_Zachistka Then
            .SposobZachistki = Shkurkoy
        Else
            .SposobZachistki = ShifovalnoyMashinkoy
        End If
        
        If frmSlesar.optPloskayaPoverkhnost_Zachistka Then
            .TipPoverhnosti = Ploskaya
        Else
            .TipPoverhnosti = Tsilindricheskaya
        End If
        
        If frmSlesar.optPryamoyShtrikh_Zachistka Then
            .TipShtrikha = Pryamoy
        Else
            .TipShtrikha = Putannyy
        End If
        
        RaschotTsht_Zachistka = .tZachistka
        
    End With
    
End Function

Private Function RaschotTsht_Pritirka()

    If OstanovitVypolneniye Then Exit Function

    Dim Pritirka As New clsPritirkaPloskikhPov
    
    With Pritirka
    
        .Otkloneniye = DblFromCtrl(frmSlesar.cboOtkloneniye_PritirPlosk)
        .Material = Material
        .ParallelnostPerpendikulyarnost = frmSlesar.chkParallel_Pritir.Value
        .PloshchadSM2 = DblFromCtrl(frmSlesar.txtPloshchad_Pritir)
        .PreryvistayaIliAzotirovaniye = frmSlesar.chkAzot_Pritir.Value
        .PritirkaNaStanke = False
        .Tolshchina = CDbl(frmSlesar.cboTolshchina_PritirPlosk.List(frmSlesar.cboTolshchina_PritirPlosk.ListIndex, 1))
        .Ra = CDbl(frmSlesar.cboRa_PritirPlosk.List(frmSlesar.cboRa_PritirPlosk.ListIndex, 3))
        
        Call ZapolnitPripusk(.DopustimyyeZnacheniyaPripuska)
        .Pripusk = DblFromCtrl(frmSlesar.cboPripusk_PritirPlosk)

        frmSlesar.lblSmax_PritirPlosk.Caption = "до " & .PloshchadMax & " cм2"

        RaschotTsht_Pritirka = .Vremya
        OstanovitVypolneniye = False

    End With
    
End Function

Private Sub ZapolnitPripusk(DopustimyyeZnacheniyaPripuska As Variant)
    
    Dim i As InstrumentyZachistki
    Dim NaydenoZnacheniye As Boolean
    
    OstanovitVypolneniye = True
    
    With frmSlesar
    
        .cboPripusk_PritirPlosk.List = DopustimyyeZnacheniyaPripuska
    
        NaydenoZnacheniye = False
        For i = 0 To .cboPripusk_PritirPlosk.ListCount - 1
            If CStr(.cboPripusk_PritirPlosk.List(i)) = .cboPripusk_PritirPlosk.Value Then
                NaydenoZnacheniye = True
                .cboPripusk_PritirPlosk.ListIndex = i
                Exit For
            End If
        Next
    
        If NaydenoZnacheniye = False Then
            If .cboPripusk_PritirPlosk.ListIndex = -1 Then
                .cboPripusk_PritirPlosk.ListIndex = 0
            Else
                .cboPripusk_PritirPlosk.ListIndex = .cboPripusk_PritirPlosk.ListCount - 1
            End If
        End If
        
    End With
        
End Sub

Private Sub OchistitStrokuObrabotkiOtverstiy()

    Dim NomerStroki As Integer
    
    With frmSlesar
        NomerStroki = CInt(Split(cmdClear.Name, "_")(1))

        .Controls("cboPerekhod_" & NomerStroki).ListIndex = -1
        .Controls("txtDotv_" & NomerStroki).text = Empty
        .Controls("txtLotv_" & NomerStroki).text = Empty
        .Controls("txtLrezb_" & NomerStroki).text = Empty
        .Controls("txtKolVoOtv_" & NomerStroki).text = Empty
        .Controls("cboITotv_" & NomerStroki).ListIndex = -1
        .Controls("chkGlukhoyeOtv_" & NomerStroki).Value = False
        .Controls("txtVremya_" & NomerStroki).text = Empty
    
    End With


End Sub






