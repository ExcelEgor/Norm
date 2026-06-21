VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrmObrabotkaOtverstiy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents cboPerekhod As MSForms.ComboBox
Attribute cboPerekhod.VB_VarHelpID = -1
Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents lblOchistit As MSForms.Label
Attribute lblOchistit.VB_VarHelpID = -1

Private Type NormyVremeni_ObrabotkaOtverstiySleasar
    tOtv As Double
    tKontrol As Double
End Type

Private Type ParametryOtverstiya
    d As Double
    KolVo As Double
    Lotv As Double
    Lnarez As Double
    Glukhoye As Boolean
    IT As Integer
End Type

Private Type ElementyVvoda
    cboPerekhod As Control
    txtD As Control
    txtLotv As Control
    txtLrezb As Control
    txtKolVoOtv As Control
    cboITotv As Control
    chkGlukhoyeOtv As Control
    txtVremya As Control
End Type

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Set ComboBoxPerekhodControl(vControl As MSForms.ComboBox): Set cboPerekhod = vControl: End Property

Public Property Set CheckBoxControl(vControl As MSForms.CheckBox): Set chk = vControl: End Property

Public Property Set LabelControl(vControl As MSForms.Label): Set lblOchistit = vControl: End Property

Private Sub cboPerekhod_Change(): Call ZapolnitKvalitet: End Sub

Private Sub lblOchistit_Click(): Call OchistitStroku(lblOchistit): End Sub

Private Sub txt_Change(): Call GlavnyyRaschot: End Sub

Private Sub cbo_Change(): Call GlavnyyRaschot: End Sub

Private Sub chk_Click(): Call GlavnyyRaschot: End Sub

Private Sub GlavnyyRaschot()

    Dim Material As EnumMaterialy
    Material = MaterialFromCboMaterial(frmObrabotkaOtverstiy.cboMaterial)

    Dim Stroka As Integer
    Dim tKontrol As Double
    For Stroka = 1 To frmObrabotkaOtverstiy.KolVoStrok
        
        Dim Elementy As ElementyVvoda
        Elementy = PrisvoitElementyVvoda(Stroka)
        
        Call SbrositFormatirovaniyeElementovVvoda(Elementy)
        
        If Elementy.cboPerekhod.ListIndex <> -1 Then
        
            Dim TipObrabotki As TipObrabotkiOtverstiya
            TipObrabotki = CInt(Elementy.cboPerekhod.List(Elementy.cboPerekhod.ListIndex, 1))
            
            Call DostupnostElementov(TipObrabotki, Elementy)
            
            Dim Otverstiye As ParametryOtverstiya
            Otverstiye = PrisvoitParametry(Stroka)
            
            Dim Normy As NormyVremeni_ObrabotkaOtverstiySleasar
            Normy = RaschotVremenyPoStroke(Material, Otverstiye, Elementy, TipObrabotki)
            tKontrol = tKontrol + Normy.tKontrol
            
            Elementy.txtVremya.text = CStr(Round(Normy.tOtv, 2))
            
        End If
        
    Next
    
    Dim tUst As Double
    If frmObrabotkaOtverstiy.chkUstanovka Then
        If DblFromCtrl(frmObrabotkaOtverstiy.txtMassa) > 0 Then
            tUst = Ustanov_NaStoleKreplBoltPlank(DblFromCtrl(frmObrabotkaOtverstiy.txtMassa), 2, 0.5)
            frmObrabotkaOtverstiy.txtTust.text = CStr(WorksheetFunction.RoundUp(tUst, 2))
        End If
    End If

    Call RaschotItoga(frmObrabotkaOtverstiy.KolVoStrok, tKontrol, tUst)
        
End Sub

Private Function PrisvoitElementyVvoda(Stroka As Integer) As ElementyVvoda

    With PrisvoitElementyVvoda
        Set .cboPerekhod = SetCtrl("cboPerekhod", Stroka)
        Set .txtD = SetCtrl("txtDotv", Stroka)
        Set .txtLotv = SetCtrl("txtLotv", Stroka)
        Set .txtLrezb = SetCtrl("txtLrezb", Stroka)
        Set .txtKolVoOtv = SetCtrl("txtKolVoOtv", Stroka)
        Set .cboITotv = SetCtrl("cboITotv", Stroka)
        Set .chkGlukhoyeOtv = SetCtrl("chkGlukhoyeOtv", Stroka)
        Set .txtVremya = SetCtrl("txtVremya", Stroka)
    End With

End Function

Private Sub DostupnostElementov(TipObrabotki As TipObrabotkiOtverstiya, Elementy As ElementyVvoda)

    If Not frmObrabotkaOtverstiy.chktObshchayaTolshchina Then
        If EstDlinaOtverstiya(TipObrabotki) = False Then
            Elementy.txtLotv.Enabled = False
            Elementy.txtLotv.text = Empty
            Elementy.txtLotv.BorderColor = &HE0E0E0
        Else
            Elementy.txtLotv.Enabled = True
            Elementy.txtLotv.BorderColor = &H808000
        End If
        
        If EstRezba(TipObrabotki) = False Then
            Elementy.txtLrezb.Enabled = False
            Elementy.txtLrezb.text = Empty
            Elementy.txtLrezb.BorderColor = &HE0E0E0
        Else
            Elementy.txtLrezb.Enabled = True
            Elementy.txtLrezb.BorderColor = &H808000
        End If
    End If

End Sub

Private Sub SbrositFormatirovaniyeElementovVvoda(Elementy As ElementyVvoda)

    frmObrabotkaOtverstiy.txttObshchayaTolshchina.Enabled = frmObrabotkaOtverstiy.chktObshchayaTolshchina.Value
    With Elementy
        .txtD.BorderColor = &H808000
        .txtLotv.BorderColor = &H808000
        .txtLrezb.BorderColor = &H808000
        .txtLotv.Enabled = Not frmObrabotkaOtverstiy.chktObshchayaTolshchina
        .txtLrezb.Enabled = Not frmObrabotkaOtverstiy.chktObshchayaTolshchina
        .txtVremya.text = Empty
    End With
    
End Sub

Private Function PrisvoitParametry(Stroka As Integer) As ParametryOtverstiya
    With PrisvoitParametry
        .d = DblFromCtrl(SetCtrl("txtDotv", Stroka))
        
        If frmObrabotkaOtverstiy.chktObshchayaTolshchina Then
            .Lotv = DblFromCtrl(frmObrabotkaOtverstiy.txttObshchayaTolshchina)
            .Lnarez = .Lotv
        Else
            .Lotv = DblFromCtrl(SetCtrl("txtLotv", Stroka))
            .Lnarez = DblFromCtrl(SetCtrl("txtLrezb", Stroka))
        End If
        
        .KolVo = DblFromCtrl(SetCtrl("txtKolVoOtv", Stroka))
        .IT = IIf(SetCtrl("cboITotv", Stroka).text = "H9", 9, 7)
        .Glukhoye = SetCtrl("chkGlukhoyeOtv", Stroka).Value
    End With
End Function

Private Function RaschotVremenyPoStroke(Material As EnumMaterialy, Otverstiye As ParametryOtverstiya, Elementy As ElementyVvoda, TipObrabotki As TipObrabotkiOtverstiya) As NormyVremeni_ObrabotkaOtverstiySleasar
        
    Dim Normy
    With Otverstiye
        If ProveritParametyOtverstiya(Material, Otverstiye, Elementy, TipObrabotki) = True Then
            Normy = ObrabotkaOtverstiySleasar(TipObrabotki, CInt(Material), .d, .KolVo, .Lotv, .Lnarez, .Glukhoye, .IT)
            RaschotVremenyPoStroke.tOtv = 1.3 * Normy(1)
            RaschotVremenyPoStroke.tKontrol = 1.3 * Normy(2)
        End If
    End With
    
End Function

Private Function ProveritParametyOtverstiya(Material As EnumMaterialy, Otverstiye As ParametryOtverstiya, Elementy As ElementyVvoda, TipObrabotki As TipObrabotkiOtverstiya) As Boolean
    
    ProveritParametyOtverstiya = True
    If TipObrabotki = ZenkovaniyeFasok Then Exit Function
    
    Dim Dmax As Long
    If TipObrabotki = TipObrabotkiOtverstiya.Sverleniye Then
        Dmax = D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA
    Else
        Dmax = D_MAX_REZBA_METCHIKOM
    End If
    
    With Otverstiye
        
        If .d <= 0 Then ProveritParametyOtverstiya = False
        
        If .d > Dmax Then
            Elementy.txtD.ForeColor = vbRed
            Elementy.txtD.ControlTipText = "Диаметр до " & Dmax & " мм"
            ProveritParametyOtverstiya = False
            Exit Function
        Else
            Elementy.txtD.ForeColor = &H404040
        End If
        
        Dim Lmax As Integer
        Lmax = LmaxObrabotkaOtverstiy(TipObrabotki, .d, Material)
        
        If EstDlinaOtverstiya(TipObrabotki) Then
            ProveritParametyOtverstiya = ProveritZnacheniye(.Lotv, Lmax, Elementy.txtLotv)
        End If
        
        If EstRezba(TipObrabotki) = True Then
            ProveritParametyOtverstiya = ProveritZnacheniye(.Lnarez, Lmax, Elementy.txtLrezb)
        End If

    End With
    
End Function

Private Function ProveritZnacheniye(L As Double, Lmax As Integer, ctrl As Control) As Boolean
    
    ProveritZnacheniye = True
    ctrl.ControlTipText = Empty
    
    If L <= 0 Or L > Lmax Then ProveritZnacheniye = False

    If L > Lmax Then
        ctrl.ForeColor = vbRed
        ctrl.ControlTipText = "Длина до " & Lmax & " мм"
    Else
        ctrl.ForeColor = &H404040
    End If
    
End Function

Private Function EstRezba(TipObrabotki As TipObrabotkiOtverstiya) As Boolean

    EstRezba = False
    If TipObrabotki = SverleniyeRezba Or TipObrabotki = NarezaniyeRezby Or TipObrabotki = KalibrovaniyeRezby Then
        EstRezba = True
    End If
    
End Function

Private Function EstDlinaOtverstiya(TipObrabotki As TipObrabotkiOtverstiya) As Boolean

    EstDlinaOtverstiya = True
    If TipObrabotki = NarezaniyeRezby Or TipObrabotki = KalibrovaniyeRezby Or TipObrabotki = ZenkovaniyeFasok Then
        EstDlinaOtverstiya = False
    End If
    
End Function

Private Sub RaschotItoga(KolVoStrok As Long, tKontrol As Double, tUst As Double)
    
    frmObrabotkaOtverstiy.txtTkontrol.text = "-"
    frmObrabotkaOtverstiy.txtItogo.text = "-"

    Dim i As Integer
    Dim tOtv As Double
    
    For i = 1 To KolVoStrok
        tOtv = tOtv + DblFromCtrl(SetCtrl("txtVremya", i))
    Next
    
    If tOtv > 0 Then
        frmObrabotkaOtverstiy.txtItogo.text = CStr(OkruglenieTsht(tOtv + tUst))
        frmObrabotkaOtverstiy.txtTkontrol.text = CStr(OkruglenieTsht(tKontrol))
    End If
    
End Sub

Private Function SetCtrl(ctrlName As String, Stroka As Integer) As Control
    Set SetCtrl = frmObrabotkaOtverstiy.Controls(ctrlName & "_" & Stroka)
End Function

Private Function LmaxObrabotkaOtverstiy(TipObrabotki As TipObrabotkiOtverstiya, d As Double, Material As EnumMaterialy) As Long
    
    Dim Lmax As Long
    If TipObrabotki = TipObrabotkiOtverstiya.Sverleniye Then
        Lmax = LmaxSverleniyeSlesar(Material, d)
    
    ElseIf TipObrabotki = NarezaniyeRezby Then
        Lmax = LmaxNarezaniyeRezbyMetchikomSlesar(Material, d)
    
    ElseIf TipObrabotki = Razvertyvaniye Then
        Lmax = LmaxRazvertyvaniyeSlesar(Material, d)
    
    ElseIf TipObrabotki = SverleniyeRezba Then
        Lmax = LmaxNarezaniyeRezbyMetchikomSlesar(Material, d)
        
    ElseIf TipObrabotki = SverleniyeRazvertyvaniye Then
        Lmax = LmaxSverleniyePlusRezba(Material, d)
    
    ElseIf TipObrabotki = Puklevka Then
        Lmax = MAKS_TOLSHCHINA_PUKLEVKA
    
    ElseIf TipObrabotki = KalibrovaniyeRezby Then
        Lmax = LmaxKalibrovaniyeVnutrenneyRezby(d)
    End If
    
    LmaxObrabotkaOtverstiy = Lmax
    
End Function

Private Sub ZapolnitKvalitet()

    Dim Stroka As Long
    Stroka = CInt(Split(cboPerekhod.Name, "_")(1))
    
    Dim cboITotv As MSForms.ComboBox
    Set cboITotv = frmObrabotkaOtverstiy.Controls("cboITotv_" & Stroka)

    With cboITotv
        .Clear
        .AddItem "-"
        .ListIndex = 0
        .Enabled = False
    End With
    
    Dim TipObrabotki As TipObrabotkiOtverstiya
    
    If cboPerekhod.ListIndex <> -1 Then
    
        TipObrabotki = CInt(cboPerekhod.List(cboPerekhod.ListIndex, 1))
        
        If TipObrabotki = NarezaniyeRezby Or TipObrabotki = Razvertyvaniye Or TipObrabotki = SverleniyeRazvertyvaniye Or TipObrabotki = SverleniyeRezba Then
        
            With cboITotv
                .Clear
                .AddItem "H9"
                .AddItem "H7"
                .ListIndex = 1
                .Enabled = True
            End With
            
        End If
        
    End If
    
    Call GlavnyyRaschot
    
End Sub

Private Sub OchistitStroku(lblOchistit)

    Dim Stroka As Integer
    
    Stroka = CInt(Split(lblOchistit.Name, "_")(1))
    
    Dim Elementy As ElementyVvoda
    Elementy = PrisvoitElementyVvoda(Stroka)
    
    Call SbrositFormatirovaniyeElementovVvoda(Elementy)
    
    With Elementy
        .cboPerekhod.ListIndex = -1
        .txtD.text = Empty
        .txtLotv.text = Empty
        .txtLrezb.text = Empty
        .txtKolVoOtv.text = Empty
        .chkGlukhoyeOtv.Value = False
        .txtVremya.text = Empty
    End With
    
End Sub


