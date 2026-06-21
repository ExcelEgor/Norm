VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatron_Frm_Raschoty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents mlt As MSForms.MultiPage
Attribute mlt.VB_VarHelpID = -1
Private WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Private WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1
Private Material As EnumMaterialy

Public Property Set TextBoxControl(vControl As MSForms.TextBox): Set txt = vControl: End Property

Public Property Set ComboBoxControl(vControl As MSForms.ComboBox): Set cbo = vControl: End Property

Public Property Set MultiPageControl(vControl As MSForms.MultiPage): Set mlt = vControl: End Property

Public Property Set OptionButtonControl(vControl As MSForms.OptionButton): Set opt = vControl: End Property

Public Property Set CheckBoxControl(vControl As MSForms.CheckBox): Set chk = vControl: End Property

Public Property Set ListBoxControl(vControl As MSForms.ListBox): Set lst = vControl: End Property

Private Sub txt_Change(): Call GlavnyyRaschot: End Sub

Private Sub cbo_Change(): Call GlavnyyRaschot: End Sub

Private Sub mlt_Change(): Call GlavnyyRaschot: End Sub

Private Sub opt_Click(): Call GlavnyyRaschot: End Sub

Private Sub chk_Click(): Call GlavnyyRaschot: End Sub

Private Sub lst_Click(): Call GlavnyyRaschot: End Sub

Private Sub GlavnyyRaschot()

    With frmDatron
        .txt_tDatron.text = "-"
        .txt_tSlesar.text = "-"
        .txt_tKontrol.text = "-"
        .txtPerekhod.text = Empty
    
        Material = CInt(.cboMaterial.List(.cboMaterial.ListIndex, 1))
        
        Select Case .mltDatron.SelectedItem.Name

            Case "ObrabotkaOtverstiy"
                Set .ParametryDatron = ObrabotkaOtverstiy

            Case "ObrabotkaRezbovykhOtverstiy"
                Set .ParametryDatron = ObrabotkaRezbovykhOtverstiy

            Case "FrezerovaniyeTortsevoyFrezoy"
                Set .ParametryDatron = FrezerovaniyeTortsevoyFrezoy
                
            Case "VyrezaniyePoKonturu"
                Set .ParametryDatron = VyrezaniyePoKonturu
                
            Case "StupenchatoyeFrezerovaniye"
                Set .ParametryDatron = StupenchatoyeFrezerovaniye

            Case "FrezPazov"
                Set .ParametryDatron = FrezPazov

            Case "FrezerovaniyeObnizheniy"
                Set .ParametryDatron = FrezerovaniyeObnizheniy
                
            Case "FrezUstup"
                Set .ParametryDatron = FrezUstup
                
            Case "Ustanovka"
                Set .ParametryDatron = Ustanovka
            
        End Select

        .ParametryDatron.id = .mltDatron.SelectedItem.Name
        .ParametryDatron.UchestSlesar = .chk_tSlec.Value
        .ParametryDatron.UchestKontrol = .chk_tKontr.Value
        .ParametryDatron.OpisaniyePerekhoda = .ParametryDatron.OpisaniyePerekhoda & " (x" & .ParametryDatron.KolVoPoverkhnostey & ")"
        
        Dim Normy As NormyVremeni
        Normy = .ParametryDatron.RaschotIzParametrov
        
        .txtPerekhod.text = .ParametryDatron.OpisaniyePerekhoda
        
        If Normy.tMekhanika > 0 Then .txt_tDatron = CStr(WorksheetFunction.RoundUp(Normy.tMekhanika, 2))
        If Normy.tSlesar > 0 Then .txt_tSlesar.text = CStr(WorksheetFunction.RoundUp(Normy.tSlesar, 2))
        If Normy.tKontrol > 0 Then .txt_tKontrol.text = CStr(WorksheetFunction.RoundUp(Normy.tKontrol, 2))
        
    End With
    
End Sub

Private Function Ustanovka() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Massa = DblFromCtrl(frmDatron.txtMassa)
        .SposobUstanovki = CInt(frmDatron.lstUstanovka.List(frmDatron.lstUstanovka.ListIndex, 1))
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_Ustanovka)
        .OpisaniyePerekhoda = "Установить"
    End With
    
    Set Ustanovka = ParametryDatron
    
End Function

Private Function FrezerovaniyeTortsevoyFrezoy() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtDlina_FrezTorts)
        .Shirina = DblFromCtrl(frmDatron.txtShirina_FrezTorts)
        .Pripusk = DblFromCtrl(frmDatron.txtPripusk_FrezTorts)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_FrezTorts)
        .Material = Material
        .OpisaniyePerekhoda = "Фрезеровать пов-ть в р-р "
    End With

    Set FrezerovaniyeTortsevoyFrezoy = ParametryDatron
    
End Function

Private Function ObrabotkaOtverstiy() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Diametr = DblFromCtrl(frmDatron.txtD_otv)
        .Glubina = DblFromCtrl(frmDatron.txtL_otv)
        .Glukhoye = frmDatron.optGlukhoye_otv Or frmDatron.optGlukhoyeSObnizh_otv
        .Dobnizheniya = DblFromCtrl(frmDatron.txtDobnizh_otv)
        .Lobnizheniya = DblFromCtrl(frmDatron.txtLobnizh_otv)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.tktKolVo_otv)
        .Material = Material
        .VysotaFaski = DblFromCtrl(frmDatron.txtVysotaFaski)
        .OpisaniyePerekhoda = "Фрезеровать отв. Ф" & CStr(.Diametr)
    End With
    
    Set ObrabotkaOtverstiy = ParametryDatron
    
End Function

Private Function ObrabotkaRezbovykhOtverstiy() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Diametr = DblFromCtrl(frmDatron.txtD_RezbOtv)
        .Glubina = DblFromCtrl(frmDatron.txtLotv_RezbOtv)
        .GlubinaRezby = DblFromCtrl(frmDatron.txtLrezb_RezbOtv)
        .Glukhoye = frmDatron.optGlukhoye_RezbOtv.Value
        .Dobnizheniya = DblFromCtrl(frmDatron.txtDobn_RezbOtv)
        .Lobnizheniya = DblFromCtrl(frmDatron.txtLobn_RezbOtv)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_RezbOtv)
        .Material = Material
        .OpisaniyePerekhoda = "Обработать резьб. отв. М" & CStr(.Diametr)
    End With

    Set ObrabotkaRezbovykhOtverstiy = ParametryDatron
    
End Function

Private Function VyrezaniyePoKonturu() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtDlina_Vyrezaniye)
        .Glubina = DblFromCtrl(frmDatron.txtGlubina_Vyrezaniye)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_Vyrezaniye)
        .Material = Material
        .OpisaniyePerekhoda = "Вырезать контур"
    End With

    Set VyrezaniyePoKonturu = ParametryDatron
    
End Function

Private Function FrezerovaniyeObnizheniy() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtDlina_Obnizheniye)
        .Shirina = DblFromCtrl(frmDatron.txtShirina_Obnizheniye)
        .Glubina = DblFromCtrl(frmDatron.txtGlubina_Obnizheniye)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_Obnizheniye)
        .Radius = DblFromCtrl(frmDatron.txtRadius_Obnizheniye)
        .Material = Material
        .BezZachistki = frmDatron.chkBezZachistki
        .OpisaniyePerekhoda = "Фрезеровать обнижение " & CStr(Round(.Dlina, 2)) & "x" & CStr(Round(.Shirina, 2)) & "x" & CStr(.Glubina)
    End With
    
    Set FrezerovaniyeObnizheniy = ParametryDatron
    
End Function

Private Function StupenchatoyeFrezerovaniye() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtLfrez)
        .DlinaKrivoy = DblFromCtrl(frmDatron.txtLkrivoy)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_StupFrez)
        .Material = Material
        .OpisaniyePerekhoda = "Фрезеровать ступенчато"
    End With
    
    Set StupenchatoyeFrezerovaniye = ParametryDatron
    
End Function

Private Function FrezPazov() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtDlina_Paz)
        .Shirina = DblFromCtrl(frmDatron.txtShirina_Paz)
        .Glubina = DblFromCtrl(frmDatron.txtGlubina_Paz)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_Paz)
        .Material = Material
        .OpisaniyePerekhoda = "Фрезеровать паз " & CStr(.Shirina) & "x" & CStr(.Glubina)
    End With
    
    Set FrezPazov = ParametryDatron

End Function

Private Function FrezUstup() As clsDatron

    Dim ParametryDatron As New clsDatron
    With ParametryDatron
        .Dlina = DblFromCtrl(frmDatron.txtDlina_Ustup)
        .Shirina = DblFromCtrl(frmDatron.txtShirina_Ustup)
        .Glubina = DblFromCtrl(frmDatron.txtVysota_Ustup)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmDatron.txtKolVo_Ustup)
        .Material = Material
        .OpisaniyePerekhoda = "Фрезеровать уступ " & CStr(.Shirina) & "x" & CStr(.Glubina)
    End With
    
    Set FrezUstup = ParametryDatron
    
End Function


