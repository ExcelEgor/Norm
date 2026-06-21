VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKIM_Ffm_Raschoty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub VremyaKIM()

    With frmKIM
    
        .txtVremya_KIM.text = Empty
        Dim tHera As Double
    
        Select Case .mltKIM.SelectedItem.Name
        
            Case "pHera"
                Select Case .mltHera.SelectedItem.Name
                    Case "pHera_Ustanovka": Set .ParameryKIM = Hera_RaschotUstanovki
                    Case "pHera_Krug":      Set .ParameryKIM = Hera_RaschotKrugloyPloskosti
                    Case "pHera_Cilindr":   Set .ParameryKIM = Hera_RaschotTsilindra
                    Case "pHera_Ploskost":  Set .ParameryKIM = Hera_RaschotPloskosti
                End Select

            Case "pKruglomer"
                Select Case .mltRoundScan.SelectedItem.Name
                    Case "pKruglomer_Ustanovka":        Set .ParameryKIM = Kruglomer_RaschotUstanovki
                    Case "pKruglomer_Vyravnivaniye":    Set .ParameryKIM = Kruglomer_RaschotVyravnivaniya
                    Case "pKruglomer_Izmereniye":       Set .ParameryKIM = Kruglomer_RaschotIzmereniya
                End Select
                
            Case "pMikroskop"
                Select Case .mltMikroskop.SelectedItem.Name
                    Case "pMikroskop_Ustanovka":    Set .ParameryKIM = Mikroskop_RaschotUstanovki
                    Case "pMikroskop_Kontur":       Set .ParameryKIM = Mikroskop_Kontur
                    Case "pMikroskop_Diametr":      Set .ParameryKIM = Mikroskop_Diametr
                End Select
                
        End Select
        
        If .ParameryKIM Is Nothing Then Exit Sub
        
        Select Case .mltKIM.SelectedItem.Name
            Case "pHera"
                .ParameryKIM.id = .mltHera.SelectedItem.Name
            Case "pKruglomer"
                .ParameryKIM.id = .mltRoundScan.SelectedItem.Name
            Case "pMikroskop"
                .ParameryKIM.id = .mltMikroskop.SelectedItem.Name
        End Select

        tHera = KIM_RaschotTshtIzParametrov(.ParameryKIM)
        .txtVremya_KIM = Format(tHera, "0.00")

    End With
    
End Sub

Private Function Hera_RaschotTsilindra() As clsKIM

    Dim ParametryKIM As New clsKIM
    
    With ParametryKIM
        .Slozhost = frmKIM.cboSloznost_Cilindr_Hera.ListIndex + 1
        .D1 = DblFromCtrl(frmKIM.txtD_Hera_Cilindr)
        .Dlina = DblFromCtrl(frmKIM.txtL_Hera_Cilindr)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Hera_Cilindr)
        .DopuskRazmera = frmKIM.chkDopuskDiametr_Hera_Cilindr
        .DopuskFormy = frmKIM.chkForma_Hera_Cilindr
        .DopuskRaspolozheniya = frmKIM.chkRaspolozenie_Hera_Cilindr
        .OtOsi = frmKIM.chkOtOsi_Hera_Cilindr
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("КИМ. Изм. цилиндр Ф" & CStr(.D1) & "x" & CStr(.Dlina), .KolVoPoverkhnostey)
    End With
    
    Set Hera_RaschotTsilindra = ParametryKIM

End Function

Private Function Hera_RaschotPloskosti() As clsKIM

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Shirina = DblFromCtrl(frmKIM.txtB_Hera_Ploskost)
        .Dlina = DblFromCtrl(frmKIM.txtL_Hera_Ploskost)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Hera_Ploskost)
        .TrigernyyRezhim = frmKIM.chkTrigger
        .DopuskFormy = frmKIM.chkForma_Hera_Ploskost
        .DopuskRaspolozheniya = frmKIM.chkRaspolozenie_Hera_Ploskost
        .DopuskRazmera = frmKIM.chkLineynyyRazmer_Hera_Ploskost
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("КИМ. Изм. плоскость " & CStr(.Shirina) & "х" & CStr(.Dlina), .KolVoPoverkhnostey)
    End With
    
    Set Hera_RaschotPloskosti = ParametryKIM
    
End Function

Private Function Hera_RaschotKrugloyPloskosti() As clsKIM
    
    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .D1 = DblFromCtrl(frmKIM.txtD_Hera_Krug)
        .d2 = DblFromCtrl(frmKIM.txtd2_Hera_Krug)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Hera_Krug)
        .TrigernyyRezhim = frmKIM.chkTrigger_Torec
        .DopuskFormy = frmKIM.chkForma_Hera_Krug
        .DopuskRaspolozheniya = frmKIM.chkRaspolozenie_Hera_Krug
        .DopuskRazmera = frmKIM.chkLineynyyRazmer_Hera_Krug
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("КИМ. Изм. плоскость Ф" & CStr(.D1) & "x" & CStr(.d2), .KolVoPoverkhnostey)
    End With
    
    Set Hera_RaschotKrugloyPloskosti = ParametryKIM

End Function

Private Function Hera_RaschotUstanovki() As clsKIM

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Massa = DblFromCtrl(frmKIM.txtMassa_Hera)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVoUst_Hera)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("КИМ. Установить", .KolVoPoverkhnostey)
    End With
    Set Hera_RaschotUstanovki = ParametryKIM
    
End Function

Private Function Kruglomer_RaschotUstanovki() As clsKIM

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Massa = DblFromCtrl(frmKIM.txtMassa_Kruglomer)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Kruglomer_Ustanovka)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Кругломер. Установить", .KolVoPoverkhnostey)
    End With
    Set Kruglomer_RaschotUstanovki = ParametryKIM
    
End Function

Private Function Kruglomer_RaschotVyravnivaniya() As clsKIM

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Dlina = DblFromCtrl(frmKIM.txtL_Kruglomer_Viverka)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Kruglomer_Viverka)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Кругломер. Выверить", .KolVoPoverkhnostey)
    End With
    Set Kruglomer_RaschotVyravnivaniya = ParametryKIM
    
End Function

Private Function Kruglomer_RaschotIzmereniya()

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Dlina = DblFromCtrl(frmKIM.txtL_Kruglomer)
        .RadialnoyeIzmereniye = frmKIM.chkKruglomer_Radialnoye
        .LieneynoyeIzmereniye = frmKIM.chkKruglomer_Lineynoye
        .Slozhost = frmKIM.cboSloznost_Kruglomer.ListIndex + 1
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Kruglomer_Izmereniye)
        
        Dim Perekhod As String
        If .RadialnoyeIzmereniye And .LieneynoyeIzmereniye Then
            Perekhod = "Измер. радиал. и линейн."
        ElseIf .RadialnoyeIzmereniye Then
            Perekhod = "Измер. радиал."
        ElseIf .LieneynoyeIzmereniye Then
            Perekhod = "Измер. линейн."
        End If
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Кругломер. " & Perekhod, .KolVoPoverkhnostey)
    End With
    
    Set Kruglomer_RaschotIzmereniya = ParametryKIM

End Function

Private Function Mikroskop_RaschotUstanovki() As clsKIM

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Massa = DblFromCtrl(frmKIM.txtMassa_Mikroskop)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVoUstanovok_Mikroskop)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Микроскоп. Установить", .KolVoPoverkhnostey)
    End With
    
    Set Mikroskop_RaschotUstanovki = ParametryKIM
    
End Function

Private Function Mikroskop_Kontur() As clsKIM
    
    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .Dlina = DblFromCtrl(frmKIM.txtKontur_Mikroskop)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Kontur_Mikroskop)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Микроскоп. Контур L=" & CStr(.Dlina), .KolVoPoverkhnostey)
    End With
    Set Mikroskop_Kontur = ParametryKIM
    
End Function

Private Function Mikroskop_Diametr()

    Dim ParametryKIM As New clsKIM
    With ParametryKIM
        .D1 = DblFromCtrl(frmKIM.txtDiametr_Mikroskop)
        .KolVoPoverkhnostey = KolVoPoverhnosteiProhodov(frmKIM.txtKolVo_Diametr_Mikroskop)
        .OpisaniyePerekhoda = SformirovatTeksPerekhoda("Микроскоп. Диаметр D=" & CStr(.D1) & CStr(.Dlina), .KolVoPoverkhnostey)
    End With
    Set Mikroskop_Diametr = ParametryKIM
    
End Function

Private Function SformirovatTeksPerekhoda(Perekhod As String, KolVo As Double) As String
    SformirovatTeksPerekhoda = Perekhod & IIf(KolVo > 1, " (x" & KolVo & ")", "")
End Function

