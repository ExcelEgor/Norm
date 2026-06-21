VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTokFrez_VspomRaschoty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private frm As frmTokFrez

Private Sub Class_Initialize()
    Set frm = frmTokFrez
End Sub

Public Sub RaschotPromyvkiObertyvaniya()
    
    With frmTokFrez

        Dim d As Double: d = DblFromCtrl(.txtPromivka_D)
        Dim L As Double: L = DblFromCtrl(.txtPromivka_L)
        
        Dim tPromivka As Double, tObertyvaniye As Double
        If d > 0 And L > 0 Then
            tPromivka = (Promivka_m64(.tabPromyvka.Value + 1, 3, d, L, 2) + Obduvka_m65(.tabPromyvka.Value + 1, d, L, 2)) * 1.3
            .txtPromivka_t.text = CStr(WorksheetFunction.RoundUp(tPromivka, 1))
            
            If .tabPromyvka.Value = 0 Then
                tObertyvaniye = 1.3 * ObertyvaniyeBumagoy_PloskieDetalie(d, L)
            Else
                tObertyvaniye = 1.3 * ObertyvaniyeBumagoy_CilindicheskiyeDetalie(d, L)
            End If
            .txtObertyvaniye.text = CStr(OkruglenieTsht(tObertyvaniye))
            
        Else
            .txtPromivka_t.text = Empty
        End If
    
    End With
    
End Sub

Public Sub RaschotGalvaniki()
    
    With frmTokFrez
    
        Dim S As Double
        S = DblFromCtrl(.txtGalvanika_S)
        
        If S > 0 Then
            
            Dim tGalvanika As Double
            tGalvanika = IIf(.chkGalvanikaPartiya, 1, 2) * Galvanika(.cboGalvanika, S)
            .txtGalvanika_t.text = CStr(OkruglenieTsht(tGalvanika))
            
            Dim tKontrol As Double
            tKontrol = VisualnyyKontrolPoPloshadi(S * 1000000) * IIf(.chkGalvanikaPartiya, 0.7, 1.3)
            .txtGalvanika_tKontrol.text = CStr(WorksheetFunction.RoundUp(tKontrol, 1))
        
        Else
            .txtGalvanika_t.text = Empty
            .txtGalvanika_tKontrol.text = Empty
        
        End If
    
    End With
    
End Sub

Public Sub RashotZachistkiShestigrannika()
    
    With frmTokFrez
    
        Dim S As Double:    S = DblFromCtrl(.txtS_ZausShestigran)
        Dim L As Double:    L = DblFromCtrl(.txtL_ZausShestigran)
        
        Dim Material As EnumMaterialy
        Material = CInt(.cboMaterial_ZausShestigran.List(.cboMaterial_ZausShestigran.ListIndex, 1))
        
        Dim tZaus As Double
        If S > 0 Then
            tZaus = 1.3 * ZachistkaZausencevNaShestigrannike(Material, S, .chkSDvukhStoron.Value, L)
            .txtVremya_ZausShestigran.text = CStr(OkruglenieTsht(tZaus))
        Else
            .txtVremya_ZausShestigran.text = Empty
        End If
    
    End With

End Sub

Public Sub RashotKopmplektovaniya()
    
    With frmTokFrez
    
        .txtKomplekt_Vremiya.text = Empty
        
        Dim tKomplekt As Double
        If IsNumeric(.txtKomplekt_KolVo) And IsNumeric(.cboKomplekt_Massa) Then
            tKomplekt = Komplektovaniye(.cboKomplekt_Massa, .txtKomplekt_KolVo, .chkKomplekt_Transport.Value)
            .txtKomplekt_Vremiya.text = Format(WorksheetFunction.RoundUp(tKomplekt, 1), "0.0")
        End If
        
    End With
    
End Sub

Public Sub FrezerovaniyePerimetra_FrezPloskKonc()

    With frmTokFrez
    
        .txtL_FrezPloskKonc.Enabled = Not .chkPerimetr_FrezPloskKonc.Value
        
        .txtPerimetrA_FrezPloskKonc.text = Empty
        .txtPerimetrB_FrezPloskKonc.text = Empty
        
        .txtPerimetrA_FrezPloskKonc.Enabled = .chkPerimetr_FrezPloskKonc.Value
        .txtPerimetrB_FrezPloskKonc.Enabled = .chkPerimetr_FrezPloskKonc.Value
        
        If .chkPerimetr_FrezPloskKonc.Value Then
            .txtPerimetrA_FrezPloskKonc.BorderColor = &H8000000D
            .txtPerimetrB_FrezPloskKonc.BorderColor = &H8000000D
        Else
            .txtPerimetrA_FrezPloskKonc.BorderColor = &H8000000A
            .txtPerimetrB_FrezPloskKonc.BorderColor = &H8000000A
        End If
        
    End With
    
End Sub

Public Sub RaschotPerimetra_FrezPloskKonc()
    
    With frmTokFrez
        If .chkPerimetr_FrezPloskKonc Then
            .txtL_FrezPloskKonc = 2 * (DblFromCtrl(.txtPerimetrA_FrezPloskKonc) + DblFromCtrl(.txtPerimetrB_FrezPloskKonc))
        End If
    End With
    
End Sub

Public Sub VklucheniyeObshchikhParametrov()

    If frm.chkObshchParam_Tok.Value = True Then
        frm.cboObshchIT_Tok.ListIndex = 4
        frm.cboObshchRaRz_Tok.ListIndex = 1
    Else
        frm.cboObshchIT_Tok.ListIndex = -1
        frm.cboObshchRaRz_Tok.ListIndex = -1
    End If
    
    frm.lblObshchIT_Tok.Enabled = frm.chkObshchParam_Tok
    frm.cboObshchIT_Tok.Enabled = frm.chkObshchParam_Tok
    
    frm.lblObshchRaRzTok.Enabled = frm.chkObshchParam_Tok
    frm.cboObshchRaRz_Tok.Enabled = frm.chkObshchParam_Tok
    
End Sub

Public Sub UstanovitObshchiyeParametry()

    Dim ObshchiyRaRz As String, ObshchiyIT As String
    Dim UstanovkaParametrov As Boolean
    If frm.chkObshchParam_Tok.Value = True And frm.cboObshchRaRz_Tok.ListIndex <> -1 And frm.cboObshchIT_Tok.ListIndex <> -1 Then
        UstanovkaParametrov = True
        ObshchiyRaRz = frm.cboObshchRaRz_Tok.List(frm.cboObshchRaRz_Tok.ListIndex, 0)
        ObshchiyIT = frm.cboObshchIT_Tok.List(frm.cboObshchIT_Tok.ListIndex, 0)
    Else
        UstanovkaParametrov = False
    End If
    
    Dim pg As MSForms.Page, ctrl As MSForms.Control
    Dim ctrlName As String
    For Each pg In frm.mltNormativTok.Pages
        For Each ctrl In pg.Controls
            
            If TypeOf ctrl Is MSForms.ComboBox Then
                ctrlName = LCase(ctrl.Name)
                
                If ctrlName Like "cborarz*" Then
                    Call SinkhronizirovatCbo(ctrl, ObshchiyRaRz, UstanovkaParametrov)
                    
                ElseIf ctrlName Like "cboit_*" Then
                    Call SinkhronizirovatCbo(ctrl, ObshchiyIT, UstanovkaParametrov)
                End If
                
            End If
        Next
    Next


End Sub

Private Sub SinkhronizirovatCbo(cbo As MSForms.Control, Znacheniye As String, UstanovkaParametrov As Boolean)
    
    Dim i As Long
    If UstanovkaParametrov Then
        For i = 0 To cbo.ListCount - 1
            If cbo.List(i, 0) = Znacheniye Then
                cbo.ListIndex = i
                Exit For
            End If
        Next i
    End If

End Sub

