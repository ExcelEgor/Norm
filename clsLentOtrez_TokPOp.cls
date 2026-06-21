VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLentOtrez_TokPOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pKolVoDetaley As Double
Private pListIndexMaterial As Integer
Private pDiametr As Double
Private pDlinaZagotovki As Double
Private pd2 As Double
Private pTocheniyeNaUdar As Boolean
Private pMassa As Double

Private pDlinaOdnoyDetali As String

Private Sub Class_Initialize()

    Dim InputKolVo
    Do
        InputKolVo = InputBox("”кажите на какое количество деталей расчитан пруток", " оличество деталей")
    Loop Until IsNumeric(InputKolVo)
    
    If IsNumeric(InputKolVo) Then
        If InputKolVo > 0 Then
            pKolVoDetaley = CDbl(InputKolVo)
        Else
            pKolVoDetaley = 1
        End If
    End If
    
End Sub

Public Sub Init(ListIndexMaterial As Integer, Massa As Double, D1 As Double, d2 As Double, DlinaZagotovki As Double, TocheniyeNaUdar As Boolean)

    pListIndexMaterial = ListIndexMaterial
    pMassa = Massa
    pDiametr = D1
    pd2 = d2
    pTocheniyeNaUdar = TocheniyeNaUdar
    
    pDlinaZagotovki = DlinaZagotovki
    pDlinaOdnoyDetali = CStr(WorksheetFunction.RoundUp(pDlinaZagotovki / pKolVoDetaley, 1))
    
End Sub

Public Sub RaschotTokarnoyOperatsii()

    If Not pKolVoDetaley > 0 Then Exit Sub
    
    frmTokFrez.cboMaterial.ListIndex = pListIndexMaterial
    
    Call Ustanovit
    Call PodrezatTorets
    Call RaschotProdolnogoTocheniya
    Call Pereustanovit
    Call PodrezatToretsDlyaSnyatiyaBobyshki
    
    Call Promyvka
    Call DlinaObrabotkiOtverstiya
    Call RaschotGalvaniki
    
    With frmTokFrez
        .cboMaterial.ListIndex = pListIndexMaterial
        .mltNormativTok.Value = .mltNormativTok.Pages("pProdolnoye").Index
        .tabOperatsii.Value = .tabOperatsii.tabs.Count - 2
        .Show
    End With
    
End Sub

Private Sub Ustanovit()

    With frmTokFrez
        If .lstTokarnoFrezer.ListCount > 0 Then Call frmTokFrez.Raschot.DobavitOperatsiyu
        .mltMain.Value = 0
        .mltNormativTok.Value = .mltNormativTok.Pages("pUst").Index
        .txtKolVoProhodov_Ust.text = 1 / pKolVoDetaley
        .lstSposobUst_Tok.ListIndex = 0
        .txtMass_Tok.text = CStr(pMassa)
        
        Call frmTokFrez.DobavitVTablitsu
        
    End With
    
End Sub

Private Sub PodrezatTorets()

    With frmTokFrez
    
        .mltNormativTok.Value = .mltNormativTok.Pages("pPoperechnoye").Index
        .txtD1_Pop.text = CStr(pDiametr)
        .txtd2_Pop.text = CStr(pd2)
        .txtPripusk_Pop.text = 1
        .txtKolVo_Pop.text = 1
        .cboIT_Pop.text = 14
        .chkUdar_Pop.Value = pTocheniyeNaUdar
        .chkTonkosten.Value = False
        .txtD1_Otrezka.text = CStr(pDiametr)
        .txtd2_Otrezka.text = CStr(pd2)
        .chkTocheniyeDiametra.Value = False
        
        Call frmTokFrez.DobavitVTablitsu
        
    End With
    
End Sub

Private Sub RaschotProdolnogoTocheniya()

    Dim EN As Double
    EN = KolVoDetIzPrutka(pDlinaZagotovki, CDbl(pKolVoDetaley), pDiametr)

    With frmTokFrez
        .txtD1_Prod.text = CStr(pDiametr)
        .chkUdar_Prod.Value = pTocheniyeNaUdar
        .txtKolVo_Prod = CStr(WorksheetFunction.Round(1 / EN, 3))
        .txtL_Prod = CStr(WorksheetFunction.RoundUp(DlinaIzPutka(pDlinaZagotovki, CDbl(pKolVoDetaley), EN), 0))
    End With

End Sub

Private Sub Pereustanovit()
    With frmTokFrez
        Call frmTokFrez.Raschot.DobavitOperatsiyu
        .cboKharakterVyverki_Tok.ListIndex = 0
        .cboTochnostVyverki_Tok.ListIndex = 0
        .tabForMlt.Value = 0
        .mltNormativTok.Value = .mltNormativTok.Pages("pUst").Index
        .txtMass_Tok.text = pMassa / pKolVoDetaley
        .txtKolVoProhodov_Ust.text = 1
        Call frmTokFrez.DobavitVTablitsu
    End With
End Sub

Private Sub PodrezatToretsDlyaSnyatiyaBobyshki()

    With frmTokFrez
    
        .mltNormativTok.Value = .mltNormativTok.Pages("pPoperechnoye").Index
        .txtD1_Pop.text = CStr(pDiametr)
        .txtd2_Pop.text = CStr(pd2)
        .txtPripusk_Pop.text = 1
        .txtKolVo_Pop.text = 1
        .cboIT_Pop.text = 11
        .chkUdar_Pop.Value = pTocheniyeNaUdar
        .chkTonkosten.Value = False
        .txtD1_Otrezka.text = CStr(pDiametr)
        .txtd2_Otrezka.text = CStr(pd2)
        
        Call frmTokFrez.DobavitVTablitsu
        
    End With
    
End Sub

Private Sub DlinaObrabotkiOtverstiya()

    With frmTokFrez
        .txtL_Sverl = pDlinaOdnoyDetali
        .txtL_Rast = pDlinaOdnoyDetali
        .txtL_ObrabOtvSNulya = pDlinaOdnoyDetali
    End With
    
End Sub

Private Sub Promyvka()

    With frmTokFrez
        .tabPromyvka.Value = 1
        .txtPromivka_D = CStr(pDiametr)
        .txtPromivka_L = CStr(pDlinaOdnoyDetali)
    End With
    
End Sub
