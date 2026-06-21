VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_Shestigrnnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pS As Double
Private pDlina As Double
Private pMaterial As EnumMaterialy
Private pTipStali As Integer

Implements ITechnologicalShape

Public Property Get Tonkostennaya() As Boolean: Tonkostennaya = False: End Property

Private Property Get ITechnologicalShape_Tonkostennaya() As Boolean: ITechnologicalShape_Tonkostennaya = Me.Tonkostennaya: End Property

Private Property Get ITechnologicalShape_tPeskostruy() As Double: ITechnologicalShape_tPeskostruy = Me.tPeskostruy: End Property

Private Property Get ITechnologicalShape_tZakalka() As Double: ITechnologicalShape_tZakalka = Me.tZakalka: End Property

Private Property Get ITechnologicalShape_tOtpusk() As Double: ITechnologicalShape_tOtpusk = Me.tOtpusk: End Property

Private Property Get ITechnologicalShape_tUpakovka() As Double: ITechnologicalShape_tUpakovka = Me.tUpakovka: End Property

Private Property Get ITechnologicalShape_tZaus() As Double: ITechnologicalShape_tZaus = ZachistkaZausencevNaShestigrannike(pMaterial, pS, False, 0): End Property

Private Property Get ITechnologicalShape_tPravka() As Double: ITechnologicalShape_tPravka = Me.tPravka: End Property

Private Property Get ITechnologicalShape_Massa() As Double: ITechnologicalShape_Massa = Me.Massa: End Property

Private Property Get ITechnologicalShape_Dlina() As Double: ITechnologicalShape_Dlina = pDlina: End Property

Private Property Get ITechnologicalShape_PerimetrIliDiametr() As Double: ITechnologicalShape_PerimetrIliDiametr = Me.DiametrOpisannoyOkruzhnosti: End Property

Private Property Get ITechnologicalShape_Material() As EnumMaterialy: ITechnologicalShape_Material = pMaterial: End Property

Private Property Get ITechnologicalShape_PloshchadPoperSech() As Double: ITechnologicalShape_PloshchadPoperSech = PloschadPoperSech: End Property

Private Property Get ITechnologicalShape_PloshchadPolnaya() As Double: ITechnologicalShape_PloshchadPolnaya = PloschadPolnaya: End Property

''''''''''''''''''''''''''''''''''''''''

Public Sub Init(Material As EnumMaterialy, S As Double, Dlina As Double)
    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)
    pS = S
    pDlina = Dlina
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

Public Function PloschadPolnaya()
    PloschadPolnaya = 3 * pS * (pS * Sqr(3) + 2 * pDlina)
End Function

Public Function PloschadPoperSech()
    PloschadPoperSech = (Sqr(3) * pS ^ 2) / 2
End Function

Public Function Obyom()
    Obyom = PloschadPoperSech * pDlina
End Function

Public Function DlinaGrani()
    DlinaGrani = (2 * (pS / 2)) / Sqr(3)
End Function

Public Function DiametrOpisannoyOkruzhnosti()
    DiametrOpisannoyOkruzhnosti = 2 * pS / Sqr(3)
End Function

''''''''''''''''''''''''''''''''''''''''

Public Function tPeskostruy() As Double
    tPeskostruy = PeskostruynayaOchistka(pS, pDlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, Pryamougolnoye, pS, Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Public Function tUpakovka()
    tUpakovka = UpakovyvaniyeShestigrannika(pS, pDlina)
End Function

Public Function tPravka() As Double
    tPravka = KOEFFITSIENT_NA_PRAVKU * PravkaPloskikhDetaley(pS, pDlina, pS, CInt(pMaterial))
End Function

Public Function tVisualKontrol() As Double
    tVisualKontrol = VisualnyyKontrol(DiametrOpisannoyOkruzhnosti, pDlina, 1, False, False)
End Function

