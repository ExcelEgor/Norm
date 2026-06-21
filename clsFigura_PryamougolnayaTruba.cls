VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_PryamougolnayaTruba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pShirina As Double
Private pVysota As Double
Private pDlina As Double
Private pTolshchinaStenki As Double
Private pMaterial As EnumMaterialy
Private pTipStali As Integer

Implements ITechnologicalShape

Public Property Get Tonkostennaya() As Boolean: Tonkostennaya = True: End Property

Private Property Get ITechnologicalShape_Tonkostennaya() As Boolean: ITechnologicalShape_Tonkostennaya = Me.Tonkostennaya: End Property

Private Property Get ITechnologicalShape_tPeskostruy() As Double: ITechnologicalShape_tPeskostruy = Me.tPeskostruy: End Property

Private Property Get ITechnologicalShape_tZakalka() As Double: ITechnologicalShape_tZakalka = Me.tZakalka: End Property

Private Property Get ITechnologicalShape_tOtpusk() As Double: ITechnologicalShape_tOtpusk = Me.tOtpusk: End Property

Private Property Get ITechnologicalShape_tUpakovka() As Double: ITechnologicalShape_tUpakovka = Me.tUpakovka: End Property

Private Property Get ITechnologicalShape_tZaus() As Double: ITechnologicalShape_tZaus = Me.tZaus: End Property

Private Property Get ITechnologicalShape_tPravka() As Double: ITechnologicalShape_tPravka = Me.tPravka: End Property

Private Property Get ITechnologicalShape_Massa() As Double: ITechnologicalShape_Massa = Me.Massa: End Property

Private Property Get ITechnologicalShape_Dlina() As Double: ITechnologicalShape_Dlina = pDlina: End Property

Private Property Get ITechnologicalShape_PerimetrIliDiametr() As Double: ITechnologicalShape_PerimetrIliDiametr = 2 * (pShirina + pVysota): End Property

Private Property Get ITechnologicalShape_Material() As EnumMaterialy: ITechnologicalShape_Material = pMaterial: End Property

Private Property Get ITechnologicalShape_PloshchadPoperSech() As Double: ITechnologicalShape_PloshchadPoperSech = PloschadPoperSech: End Property

Private Property Get ITechnologicalShape_PloshchadPolnaya() As Double: ITechnologicalShape_PloshchadPolnaya = PloschadPolnaya: End Property

''''''''''''''''''''''''''''''''''''''''

Public Sub Init(Material As EnumMaterialy, Shirina As Double, Vysota As Double, TolshchinaStenki As Double, Dlina As Double)
    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)
    pShirina = Shirina
    pVysota = Vysota
    pTolshchinaStenki = TolshchinaStenki
    pDlina = Dlina
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

Public Function PloschadPolnaya()
    PloschadPolnaya = 4 * (pDlina + pTolshchinaStenki) * (pShirina + pVysota - 2 * pTolshchinaStenki)
End Function

Public Function PloschadPoperSech()
    PloschadPoperSech = 2 * pTolshchinaStenki * (pShirina + pVysota - 2 * pTolshchinaStenki)
End Function

Public Function Obyom()
    Obyom = PloschadPoperSech * pDlina
End Function

''''''''''''''''''''''''''''''''''''''''

Public Function tPeskostruy() As Double
    tPeskostruy = PeskostruynayaOchistka(Gabarity.Shirina, pDlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, Pryamougolnoye, pTolshchinaStenki, Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Private Property Get Gabarity() As GabarityZagotovki
    Gabarity = RaschotGabaritov(pDlina, pShirina, pVysota)
End Property

Public Function tZaus() As Double
    tZaus = 2 * ZachistkaZausencev(pMaterial, 4 * pShirina + 4 * pVysota - 8 * pTolshchinaStenki)
End Function

Public Function tPravka() As Double
    tPravka = KOEFFITSIENT_NA_PRAVKU * PravkaPloskikhDetaley(pTolshchinaStenki, pDlina, 2 * (pShirina + pVysota), CInt(pMaterial))
End Function

Public Function tUpakovka()
    tUpakovka = UpakovyvaniyePramougolnika(pDlina, pShirina, pVysota)
End Function

Public Function tVisual() As Double
    tVisual = VisualnyyKontrol(2 * (pShirina + pVysota), pDlina, 1, False, False)
End Function
