VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_Ugolnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDlina As Double
Private pShirina As Double
Private pVysota As Double
Private pTolshchinaStenki_a As Double
Private pTolshchinaStenki_b As Double
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

Public Sub Init(Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal TolshchinaStenki_a As Double, ByVal TolshchinaStenki_b As Double, Dlina As Double)
    
    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)

    pShirina = WorksheetFunction.Max(Shirina, Vysota)
    pVysota = WorksheetFunction.Min(Shirina, Vysota)
    pTolshchinaStenki_a = TolshchinaStenki_a
    pTolshchinaStenki_b = TolshchinaStenki_b

    pDlina = Dlina
    
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

Public Function PloschadPolnaya()
    Dim S1 As Double, S2 As Double, S3 As Double, S4 As Double
    S1 = 2 * PloschadPoperSech
    S2 = pDlina * (pShirina + pVysota)
    S3 = pDlina * (pShirina + pVysota - pTolshchinaStenki_a - pTolshchinaStenki_b)
    S4 = pTolshchinaStenki_b * pDlina + pTolshchinaStenki_a * pDlina
    PloschadPolnaya = S1 + S2 + S3 + S4
End Function

Public Function PloschadPoperSech()
    PloschadPoperSech = pShirina * pTolshchinaStenki_b + (pVysota - pTolshchinaStenki_b) * pTolshchinaStenki_a
End Function

Public Function Obyom()
    Obyom = PloschadPoperSech * pDlina
End Function

Public Function tPeskostruy() As Double
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(pDlina, pShirina, pVysota)
    tPeskostruy = PeskostruynayaOchistka(Gabarity.Shirina, Gabarity.Dlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, Pryamougolnoye, WorksheetFunction.Max(pTolshchinaStenki_a, pTolshchinaStenki_b), Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Public Function tZaus() As Double
    tZaus = 2 * ZachistkaZausencev(pMaterial, 2 * (pVysota + pShirina))
End Function

Public Function tPravka() As Double
    tPravka = KOEFFITSIENT_NA_PRAVKU * (PravkaPloskikhDetaley(pTolshchinaStenki_a, pDlina, pShirina, CInt(pMaterial)) + PravkaPloskikhDetaley(pTolshchinaStenki_b, pDlina, pVysota, CInt(pMaterial)))
End Function
        
Public Function tUpakovka()
    tUpakovka = UpakovyvaniyePramougolnika(pDlina, pShirina, pVysota)
End Function
        
        
