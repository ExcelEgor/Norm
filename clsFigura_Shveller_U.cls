VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_Shveller_U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDlina As Double
Private pShirina As Double
Private pVysota As Double
Private pNomer As String
Private pTolshchinaStenki As Double
Private pMaterial As EnumMaterialy
Private pTipStali As Integer
Private pPloschadPoperSech As Double

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

Private Property Get ITechnologicalShape_PloshchadPolnaya() As Double: ITechnologicalShape_PloshchadPolnaya = 0: End Property

''''''''''''''''''''''''''''''''''''''''

Public Sub Init(Material As EnumMaterialy, Nomer As String, Dlina As Double)

    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)
    
    pDlina = Dlina
    
    pNomer = Nomer
    
    Select Case pNomer
        Case "5г"
            pPloschadPoperSech = 6.16: pShirina = 50: pTolshchinaStenki = 4.4: pVysota = 32
        Case "6.5г"
            pPloschadPoperSech = 7.51: pShirina = 65: pTolshchinaStenki = 4.4: pVysota = 36
        Case "8г"
            pPloschadPoperSech = 8.98: pShirina = 80: pTolshchinaStenki = 4.5: pVysota = 40
        Case "10г"
            pPloschadPoperSech = 10.9: pShirina = 100: pTolshchinaStenki = 4.5: pVysota = 46
        Case "12г"
            pPloschadPoperSech = 13.3: pShirina = 120: pTolshchinaStenki = 4.8: pVysota = 52
        Case "14г"
            pPloschadPoperSech = 15.6: pShirina = 140: pTolshchinaStenki = 4.9: pVysota = 58
        Case "16г"
            pPloschadPoperSech = 18.1: pShirina = 160: pTolshchinaStenki = 5: pVysota = 64
        Case "16рг"
            pPloschadPoperSech = 19.5: pShirina = 160: pTolshchinaStenki = 5: pVysota = 68
        Case "18г"
            pPloschadPoperSech = 20.7: pShirina = 180: pTolshchinaStenki = 5.1: pVysota = 70
        Case "18рг"
            pPloschadPoperSech = 22.2: pShirina = 180: pTolshchinaStenki = 5.1: pVysota = 74
        Case "20г"
            pPloschadPoperSech = 23.4: pShirina = 200: pTolshchinaStenki = 5.2: pVysota = 76
        Case "22г"
            pPloschadPoperSech = 26.7: pShirina = 220: pTolshchinaStenki = 5.4: pVysota = 82
        Case "24г"
            pPloschadPoperSech = 30.6: pShirina = 240: pTolshchinaStenki = 5.6: pVysota = 90
        Case "27г"
            pPloschadPoperSech = 35.2: pShirina = 270: pTolshchinaStenki = 6: pVysota = 95
        Case "30г"
            pPloschadPoperSech = 40.5: pShirina = 300: pTolshchinaStenki = 6.5: pVysota = 100
        Case "33г"
            pPloschadPoperSech = 46.5: pShirina = 330: pTolshchinaStenki = 7: pVysota = 105
        Case "36г"
            pPloschadPoperSech = 53.4: pShirina = 360: pTolshchinaStenki = 7.5: pVysota = 110
        Case "40г"
            pPloschadPoperSech = 61.5: pShirina = 400: pTolshchinaStenki = 8: pVysota = 115
    End Select
    
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

Public Function PloschadPoperSech()
    PloschadPoperSech = pPloschadPoperSech * 100
End Function

Public Function Obyom()
    Obyom = PloschadPoperSech * pDlina
End Function

Public Function tPeskostruy() As Double
    tPeskostruy = PeskostruynayaOchistka(pShirina, pDlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, Pryamougolnoye, pTolshchinaStenki, Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Public Function Shirina()
    Shirina = pShirina
End Function

Public Function Vysota()
    Vysota = pVysota
End Function

Public Function tZaus() As Double
    tZaus = 2 * ZachistkaZausencev_PoKonturu_Napilnikom(pMaterial, 2 * (pVysota + pShirina), False, False)
End Function

Public Function tPravka() As Double
    tPravka = KOEFFITSIENT_NA_PRAVKU * (PravkaZagIzListMatVRuch(pShirina, pDlina, pTolshchinaStenki, CInt(pMaterial)) + PravkaZagIzListMatVRuch(pVysota, pDlina, pTolshchinaStenki, CInt(pMaterial)))
End Function

Public Function tUpakovka()
    tUpakovka = UpakovyvaniyePramougolnika(pDlina, pShirina, pVysota)
End Function
