VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_Krug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDiametr As Double
Private pDlina As Double
Private pMaterial As EnumMaterialy
Private pTipStali As Integer
Private pRadius As Double

Implements ITechnologicalShape

Public Property Get Tonkostennaya() As Boolean: Tonkostennaya = False: End Property

Private Property Get ITechnologicalShape_Tonkostennaya() As Boolean: ITechnologicalShape_Tonkostennaya = Me.Tonkostennaya: End Property

Private Property Get ITechnologicalShape_tPeskostruy() As Double: ITechnologicalShape_tPeskostruy = Me.tPeskostruy: End Property

Private Property Get ITechnologicalShape_tZakalka() As Double: ITechnologicalShape_tZakalka = Me.tZakalka: End Property

Private Property Get ITechnologicalShape_tOtpusk() As Double: ITechnologicalShape_tOtpusk = Me.tOtpusk: End Property

Private Property Get ITechnologicalShape_tUpakovka() As Double: ITechnologicalShape_tUpakovka = Me.tUpakovka: End Property

Private Property Get ITechnologicalShape_tZaus() As Double: ITechnologicalShape_tZaus = 0: End Property

Private Property Get ITechnologicalShape_tPravka() As Double: ITechnologicalShape_tPravka = 0: End Property

Private Property Get ITechnologicalShape_Massa() As Double: ITechnologicalShape_Massa = Me.Massa: End Property

Private Property Get ITechnologicalShape_Dlina() As Double: ITechnologicalShape_Dlina = pDlina: End Property

Private Property Get ITechnologicalShape_PerimetrIliDiametr() As Double: ITechnologicalShape_PerimetrIliDiametr = pDiametr: End Property

Private Property Get ITechnologicalShape_Material() As EnumMaterialy: ITechnologicalShape_Material = pMaterial: End Property

Private Property Get ITechnologicalShape_PloshchadPoperSech() As Double: ITechnologicalShape_PloshchadPoperSech = PloschadPoperSech: End Property

Private Property Get ITechnologicalShape_PloshchadPolnaya() As Double: ITechnologicalShape_PloshchadPolnaya = PloschadPolnaya: End Property

''''''''''''''''''''''''''''''''''''''''

Public Sub Init(Material As EnumMaterialy, Diametr As Double, Dlina As Double)

    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)
    
    pDiametr = Diametr
    pRadius = 0.5 * pDiametr
    
    pDlina = Dlina
    
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function PloschadPoperSech() As Double
    PloschadPoperSech = PI * pRadius ^ 2
End Function

Public Function PloschadPolnaya() As Double
    PloschadPolnaya = 2 * PI * pRadius * pDlina + 2 * PI * pRadius ^ 2
End Function

Private Function Obyom() As Double
    Obyom = PloschadPoperSech * pDlina
End Function

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

''''''''''''''''''''''''''''''''''''''''

Public Function tPeskostruy() As Double
    tPeskostruy = PeskostruynayaOchistka(pDiametr, pDlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, Krugloe, pDiametr, Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Public Function tUpakovka() As Double
    tUpakovka = UpakovyvaniyeKruglogoPrutka(pDiametr, pDlina)
End Function

Public Function tVisualKontrol() As Double
    tVisualKontrol = VisualnyyKontrol(pDiametr, pDlina, 1, False, False)
End Function

