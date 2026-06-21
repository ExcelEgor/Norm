VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFigura_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDlina As Double
Private pShirina As Double
Private pTolshchina As Double
Private pMaterial As EnumMaterialy
Private pTipStali As Integer

Implements ITechnologicalShape

Public Property Get Tonkostennaya() As Boolean: Tonkostennaya = False: End Property

Private Property Get ITechnologicalShape_Tonkostennaya() As Boolean: ITechnologicalShape_Tonkostennaya = Me.Tonkostennaya: End Property

Private Property Get ITechnologicalShape_tPeskostruy() As Double: ITechnologicalShape_tPeskostruy = Me.tPeskostruy: End Property

Private Property Get ITechnologicalShape_tZakalka() As Double: ITechnologicalShape_tZakalka = Me.tZakalka: End Property

Private Property Get ITechnologicalShape_tOtpusk() As Double: ITechnologicalShape_tOtpusk = Me.tOtpusk: End Property

Private Property Get ITechnologicalShape_tUpakovka() As Double: ITechnologicalShape_tUpakovka = Me.tUpakovka: End Property

Private Property Get ITechnologicalShape_tZaus() As Double: ITechnologicalShape_tZaus = ZachistkaZausencev_PoKonturu_Napilnikom(pMaterial, 4 * (pShirina + pDlina) + 4 * pTolshchina, False, False): End Property

Private Property Get ITechnologicalShape_tPravka() As Double: ITechnologicalShape_tPravka = Me.tPravka: End Property

Private Property Get ITechnologicalShape_Massa() As Double: ITechnologicalShape_Massa = Me.Massa: End Property

Private Property Get ITechnologicalShape_Dlina() As Double: ITechnologicalShape_Dlina = pDlina: End Property

Private Property Get ITechnologicalShape_PerimetrIliDiametr() As Double: ITechnologicalShape_PerimetrIliDiametr = 2 * (pShirina + pTolshchina): End Property

Private Property Get ITechnologicalShape_Material() As EnumMaterialy: ITechnologicalShape_Material = pMaterial: End Property

Private Property Get ITechnologicalShape_PloshchadPoperSech() As Double: ITechnologicalShape_PloshchadPoperSech = PloschadPoperSech: End Property

Private Property Get ITechnologicalShape_PloshchadPolnaya() As Double: ITechnologicalShape_PloshchadPolnaya = PloschadPolnaya: End Property

''''''''''''''''''''''''''''''''''''''''

Public Sub Init(Material As EnumMaterialy, Shirina As Double, Tolshchina As Double, Dlina As Double)

    pMaterial = Material
    pTipStali = IIf(pMaterial = STAL_NERZHAVEYUSHCHAYA, 2, 1)
    
    pDlina = Dlina
    
    pShirina = Shirina
    
    pTolshchina = Tolshchina
    
End Sub

''''''''''''''''''''''''''''''''''''''''

Public Function PloschadPoperSech()
    PloschadPoperSech = pShirina * pTolshchina
End Function

Public Function PloschadPolnaya()
    PloschadPolnaya = 2 * (pDlina * pShirina + pDlina * pTolshchina + pShirina * pTolshchina)
End Function

Public Function Obyom()
    Obyom = pShirina * pTolshchina * pDlina
End Function

Public Function Massa()
    Massa = RaschotMassyPoMaterialu_i_Obyomu(pMaterial, Obyom)
End Function

''''''''''''''''''''''''''''''''''''''''

Public Function tOtrezkiNaGilotinnykhNozhnitsakhPoRazmetke() As Double
    tOtrezkiNaGilotinnykhNozhnitsakhPoRazmetke = _
        1.2 * OtrezkaNaGilotinNozhnitsZaDvaReza(CInt(pMaterial), pDlina, pShirina, pTolshchina)
End Function

Public Function tPeskostruy() As Double
    tPeskostruy = PeskostruynayaOchistka(Gabarity.Shirina, Gabarity.Dlina, 2, 2, Massa, pMaterial)
End Function

Public Function tZakalka() As Double
    tZakalka = Zakalka(pTipStali, IIf(Gabarity.Tolshchina = Gabarity.Shirina, Kvadratnoye, Pryamougolnoye), Gabarity.Tolshchina, Massa)
End Function

Public Function tOtpusk() As Double
    tOtpusk = Otpusk(pTipStali, Massa)
End Function

Private Property Get Gabarity() As GabarityZagotovki
    Gabarity = RaschotGabaritov(pDlina, pShirina, pTolshchina)
End Property

Public Function tUpakovka()
    tUpakovka = UpakovyvaniyePramougolnika(pDlina, pShirina, pTolshchina)
End Function

Public Function tZaus(Optional KolVoStoron As Integer = 2) As Double
    
    Dim DlinaZachistki As Double
    DlinaZachistki = IIf(KolVoStoron = 2, 4, 2) * (pDlina + pShirina) + 4 * pTolshchina
    
    If DlinaZachistki <= 0 Then Exit Function
    
    Dim pMassa As Double
    pMassa = Massa
    
    Dim tUst As Double
    tUst = UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(pMassa) + 3 * UstanovkaSlesar_NaStoleBezKrepleniya_Povorot(pMassa)

    tZaus = tUst + ZachistkaZausencev(pMaterial, DlinaZachistki, 2, False, pMassa > 20)
    
End Function

Public Function tPravka() As Double

    If Not pDlina > 0 Then Exit Function
    tPravka = KOEFFITSIENT_NA_PRAVKU * PravkaPloskikhDetaley(pTolshchina, pDlina, pShirina, CInt(pMaterial))
    
End Function

Public Function tVisual() As Double
    tVisual = VisualnyyKontrol(2 * (pShirina + pTolshchina), pDlina, 1, False, False)
End Function

Public Function tKontrol() As Double
    tKontrol = IzmerShtangenCircul_CPL(pDlina, pShirina)
End Function

Public Function OtrezkaNaTsirkulyarnykhPilakh_Marshrut(KolVoDetVzagotovke As Double) As Variant
    
    Const K_TIP_PROIZV As Double = 1.3
    
    Dim Operations As Variant
    Operations = Array( _
                Array("Отрезная", KolVoDetVzagotovke, 2, 10, K_TIP_PROIZV * OtrezkaNaTsirkulyarnykhPilakhZaDvaReza), _
                Array("Контрольная", KolVoDetVzagotovke, 3, 0, K_TIP_PROIZV * tKontrol), _
                Array("Маркирование (этикеткой)", KolVoDetVzagotovke, 2, 0, T_MARKIROVKA), _
                Array("Упаковывание", KolVoDetVzagotovke, 2, 0, K_TIP_PROIZV * tUpakovka), _
                Array("Контрольная", KolVoDetVzagotovke, 3, 0, K_TIP_PROIZV * tVisual))
                
    Dim i As Integer
    For i = LBound(Operations) To UBound(Operations)
        Operations(i)(4) = OkruglenieTsht(CDbl(Operations(i)(4)))
    Next
                
    OtrezkaNaTsirkulyarnykhPilakh_Marshrut = Operations

End Function

Private Function OtrezkaNaTsirkulyarnykhPilakhZaDvaReza() As Variant

    Dim tRazmetka As Double
    tRazmetka = RazmetkaKonturaBezShablona_Lineykoy(pMaterial, MassaLista(pDlina, pShirina, pTolshchina, pMaterial), 2 * (pDlina + pShirina))
    
    Dim tOtrezka As Double
    tOtrezka = OtrezkaTsirkulyarnoyPiloy_List(pMaterial, pDlina, pTolshchina) + OtrezkaTsirkulyarnoyPiloy_List(pMaterial, pShirina, pTolshchina)
    
    OtrezkaNaTsirkulyarnykhPilakhZaDvaReza = tRazmetka + tOtrezka

End Function
