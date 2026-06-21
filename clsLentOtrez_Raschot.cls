VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLentOtrez_Raschot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const kLentOtrez As Integer = 4
Private Const tObs As Double = 1.1

Private Type SpisokNorm
    tLentOtrez As Double
    tPeskostruy As Double
    tZakalka As Double
    tOtpusk As Double
    tKontrolLineynogoRazmera As Double
    tVisualnyyKontrol As Double
    tZaus As Double
    tPravka As Double
    tUpakovka As Double
End Type

Public Function tLentOtrez(ByVal Figura As ITechnologicalShape) As Double
    With Figura
        tLentOtrez = LentOtrezPoPloshchadiIMasse(.Material, .PloshchadPoperSech, .Massa, .Tonkostennaya)
    End With
End Function

Public Function RaschotMarshrutaFigury(ByVal Figura As ITechnologicalShape, ByVal Rikhtovka As Boolean) As Variant
    Dim Normy As SpisokNorm
    
    With Figura
        Normy.tLentOtrez = Me.tLentOtrez(Figura)
        
        Normy.tPeskostruy = .tPeskostruy
        Normy.tZakalka = .tZakalka
        Normy.tOtpusk = .tOtpusk
        Normy.tUpakovka = .tUpakovka
        Normy.tZaus = .tZaus
        Normy.tPravka = .tPravka
        
        RaschotMarshrutaFigury = SostavitMarshrut(.Dlina, .PerimetrIliDiametr, Rikhtovka, Normy, .Massa)
    End With
End Function

Private Function SostavitMarshrut(ByVal Dlina As Double, ByVal PerimetrIliDiametr As Double, ByVal Rikhtovka As Boolean, ByRef Normy As SpisokNorm, ByVal Massa As Double) As Variant
    Const K_TIP_PROIZV As Double = 1.3
   
    Normy.tKontrolLineynogoRazmera = IzmereniyeLineykoyRuletkoy(Dlina)
    Normy.tVisualnyyKontrol = VisualnyyKontrol(PerimetrIliDiametr, Dlina)
    
    Dim Operations As Variant
    Operations = Array( _
                Array("Ленточно-отрезная", 10, Normy.tLentOtrez), _
                Array("Контрольная", 0, K_TIP_PROIZV * Normy.tKontrolLineynogoRazmera), _
                Array("Маркирование (этикеткой)", 0, 3), _
                Array("Упаковывание", 0, K_TIP_PROIZV * Normy.tUpakovka), _
                Array("Контрольная", 0, K_TIP_PROIZV * Normy.tVisualnyyKontrol), _
                Array("Закалка", 60, Normy.tZakalka), _
                Array("Отпуск", 30, Normy.tOtpusk), _
                Array("Контрольная", 0, 2), _
                Array("Пескоструйная", 5, K_TIP_PROIZV * Normy.tPeskostruy), _
                Array("Контрольная", 0, K_TIP_PROIZV * Normy.tVisualnyyKontrol))
    
    Dim tSlesar As Double
    tSlesar = Normy.tZaus + IIf(Rikhtovka, Normy.tPravka, 0) + UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(Massa)
    
    If tSlesar > 0 Then
        ReDim Preserve Operations(UBound(Operations) + 1)
        Operations(UBound(Operations)) = Array("Слесарная", 5, K_TIP_PROIZV * tSlesar)
    End If
    
    Dim i As Integer
    For i = LBound(Operations) To UBound(Operations)
        Operations(i)(2) = OkruglenieTsht(CDbl(Operations(i)(2)))
    Next
    
    ReDim Preserve Operations(UBound(Operations) + 1)
    Operations(UBound(Operations)) = Array("Масса", Massa, "")
    
    SostavitMarshrut = Operations
End Function

Private Function LentOtrezPoPloshchadiIMasse(ByVal Material As EnumMaterialy, ByVal Ploshchad As Double, ByVal Massa As Double, ByVal Tonkosten As Boolean) As Double
    Dim tMash As Double
    tMash = tMash_LentOtrez(Material, Ploshchad, Tonkosten)
    
    Dim tUst As Double
    tUst = Ustanov_NaStoleKreplBoltPlank(Massa, 1, 0)
    
    LentOtrezPoPloshchadiIMasse = tObs * (tMash + tUst)
End Function
 
Private Function tMash_LentOtrez(ByVal Material As EnumMaterialy, ByVal Ploshchad As Double, ByVal Tonkosten As Boolean) As Double
    If Ploshchad <= 0 Then Exit Function

    Dim K As Double
    Select Case Material
        Case 1: K = 1
        Case 2: K = 1.5
        Case 3: K = 2
        Case 4: K = 3
        Case 5: K = 4
    End Select
    
    tMash_LentOtrez = IIf(Tonkosten, kLentOtrez, 1) * K * 8 * 10 ^ -4 * Ploshchad
End Function

