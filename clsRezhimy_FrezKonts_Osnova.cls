VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRezhimy_FrezKonts_Osnova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

Public Enum EnumMillingStage
    Roughing = 1  ' Черновая
    Finishing = 2 ' Чистовая
End Enum

Private Const DOPUSTIOMOYE_UVELICHENIYE As Double = 1.05
Private Const MIN_GLUBINA_REZANIYA As Integer = 1

' Массивы для числовых рядов станка (вместо строк)
Private ArrPodachi() As Variant
Private ArrOboroty() As Variant

Private Sub Class_Initialize()
    ArrPodachi = Array(8, 10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400)
    ArrOboroty = Array(50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500)
End Sub

Public Function RaschotMinutnoyPodachi(Stadiya As EnumMillingStage, Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double, _
    ByVal GlubinaRezaniya As Double, Korka As Boolean) As clsRezhimIndicators
    
    Dim Pokazateli As New clsRezhimIndicators
    
    If GlubinaRezaniya < MIN_GLUBINA_REZANIYA Then
        GlubinaRezaniya = MIN_GLUBINA_REZANIYA
    End If
    
    Dim Sz As Double
    Sz = RaschotPodachiNaZub(Stadiya, Material, Dfrezy, ShirinaFrezerovaniya, GlubinaRezaniya)
    
    Pokazateli.N = RaschotOborotov(Material, Dfrezy, ShirinaFrezerovaniya, GlubinaRezaniya, Sz, Korka)
    Pokazateli.V = RaschotFakticheskoySkorostiRezaniya(Dfrezy, Pokazateli.N)
    
    Dim Sraschot As Double
    Sraschot = Sz * RaschotChislaZubyev(Dfrezy) * Pokazateli.N
    
    Pokazateli.S = VybratZnaceniyeRezhimaIzMassiva(Sraschot, ArrPodachi)
    
    Set RaschotMinutnoyPodachi = Pokazateli
    
End Function

Private Function RaschotPodachiNaZub(Stadiya As EnumMillingStage, Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double, GlubinaRezaniya As Double) As Double

    If Stadiya = Roughing Then
        RaschotPodachiNaZub = RaschotChernovoyPodachiNaZub(Material, Dfrezy, ShirinaFrezerovaniya, GlubinaRezaniya)
    ElseIf Stadiya = Finishing Then
        RaschotPodachiNaZub = RaschotChistovoyPodachiNaZub(Material, Dfrezy, ShirinaFrezerovaniya, GlubinaRezaniya)
    End If
    
End Function

Private Function RaschotChislaZubyev(Dfrezy As Double) As Integer
    Select Case Dfrezy
        Case Is <= 25:  RaschotChislaZubyev = 3
        Case Is <= 50:  RaschotChislaZubyev = 4
        Case Else:      RaschotChislaZubyev = 5
    End Select
End Function

Private Function RaschotChernovoyPodachiNaZub(Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double, GlubinaRezaniya As Double) As Double

    Dim Pokazateli As Variant
    If Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Then
        Pokazateli = Array(0.0397, 0.83, 0.41, 0.42)
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        Pokazateli = Array(0.0241, 0.89, 0.43, 0.46)
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        Pokazateli = Array(0.0509, 0.85, 0.41, 0.44)
    Else
        Exit Function
    End If
    
    Dim Cs As Double, ns1 As Double, ns2 As Double, ns3 As Double
    Cs = Pokazateli(0)
    ns1 = Pokazateli(1)
    ns2 = Pokazateli(2)
    ns3 = Pokazateli(3)
    
    RaschotChernovoyPodachiNaZub = Cs * (Dfrezy ^ ns1) / (GlubinaRezaniya ^ ns2 * ShirinaFrezerovaniya ^ ns3)
    
End Function

Private Function RaschotChistovoyPodachiNaZub(Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double, GlubinaRezaniya As Double) As Double

    Dim Pokazateli As Variant
    If Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Then
        Pokazateli = Array(0.0296, 0.75, 0.36, 0.38)
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        Pokazateli = Array(0.0214, 0.75, 0.36, 0.41)
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        Pokazateli = Array(0.0358, 0.79, 0.38, 0.4)
    Else
        Exit Function
    End If
    
    Dim Cs As Double, ns1 As Double, ns2 As Double, ns3 As Double
    Cs = Pokazateli(0)
    ns1 = Pokazateli(1)
    ns2 = Pokazateli(2)
    ns3 = Pokazateli(3)
    
    Dim Sz As Double
    Sz = Cs * (Dfrezy ^ ns1) / (GlubinaRezaniya ^ ns2 * ShirinaFrezerovaniya ^ ns3)
    Sz = WorksheetFunction.Min(Sz, PodachaDopustiomayaPoSherokhovatosti(Dfrezy))
    
    RaschotChistovoyPodachiNaZub = Sz
    
End Function

Private Function PodachaDopustiomayaPoSherokhovatosti(Dfrezy As Double) As Double
    ' Подача допустимая по шероховатости Ra2.5
    Dim Sz As Double
    Select Case Dfrezy
        Case Is <= 6: Sz = 0.04
        Case Is <= 8: Sz = 0.05
        Case Is <= 10: Sz = 0.05
        Case Is <= 12: Sz = 0.06
        Case Is <= 16: Sz = 0.07
        Case Is <= 20: Sz = 0.08
        Case Is <= 25: Sz = 0.08
        Case Is <= 32: Sz = 0.07
        Case Is <= 40: Sz = 0.08
        Case Is <= 50: Sz = 0.09
        Case Is <= 63: Sz = 0.08
        Case Else: Exit Function
    End Select
    PodachaDopustiomayaPoSherokhovatosti = Sz
End Function

Public Function RaschotOborotov(Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double, GlubinaRezaniya As Double, PodachaNaZub As Double, Korka As Boolean) As Double

    Dim Pokazateli As Variant
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        Pokazateli = Array(21.5, 0.5, 0.31, 0.21)
    ElseIf Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA _
        Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        Pokazateli = Array(5.88, 0.48, 0.23, 0.21)
    Else
        Exit Function
    End If

    Dim Cv As Double, nv1 As Double, nv2 As Double, nv3 As Double
    Cv = Pokazateli(0)
    nv1 = Pokazateli(1)
    nv2 = Pokazateli(2)
    nv3 = Pokazateli(3)

    Dim Kvm As Double
    If Material = STAL_LEGIROVANNAYA Then
        Kvm = 0.8
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        Kvm = 0.5
    ElseIf Material = TITANOVYYE_SPLAVY Then
        Kvm = 0.4
    Else
        Kvm = 1
    End If
    
    Dim Kvk As Double
    Kvk = IIf(Korka, 0.8, 1)
    
    Dim Kvb As Double
    Kvb = RaschotKoefShirinyFrez(Material, Dfrezy, ShirinaFrezerovaniya)
    
    Dim vRaschot As Double
    vRaschot = Kvm * Kvk * Kvb * (Cv * Dfrezy ^ nv1) / (GlubinaRezaniya ^ nv2 * PodachaNaZub ^ nv3)
    
    Dim nRaschot As Double
    nRaschot = Round((1000 * vRaschot) / (PI * Dfrezy), 0)
    
    RaschotOborotov = VybratZnaceniyeRezhimaIzMassiva(nRaschot, ArrOboroty)
    
End Function

Public Function RaschotFakticheskoySkorostiRezaniya(Dfrezy As Double, N As Double) As Double
    RaschotFakticheskoySkorostiRezaniya = Round((PI * Dfrezy * N) / 1000, 0)
End Function

Private Function RaschotKoefShirinyFrez(Material As EnumMaterialy, Dfrezy As Double, ShirinaFrezerovaniya As Double) As Double
    
    Dim Bnorm As Double
    Select Case Dfrezy
        Case Is <= 8:   Bnorm = 5
        Case Is <= 16:  Bnorm = 10
        Case Is <= 32:  Bnorm = 20
        Case Is <= 63:  Bnorm = 30
        Case Else:      Exit Function
    End Select
    
    Dim mv As Double
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        mv = -0.11
    ElseIf Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA _
        Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        mv = -0.254
    Else
        Exit Function
    End If
    
    RaschotKoefShirinyFrez = (ShirinaFrezerovaniya / Bnorm) ^ mv
    
End Function

' Обновленная математически точная функция поиска по числовому массиву
Private Function VybratZnaceniyeRezhimaIzMassiva(RezhimRaschot As Double, ByRef ArrRezhimov() As Variant) As Double

    Dim RezhimMaks As Double
    RezhimMaks = ArrRezhimov(UBound(ArrRezhimov))
    
    Dim RezhimDopustiomyy As Double
    RezhimDopustiomyy = DOPUSTIOMOYE_UVELICHENIYE * RezhimRaschot
    
    If RezhimDopustiomyy >= RezhimMaks Then
        VybratZnaceniyeRezhimaIzMassiva = RezhimMaks
        Exit Function
    End If
    
    Dim i As Long
    For i = UBound(ArrRezhimov) To LBound(ArrRezhimov) Step -1
        If ArrRezhimov(i) <= RezhimDopustiomyy Then
            VybratZnaceniyeRezhimaIzMassiva = ArrRezhimov(i)
            Exit Function
        End If
    Next i
    
    VybratZnaceniyeRezhimaIzMassiva = ArrRezhimov(LBound(ArrRezhimov))
    
End Function

