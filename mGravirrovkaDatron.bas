Attribute VB_Name = "mGravirrovkaDatron"
Option Explicit
Const tSmenaInstr As Double = 15 / 60
Const tOdnogoPromera As Double = 5 / 60
Const kOtdObs As Double = 1.08
Const SmmMin As Integer = 800
Const SmmMin_Pritupleniye As Integer = 2000

Const tUst As Double = 1
Const tPerem As Double = 4 / 60         '4 сек время перемещения
Const tVrezaniye As Double = 4 / 60     '4 сек время предварительного врезания
Const tPramoyUgol = 4 / 60              '4 сек один угол

Private Const TOLSHCHINA_LISTA As Double = 0.8

Public Type NormyMetallografika
    tDatron As Double
    tSles As Double
    tKontrol As Double
    tVyrezaniye As Double
End Type

Private Function SformirovatMassivNorm(Normy As NormyMetallografika) As Variant
    
    Dim ArrNormy(1 To 4) As Double
    
    ArrNormy(1) = Normy.tDatron
    ArrNormy(2) = Normy.tSles
    ArrNormy(3) = Normy.tKontrol
    ArrNormy(4) = Normy.tVyrezaniye
    
    SformirovatMassivNorm = ArrNormy

End Function

Function Datron_tMashPerimetr(a As Double, b As Double, Optional PramoyUgol As Boolean = False, Optional Pritupleniye As Boolean = True, _
    Optional ByVal KolVoOkon As Double = 1) As Variant
    
    Dim Normy As NormyMetallografika

    Dim L As Double
    L = 2 * (a + b)
    
    If Not L > 0 Then Exit Function
    
    If KolVoOkon <= 0 Then KolVoOkon = 1
    
    Dim tMash As Double
    tMash = Datron_tMashFrezerovaniye(L, Pritupleniye)
    
    Dim tUgly As Double
    If PramoyUgol Then
        tUgly = tPerem + 4 * tPramoyUgol
    Else
        tUgly = 0
    End If

    Normy.tDatron = KolVoOkon * (tMash + tUgly)

    Normy.tSles = 0.5 + SnyatiyeZausentsevPoKonturu(L, ALUMINIYEVYYE_SPLAVY, True, NADFIL, True)
    If PramoyUgol Then
        Normy.tSles = Normy.tSles + 4 * VypilivaniyeVnutrennegoRadiusa(TOLSHCHINA_LISTA, 1, ALUMINIYEVYYE_SPLAVY)
    End If
    Normy.tSles = KolVoOkon * Normy.tSles
    
    Normy.tKontrol = IzmerShtangenCircul_CPL(a, b, False, KolVoOkon)
    Normy.tVyrezaniye = KolVoOkon * VyrezaniyeProkladokVruchnuyu(ALUMINIYEVYYE_SPLAVY, TOLSHCHINA_LISTA, L, , True, 2)

    Datron_tMashPerimetr = SformirovatMassivNorm(Normy)
    
End Function
Function Datron_tMashUgol(KolVoUglov As Double) As Variant

    Dim Normy As NormyMetallografika
    
    If KolVoUglov <= 0 Then Exit Function
    
    Normy.tDatron = KolVoUglov * (tPerem + tPramoyUgol)
    Normy.tSles = KolVoUglov * 0.5 + VypilivaniyeVnutrennegoRadiusa(TOLSHCHINA_LISTA, 1, ALUMINIYEVYYE_SPLAVY)
    Normy.tKontrol = 0
    Normy.tVyrezaniye = 0
    
    Datron_tMashUgol = SformirovatMassivNorm(Normy)

End Function

Function Datron_tOtverstiye(d As Double, Optional Zamok As Boolean = False, Optional Pritupleniye As Boolean = True, _
    Optional ByVal KolVoOtverstiy As Double = 1) As Variant
    
    If Not d > 0 Then Exit Function
    
    Dim Normy As NormyMetallografika
    
    If KolVoOtverstiy <= 0 Then KolVoOtverstiy = 1
    
    Dim tMash As Double
    tMash = Datron_tMashFrezerovaniye(PI * d, Pritupleniye)

    Dim tZamok As Double
    If Zamok Then
        tZamok = tPerem + 3 / 60
    Else
        tZamok = 0
    End If
    
    Normy.tDatron = KolVoOtverstiy * (tMash + tZamok)
    Normy.tSles = KolVoOtverstiy * (0.5 + SnyatiyeZausentsevPoKonturu(d * PI, ALUMINIYEVYYE_SPLAVY, 1, NADFIL, 1))
    Normy.tKontrol = IzmerenieOtverstiiProbkoi(d, TOLSHCHINA_LISTA, 11, KolVoOtverstiy)
    Normy.tVyrezaniye = KolVoOtverstiy * VyrezaniyeProkladokVruchnuyu(ALUMINIYEVYYE_SPLAVY, TOLSHCHINA_LISTA, PI * d, , True, 2)
    
    Datron_tOtverstiye = SformirovatMassivNorm(Normy)

End Function

Private Function Datron_tMashFrezerovaniye(L As Double, Pritupleniye As Boolean) As Double
    
    Dim tFrez As Double
    tFrez = tPerem + L / SmmMin
    
    Dim tPritupleniye As Double
    If Pritupleniye Then tPritupleniye = tPerem + L / SmmMin_Pritupleniye
    
    Datron_tMashFrezerovaniye = tVrezaniye + tFrez + tPritupleniye
    
End Function

Function GravirovkaDatron_Sivmoly(Shrift As Double, KolVoSimvolov As Double) As Double

    'Для шрифтов >=8 используется фреза (фрезеруется наружный и внутренний контур символа).
    'Для меньших шрифтов используетсфя резец (контур символа не описывается, необходимая ширина символа достигается за счёт углубления резца).
    
    Dim tSimvol_sek As Double
    Select Case Shrift
        Case Is < 3:            tSimvol_sek = Shrift * (5.3 / 3)
        Case Is <= 3:           tSimvol_sek = 5.3
        Case Is <= 5:           tSimvol_sek = 8.7
        Case Is <= 6:           tSimvol_sek = 12.4
        Case Is <= 8:           tSimvol_sek = 40.7
        Case Else:              tSimvol_sek = Shrift * (40.7 / 8)
    End Select
    
    GravirovkaDatron_Sivmoly = (tSimvol_sek * KolVoSimvolov) / 60
    
End Function
Function GravirovkaDatron_Molniya(Vysota As Double) As Double
    
    Dim tGrav_Min As Double
    tGrav_Min = 6.4 'Фактическое время граирования молнии высотой 30
    
    If Vysota <> 30 Then
        tGrav_Min = Vysota * (tGrav_Min / 30)

    End If
    GravirovkaDatron_Molniya = tGrav_Min
    
End Function
Function GravirovkaDatron_Zemlya(Vysota As Double) As Double
    
    Dim tGrav_Min As Double
    tGrav_Min = 0.3 'Фактическое время граирования молнии высотой 10
    
    If Vysota <> 10 Then
        tGrav_Min = Vysota * (tGrav_Min / 10)

    End If
    GravirovkaDatron_Zemlya = tGrav_Min
    
End Function


Function GravirovkaDatron_Linii(Dlina As Double) As Double
    
    Dim Skorost_mm_sek As Double
    Skorost_mm_sek = 4

    GravirovkaDatron_Linii = Dlina / Skorost_mm_sek / 60
    
End Function

Function Tsht_GravirovkaDatron(tMash As Double, KolVoRazhykhPerekhodov As Integer, KolVoUstanovok As Integer) As Double
    
    If KolVoUstanovok < 1 Then KolVoUstanovok = 1
    
    Dim KolVoInstrumentov As Integer
    KolVoInstrumentov = KolVoRazhykhPerekhodov + 1 '+1 - это щуп для промера
    
    Tsht_GravirovkaDatron = kOtdObs * (KolVoUstanovok * tUst + KolVoInstrumentov * tSmenaInstr + tMash)
    
End Function

Function Tsht_GravirovkaDatron_Metallografika(Dlina As Double, Shirina As Double, tMash As Double, KolVoRazhykhPerekhodov As Integer, KolVoUstanovok As Integer) As Double
    
    If KolVoUstanovok < 1 Then KolVoUstanovok = 1
    
    Dim tPromer As Double
    tPromer = RaschotVremeniPromera(Dlina, Shirina)
    
    Dim KolVoInstrumentov As Integer
    KolVoInstrumentov = KolVoRazhykhPerekhodov + 1 '+1 - это щуп для промера
    
    Tsht_GravirovkaDatron_Metallografika = kOtdObs * (KolVoUstanovok * (tUst + tPromer) + KolVoInstrumentov * tSmenaInstr + tMash)
    
End Function

Private Function RaschotVremeniPromera(Dlina As Double, Shirina As Double) As Double
    
    Dim KolVoPromerov As Integer
    With WorksheetFunction
        KolVoPromerov = .RoundUp(Dlina / 50, 0) * .RoundUp(Shirina / 50, 0)
    End With
        
    RaschotVremeniPromera = tOdnogoPromera * KolVoPromerov
    
End Function
