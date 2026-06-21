Attribute VB_Name = "mTokarnyye_Normativ"
Option Explicit

Public Const DMAX_TOKARNYY As Integer = 400
Public Const DMAX_GALTEL As Integer = 300
Public Const DMAX_SVERLENIYE_TOKARNYY As Integer = 40
Public Const DMAX_DOVODKA As Integer = 100
Public Const DMAX_KONUS As Integer = 100
Public Const DMAX_NAKATYVANIYE As Integer = 100
Public Const DMAX_METRREZBAREZTSOM As Integer = 120
Public Const DMAX_METCHIK As Integer = 20
Public Const DMAX_PLASHKA As Integer = 20
Public Const MAX_SHIRINA_KANAVKI As Integer = 20

Public Enum TipObrabotkiOtverstiya_Tokar
    Sverleniye = 1
    Rassverlivaniye = 2
    Zenkerovaniye = 3
    Razvertyvaniye_Ra2_5 = 4
    Razvertyvaniye_Ra1_25 = 5
End Enum

Public Enum EnumSposobNarezkiRezby
    Reztsom = 1
    PlashkoyIliMetchikom = 2
End Enum

Private Const MIN_RA As Double = 0.63
Private Const MIN_IT As Integer = 6

'Материалы
'1 - ЛС59-1
'2 - АЛ2
'3 - Сталь 35
'4 - 25Х13Н2
'5 - Титан ВТ1-0
'6 - 12Х18Н9Т
'7 - ВНМ 3-2

Function PreobrazovatMaterial(ByVal Material As Integer) As Integer
    Select Case Material
        Case EnumMaterialy.MEDNYYE_SPLAVY:           Material = 1
        Case EnumMaterialy.ALUMINIYEVYYE_SPLAVY:     Material = 2
        Case EnumMaterialy.STAL_UGLERODISTAYA:       Material = 3
        Case EnumMaterialy.STAL_LEGIROVANNAYA:       Material = 4
        Case EnumMaterialy.STAL_NERZHAVEYUSHCHAYA:   Material = 6
        Case EnumMaterialy.TITANOVYYE_SPLAVY:        Material = 7
    End Select
    PreobrazovatMaterial = Material
End Function

Private Function ProverkaOsnovnykhParametrov(Materil As EnumMaterialy, d As Double, L As Double, Optional Dmax = DMAX_TOKARNYY) As Boolean

    If d > Dmax Or d <= 0 Then
        ProverkaOsnovnykhParametrov = False
        
    ElseIf L <= 0 Then
        ProverkaOsnovnykhParametrov = False
        
    ElseIf Materil = EnumMaterialy.ALUMINIYEVYYE_SPLAVY Or _
        Materil = EnumMaterialy.MEDNYYE_SPLAVY Or _
        Materil = EnumMaterialy.STAL_LEGIROVANNAYA Or _
        Materil = EnumMaterialy.STAL_NERZHAVEYUSHCHAYA Or _
        Materil = EnumMaterialy.STAL_UGLERODISTAYA Or _
        Materil = EnumMaterialy.TITANOVYYE_SPLAVY Then
        
        ProverkaOsnovnykhParametrov = True
    
    Else
        ProverkaOsnovnykhParametrov = False
    End If
    
End Function

Function ProdolnoeTochenie(Material As EnumMaterialy, DiametrZagotovki As Double, DiametrDetali As Double, L As Double, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, _
    Optional ByVal Korka As Boolean = False, Optional Udar As Boolean = False, _
    Optional TonkostennayaKonfiguraciya As Boolean = False, Optional PodrezkaTorca As Boolean = False, Optional PoRezhimam As Boolean = True)
    
    If Not (L > 0 And DiametrDetali > 0) Then Exit Function
    If DiametrDetali > DiametrZagotovki Then Exit Function
    
    If IT = 0 Then IT = 12
    If Ra = 0 Then Ra = 10
    If GlubinaRezaniya = 0 Then GlubinaRezaniya = 2
    
    Dim tTok As Double
    If PoRezhimam Then
        tTok = ProdolnoyeTocheniye_PoRezhimam(Material, DiametrZagotovki, DiametrDetali, GlubinaRezaniya, L, Ra, IT, Korka, Udar, PodrezkaTorca)
    
    Else
        tTok = ProdolnoyeTocheniye_Normativ(CInt(Material), DiametrZagotovki, DiametrDetali, L, IT, Ra, GlubinaRezaniya, Korka, Udar, TonkostennayaKonfiguraciya, PodrezkaTorca)
    
    End If
    
    Dim tKontrol As Double
    tKontrol = KontrolNaryzhnogoDiametra(DiametrDetali, L, IT)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    ProdolnoeTochenie = Normy

End Function

Private Function ProdolnoyeTocheniye_Normativ(Material As EnumMaterialy, DiametrZagotovki As Double, DiametrDetali As Double, _
    L As Double, IT As Integer, Ra As Double, GlubinaRezaniya As Double, Korka As Boolean, _
    Udar As Boolean, TonkostennayaKonfiguraciya As Boolean, PodrezkaTorca As Boolean) As Double

    Dim PripuskNaChistovuyuObrabotku As Double
    PripuskNaChistovuyuObrabotku = OpredelitPripuskNaChistovuyuObrabotku(CInt(Material), IT)

    Dim tChernovoiObrabotki As Double
    tChernovoiObrabotki = ProdolnoeTochenie_Chernovoye(CInt(Material), DiametrZagotovki, DiametrDetali + PripuskNaChistovuyuObrabotku, L, GlubinaRezaniya, Korka, Udar)

    Dim tChistovoiObrabotki As Double
    tChistovoiObrabotki = ProdolnoeTochenie_Chistovoe(Material, DiametrDetali, L, IT, Ra, TonkostennayaKonfiguraciya, Udar)
    
    Dim tPodrezki As Double, lPoperechnoye As Double
    lPoperechnoye = 0.5 * (DiametrZagotovki - DiametrDetali)
    If PodrezkaTorca Then
        tPodrezki = PoperechnoyeTochenieye_Chistovoe(Material, DiametrZagotovki, lPoperechnoye, Ra, IT, Udar, TonkostennayaKonfiguraciya)
    Else
        tPodrezki = 0
    End If
    
    ProdolnoyeTocheniye_Normativ = tChernovoiObrabotki + tChistovoiObrabotki + tPodrezki
    
End Function

Private Function OpredelitPripuskNaChistovuyuObrabotku(DiametrDetali As Double, IT As Integer) As Double

    Dim Pripusk As Double
    If IT > 11 Then
        If DiametrDetali > 70 Then
            Pripusk = 3
        Else
            Pripusk = 1.5
        End If
    Else
        If DiametrDetali > 70 Then
            Pripusk = 0.5
        Else
            Pripusk = 0.4
        End If
    End If
    
    OpredelitPripuskNaChistovuyuObrabotku = Pripusk
        
End Function

Private Function ProdolnoeTochenie_Chernovoye(Material As EnumMaterialy, DiametrZagotovki As Double, DiametrDetali As Double, _
    L As Double, GlubinaRezaniya As Double, Korka As Boolean, Udar As Boolean) As Double
    
    Dim MaksGlubinaRezaniya As Integer
    MaksGlubinaRezaniya = 6
    
    If ProverkaOsnovnykhParametrov(CInt(Material), DiametrZagotovki, L) = False Or GlubinaRezaniya > MaksGlubinaRezaniya Or DiametrDetali > DiametrZagotovki Then Exit Function
     
    Dim tNaProhod As Double, tChernovoiObrabotki As Double
    Dim ObrabotkaSPromerom As Boolean

    Dim i As Integer
    Dim DiametrObrabotki As Double
    For DiametrObrabotki = DiametrZagotovki To DiametrDetali Step -GlubinaRezaniya * 2
        
        If DiametrObrabotki <= (DiametrDetali + 2 * GlubinaRezaniya) Then ObrabotkaSPromerom = True
        
        tNaProhod = NaytyVMassive_ChernovoyeTocheniye(Material, DiametrObrabotki, L, GlubinaRezaniya, ObrabotkaSPromerom, Udar, Korka)

        tChernovoiObrabotki = tChernovoiObrabotki + tNaProhod
        
        Korka = False
        
    Next

    ProdolnoeTochenie_Chernovoye = tChernovoiObrabotki
        
End Function

Private Function NaytyVMassive_ChernovoyeTocheniye(Material As EnumMaterialy, DiametrObrabotki As Double, L As Double, GlubinaRezaniya As Double, _
    ObrabotkaSPromerom As Boolean, Udar As Boolean, Korka As Boolean) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_prodolnoye_chernovoye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= DiametrObrabotki And _
            Normativ(i, 3) >= GlubinaRezaniya And _
            Normativ(i, 4) = ObrabotkaSPromerom Then
            a = Normativ(i, 5)
            b = Normativ(i, 6)
            Exit For
        End If
    Next
    
    NaytyVMassive_ChernovoyeTocheniye = PoprKoefPriChernTocheniiRastachivanii(L, L, DiametrObrabotki, Udar, Korka, False) * (a * L + b)
        
End Function

Function ProdolnoeTochenie_Chistovoe(Material As EnumMaterialy, d As Double, L As Double, _
    IT As Integer, Ra As Double, TonkostennayaKonfiguraciya As Boolean, NaUdar As Boolean) As Double

    If ProverkaOsnovnykhParametrov(CInt(Material), d, L) = False Or IT < MIN_IT Or Ra < MIN_RA Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("tok_prodolnoye_chistovoye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) <= IT And _
            Normativ(i, 4) <= Ra Then
            a = Normativ(i, 5)
            b = Normativ(i, 6)
            Exit For
        End If
    Next
    
    Dim kLD As Double, kUdar As Double, kTonksosten As Double
    kLD = PoprKoefLD(d, L)
    kUdar = IIf(NaUdar, 1.2, 1)
    kTonksosten = IIf(TonkostennayaKonfiguraciya, 1.3, 1)
    
    ProdolnoeTochenie_Chistovoe = kLD * kUdar * kTonksosten * (a * L + b)
        
End Function

Function Dovodka(ByVal Material As Integer, TipPov As Integer, VUpor As Boolean, d As Double, L As Double, Ra As Double, Otkl As Double, TonkSten As Boolean)
    
    Material = PreobrazovatMaterial(Material)
    If ProverkaOsnovnykhParametrov(CInt(Material), d, L, DMAX_DOVODKA) = False Then Exit Function
    
    'TipPov
    '1 - Наружная
    '2 - Внутренняя
    '3 - Коническая
    
    Dim ct As Integer
    Select Case TipPov
        Case 1: ct = 20
        Case 2: ct = 44
        Case 3: ct = 68
    End Select
    ct = ct + IIf(TonkSten, 6, 0)
    If TipPov <> 3 Then ct = ct + IIf(VUpor, 12, 0)
    
    Dim rTime As Range 'Диапазон времени
    With ThisWorkbook.Worksheets("Ток_29-38")
        Set rTime = .Range(.Cells(ct, 2), .Cells(ct + 4, 23))
    End With
    
    Dim iOtkl As Integer
    Select Case Otkl
        Case Is <= 0.001:   iOtkl = 1
        Case Is <= 0.002:   iOtkl = 2
        Case Is <= 0.005:   iOtkl = 3
        Case Is <= 0.01:    iOtkl = 4
        Case Is <= 0.05:    iOtkl = 5
    End Select
    
    Dim ArrD()          'Массив диаметров
    Dim id As Integer   'Позиция ближайшего большего диаметра
    ArrD = Array(5, 10, 15, 20, 30, 40, 50, 70, 100)
    id = BlizhBolshRavn_Pozic_Array(d, ArrD)
    
    Dim iMat As Integer
    Select Case Material
        Case 3:         iMat = 0    'Сталь как Сталь 50, АЛ2
        Case 2:         iMat = 1    'Алюминиевые, бронзовые и медные сплавы как Л96
        Case 4, 6, 7:   iMat = 3    'Легированные стали, нержавеющие стали, титановые сплавы как 40Х13
    End Select
    
    Dim iRa
    Select Case Ra
        Case Is <= 0.04:    iRa = 4
        Case Is <= 0.08:    iRa = 3
        Case Is <= 0.16:    iRa = 2
        Case Is <= 0.32:    iRa = 1
        Case Else:          iRa = 0
    End Select
    
    Dim tTok As Double
    
    Dim iL As Integer   'Позиция ближайшей большей длины
    If L > 50 Then
        Dim tSmall As Double
        Dim tLarge As Double
        tSmall = rTime(iOtkl, (1 + id - 1) + iMat + iRa)
        tLarge = rTime(iOtkl, (7 + id - 1) + iMat + iRa)
        tTok = LineynayaInterpolyatisiya(tLarge, tSmall, 50, 5, L)
    Else
        Select Case L
            Case Is <= 5:   iL = 1
            Case Is <= 10:  iL = 2
            Case Is <= 15:  iL = 3
            Case Is <= 20:  iL = 4
            Case Is <= 30:  iL = 5
            Case Is <= 40:  iL = 6
            Case Is <= 50:  iL = 7
        End Select
        iL = iL + (id - 1)
        tTok = rTime(iOtkl, iL + iMat + iRa)
    End If
    
    Dim tKontrol As Double
    If TipPov = 1 Then
        tKontrol = IzmereniyeMikrometrom(d, L)
    Else
        tKontrol = IzmereniyeNutromerom(d, L)
    End If
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    Dovodka = Normy

End Function

Function DovodkaTsentrov() As Integer
    DovodkaTsentrov = 4
End Function

Function ObtachivaniyeGalteley(Material As EnumMaterialy, d As Double, R As Double, Ra As Double)

    If ProverkaOsnovnykhParametrov(CInt(Material), d, R, DMAX_GALTEL) = False Then Exit Function
    
    Dim Lmax As Integer
    Lmax = Lmax_Galtel(Material, d)
    
    If R > Lmax Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_galtel")
    
    Dim i As Long, tTok As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
             Normativ(i, 2) >= d And Normativ(i, 3) >= R And Normativ(i, 4) <= Ra Then
             tTok = Normativ(i, 5)
             Exit For
        End If
    Next
    
    Dim tKontrol As Double
    tKontrol = ProverkaRadiusomerom(R)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    ObtachivaniyeGalteley = Normy
    
End Function

Function Lmax_Galtel(Material As EnumMaterialy, d As Double) As Integer

    Dim Normativ
    Normativ = ZagruzitNormativ("tok_galtel")
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
             Normativ(i, 2) >= d Then
        
            If Normativ(i, 3) > Lmax Then
                Lmax = Normativ(i, 3)
                If PoslednyayaStrokaVMassive(Normativ, i, 2) = True Then Exit For
            End If
            
        End If
    Next
   
    Lmax_Galtel = Lmax
    
End Function

Function Konus(ByVal Material As Integer, d As Double, L As Double, Ra As Double, IT As Integer, TipPoverkhnosti As Integer, Optional TonkSten As Boolean = False) As Double

    'TipPoverkhnosti - Тип поверхности, 1 - Наружная, 2 - Внутренняя
    
    Dim Lmax As Integer
    Lmax = Lmax_Konus(d)
    If d > DMAX_KONUS Or L > Lmax Or (Ra = 3 And TonkSten) Then Exit Function
    
    Material = PreobrazovatMaterial(Material)
    
    Dim rD As Range     'Диапазон диамтров
    Dim rL As Range     'Диапазон длин
    Dim rTime As Range  'Диапазон времени
    
    Dim NumCol As Integer
    Select Case Material
        Case 1: NumCol = 6
        Case 2: NumCol = 6
        Case 3: NumCol = 5
        Case 4: NumCol = 4
        Case 5: NumCol = 3
        Case 6: NumCol = 2
        Case 7: NumCol = 1
    End Select
    
    With ThisWorkbook.Worksheets("Ток_" & CStr(IIf(TonkSten, 15, 14)))
        Set rD = .Range("A8:F22").Offset(0, IIf(TipPoverkhnosti = 1, 6, 0)).Columns(NumCol)
        Set rL = .Range("M8:AH22")
        Select Case Ra
            Case Is > 1.25:     Set rTime = .Range("M25:AH27")
            Case Is > 0.63:     Set rTime = .Range("M28:AH30")
            Case Else:          Set rTime = .Range("M31:AH33")
        End Select
    End With
    
    Dim ArrD                'Массив диаметров
    Dim DLarge As Integer   'Ближайший больший диаметр
    ArrD = Array(5, 10, 15, 20, 30, 40, 50, 70, 100)
    DLarge = BlizhBolshRavn_Znachenie_Array(L, ArrD)

    Dim ArrL()
    ArrL = Array(5, 10, 15, 20, 30, 40, 50, 70, 100)
    L = BlizhBolshRavn_Znachenie_Array(L, ArrL)
    
    Dim id As Integer   'Позиция ближайшего большего диаметра
    id = WorksheetFunction.Match(DLarge, rD, 0)
    
    Dim iL As Integer   'Позиция ближайшей большей длины
    iL = WorksheetFunction.Match(L, rL.Rows(id), 0)
    
    Dim NomerStroki As Integer
    Select Case IT
        Case Is > 9:    NomerStroki = 1
        Case Is > 7:    NomerStroki = 2
        Case Else:      NomerStroki = 3
    End Select
    Konus = rTime(NomerStroki, iL) * IIf(Material = 1, 0.9, 1)
    
End Function

Function Kanavka(Material As EnumMaterialy, TipKanavki As Integer, ByVal d As Double, ByVal GlubinaKanavki As Double, _
    ByVal ShirinaKanavki As Double, ByVal IT As Integer, ByVal Ra As Double)
    'Карта 19. Протачивание канавок

    If ProverkaOsnovnykhParametrov(CInt(Material), d, GlubinaKanavki) = False Or Not ShirinaKanavki > 0 Then Exit Function
    
    Dim kB As Double
    If ShirinaKanavki > MAX_SHIRINA_KANAVKI Then
        kB = WorksheetFunction.RoundUp(ShirinaKanavki / MAX_SHIRINA_KANAVKI, 0)
        ShirinaKanavki = MAX_SHIRINA_KANAVKI
    Else
        kB = 1
    End If
    
    Dim kL As Double
    Dim Lmax As Double
    Lmax = Lmax_Kanavka(Material, TipKanavki, d, ShirinaKanavki, Ra, IT)
    If GlubinaKanavki > Lmax Then
        kL = KoefficientLmaxL(Lmax, GlubinaKanavki)
        GlubinaKanavki = Lmax
    Else
        kL = 1
    End If
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_kanavka")
    
    Dim i As Long
    Dim tTok As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
             Normativ(i, 2) = TipKanavki And Normativ(i, 3) >= d And _
             Normativ(i, 4) >= ShirinaKanavki And Normativ(i, 5) <= Ra And Normativ(i, 6) <= IT And _
             Normativ(i, 7) >= GlubinaKanavki Then
        
            tTok = kB * kL * Normativ(i, 8)
            Exit For
        End If
    Next
    
    Dim tKontrol As Double
    If TipKanavki = 1 Then
        tKontrol = KontrolNaruzhnoyKanavki(d, ShirinaKanavki, IT)
    Else
        tKontrol = KontrolVnutrenneyKanavki(d, ShirinaKanavki)
    End If
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    Kanavka = Normy

End Function

Private Function Lmax_Kanavka(Material As EnumMaterialy, TipKanavki As Integer, d As Double, _
    ByVal ShirinaKanavki As Double, Ra As Double, IT As Integer) As Integer
    
    If ProverkaOsnovnykhParametrov(CInt(Material), d, 1) = False Or ShirinaKanavki <= 0 Then Exit Function
    
    If ShirinaKanavki > MAX_SHIRINA_KANAVKI Then ShirinaKanavki = MAX_SHIRINA_KANAVKI

    Dim Normativ
    Normativ = ZagruzitNormativ("tok_kanavka")

    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
             Normativ(i, 2) = TipKanavki And Normativ(i, 3) >= d And _
             Normativ(i, 4) >= ShirinaKanavki And Normativ(i, 5) <= Ra And Normativ(i, 6) <= IT Then
        
            If Normativ(i, 7) > Lmax Then
                Lmax = Normativ(i, 7)
                If PoslednyayaStrokaVMassive(Normativ, i, 6) = True Then Exit For
            End If
            
        End If
    Next
   
    Lmax_Kanavka = Lmax
    
End Function

Function NakativaniyeRifleniy(Material As EnumMaterialy, d As Double, L As Double) As Double
    
    Dim Lmax As Integer
    Lmax = Lmax_NakativaniyeRifleniy(d)
    
    If L > Lmax Or d > DMAX_NAKATYVANIYE Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_nakatyvaniye")
    
    Dim i As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) >= Normativ(i, 2) And Normativ(i, 3) >= L Then
            NakativaniyeRifleniy = Normativ(i, 4)
            Exit For
        End If
    Next
    
End Function

Function KompleksNaruzhnayaRezba(Material As EnumMaterialy, Dzagotovki As Double, dRezby As Double, Ltocheniya As Double, Lrezby As Double, _
    Optional IT As Integer = 7, Optional VUpor As Boolean = False, Optional Shag As Double = 0)
    
    Dim tTocheniya As Double
    tTocheniya = ProdolnoeTochenie(Material, Dzagotovki, dRezby, Ltocheniya, IIf(IT = 7, 11, 9), , , , , , , VUpor)
    
    Dim tFaski As Double
    tFaski = ObrabotkaFasok(Material, dRezby)
    
    Dim tRezby As Double
    tRezby = NarezanieNaruzhnoyRezby(Material, dRezby, Lrezby, IT, VUpor, Shag)
    
    KompleksNaruzhnayaRezba = tTocheniya + tFaski + tRezby
    
End Function

Function NarezaniyeRezbyNaTokarnomStankeAvtoVyborSposoba(Material As EnumMaterialy, Diametr As Double, Dlina As Double, TipRezby As Integer, _
    Optional VUpor As Boolean = False, Optional Kvalitet As Integer = 7, Optional ShagRezby As Double = 0)
    
    If ProverkaOsnovnykhParametrov(Material, Diametr, Dlina, DMAX_METRREZBAREZTSOM) = False Then Exit Function

    Dim SposobNarezki As EnumSposobNarezkiRezby
    SposobNarezki = OpredelitSposobNarezkiRezbyNaTokarnomStanke(TipRezby, Diametr)
    
    Dim Normy
    If SposobNarezki = Reztsom Then
        
        Normy = NarezanieMetricheskoiRezbiRezcom(TipRezby, Material, Diametr, Dlina, Kvalitet, VUpor, ShagRezby)
    
    ElseIf SposobNarezki = PlashkoyIliMetchikom Then
         
         If TipRezby = 1 Then   'Наружная
            Normy = NarezaniyeRezbyPlashkoy(Material, Diametr, Dlina, ShagRezby, Kvalitet)
            
         ElseIf TipRezby = 2 Then   'Внутренняя
            Normy = NarezaniyeRezbyMetchikom(Material, Diametr, Dlina, Kvalitet, VUpor, ShagRezby)
         
         Else
            Exit Function
            
         End If
    
    Else
        Exit Function
    End If
     
    NarezaniyeRezbyNaTokarnomStankeAvtoVyborSposoba = Normy
    
End Function

Function OpredelitSposobNarezkiRezbyNaTokarnomStanke(TipRezby As Integer, Diametr As Double) As EnumSposobNarezkiRezby

    If Diametr > DMAX_METRREZBAREZTSOM Then Exit Function
    
    Dim SposobNarezki As EnumSposobNarezkiRezby
    
    Dim dMaxPlashkaMetchik As Integer
    dMaxPlashkaMetchik = 8
    
    If TipRezby = 1 Then    'Наружная
        
        If Diametr > DMAX_PLASHKA Then
            SposobNarezki = Reztsom
        Else
            If Diametr > dMaxPlashkaMetchik Then
                SposobNarezki = Reztsom
            Else
                SposobNarezki = PlashkoyIliMetchikom
            End If
        End If
        
    ElseIf TipRezby = 2 Then    'Внутренняя
    
        If Diametr > DMAX_METCHIK Then
            SposobNarezki = Reztsom
        Else
            If Diametr < dMaxPlashkaMetchik Then
                SposobNarezki = PlashkoyIliMetchikom
            Else
                SposobNarezki = Reztsom
            End If
        End If
        
    Else
        Exit Function
    End If
    
    OpredelitSposobNarezkiRezbyNaTokarnomStanke = SposobNarezki
    
End Function
Function NarezanieNaruzhnoyRezby(Material As EnumMaterialy, d As Double, L As Double, _
    Optional IT As Integer = 7, Optional VUpor As Boolean = False, Optional Shag As Double = 0)
    
    NarezanieNaruzhnoyRezby = NarezanieMetricheskoiRezbiRezcom(1, Material, d, L, IT, VUpor, Shag)

End Function

Function NarezanieVnutrenneyRezbyReztsom(Material As Integer, d As Double, L As Double, _
    Optional IT As Integer = 7, Optional VUpor As Boolean = False, Optional Shag As Double = 0)
    
    NarezanieVnutrenneyRezbyReztsom = NarezanieMetricheskoiRezbiRezcom(2, Material, d, L, IT, VUpor, Shag)

End Function


Function NarezanieMetricheskoiRezbiRezcom(TipRezby As Integer, ByVal Material As EnumMaterialy, d As Double, ByVal L As Double, _
    Optional IT As Integer = 7, Optional VUpor As Boolean = False, Optional ByVal Shag As Double = 0)
    
    'TipRezby: 1-Наружная; 2-Внутренняя
    
    Dim Dmax As Double
    Dmax = 120
    
    If ProverkaOsnovnykhParametrov(CInt(Material), d, L, Dmax) = False Or IT < 6 Or TipRezby < 1 Or TipRezby > 2 Then Exit Function
    
    Dim ShagMax As Double
    ShagMax = Smax_NarezanieMetricheskoiRezbiRezcom(d)
    If Shag > ShagMax Then Exit Function
    If Shag = 0 Then Shag = ShagMax
    
    Dim Lmax As Double
    Lmax = Lmax_NarezanieMetricheskoiRezbiRezcom(Material, d, Shag, TipRezby, IT, VUpor)
    
    Dim tTok As Double
    tTok = KoefficientLmaxL(Lmax, L) * NaytyVMassive_NarezanieMetricheskoiRezbiRezcom(Material, d, Shag, TipRezby, IT, VUpor, L, Lmax)
    
    Dim tKontrol As Double
    tKontrol = Kontrol_NaruzhnayaRezba(d, Shag, IT < 7)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    NarezanieMetricheskoiRezbiRezcom = Normy

End Function

Private Function Smax_NarezanieMetricheskoiRezbiRezcom(d As Double) As Double
    Dim ShagMax As Double
    Select Case d
        Case Is <= 10:  ShagMax = 1.5
        Case Is <= 16:  ShagMax = 2
        Case Is <= 24:  ShagMax = 3
        Case Is <= 30:  ShagMax = 3.5
        Case Is <= 36:  ShagMax = 4
        Case Is <= 52:  ShagMax = 5
        Case Is <= 120: ShagMax = 6
    End Select
    
    Smax_NarezanieMetricheskoiRezbiRezcom = ShagMax
End Function

Private Function Lmax_NarezanieMetricheskoiRezbiRezcom(Material As EnumMaterialy, d As Double, Shag As Double, TipRezby As Integer, IT As Integer, VUpor As Boolean) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_rezba_metrich_reztsom")
    
    Dim i As Long
    Dim Lmax As Double
    
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) >= Shag And _
            Normativ(i, 4) = TipRezby And _
            Normativ(i, 5) = IIf(IT > 7, 7, IT) And _
            Normativ(i, 6) = VUpor Then

            If Normativ(i, 7) > Lmax Then
                Lmax = Normativ(i, 7)
                If i = UBound(Normativ) Then
                    Exit For
                ElseIf Normativ(i, 6) <> Normativ(i + 1, 6) Then
                    Exit For
                End If
            End If

        End If
    Next
    
    Lmax_NarezanieMetricheskoiRezbiRezcom = Lmax
    
End Function

Private Function NaytyVMassive_NarezanieMetricheskoiRezbiRezcom(Material As EnumMaterialy, d As Double, Shag As Double, TipRezby As Integer, IT As Integer, VUpor As Boolean, L As Double, Lmax As Double) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_rezba_metrich_reztsom")
    
    Dim i As Integer
    
    For i = 1 To UBound(Normativ)
        
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) >= Shag And _
            Normativ(i, 4) = TipRezby And _
            Normativ(i, 5) = IIf(IT > 7, 7, IT) And _
            Normativ(i, 6) = VUpor And _
            Normativ(i, 7) >= IIf(L > Lmax, Lmax, L) Then
            NaytyVMassive_NarezanieMetricheskoiRezbiRezcom = Normativ(i, 8)
            Exit For
        End If
        
    Next
    
End Function

Function NarezaniyeRezbyMetchikom(Material As EnumMaterialy, ByVal d As Double, ByVal L As Double, _
    Optional IT As Integer = 7, Optional VUpor As Boolean = False, Optional Shag As Double = 0)
    
    If ProverkaOsnovnykhParametrov(Material, d, L, DMAX_METCHIK) = False Or IT < 6 Then Exit Function
        
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_metchik")
    
    If Shag = 0 Then Shag = KrupShagRezb(d)
    
    Dim i As Long
    Dim tTok As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) >= Shag And _
            Normativ(i, 4) = IIf(IT > 7, 7, IT) And _
            Normativ(i, 5) = VUpor Then
            tTok = Normativ(i, 6) * L + Normativ(i, 7)
        End If
    Next

    Dim tKontrol As Double
    tKontrol = Kontrol_RezbovayaProbka(d, Shag, IT)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    NarezaniyeRezbyMetchikom = Normy
    
End Function

Function NarezaniyeVnutrenneyRezby(Material As Integer, d As Double, Glukhoye As Boolean, DlinaNarezki As Double, _
    Optional IT As Integer = 7, Optional ShagRezbi As Double = 0)
    
    If IT = 0 Then IT = 7
    
    Dim tNarez As Double
    If d <= 10 Then
        tNarez = NarezaniyeRezbyMetchikom(CInt(Material), d, DlinaNarezki, IT, Glukhoye, ShagRezbi)
    Else
        tNarez = NarezanieMetricheskoiRezbiRezcom(2, Material, d, DlinaNarezki, IT, Glukhoye, ShagRezbi)
    End If
    NarezaniyeVnutrenneyRezby = tNarez
    
End Function

Function ObrabotkaFasok(ByVal Material As Integer, d As Double) As Double

    'Material - Материал:
    '1 - ЛС59-1
    '2 - АЛ2
    '3 - Сталь 35
    '4 - 25Х13Н2
    '5 - Титан ВТ1-0
    '6 - 12Х18Н9Т
    '7 - ВНМ 3-2
    
    If d > DMAX_TOKARNYY Or Not d > 0 Then Exit Function
    
    Material = PreobrazovatMaterial(Material)

    Dim rD As Range    'Диапазон диамтров
    Dim rTime As Range 'Диапазон времени
    
    With ThisWorkbook.Worksheets("Ток_16")
        Set rD = .Range("A1:G1")
        Set rTime = .Range("A6:L6")
    End With

    Dim NumCol As Integer   'Позиция ближайшего большего диаметра с учетом материала
    NumCol = NomerStolbca_BlizhBolshRavn(d, rD) + NomerStolbcaMateriala_Tokarnyye(Material)

    ObrabotkaFasok = rTime.Columns(NumCol) * IIf(Material = 1, 0.9, 1)

End Function

Function ZachistkaTortsaShkurkoy(ByVal Material As Integer, d As Double) As Double
    
    If d > DMAX_TOKARNYY Or Not d > 0 Then Exit Function
    
    Material = PreobrazovatMaterial(Material)

    Dim rD As Range    'Диапазон диамтров
    Dim rTime As Range 'Диапазон времени
    
    With ThisWorkbook.Worksheets("Ток_16")
        Set rD = .Range("A1:G1")
        Set rTime = .Range("A8:L8")
    End With

    Dim NumCol As Integer   'Позиция ближайшего большего диаметра с учетом материала
    NumCol = NomerStolbca_BlizhBolshRavn(d, rD) + NomerStolbcaMateriala_Tokarnyye(Material)

    ZachistkaTortsaShkurkoy = rTime.Columns(NumCol) * IIf(Material = 1, 0.9, 1)

End Function

Function PritupleniyeKromok(ByVal Material As Integer, d As Double) As Double
    
    If d > DMAX_TOKARNYY Or Not d > 0 Then Exit Function
    
    Material = PreobrazovatMaterial(Material)

    Dim rD As Range    'Диапазон диамтров
    Dim rTime As Range 'Диапазон времени
    
    With ThisWorkbook.Worksheets("Ток_16")
        Set rD = .Range("A1:G1")
        Set rTime = .Range("A7:L7")
    End With

    Dim NumCol As Integer   'Позиция ближайшего большего диаметра с учетом материала
    NumCol = NomerStolbca_BlizhBolshRavn(d, rD) + NomerStolbcaMateriala_Tokarnyye(Material)

    PritupleniyeKromok = rTime.Columns(NumCol) * IIf(Material = 1, 0.9, 1)

End Function

Function Tsentrovaniye(ByVal Material As Integer, d As Double, Optional TipTsentrovki As Integer = 3, Optional Ra As Double = 1.25) As Double

    'Material - Материал:
    '1 - ЛС59-1
    '2 - АЛ2
    '3 - Сталь 35
    '4 - 25Х13Н2
    '5 - Титан ВТ1-0
    '6 - 12Х18Н9Т
    '7 - ВНМ 3-2
    
    If Not d > 0 Or d > DMAX_TOKARNYY Or TipTsentrovki > 3 Or TipTsentrovki < 1 Then Exit Function
    
    Material = PreobrazovatMaterial(Material)
    
    Dim rD As Range         'Диапазон диамтров
    Dim rTime As Range      'Диапазон времени
    
    With ThisWorkbook.Worksheets("Ток_16")
        Set rD = .Range("A1:G1")
        Set rTime = .Range("A2:L5")
    End With
        
    Dim ArrD()              'Массив диаметров
    Dim DLarge As Integer   'Ближайший больший диаметр
    ArrD = Array(10, 20, 40, 70, 170, 300, 400)
    DLarge = BlizhBolshRavn_Znachenie_Array(d, ArrD)
        
    Dim id As Integer   'Позиция ближайшего большего диаметра
    id = WorksheetFunction.Match(DLarge, rD, 0)
        
    Dim iMat As Integer
    iMat = NomerStolbcaMateriala_Tokarnyye(Material)
    
    Dim NomerStroki As Integer
    Select Case TipTsentrovki
        Case 1, 2:  NomerStroki = TipTsentrovki
        Case 3:     NomerStroki = IIf(Ra < 1.25, 3, 4)
    End Select

    Tsentrovaniye = rTime(NomerStroki, id + iMat) * IIf(Material = 1, 0.9, 1)

End Function

Function ObrabotkaOtverstiySNulya(Material As EnumMaterialy, ByVal d As Double, ByVal L As Double, Glukhoye As Boolean, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, _
    Optional TonkostennayaKonfiguraciya As Boolean = False, Optional PoRezhimam As Boolean = True) As Variant
    
    If d <= 0 Or L <= 0 Then Exit Function
    
    Dim dSverleniya As Double, tSverleniye As Double, LmaxSverl As Double
    dSverleniya = IIf(d > 20, 20, d)
    LmaxSverl = Lmax_Sverleniye_Tokarnyy(Material, dSverleniya, Glukhoye)
    tSverleniye = KoefficientLmaxL(LmaxSverl, L) * ObrabotkaOtverstiy_Tokarnyy(Sverleniye, Material, dSverleniya, IIf(L > LmaxSverl, LmaxSverl, L), Glukhoye)(1)
    
    Dim dRassverlivaniye As Double, tRassverlivaniye As Double, LmaxRassverl As Double
    dRassverlivaniye = IIf(d > DMAX_SVERLENIYE_TOKARNYY, DMAX_SVERLENIYE_TOKARNYY, 0)
    If dRassverlivaniye > 0 Then
        LmaxRassverl = Lmax_Rassverlivaniye_Tokarnyy(Material, dRassverlivaniye, Glukhoye)
        tRassverlivaniye = KoefficientLmaxL(LmaxRassverl, L) * ObrabotkaOtverstiy_Tokarnyy(Rassverlivaniye, Material, dRassverlivaniye, IIf(L > LmaxSverl, LmaxSverl, L), Glukhoye)(1)
    End If
    
    Dim Dzagotovki As Double, tRastachivaniye As Double
    Dzagotovki = IIf(tRassverlivaniye > 0, dRassverlivaniye, dSverleniya)
    tRastachivaniye = Rastachivanie(CInt(Material), d, Dzagotovki, L, IT, Ra, GlubinaRezaniya, False, Glukhoye, False, TonkostennayaKonfiguraciya, Glukhoye, PoRezhimam)(1)

    Dim tTok As Double
    tTok = tSverleniye + tRassverlivaniye + tRastachivaniye
    
    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(d, L, IT)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    ObrabotkaOtverstiySNulya = Normy
    
End Function

Function OtrezkaNaTokarnomStanke(ByVal Material As EnumMaterialy, ByVal d As Double, ByVal L As Double) As Double
    
    If ProverkaOsnovnykhParametrov(Material, d, L) = False Or L > 0.5 * d Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_otrezka")
    
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d Then
            OtrezkaNaTokarnomStanke = Normativ(i, 3) * L + Normativ(i, 4)
            Exit For
        End If
    Next

End Function

Function PoperechnoeTocheniye(Material As EnumMaterialy, D1 As Double, d2 As Double, Pripusk As Double, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, _
    Optional Korka As Boolean = False, Optional Udar As Boolean = False, _
    Optional TonkostennayaKonfiguraciya As Boolean = False, Optional PostoyannayaSkorostRezaniya As Boolean = False, Optional TocheniyeDiametra As Boolean = False, _
    Optional PoRezhimam As Boolean = True) As Variant
    
    If Not D1 > 0 Then Exit Function
    
    If IT = 0 Then IT = 12
    If Ra = 0 Then Ra = 10
    If GlubinaRezaniya = 0 Then GlubinaRezaniya = 2
    Dim tTok As Double
    
    If PoRezhimam Then
        tTok = PoperechnoyeTocheniyeKompleks_PoRezhimam(Material, D1, d2, Pripusk, GlubinaRezaniya, Ra, IT, RAZMER_MIN_POP, TocheniyeDiametra, Korka, Udar)
    
    Else
    
        Dim PripuskNaChistovuyuObrabotku As Double
        PripuskNaChistovuyuObrabotku = IIf(D1 > 70, 3, 1.5)
        
        Dim PripuskNaСhernovuyuObrabotku As Double
        PripuskNaСhernovuyuObrabotku = Pripusk - PripuskNaChistovuyuObrabotku
        
        'Расчет черновой обработки
        Dim tChernovoiObrabotki As Double
        Dim tObrabotkiPoKorke As Double
        Dim tObrabotkiSPromerom As Double
        Dim KolVoProhodov As Double
        
        Dim L As Double: L = (D1 - d2) / 2
    
        If PripuskNaСhernovuyuObrabotku > 0 Then
            KolVoProhodov = WorksheetFunction.RoundUp(PripuskNaСhernovuyuObrabotku / GlubinaRezaniya, 0)
            tObrabotkiPoKorke = PoperechnoyeTocheniye_Chernovoye(Material, D1, L, Udar, Korka, IIf(KolVoProhodov = 1, True, False), PostoyannayaSkorostRezaniya)
            KolVoProhodov = KolVoProhodov - 1
            If KolVoProhodov > 0 Then
                tObrabotkiSPromerom = PoperechnoyeTocheniye_Chernovoye(Material, D1, L, Udar, False, True, PostoyannayaSkorostRezaniya)
                KolVoProhodov = KolVoProhodov - 1
                If KolVoProhodov > 0 Then
                    tChernovoiObrabotki = PoperechnoyeTocheniye_Chernovoye(Material, D1, L, Udar, False, False, PostoyannayaSkorostRezaniya) * KolVoProhodov
                End If
            End If
            tChernovoiObrabotki = tChernovoiObrabotki + tObrabotkiSPromerom + tObrabotkiPoKorke
        End If
        
        'Расчет чистовой обработки
        Dim tChistovoiObrabotki As Double
        tChistovoiObrabotki = PoperechnoyeTochenieye_Chistovoe(Material, D1, L, Ra, IT, Udar, TonkostennayaKonfiguraciya)
        
        tTok = tChernovoiObrabotki + tChistovoiObrabotki
        
    End If
        
    'Расчет времени на точение диаметра
    Dim tDiametr As Double
    If TocheniyeDiametra Then
        If PoRezhimam Then
            tDiametr = ProdolnoyeTocheniye_PoRezhimam(CInt(Material), D1, D1, 1, Pripusk, Ra, IT, False, Udar, False)
        Else
            tDiametr = ProdolnoeTochenie_Chistovoe(CInt(Material), d2, Pripusk, IT, Ra, TonkostennayaKonfiguraciya, Udar)
        End If
    Else
        tDiametr = 0
    End If

    tTok = tTok + tDiametr
      
    
    
    Dim tKontrol As Double
    tKontrol = IzmerShtangenCircul(, D1, IT < 11)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    PoperechnoeTocheniye = Normy

End Function

Function PoperechnoyeTochenieye_Chistovoe(ByVal Material As Integer, ByVal d As Double, ByVal L As Double, Ra As Double, _
    Optional IT As Integer = 0, Optional NaUdar As Boolean = False, _
    Optional TonkostennayaKonfiguraciya As Boolean = False, Optional PostoyannayaSkorostRezaniya As Boolean = False) As Double
    
    If ProverkaOsnovnykhParametrov(CInt(Material), d, L) = False Or IT < MIN_IT Or Ra < MIN_RA Or L > 0.5 * d Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_poperchnoye_chistovoye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) <= IT And _
            Normativ(i, 4) <= Ra Then
            a = Normativ(i, 5)
            b = Normativ(i, 6)
            Exit For
        End If
    Next
    
    Dim kUdar As Double, kTonksosten As Double
    kUdar = IIf(NaUdar, 1.2, 1)
    kTonksosten = IIf(TonkostennayaKonfiguraciya, 1.3, 1)
    
    PoperechnoyeTochenieye_Chistovoe = kUdar * kTonksosten * (a * L + b)
    
End Function

Private Function PoperechnoyeTocheniye_Chernovoye(ByVal Material As EnumMaterialy, ByVal d As Double, ByVal L As Double, _
    Optional TochenieNaUdar As Boolean = False, Optional TocheniePoKorke As Boolean = False, Optional ObrabotkaSPromerom As Boolean = False, _
    Optional PostoyannayaSkorostRezaniya As Boolean = False) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_poperechnoye_chernovoye")
        
    If ProverkaOsnovnykhParametrov(Material, d, L) = False Or L > d / 2 Then Exit Function
    If PostoyannayaSkorostRezaniya Then L = L / 2
    
    Dim i As Long
    Dim PoprKoef As Double
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) = ObrabotkaSPromerom Then
            a = Normativ(i, 4)
            b = Normativ(i, 5)
            Exit For
        End If
    Next
    
    PoprKoef = PoprKoefPriChernTocheniiRastachivanii(L, L, d, TochenieNaUdar, TocheniePoKorke, False)
    
    PoperechnoyeTocheniye_Chernovoye = PoprKoef * (a * L + b)
    
End Function

Function ProdolnoeTochenie_IzPrutka(Material As EnumMaterialy, DiametrZagotovki As Double, DiametrDetali As Double, Lzagotovki As Double, KolVoDetVPrutke As Double, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, _
    Optional TocheniePoKorke As Boolean = False, Optional Udar As Boolean = False, _
    Optional TonkostennayaKonfiguraciya As Boolean = False, Optional PodrezkaTorca As Boolean = False) As Double
    
    Dim Dlina As Double
    Dim KolVoDet_Fact As Double
    KolVoDet_Fact = KolVoDetIzPrutka(Lzagotovki, KolVoDetVPrutke, DiametrDetali)
    Dlina = DlinaIzPutka(Lzagotovki, KolVoDetVPrutke, KolVoDet_Fact)
    
    ProdolnoeTochenie_IzPrutka = ProdolnoeTochenie(Material, DiametrZagotovki, DiametrDetali, Dlina, IT, Ra, GlubinaRezaniya, TocheniePoKorke, Udar, _
        TonkostennayaKonfiguraciya, PodrezkaTorca)(1) / KolVoDet_Fact
    
End Function

Function Rastachivanie(Material As EnumMaterialy, DiametrDetali As Double, DiametrZagotovki As Double, L As Double, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, Optional ByVal Korka As Boolean = False, _
    Optional Glukhoye As Boolean = False, Optional Udar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False, _
    Optional PodrezkaTorca As Boolean = False, _
    Optional PoRezhimam As Boolean = True)

    If L <= 0 Or DiametrZagotovki > DiametrDetali Or DiametrDetali <= 0 Or DiametrZagotovki <= 0 Then Exit Function
    
    If IT = 0 Then IT = 12
    If Ra = 0 Then Ra = 10
    If GlubinaRezaniya = 0 Then GlubinaRezaniya = 2
            
    Dim tTok As Double
    
    If PoRezhimam Then
        tTok = RastachivaniyeKompleks_PoRezhimam(CInt(Material), DiametrDetali, DiametrZagotovki, GlubinaRezaniya, L, Ra, IT, Korka, Udar, PodrezkaTorca)
        
    Else
    
        Dim MaksGlubinaRezaniya As Integer
        MaksGlubinaRezaniya = 5
    
        If ProverkaOsnovnykhParametrov(CInt(Material), DiametrZagotovki, L) = False Or GlubinaRezaniya > MaksGlubinaRezaniya Then Exit Function

        tTok = RastachivaniyeKompleks_PoNormativu(Material, DiametrDetali, DiametrZagotovki, L, IT, Ra, GlubinaRezaniya, Korka, _
            Glukhoye, Udar, TonkostennayaKonfiguraciya, PodrezkaTorca)
    
    End If

    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(DiametrDetali, L, IT)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    Rastachivanie = Normy
    
End Function

Private Function RastachivaniyeKompleks_PoNormativu(Material As EnumMaterialy, DiametrDetali As Double, DiametrZagotovki As Double, L As Double, _
    Optional IT As Integer = 12, Optional Ra As Double = 10, Optional GlubinaRezaniya As Double = 2, Optional ByVal Korka As Boolean = False, _
    Optional Glukhoye As Boolean = False, Optional Udar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False, _
    Optional PodrezkaTorca As Boolean = False)

    Dim PripuskNaChistovuyuObrabotku As Double
    If IT > 11 Then
        If DiametrDetali > 70 Then
            PripuskNaChistovuyuObrabotku = 3
        Else
            PripuskNaChistovuyuObrabotku = 1.5
        End If
    Else
        If DiametrDetali > 70 Then
            PripuskNaChistovuyuObrabotku = 0.5
        Else
            PripuskNaChistovuyuObrabotku = 0.4
        End If
    End If

    'Расчет черновой обработки
    Dim tChernovoiObrabotki As Double
    Dim DiametrPredvaritelnyi As Double
    DiametrPredvaritelnyi = DiametrDetali - PripuskNaChistovuyuObrabotku
    tChernovoiObrabotki = Rastachivanie_Chernovoe(Material, DiametrZagotovki, DiametrPredvaritelnyi, L, GlubinaRezaniya, Glukhoye, Udar, Korka)
       
    'Расчет чистовой обработки
    Dim tChistovoiObrabotki As Double
    tChistovoiObrabotki = Rastachivanie_Chistovoe(Material, DiametrDetali, L, Ra, IT, Glukhoye, Udar, TonkostennayaKonfiguraciya)
    
    'Расчет времени на подрезку внутреннего торца
    Dim tPodrezki As Double
    If PodrezkaTorca Then
        tPodrezki = PoperechnoyeTochenieye_Chistovoe(Material, DiametrDetali, (DiametrDetali - DiametrZagotovki), Ra, IT, Udar, TonkostennayaKonfiguraciya, False)
    Else
        tPodrezki = 0
    End If
    
    RastachivaniyeKompleks_PoNormativu = tChernovoiObrabotki + tChistovoiObrabotki + tPodrezki
        
End Function

Function NarezaniyeRezbyPlashkoy(Material As EnumMaterialy, d As Double, L As Double, Optional Shag As Double, Optional IT As Integer = 7)
    
    If ProverkaOsnovnykhParametrov(Material, d, L, DMAX_PLASHKA) = False Or IT < 6 Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_plashka")
    
    If Shag = 0 Then Shag = KrupShagRezb(d)
    
    Dim i As Long
    Dim tTok As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) >= Shag And _
            Normativ(i, 4) = IIf(IT > 7, 7, IT) Then
            tTok = Normativ(i, 5) * L + Normativ(i, 6)
            Exit For
        End If
    Next

    Dim tKontrol As Double
    tKontrol = Kontrol_NaruzhnayaRezba(d, Shag, IT < 7)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tTok
    Normy(2) = 0
    Normy(3) = tKontrol
    
    NarezaniyeRezbyPlashkoy = Normy

End Function

Function Rastachivanie_Chernovoe(Material As EnumMaterialy, DiametrZagotovki As Double, DiametrDetali As Double, L As Double, GlubinaRezaniya As Double, _
    Optional Glukhoye As Boolean = False, Optional Udar As Boolean = False, Optional Korka As Boolean = False) As Double
    
    Dim MaksGlubinaRezaniya As Integer
    MaksGlubinaRezaniya = 5
    
    If ProverkaOsnovnykhParametrov(CInt(Material), DiametrZagotovki, L, IIf(Material = TITANOVYYE_SPLAVY, 300, DMAX_TOKARNYY)) = False Or GlubinaRezaniya > MaksGlubinaRezaniya Or DiametrZagotovki > DiametrDetali Then Exit Function
    
    Dim tNaProhod As Double, tChernovoiObrabotki As Double
    Dim ObrabotkaSPromerom As Boolean
    
    Dim i As Integer
    Dim DiametrObrabotki As Double
    For DiametrObrabotki = DiametrZagotovki To DiametrDetali Step GlubinaRezaniya * 2
        
        If DiametrObrabotki >= (DiametrDetali - 2 * GlubinaRezaniya) Then ObrabotkaSPromerom = True
        
        tNaProhod = NaytyVMassive_ChernovoyeRastachivaniye(Material, DiametrObrabotki, L, GlubinaRezaniya, ObrabotkaSPromerom, Udar, Korka, Glukhoye)
        tChernovoiObrabotki = tChernovoiObrabotki + tNaProhod
        
        Korka = False
        
    Next

    Rastachivanie_Chernovoe = tChernovoiObrabotki
    
End Function

Private Function NaytyVMassive_ChernovoyeRastachivaniye(Material As EnumMaterialy, DiametrObrabotki As Double, L As Double, GlubinaRezaniya As Double, _
    ObrabotkaSPromerom As Boolean, Udar As Boolean, Korka As Boolean, Glukhoye As Boolean) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_rastachivaniye_chernovoye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= DiametrObrabotki And _
            Normativ(i, 3) >= GlubinaRezaniya And _
            Normativ(i, 4) = ObrabotkaSPromerom And _
            Normativ(i, 5) = Glukhoye Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next
    
    NaytyVMassive_ChernovoyeRastachivaniye = PoprKoefPriChernTocheniiRastachivanii(L, L, DiametrObrabotki, Udar, Korka, False) * (a * L + b)
        
End Function

Function Rastachivanie_Chistovoe(ByVal Material As Integer, ByVal d As Double, ByVal L As Double, Ra As Double, _
    Optional IT As Integer = 0, Optional Glukhoye As Boolean = False, Optional RastachivanieNaUdar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Double
    'Чистовое растачивание отверстий. Карты 10-12
    
    'Материалы
    '1 - ЛС59-1
    '2 - АЛ2
    '3 - Сталь 35
    '4 - 25Х13Н2
    '5 - Титан ВТ1-0
    '6 - 12Х18Н9Т
    '7 - ВНМ 3-2

    If d > DMAX_TOKARNYY Then Exit Function
    
    Material = PreobrazovatMaterial(Material)
    
    Dim rIT As Range    'Диапазон квалитетов
    Dim rD As Range     'Диапазон диамтров
    Dim rL As Range     'Диапазон длин
    Dim rTime As Range  'Диапазон времени
    With ThisWorkbook.Worksheets("Ток_" & IIf(TonkostennayaKonfiguraciya, 11, 10) + IIf(Glukhoye, 2, 0))
        Set rD = .Range("C1:C14")
        Set rIT = .Range("C15:C27")
        Set rL = .Range("D1:AF14")
        Set rTime = .Range("D15:AJ27")
    End With

    Dim Lmax As Double
    Lmax = Lmax_ProdolnoyeTocheniye(d)

    Dim NumCol As Integer
    NumCol = NomerStolbcaPoDiametruIDline(Material, d, rD, L, rL, Lmax) 'Номер столбца ближайшей большей длины в зависимости от диаметра и материала
        
    'Номер строки шероховатости и точности
    'Если пользователь не указал квалитет, то принимается "предварительная" обработка
    If IT = 0 Then IT = 12
    
    Dim NumRow As Integer   'Номер строки в диапазоне времени в зависимости от шероховатости и квалитета
    NumRow = NomerStrokiPriTochenii(Ra, IT)
    
    Dim kL As Double                'Общий поправочный коэффициент на длину точения
    kL = KoefficientLmaxL(Lmax, L)  'Поправочный коэффициент на длину точения если фактическая длина больше максимальной по карте
    kL = kL * PoprKoefLD(d, L)      'Поправочный коэффициент на длину точения в зависимости от отношения длины к диаметру
    
    Dim K1 As Double, K2 As Double
    K1 = IIf(Material = 1, 0.9, 1)
    K2 = IIf(RastachivanieNaUdar, 1.2, 1)
    
    Rastachivanie_Chistovoe = rTime(NumRow, NumCol) * K1 * K2 * kL
    
End Function

Function Lmax_Sverleniye_Tokarnyy(Material As EnumMaterialy, Diametr As Double, Glukhoye As Boolean) As Double
    Lmax_Sverleniye_Tokarnyy = Lmax_ObrabotkaOtverstiy_Tokarnyy(Sverleniye, Material, Diametr, Glukhoye)
End Function

Function Lmax_Rassverlivaniye_Tokarnyy(Material As EnumMaterialy, Diametr As Double, Glukhoye As Boolean) As Double
    Lmax_Rassverlivaniye_Tokarnyy = Lmax_ObrabotkaOtverstiy_Tokarnyy(Rassverlivaniye, Material, Diametr, Glukhoye)
End Function

Function ObrabotkaOtverstiy_Tokarnyy(Perekhod As TipObrabotkiOtverstiya_Tokar, Material As EnumMaterialy, Diametr As Double, ByVal Dlina As Double, Optional Glukhoye As Boolean = False) As Variant
    
    Dim Normativ As Variant
    Normativ = ZagruzitNormativ("tok_obrabotka_otverstiy")
    
    Dim Lmax As Double, kL As Double
    Lmax = Lmax_ObrabotkaOtverstiy_Tokarnyy(Perekhod, Material, Diametr, Glukhoye)
    
    kL = 1
    If Dlina > Lmax Then
        kL = Dlina / Lmax
        Dlina = Lmax
    End If
    
    Dim i As Long
    Dim tObrOtv As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Perekhod And _
            Normativ(i, 2) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 3) >= Diametr And _
            Normativ(i, 4) = Glukhoye And _
            Normativ(i, 5) >= Dlina Then

            tObrOtv = Normativ(i, 6)
            Exit For
        End If

    Next
    
    Dim tKontrol As Double, IT As Integer
    If Perekhod = Razvertyvaniye_Ra1_25 Then
        IT = 7
    ElseIf Perekhod = Razvertyvaniye_Ra2_5 Then
        IT = 9
    Else
        IT = 11
    End If
    tKontrol = KontrolOtverstiy(Diametr, Dlina, IT, Glukhoye)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = kL * tObrOtv
    Normy(2) = 0
    Normy(3) = tKontrol
    
    ObrabotkaOtverstiy_Tokarnyy = Normy
    
End Function

Function Lmax_ObrabotkaOtverstiy_Tokarnyy(Perekhod As TipObrabotkiOtverstiya_Tokar, Material As EnumMaterialy, Diametr As Double, Glukhoye As Boolean) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tok_obrabotka_otverstiy")
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Perekhod And _
            Normativ(i, 2) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 3) >= Diametr And _
            Normativ(i, 4) = Glukhoye Then
            
            If Normativ(i, 5) > Lmax Then
                Lmax = Normativ(i, 5)
                If i = UBound(Normativ) Then
                    Exit For
                ElseIf Normativ(i, 4) <> Normativ(i + 1, 4) Then
                     Exit For
                End If
            End If

        End If

    Next
   
    Lmax_ObrabotkaOtverstiy_Tokarnyy = Lmax
    
End Function


Function Lmax_ProdolnoyeTocheniye(d As Double) As Integer
    
    Select Case d
        Case Is <= 5:   Lmax_ProdolnoyeTocheniye = 50
        Case Is <= 10:  Lmax_ProdolnoyeTocheniye = 100
        Case Is <= 15:  Lmax_ProdolnoyeTocheniye = 310
        Case Is <= 70:  Lmax_ProdolnoyeTocheniye = 550
        Case Else:      Lmax_ProdolnoyeTocheniye = 750
    End Select
    
End Function



Function Lmax_Konus(d As Double) As Integer
    Lmax_Konus = IIf(d > 5, 100, 50)
End Function

Function Lmax_Metchik(d As Double) As Integer
    
    Select Case d
        Case Is <= 3:   Lmax_Metchik = 10
        Case Is <= 5:   Lmax_Metchik = 50
        Case Is <= 16:  Lmax_Metchik = 80
        Case Is <= 20:  Lmax_Metchik = 60
    End Select
    
End Function

Function Lmax_NakativaniyeRifleniy(d As Double) As Integer
    Dim Lmax As Integer
    Select Case d
        Case Is <= 5:   Lmax = 15
        Case Is <= 10:  Lmax = 30
        Case Is <= 20:  Lmax = 70
        Case Else:      Lmax = 100
    End Select
    Lmax_NakativaniyeRifleniy = Lmax
End Function

Function Koefficient_Latun_LmaxL(Material As Integer, Lmax As Double, L As Double)
    Koefficient_Latun_LmaxL = IIf(Material = 1, 0.9, 1) * KoefficientLmaxL(Lmax, L)
End Function

Private Function NomerStolbcaMateriala_Tokarnyye(MaterialID As Integer) As Integer
    
    Select Case MaterialID
        Case 1: NomerStolbcaMateriala_Tokarnyye = 0 'ЛС59-1
        Case 2: NomerStolbcaMateriala_Tokarnyye = 0 'АЛ2
        Case 3: NomerStolbcaMateriala_Tokarnyye = 1 'Сталь 35
        Case 4: NomerStolbcaMateriala_Tokarnyye = 2 '25Х13Н2
        Case 5: NomerStolbcaMateriala_Tokarnyye = 3 'Титан ВТ1-0
        Case 6: NomerStolbcaMateriala_Tokarnyye = 4 '12Х18Н9Т
        Case 7: NomerStolbcaMateriala_Tokarnyye = 5 'ВНМ 3-2
    End Select
    
End Function

Private Function NomerStolbcaPoDiametruIDline(Material As Integer, d As Double, rD As Range, L As Double, rL As Range, Lmax As Double) As Integer
    
    Dim NumRow As Integer
    NumRow = NomerStroki_BlizhBolshRavn(d, rD)
    
    Dim NumCol As Integer
    NumCol = NomerStolbca_BlizhBolshRavn(L, rL.Rows(NumRow), Lmax)
    
    NomerStolbcaPoDiametruIDline = NumCol + NomerStolbcaMateriala_Tokarnyye(Material)

End Function

Private Function NomerStrokiPriTochenii(Ra As Double, IT As Integer)

    Select Case Ra
        Case Is > 5
            NomerStrokiPriTochenii = IIf(IT > 11, 1, 2)
        Case Is > 1.25
            Select Case IT
                Case Is > 9:    NomerStrokiPriTochenii = 3
                Case Is > 7:    NomerStrokiPriTochenii = 4
                Case Else:      NomerStrokiPriTochenii = 5
            End Select
        Case Is > 0.63
            Select Case IT
                Case Is > 9:    NomerStrokiPriTochenii = 6
                Case Is > 7:    NomerStrokiPriTochenii = 7
                Case Is > 5:    NomerStrokiPriTochenii = 8
                Case Else:      NomerStrokiPriTochenii = 9
            End Select
        Case Else
            Select Case IT
                Case Is > 9:    NomerStrokiPriTochenii = 10
                Case Is > 7:    NomerStrokiPriTochenii = 11
                Case Is > 5:    NomerStrokiPriTochenii = 12
                Case Else:      NomerStrokiPriTochenii = 13
            End Select
    End Select

End Function

Function PoprKoefLD(d As Double, L As Double)
    If d > 0 Then
        Select Case L / d
            Case Is < 12: PoprKoefLD = 1
            Case Is < 15: PoprKoefLD = 1 / 0.7
            Case Is < 20: PoprKoefLD = 1 / 0.6
            Case Is < 30: PoprKoefLD = 1 / 0.5
            Case Is < 40: PoprKoefLD = 1 / 0.4
            Case Else: PoprKoefLD = 1 / 0.3
        End Select
    End If
End Function

Private Function PoprKoefPriChernTocheniiRastachivanii(Lmax As Double, L As Double, d As Double, TochenieNaUdar As Boolean, _
    TocheniePoKorke As Boolean, KonicheskayaPoverhnost As Boolean)
    
    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double, K5 As Double
    
    K1 = IIf(TochenieNaUdar, 1.2, 1)
    K2 = IIf(TocheniePoKorke, 1.3, 1)
    K3 = IIf(KonicheskayaPoverhnost, 1.3, 1)
    K4 = KoefficientLmaxL(Lmax, L)  'Поправочный коэффициент на длину точения если фактическая длина больше максимальной по карте
    K5 = PoprKoefLD(d, L)           'Поправочный коэффициент на длину точения в зависимости от отношения длины к диаметру
    
    PoprKoefPriChernTocheniiRastachivanii = K1 * K2 * K3 * K4 * K5
    
End Function

