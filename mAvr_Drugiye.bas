Attribute VB_Name = "mAvr_Drugiye"
Option Explicit

Function DlinaRezaniyaPriFrezOtverstiy(Dotv As Double, Dpred As Double, Dfrezy As Double, GlubinaRezaniya As Double) As Double
      
    If Dotv <= Dfrezy Or GlubinaRezaniya <= 0 Then
        DlinaRezaniyaPriFrezOtverstiy = 0
        Exit Function
    End If
    
    Dim Dnachalnyy As Double
    If Dpred > Dfrezy Then
        Dnachalnyy = Dpred
    Else
        Dnachalnyy = Dfrezy
    End If

    Dim Shag As Double
    Shag = 2 * GlubinaRezaniya
    
    Dim KolVoProkhodov As Long
    KolVoProkhodov = Int((Dotv - Dnachalnyy) / Shag)
    
    ' Расчет длины полных проходов по формуле арифметической прогрессии
    Dim DinaRezaniya As Double
    If KolVoProkhodov > 0 Then
        DinaRezaniya = PI * KolVoProkhodov * (Dnachalnyy - Dfrezy + GlubinaRezaniya * (KolVoProkhodov + 1))
    Else
        DinaRezaniya = 0
    End If
    
    ' Расчет диаметра после выполнения полных проходов
    Dim DotvPosleFrez As Double
    DotvPosleFrez = Dnachalnyy + KolVoProkhodov * Shag
    
    ' Если остался недочищенный припуск, добавляем финальный проход в размер Dotv
    If Dotv > DotvPosleFrez Then
        DinaRezaniya = DinaRezaniya + PI * (Dotv - Dfrezy)
    End If
    
    DlinaRezaniyaPriFrezOtverstiy = Round(DinaRezaniya, 2)
    
End Function

Function SborkaVPaket(Material As EnumMaterialy, DlinaPaketa As Double, ShirinaPaketa As Double, TolshchinaOdnoyDetali As Double, KolVoDetaley As Double) As Double
        
    Const dSverl As Double = 2.5
    Const KolVoOtv As Integer = 4
    Const KolVoTekhnProkladok As Integer = 4
    Const TolshchinaTekhnProkladki As Integer = 2
    Const tSklepat As Double = 1.5
    
    Dim TolshchinaPaketa As Double
    TolshchinaPaketa = KolVoDetaley * TolshchinaOdnoyDetali + KolVoTekhnProkladok * TolshchinaTekhnProkladki
    
    Dim Massa As Double
    Massa = MassaLista(DlinaPaketa, ShirinaPaketa, TolshchinaPaketa, Material)

    Dim tSborka As Double
    tSborka = 0.1 * KolVoDetaley

    Dim tUst As Double
    tUst = Ustanov_NaStoleKreplBoltPlank(Massa, 2, 0.5)
    
    Dim NormySverl
    NormySverl = ObrabotkaOtverstiySleasar(1, Material, dSverl, KolVoOtv, TolshchinaPaketa, 0, False, False, 1)
    
    Dim tSverl As Double, tZausOtv As Double
    tSverl = NormySverl(1)
    tZausOtv = NormySverl(2)

    SborkaVPaket = tSborka + tSklepat + tUst + tSverl + tZausOtv

End Function

Function SborkaVPaket2(Material As EnumMaterialy, Dlina As Double, Shirina As Double, Tolshchina As Double, KolVoDetaley As Double)
        
    Const dSverl As Double = 2.5
    Const KolVoOtv As Integer = 4
    Const KolVoTekhnProkladok As Integer = 4
    Const TolshchinaTekhnProkladki As Integer = 2
    Const tSklepat As Double = 1.5
    
    Dim tZausProkladok As Double, tPravkaProkladok As Double, tZausDetaley As Double
    tZausProkladok = KolVoTekhnProkladok * SnyatiyeZausentsevList(Dlina, Shirina, TolshchinaTekhnProkladki, Material)
    tPravkaProkladok = KolVoTekhnProkladok * PravkaPloskikhDetaley(TolshchinaTekhnProkladki, Dlina, Shirina, Material)
    tZausDetaley = KolVoDetaley * SnyatiyeZausentsevList(Dlina, Shirina, Tolshchina, Material)
    
    Dim TolshchinaPaketa As Double
    TolshchinaPaketa = KolVoDetaley * Tolshchina + KolVoTekhnProkladok * TolshchinaTekhnProkladki
    
    Dim Massa As Double
    Massa = MassaLista(Dlina, Shirina, TolshchinaPaketa, Material)

    Dim tSborka As Double
    tSborka = 0.1 * KolVoDetaley
    
    Dim NormySverl
    NormySverl = ObrabotkaOtverstiySleasar(1, Material, dSverl, KolVoOtv, TolshchinaPaketa, 0, False, False, 1)
    
    Dim tSverl As Double, tZausOtv As Double, tKontrolOtv As Double
    tSverl = NormySverl(1)
    tZausOtv = NormySverl(2)
    tKontrolOtv = NormySverl(3)

    Dim tUst As Double
    tUst = Ustanov_NaStoleKreplBoltPlank(Massa, 2, 0.5)
    
    Dim tPravka As Double
    
    SborkaVPaket2 = tSborka + tZausOtv + tPravka + tSverl + tSklepat

End Function



Function OkrashivaniyePlanki_s_Nadpisyu(Dlina As Double, Shirina As Double, Tolshchina As Double, OdnaStorona As Boolean) As Double
    
    Dim tGrunt As Double, tOkras As Double, tIzol As Double
    Dim Ploshchad_dm2 As Double
    
    Ploshchad_dm2 = 2 * (Dlina * Shirina + Shirina * Tolshchina + Dlina * Tolshchina) * 10 ^ -4
    
    If OdnaStorona Then
        Ploshchad_dm2 = 0.5 * Ploshchad_dm2
        tIzol = fIzolPloskPov(5, Ploshchad_dm2, 1)
    End If
    
    tGrunt = Gruntovaniye(Ploshchad_dm2)
    tOkras = Okrashivaniye(Ploshchad_dm2)
    
    OkrashivaniyePlanki_s_Nadpisyu = OKRUGLVVERKH(1.3 * (tIzol + tGrunt + tOkras))
    
End Function



Function DlinaGraniShestigrannika(S As Double) As Double
    DlinaGraniShestigrannika = (2 * (S / 2)) / Sqr(3)
End Function

Function ObrabotkaOtverstiySleasar(TipObrabotki As TipObrabotkiOtverstiya, Material As EnumMaterialy, d As Double, _
    ByVal KolVo As Double, Lotv As Double, Lnarez As Double, Glukhoye As Boolean, Optional IT As Integer = 7, _
    Optional Sposob As Integer = 1)
    
    If IT = 0 Then IT = 7
    If KolVo = 0 Then KolVo = 1
    
    Dim tOtv As Double
    If TipObrabotki = TipObrabotkiOtverstiya.Sverleniye Then
        tOtv = SverleniyeSlesar(CInt(Material), d, Lotv, CInt(Sposob), Glukhoye)
    ElseIf TipObrabotki = NarezaniyeRezby Then
        tOtv = NarezaniyeRezbyMetchikomSlesar(CInt(Material), d, Lnarez, IT, Glukhoye)
    ElseIf TipObrabotki = SverleniyeRezba Then
        tOtv = SverleniyePlusRezba(CInt(Material), d, Lotv, Lnarez, IT, Glukhoye, CInt(Sposob))
    ElseIf TipObrabotki = Razvertyvaniye Then
        tOtv = RazvertyvaniyeSlesar(CInt(Material), d, Lotv, IT, Glukhoye)
    ElseIf TipObrabotki = SverleniyeRazvertyvaniye Then
        tOtv = SverleniyePlusRazvertyvaniye(CInt(Material), d, Lotv, Glukhoye, CInt(Sposob), IT)
    ElseIf TipObrabotki = KalibrovaniyeRezby Then
        tOtv = KalibrovaniyeVnutrenneyRezby(d, Lnarez, Glukhoye)
    ElseIf TipObrabotki = Puklevka Then
        tOtv = PuklevkaOtverstiiVruchnuyu(CInt(Material), Lotv, KolVo)
    ElseIf TipObrabotki = ZenkovaniyeFasok Then
        tOtv = ZenkovaniyeFasokPnevmaticheskoySverlilnoyMashinoy(CInt(Material), d, KolVo)
    End If
    
    If TipObrabotki = ZenkovaniyeFasok Or TipObrabotki = Puklevka Then
        tOtv = tOtv
    Else
        tOtv = tOtv * KolVo
    End If
    
    Dim tKontrol As Double
    If TipObrabotki = NarezaniyeRezby Or TipObrabotki = SverleniyeRezba Or TipObrabotki = KalibrovaniyeRezby Then
        tKontrol = Kontrol_RezbovayaProbka(d, , IT)
    ElseIf TipObrabotki = ZenkovaniyeFasok Then
        tKontrol = 0
    Else
        tKontrol = KontrolOtverstiy(d, Lotv, IT, Glukhoye)
    End If
    
    Dim Normy(1 To 2)
    Normy(1) = tOtv
    Normy(2) = tKontrol * KolVo
    
    ObrabotkaOtverstiySleasar = Normy
    
End Function




Function GravirovkaLiniy(DlinaLinii As Double, KolichestvoLiniy As Double) As Double
    GravirovkaLiniy = 0.0937 * DlinaLinii ^ 0.297 * 0.2 ^ 0.384
    GravirovkaLiniy = GravirovkaLiniy * KolichestvoLiniy
End Function

Function GravirovkaLiniy_Range(RangeDlinaLinii As Range, RangeKolichestvoLiniy As Range) As Double
    Dim i As Long
    Dim DlinaLinii As Double, KolichestvoLiniy As Double, t As Double
    For i = 1 To RangeDlinaLinii.Rows.Count
        If IsNumeric(RangeDlinaLinii(i, 1)) Then
            DlinaLinii = RangeDlinaLinii(i, 1)
            KolichestvoLiniy = RangeKolichestvoLiniy(i, 1)
            If DlinaLinii > 0 Then
                If KolichestvoLiniy = 0 Then KolichestvoLiniy = 1
                t = GravirovkaLiniy(DlinaLinii, KolichestvoLiniy)
                GravirovkaLiniy_Range = GravirovkaLiniy_Range + t
            End If
        End If
    Next
End Function

Function GravirovkaSimvolov(VisotaShrifta As Double, KolichestvoSimvolov As Double) As Double
    GravirovkaSimvolov = 0.107 * VisotaShrifta ^ 0.314
    GravirovkaSimvolov = GravirovkaSimvolov * KolichestvoSimvolov
End Function

Function GravirovkaSimvolov_Range(rVisotaShrifta As Range, rKolichestvoSimvolov As Range) As Double
    Dim i As Long
    For i = 1 To rVisotaShrifta.Rows.Count
        If rVisotaShrifta(i) > 0 And rKolichestvoSimvolov(i) > 0 Then
            GravirovkaSimvolov_Range = GravirovkaSimvolov_Range + GravirovkaSimvolov(rVisotaShrifta(i), rKolichestvoSimvolov(i))
        End If
    Next
End Function

Function KontaktnayaSvarka_TTP_60_Slesarnaya(KolVoTochek As Double, Dlina As Range, Shirina As Range, KolVoPozitsiy As Range, IspolzovaniyeKleya As Boolean) As Double
    
    Dim i As Integer
    
    Dim tObezzhirivaniye As Double
    For i = 1 To Dlina.Rows.Count
        If Dlina.Rows(i) > 0 Then tObezzhirivaniye = tObezzhirivaniye + (2 * KolVoPozitsiy.Rows(i) * Galvanika("Обезжиривание", (Shirina.Rows(i) * Dlina.Rows(i)) / 1000000))
    Next
    
    Dim tSkleivaniye As Double
    If IspolzovaniyeKleya Then
        For i = 1 To Dlina.Rows.Count
            If Dlina.Rows(i) > 0 Then tSkleivaniye = tSkleivaniye + (UstanovkaDetaleyNaKley(Ploskaya, Dlina.Rows(i), Shirina.Rows(i), 1, 1, 1, False) * KolVoPozitsiy.Rows(i))
        Next
    End If
    
    Dim tRazmetki As Double
    tRazmetki = 0.3 * KolVoTochek
    
    KontaktnayaSvarka_TTP_60_Slesarnaya = tObezzhirivaniye + tSkleivaniye + tRazmetki
    
    
End Function

Function KontaktnayaSvarka_TTP_65_Svarka(KolVoTochek As Double) As Double
    KontaktnayaSvarka_TTP_65_Svarka = WorksheetFunction.RoundUp((0.5 * KolVoTochek) + 1, 0)
End Function

Function KontaktnayaSvarka_TTP_70_Slesarnaya(KolVoTochek As Double, Dlina As Range, Shirina As Range, KolVoPozitsiy As Range, IspolzovaniyeKleya As Boolean) As Double
    
    Dim i As Integer
       
    Dim tSkleivaniye As Double
    If IspolzovaniyeKleya Then
        For i = 1 To Dlina.Rows.Count
            If Dlina.Rows(i) > 0 Then tSkleivaniye = tSkleivaniye + (UstanovkaDetaleyNaKley(Ploskaya, Dlina.Rows(i), Shirina.Rows(i), 1, 1, 1, False) * KolVoPozitsiy.Rows(i))
        Next
    End If
    Dim tUdalitPodtekiKleya
    tUdalitPodtekiKleya = tSkleivaniye / 2
    
    Dim tPravkaPosleSvarki As Double
    For i = 1 To Dlina.Rows.Count
        If Dlina.Rows(i) > 0 Then tPravkaPosleSvarki = tPravkaPosleSvarki + (PravkaZagIzListMatVRuch(Shirina.Rows(i), Dlina.Rows(i), 3, 1) * 1.3 * KolVoPozitsiy.Rows(i))
    Next
    
    Dim tZachistki As Double
    tZachistki = 0.3 * KolVoTochek
    
    KontaktnayaSvarka_TTP_70_Slesarnaya = WorksheetFunction.RoundUp(tUdalitPodtekiKleya + tPravkaPosleSvarki + tZachistki, 0)
    
    
End Function



Function KrupShagRezb(d As Double) As Double

    Dim tb As ListObject
    Set tb = ThisWorkbook.Worksheets("Ток_23, 24, 25, 26").ListObjects("tbKrupShagRezb")
    Dim i As Integer
    For i = 1 To tb.ListRows.Count
        If tb.DataBodyRange(i, 1) >= d Then
            KrupShagRezb = tb.DataBodyRange(i, 2)
            Exit For
        End If
    Next
    
End Function

Function MaksZnacheniyePoNaboruUsloviy(DiapazonZnacheniy As Range, ParamArray CriteriiUsloviye() As Variant) As Variant

    Dim ArrZnach() As Variant
    ArrZnach = DiapazonZnacheniy.Value

    Dim ArrUsloviy() As Variant
    Dim i As Long, j As Long
    Dim Uslovie As Variant
    Dim MaksZnacheniye As Double
    Dim EstSovpadeniye As Boolean
    Dim UsloviyaSoblyuDeny As Boolean

    If UBound(CriteriiUsloviye) Mod 2 <> 1 Then
        MaksZnacheniyePoNaboruUsloviy = CVErr(xlErrValue)
        Exit Function
    End If

    ' Преобразуем все диапазоны условий в массивы
    Dim KolUsloviy As Long
    KolUsloviy = (UBound(CriteriiUsloviye) + 1) \ 2
    ReDim ArrUsloviy(1 To KolUsloviy)

    For j = 0 To UBound(CriteriiUsloviye) Step 2
        ArrUsloviy(j \ 2 + 1) = CriteriiUsloviye(j).Value
    Next j

    For i = 1 To UBound(ArrZnach, 1)
        UsloviyaSoblyuDeny = True

        For j = 0 To UBound(CriteriiUsloviye) Step 2
            Uslovie = CriteriiUsloviye(j + 1)
            If ArrUsloviy(j \ 2 + 1)(i, 1) <> Uslovie Then
                UsloviyaSoblyuDeny = False
                Exit For
            End If
        Next j

        If UsloviyaSoblyuDeny Then
            If ArrZnach(i, 1) > MaksZnacheniye Then
                MaksZnacheniye = ArrZnach(i, 1)
                EstSovpadeniye = True
            End If
        End If
    Next i

    If EstSovpadeniye Then
        MaksZnacheniyePoNaboruUsloviy = MaksZnacheniye
    End If

End Function

Function MelkiyShagRezby(d As Double) As Double

    Dim tb As ListObject
    Set tb = ThisWorkbook.Worksheets("Ток_23, 24, 25, 26").ListObjects("tbMetrRezb_Shag")
    Dim i As Integer
    For i = 1 To tb.ListRows.Count
        If tb.DataBodyRange(i, 1) >= d Then
            If tb.DataBodyRange(i, 1) <> tb.DataBodyRange(i + 1, 1) Then
                MelkiyShagRezby = tb.DataBodyRange(i, 2)
                Exit For
            End If
        End If
    Next
    
End Function

Function OkruglenieTsht(Tsht As Double) As Double
    
    If Tsht > 0 Then
        Select Case Tsht
            Case Is <= 1.8:     OkruglenieTsht = WorksheetFunction.RoundUp(Tsht, 1)
            Case Is <= 5:       OkruglenieTsht = WorksheetFunction.Ceiling(Tsht, 0.5)
            Case Else:          OkruglenieTsht = WorksheetFunction.RoundUp(Tsht, 0)
        End Select
    End If
    
End Function

Function OkrugleniyeTsht_5(Tsht As Double) As Double
    
    If Tsht > 12 Then
        OkrugleniyeTsht_5 = WorksheetFunction.Ceiling(Tsht, 5)
    Else
        OkrugleniyeTsht_5 = OkruglenieTsht(Tsht)
    End If
    
End Function

Function PokazateliStepennoiFunkcii(rY As Range, rX As Range)
    
    Dim NewRY() As Double, NewRX() As Double
    Dim R As Range, i As Integer, j As Integer
    
    i = 0
    For Each R In rY.Cells
        ReDim Preserve NewRY(i)
        NewRY(i) = R.Value
        i = i + 1
    Next
    
    i = 0
    For Each R In rX.Cells
        ReDim Preserve NewRX(i)
        NewRX(i) = R.Value
        i = i + 1
    Next
    
    Dim ArrY() As Variant, ArrX() As Variant
    For i = 0 To UBound(NewRY)
        If NewRY(i) > 0 Then
            ReDim Preserve ArrY(j)
            ReDim Preserve ArrX(j)
            ArrY(j) = NewRY(i)
            ArrX(j) = NewRX(i)
            j = j + 1
        End If
    Next

    Dim Pokazateli(1 To 2) As Double
    Pokazateli(1) = Exp(Application.Index(Application.LinEst(Application.Ln(ArrY), Application.Ln(ArrX), 1), 2))
    Pokazateli(2) = Application.Index(Application.LinEst(Application.Ln(ArrY), Application.Ln(ArrX)), 1)
    
    PokazateliStepennoiFunkcii = Pokazateli
    
End Function


Function ProverGermetSvarShvaKerosinom(L As Double, Optional Sloznost As Integer = 3) As Double
    'Нормативы времени на контроль операций в слесарно-сборочных и сварочных цехах и участках котельного производства.
    'Единичное и мелкосерийное производства. Таганрок 1984.
    'Карта 63. Проверка герметичности сварного шва при исытании керосином.
    
    If L <= 0 Then Exit Function
    
    Dim a As Double:    a = 9.24415544081723E-04
    Dim b As Double:    b = 1.01907775713926
    
    Dim K1 As Double
    Select Case Sloznost
        Case 1:      K1 = 0.8
        Case 2:      K1 = 1
        Case 3:      K1 = 1.3
    End Select
    
    ProverGermetSvarShvaKerosinom = (a * L + b) * K1

End Function

Function PrikleivaniyeMontazhnoyLenty_k_Planke(Dlina As Double, Shirina As Double) As Double
    
    Dim Perimetr As Double
    Perimetr = (Dlina + Shirina) * 2
    
    Dim tRazmetka As Double
    tRazmetka = RazmetkaKonturaBezShablona_Lineykoy(2, 1, Perimetr)
    
    Dim tVyrezaniye As Double
    tVyrezaniye = VyrezaniyeProkladokVruchnuyu(1, 1, Perimetr)
    
    Dim tSkleivaniye
    tSkleivaniye = UstanovkaDetaleyNaKley(Ploskaya, Dlina, Shirina, 1, 1, 1, False) * 0.5
    
    PrikleivaniyeMontazhnoyLenty_k_Planke = (tRazmetka + tVyrezaniye + tSkleivaniye)
    
End Function


Function Sborka(ByVal Naimenovaniye As String, Optional Zapodlitso As Boolean = False) As Double
    
    Dim NomerStolbtsa As Integer
    NomerStolbtsa = IIf(Zapodlitso, 3, 2)

    If Naimenovaniye = "" Then Exit Function

    Dim Naimenovaniye_1 As String
    Naimenovaniye_1 = Split(Naimenovaniye, " ")(0)
    
    Dim Naimenovaniye_2 As String
    If UBound(Split(Naimenovaniye, " ")) > 0 Then
        Naimenovaniye_2 = Naimenovaniye_1 & " " & Split(Naimenovaniye, " ")(1)
    End If
    
    Dim Naimenovaniye_3 As String
    If UBound(Split(Naimenovaniye, " ")) = 2 Then
        Naimenovaniye_3 = Naimenovaniye_2 & " " & Split(Naimenovaniye, " ")(2)
    End If
    
    Dim i As Long
    Dim tbSborka As ListObject
    Set tbSborka = wsSborka.ListObjects("tbSborka")
    
    With tbSborka
    
        If Naimenovaniye_3 <> "" Then
            For i = 1 To .ListRows.Count
                If LCase(Naimenovaniye_3) = LCase(.DataBodyRange(i, 1)) Then
                    Sborka = .DataBodyRange(i, NomerStolbtsa)
                    Exit For
                End If
            Next
        End If
    
        If Naimenovaniye_2 <> "" Then
            For i = 1 To .ListRows.Count
                If LCase(Naimenovaniye_2) = LCase(.DataBodyRange(i, 1)) Then
                    Sborka = .DataBodyRange(i, NomerStolbtsa)
                    Exit For
                End If
            Next
        End If

        If Sborka = 0 Then
            For i = 1 To .ListRows.Count
                If LCase(Naimenovaniye_1) = LCase(.DataBodyRange(i, 1)) Then
                    Sborka = .DataBodyRange(i, NomerStolbtsa)
                    Exit For
                End If
            Next
        End If
    
    End With

End Function

Function TshtLaserPress(ByVal DlinaZagotovki As Double, ByVal ShirinaZagotovki As Double, TolschinaZagotovkiotovki As Double, _
    ByVal DlinaDetali As Double, ByVal ShirinaDetali As Double, tMin As Double, tSec As Double, Material As EnumMaterialy, Optional KolVo As Double = 1)

    Dim tMashMin As Double
    tMashMin = (tSec / 60) + tMin
    
    Dim DlinaShirina
    
    DlinaShirina = Array(DlinaZagotovki, ShirinaZagotovki)
    DlinaZagotovki = WorksheetFunction.Large(DlinaShirina, 1)
    ShirinaZagotovki = WorksheetFunction.Large(DlinaShirina, 2)
    
    DlinaShirina = Array(DlinaDetali, ShirinaDetali)
    DlinaDetali = WorksheetFunction.Large(DlinaShirina, 1)
    ShirinaDetali = WorksheetFunction.Large(DlinaShirina, 2)
    
    If DlinaZagotovki > 0 And ShirinaZagotovki > 0 And DlinaDetali > 0 And ShirinaDetali > 0 And tMashMin > 0 And DlinaZagotovki >= DlinaDetali And ShirinaZagotovki >= ShirinaDetali Then
    
        Dim KolvoDetVZag As Double
        KolvoDetVZag = WorksheetFunction.RoundDown(DlinaZagotovki / DlinaDetali, 0) * WorksheetFunction.RoundDown(ShirinaZagotovki / ShirinaDetali, 0)
        If KolVo > 1 Then KolvoDetVZag = KolvoDetVZag * KolVo
        
        Dim Massa As Double
        Massa = MassaLista(DlinaZagotovki, ShirinaZagotovki, TolschinaZagotovkiotovki, Material)
        
        Dim tUstanovka As Double, a As Double, X As Double
        Select Case Massa
            Case Is <= 20
                a = 0.73
                X = 0.26
            Case Is <= 3000
                a = 2
                X = 0.22
            Case Is <= 30000
                a = 1.05
                X = 0.3
            Case Else
                a = 0.233
                X = 0.45
        End Select
        tUstanovka = a * Massa ^ X
        
        Dim Tsht As Double
        Tsht = (tMashMin + tUstanovka / KolvoDetVZag) * 1.1
        
        TshtLaserPress = OkruglenieTsht(Tsht)
        
    End If
    
    
End Function
Function KolVoDetIzPrutka(DlinaZagotovki As Double, KolVoDetaleyVZagotovke As Double, DiametrPrutka As Double) As Double
    'Функция для расчета количества деталей на которое мы точим наружный диаметр при работе из прутка
    Dim maxL As Double
    Dim dFive As Double
    
    If KolVoDetaleyVZagotovke = 1 Then
        KolVoDetIzPrutka = 1
    Else
        maxL = (DlinaZagotovki / KolVoDetaleyVZagotovke) * (KolVoDetaleyVZagotovke - 1)
        dFive = 5 * DiametrPrutka
        KolVoDetIzPrutka = Fix((KolVoDetaleyVZagotovke / DlinaZagotovki) * WorksheetFunction.Min(dFive, maxL))
        If KolVoDetIzPrutka = 0 Then KolVoDetIzPrutka = 1
    End If
End Function

Function DlinaIzPutka(Lzag As Double, KolDet_v_Prutke As Double, KolDet_Fact As Double) As Double
    'Функция для расчета длина точения при работе из прутка
    DlinaIzPutka = (Lzag / KolDet_v_Prutke) * KolDet_Fact
End Function


Function ZachistkaZausencevNaShestigrannike(Material As EnumMaterialy, S As Double, ZachistkaSDvuhStoron As Boolean, _
    Optional ProdolnayaDlina As Double = 0) As Double
    
    ZachistkaZausencevNaShestigrannike = 0
    
    If Not S > 0 Then Exit Function

    Dim KolVoStoron As Integer
    KolVoStoron = IIf(ZachistkaSDvuhStoron, 2, 1)
    
    Dim t1 As Double
    t1 = KolVoStoron * ZachistkaZausencev_PoKonturu_Napilnikom(Material, 6 * DlinaGraniShestigrannika(S), False, False)
    
    Dim t2 As Double
    If ProdolnayaDlina > 0 Then t2 = ZachistkaZausencev_PoKonturu_Napilnikom(Material, 6 * ProdolnayaDlina, False, False)
    
    Dim tUst As Double
    tUst = 0.32
    
    ZachistkaZausencevNaShestigrannike = KolVoStoron * tUst + t1 + t2

End Function



Function ZamenElektrElemnt(KolVivodov As Integer) As Double
    If KolVivodov > 0 Then
        ZamenElektrElemnt = KolVivodov * 1.421052632 + 3.157894737
        If KolVivodov > 40 Then ZamenElektrElemnt = 60
    End If
End Function

Function ObortkaZhguta(lkmM2 As Double, GruppaSlozhnosti As Integer, DlinaObmotkiMetr As Double)
    
    Dim tNaMetrZhguta As Double
    'Норма времени на обмотку 1 метра жгута (Электромонтажные работы Часть 1 Карта 46, лист 100)
    Select Case GruppaSlozhnosti
        Case 1: tNaMetrZhguta = 21.4
        Case 2: tNaMetrZhguta = 27.8
        Case 3: tNaMetrZhguta = 32
    End Select
    
    Dim lkmMM As Double
    lkmMM = lkmM2 * 1000
    
    Dim ShirinaPoloski As Double
    ShirinaPoloski = 15

    Dim tZagotovka As Double
    tZagotovka = 0.0147 * (lkmMM ^ 0.43)  'Норма времени на 1 заготовку (Электромонтажные работы Часть 1 Карта 4 поз. 6, лист 29)
    
    Dim KolVoPolosok As Integer
    KolVoPolosok = WorksheetFunction.RoundDown(lkmMM / ShirinaPoloski, 0)
    
    ObortkaZhguta = (tZagotovka * KolVoPolosok + DlinaObmotkiMetr * tNaMetrZhguta) * 1.5
    
    
End Function

Function ObortkaZhgutaAZT(DlinaObmotkiMetr As Double, Optional GruppaSlozhnosti As Integer = 1)
            
    Dim tNaMetrZhguta As Double
    'Норма времени на обмотку 1 метра жгута (Электромонтажные работы Часть 1 Карта 46, лист 100)

    Select Case GruppaSlozhnosti
        Case 1: tNaMetrZhguta = 18.4
        Case 2: tNaMetrZhguta = 24
        Case 3: tNaMetrZhguta = 27.6
    End Select
    
    Dim tZagotovka As Double
    tZagotovka = 0.348  'Норма времени на 1 заготовку длиной 0,8 м (ширина рулона материи облиц.) (Электромонтажные работы Часть 1 Карта 4 поз. 3, лист 29)
    
    Dim KolVoPolosok As Integer
    KolVoPolosok = WorksheetFunction.RoundUp(DlinaObmotkiMetr / 0.8, 0) 'Ширна рулона АЗТс (0,8 м)
        
    ObortkaZhgutaAZT = (tZagotovka * KolVoPolosok + DlinaObmotkiMetr * tNaMetrZhguta) * 1.5
    
    
End Function









