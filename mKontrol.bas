Attribute VB_Name = "mKontrol"
' ÕŒ–Ã¿“»¬€ ¬–≈Ã≈Õ» Õ¿  ŒÕ“–ŒÀ‹Õ€≈ Œœ≈–¿÷»» œ–» Ã≈’¿Õ»◊≈— Œ… Œ¡–¿¡Œ“ ≈ ƒ≈“¿À≈…, —¡Œ– ≈ » »—œ€“¿Õ»ﬂ’
' ≈‰ËÌË˜ÌÓÂ Ë ÏÂÎÍÓÒÂËÈÌÓÂ ÔÓËÁ‚Ó‰ÒÚ‚Ó

Option Explicit

Private Const MIN_DIAMETR_PERIMETR_VISUAL_KONTROL As Double = 25

Public Const MIN_IZMER_RAZMER As Double = 25
Public Const DMAX_IZMERENIE_PROBKOI As Double = 165

Public Enum ePolozheniyeShva
    Nizhnee = 1
    Vertikalnoe = 2
    Potolochnoe = 3
End Enum

Function KoefPartiynosti_Kontrol(KolVoDetaleyVPartii) As Double
    
    Select Case KolVoDetaleyVPartii
        Case Is <= 3:   KoefPartiynosti_Kontrol = 1
        Case Is <= 6:   KoefPartiynosti_Kontrol = 0.95
        Case Is <= 10:  KoefPartiynosti_Kontrol = 0.9
        Case Else:      KoefPartiynosti_Kontrol = 0.85
    End Select

End Function


Function IzmerShtangenCircul(Optional ByVal IzmerRazmer As Double = MIN_IZMER_RAZMER, Optional ByVal DlinaIzmer As Double = 25, Optional isHighPrecision As Boolean = False, Optional kolVoPoverhnostei As Double = 1) As Double
    
    If IzmerRazmer <= 0 Or DlinaIzmer <= 0 Then Exit Function
    
    IzmerRazmer = MinimalnoeZnachenie(IzmerRazmer, 25)
    DlinaIzmer = MinimalnoeZnachenie(DlinaIzmer, 25)
    
    Dim tBase As Double
    tBase = 0.009 * (DlinaIzmer ^ 0.34) * (IzmerRazmer ^ 0.35)
    
    If isHighPrecision Then tBase = tBase * 1.1
    
    IzmerShtangenCircul = PereodichnostKontrola(tBase, kolVoPoverhnostei)
End Function

Function KontrolKazhdoyTretyeyDetali(Tsht As Double, KolVoDetaley As Double) As Double
    If KolVoDetaley <= 0 Then Exit Function
    
    Dim kolVoProverok As Double
    kolVoProverok = OKRUGLVVERKH(KolVoDetaley / 3)
    
    Dim tProverok As Double
    tProverok = Tsht * kolVoProverok
    
    KontrolKazhdoyTretyeyDetali = tProverok / KolVoDetaley
End Function

Function PereodichnostKontrola(t As Double, KolVo As Double) As Double
    If KolVo <= 0 Then Exit Function
    
    Dim tTotal As Double
    tTotal = t * KolVo
    
    If KolVo >= 6 Then
        Dim tThird As Double
        tThird = t * (KolVo / 3)
        PereodichnostKontrola = IIf(tThird < tTotal, tTotal, tThird)
    Else
        PereodichnostKontrola = tTotal
    End If
End Function


' --- »«Ã≈–≈Õ»≈ À»Õ≈…Õ€’ –¿«Ã≈–Œ¬ ---

Function KontrolNaryzhnogoDiametra(d As Double, L As Double, IT As Integer) As Double
    Dim currentIT As Integer
    currentIT = IIf(IT = 0, 11, IT)
    
    Dim dopusk As Double
    dopusk = DopuskPoKvalitetu(d, currentIT)
    
    If dopusk < 100 Then
        KontrolNaryzhnogoDiametra = IzmereniyeMikrometrom(d, L)
    Else
        KontrolNaryzhnogoDiametra = IzmerShtangenCircul(d, L)
    End If
End Function

Function IzmereniyeLineykoyRuletkoy(razmer As Double) As Double
    If razmer <= 0 Then Exit Function
    
    Dim baseKoef As Double
    Dim constantAdded As Double
    
    If razmer > 1000 Then
        baseKoef = 0.0001
        constantAdded = 0.5
    Else
        baseKoef = 0.0002
        constantAdded = 0.2
    End If
   
    IzmereniyeLineykoyRuletkoy = baseKoef * razmer + constantAdded
End Function

Function IzmereniyeMikrometrom(ByVal razmer As Double, ByVal Dlina As Double, Optional limitToMax As Boolean = True) As Double
    Const MIN_LEN As Double = 20
    Const MAX_LEN As Double = 2800
    Const MIN_SIZE As Double = 10
    Const MAX_SIZE As Double = 700

    If limitToMax And (Dlina > MAX_LEN Or razmer > MAX_SIZE) Then Exit Function
    
    Dlina = MinimalnoeZnachenie(Dlina, MIN_LEN)
    razmer = MinimalnoeZnachenie(razmer, MIN_SIZE)
    
    IzmereniyeMikrometrom = 0.0342 * (razmer ^ 0.3) * (Dlina ^ 0.23)
End Function

Function IzmerShtangenCircul_Range(rIzmerRazmer As Range, rKolvo As Range, Optional rDlinaIzmerPoverh As Range, Optional rTochnostIzmereniya_005 As Range) As Double
    
    Dim i As Long
    Dim TochnostIzmereniya_005 As Boolean
    Dim Dlina As Double
    For i = 1 To rIzmerRazmer.Rows.Count
        If rIzmerRazmer(i) > 0 Then
            If Not rTochnostIzmereniya_005 Is Nothing Then
                If rTochnostIzmereniya_005(i) <> 0 Then
                    TochnostIzmereniya_005 = True
                Else
                    TochnostIzmereniya_005 = False
                End If
            End If
            If Not rDlinaIzmerPoverh Is Nothing Then
                Dlina = rDlinaIzmerPoverh(i)
            Else
                Dlina = 0
            End If
            IzmerShtangenCircul_Range = IzmerShtangenCircul_Range + (IzmerShtangenCircul(rIzmerRazmer(i), Dlina, TochnostIzmereniya_005) * IIf(rKolvo(i) = 0, 1, rKolvo(i)))
        End If
    Next

End Function

Function IzmereniyeShtangenGlubinomerom(izmeryaemyyRazmer As Double, Optional isHighPrecision As Boolean = False) As Double
    If izmeryaemyyRazmer <= 0 Then Exit Function
    
    Dim precisionMultiplier As Double
    precisionMultiplier = IIf(isHighPrecision, 1.1, 1)
    
    IzmereniyeShtangenGlubinomerom = precisionMultiplier * 0.0004 * izmeryaemyyRazmer + 0.078
End Function


' --- »«Ã≈–≈Õ»≈ Œ“¬≈–—“»… » ¬Õ”“–≈ÕÕ»’ œŒ¬≈–’ÕŒ—“≈… ---

Function KontrolOtverstiy(d As Double, Dlina As Double, Optional IT As Integer = 11, Optional isGlukhoye As Boolean = False) As Double
    If d <= 0 Or Dlina <= 0 Then Exit Function
    
    Dim currentIT As Integer
    currentIT = IIf(IT = 0, 11, IT)
    
    Dim dopusk As Double
    dopusk = DopuskPoKvalitetu(d, currentIT)
    Dim tDiametr As Double
    
    If dopusk < 100 Then
        If d > 450 Then
            tDiametr = IzmereniyeShtikhmassom(d, Dlina)
        ElseIf d > DMAX_IZMERENIE_PROBKOI Then
            tDiametr = IzmereniyeNutromerom(d, Dlina)
        Else
            tDiametr = IIf(Dlina > 250, IzmereniyeNutromerom(d, Dlina), IzmerenieOtverstiiProbkoi(d, Dlina, currentIT))
        End If
    Else
        tDiametr = IIf(d < 25, IzmerenieOtverstiiProbkoi(d, Dlina), IzmerShtangenCircul(d, Dlina))
    End If
    
    Dim tGlubina As Double
    If isGlukhoye Then tGlubina = IzmereniyeShtangenGlubinomerom(Dlina, currentIT < 11)
    
    KontrolOtverstiy = tDiametr + tGlubina
End Function

Function IzmereniyeNutromerom(ByVal razmer As Double, ByVal Dlina As Double) As Double
    Const MIN_LEN As Double = 15
    Const MAX_LEN As Double = 650
    Const MIN_SIZE As Double = 20
    Const MAX_SIZE As Double = 450

    If Dlina > MAX_LEN Or razmer > MAX_SIZE Then Exit Function
    
    Dlina = MinimalnoeZnachenie(Dlina, MIN_LEN)
    razmer = MinimalnoeZnachenie(razmer, MIN_SIZE)
    
    IzmereniyeNutromerom = 0.0255 * (razmer ^ 0.33) * (Dlina ^ 0.3)
End Function

Function IzmereniyeShtikhmassom(ByVal razmer As Double, ByVal Dlina As Double) As Double
    Const MIN_LEN As Double = 50
    Const MAX_LEN As Double = 520
    Const MIN_SIZE As Double = 50
    Const MAX_SIZE As Double = 1000

    If Dlina > MAX_LEN Or razmer > MAX_SIZE Then Exit Function
    
    Dlina = MinimalnoeZnachenie(Dlina, MIN_LEN)
    razmer = MinimalnoeZnachenie(razmer, MIN_SIZE)
    
    IzmereniyeShtikhmassom = 0.0293 * (razmer ^ 0.33) * (Dlina ^ 0.3)
End Function

Function IzmerenieOtverstiiProbkoi(ByVal Diametr As Double, Optional ByVal Glubina As Double = 0, Optional IT As Integer = 11, Optional KolVo As Double = 1) As Double
    If Diametr <= 0 Or Diametr > DMAX_IZMERENIE_PROBKOI Or Glubina > 250 Or IT < 6 Or KolVo < 1 Then Exit Function
    
    Dim currentIT As Integer
    currentIT = IIf(IT = 0, 11, IT)
    
    Diametr = MinimalnoeZnachenie(Diametr, 10)
    Glubina = MinimalnoeZnachenie(Glubina, 10)
    
    Dim koefProhod As Double
    Select Case currentIT
        Case Is > 9: koefProhod = 0.01728
        Case Is > 6: koefProhod = 0.0243
        Case Else:   koefProhod = 0.04374
    End Select
    
    Dim tProhod As Double
    tProhod = koefProhod * (Diametr ^ 0.35) * (Glubina ^ 0.3)
    
    Dim koefNeProhod As Double
    koefNeProhod = IIf(currentIT > 9, 0.02115, 0.0315)
    
    Dim tNeProhod As Double
    tNeProhod = koefNeProhod * (Diametr ^ 0.35)
    
    Dim tBase As Double
    tBase = tProhod + tNeProhod
    
    IzmerenieOtverstiiProbkoi = PereodichnostKontrola(tBase, KolVo)
End Function


' --- –≈«‹¡€,  ¿Õ¿¬ » » ÿÀ»÷€ ---

Function Kontrol_RezbovayaProbka(Diametr As Double, Optional Shag As Double = 0, Optional Kvalitet As Integer = 7) As Double
    If Diametr > 120 Or Shag > 6 Or Diametr <= 0 Then Exit Function

    If Shag <= 0 Then Shag = KrupShagRezb(Diametr)
    
    Dim dlinaVvertyvaniya As Double
    dlinaVvertyvaniya = MinimalnoeZnachenie(Shag * 3, 5)
    
    Dim tBase As Double
    tBase = (0.05058 * (Diametr ^ 0.27) * (dlinaVvertyvaniya ^ 0.75)) / (Shag ^ 0.77)
    
    Dim currentKvalitet As Integer
    currentKvalitet = IIf(Kvalitet = 0, 7, Kvalitet)
    
    Kontrol_RezbovayaProbka = IIf(currentKvalitet < 7, tBase * 1.2, tBase)
End Function

Function Kontrol_NaruzhnayaRezba(ByVal Diametr As Double, Optional ByVal Shag As Double = 0, Optional isVtoroiKlass As Boolean = False) As Double
    If Diametr > 120 Or Shag > 6 Or Diametr <= 0 Then Exit Function
    
    Diametr = MinimalnoeZnachenie(Diametr, 3.5)
    If Shag <= 0 Then Shag = KrupShagRezb(Diametr)
    Shag = MinimalnoeZnachenie(Shag, 0.35)
    
    Dim dlinaRezbi As Double
    dlinaRezbi = MinimalnoeZnachenie(Shag * 3, 5)
    
    Dim tBase As Double
    tBase = (0.0346 * (Diametr ^ 0.27) * (dlinaRezbi ^ 0.75)) / (Shag ^ 0.77)
    tBase = tBase + IIf(Diametr > 48, 0.08, 0.06)
    
    Kontrol_NaruzhnayaRezba = IIf(isVtoroiKlass, tBase * 1.2, tBase)
End Function

Function KontrolNaruzhnoyKanavki(d As Double, Shirina As Double, IT As Integer) As Double
    KontrolNaruzhnoyKanavki = KontrolNaryzhnogoDiametra(d, Shirina, IT) + IzmerShtangenCircul(Shirina)
End Function

Function KontrolVnutrenneyKanavki(d As Double, Shirina As Double) As Double
    Dim tDiametr As Double
    If d > 110 Then
        tDiametr = IzmereniyeNutromerom(d, Shirina)
    Else
        tDiametr = IzmereniyeShtangencirculemKanavochnym(d)
    End If
    
    KontrolVnutrenneyKanavki = tDiametr + IzmerShtangenCircul(Shirina)
End Function

Function IzmereniyeShtangencirculemKanavochnym(razmer As Double, Optional L As Double = 34) As Double
    If L <= 0 Or razmer <= 0 Or razmer > 110 Or L > 60 Then Exit Function
    
    razmer = MinimalnoeZnachenie(razmer, 10)
    L = MinimalnoeZnachenie(L, 8)
    
    IzmereniyeShtangencirculemKanavochnym = WorksheetFunction.Round(0.0157 * (razmer ^ 0.354) * (L ^ 0.329), 2)
End Function

Function IzmereniyeShlitsevKalibrProbkoy(razmer As Double, Dlina As Double) As Double
    If razmer <= 0 Or Dlina <= 0 Then Exit Function
    IzmereniyeShlitsevKalibrProbkoy = 0.0222 * (razmer ^ 0.32) * (Dlina ^ 0.26)
End Function


' --- —¬¿–Õ€≈ ÿ¬€, –¿ƒ»”—€ » ¬»«”¿À‹Õ€…  ŒÕ“–ŒÀ‹ ---

Function ProverkaRadiusomerom(R As Double) As Double
    ProverkaRadiusomerom = 0.147 * (R ^ 0.23)
End Function

Function KontrolSvarnykhShvov(dlinaShva As Double, Optional tipShva As Integer = 2, Optional polozheniyeShva As ePolozheniyeShva = Nizhnee, Optional isLupa As Boolean = False, Optional isPosleAvtomata As Boolean = False) As Double
    If dlinaShva <= 0 Or tipShva < 1 Or tipShva > 2 Or polozheniyeShva < 1 Or polozheniyeShva > 3 Then Exit Function
    
    Dim koefA As Double, koefB As Double
    
    If tipShva = 1 Then
        Select Case polozheniyeShva
            Case 1: koefA = 0.0002: koefB = 0.09
            Case 2: koefA = 0.0002: koefB = 0.12
            Case 3: koefA = 0.0003: koefB = 0.135
        End Select
    Else ' tipShva = 2
        Select Case polozheniyeShva
            Case 1: koefA = 0.0003: koefB = 0.135
            Case 2: koefA = 0.0003: koefB = 0.185
            Case 3: koefA = 0.0004: koefB = 0.21
        End Select
    End If

    Dim tBase As Double
    tBase = koefA * dlinaShva + koefB
    
    Dim multiplier As Double
    multiplier = 1
    If isLupa Then multiplier = multiplier * 1.4
    If isPosleAvtomata Then multiplier = multiplier * 0.8

    KontrolSvarnykhShvov = tBase * multiplier
End Function

Public Function VisualnyyKontrol(ByVal DiametrIliPerimetr As Double, ByVal Dlina As Double, Optional Slozhnost As Integer = 1, _
    Optional CherezLupu As Boolean = False, Optional SravnenieSEtalonom As Boolean = False) As Double
    
    If DiametrIliPerimetr <= 0 Or Dlina <= 0 Then Exit Function

    DiametrIliPerimetr = MinimalnoeZnachenie(DiametrIliPerimetr, 25)
    Dlina = MinimalnoeZnachenie(Dlina, 65)
    
    Dim koefA As Double
    Select Case Slozhnost
        Case 1: koefA = 0.01
        Case 2: koefA = 0.0161
        Case 3: koefA = 0.0199
        Case Else: Exit Function
    End Select
    
    Dim kLupa As Double, kEtalon As Double
    kLupa = IIf(CherezLupu, 1.4, 1)
    kEtalon = IIf(SravnenieSEtalonom, 1.2, 1)

    Dim tBase As Double
    tBase = koefA * kLupa * kEtalon * (DiametrIliPerimetr * Dlina) ^ 0.3

    VisualnyyKontrol = tBase
    
End Function

Function VisualKontrolZatirkiLiniy(ByVal summarnayaDlina As Double, Optional Slozhnost As Integer = 1) As Double
    VisualKontrolZatirkiLiniy = VisualnyyKontrol(MIN_DIAMETR_PERIMETR_VISUAL_KONTROL, summarnayaDlina, Slozhnost)
End Function

Function VisualnyyKontrol_Zatirka(Shrift As Double, kolSimvolov As Double, Optional KolOsmotrov As Double = 0, Optional Slozhnost As Integer = 3) As Double
    If kolSimvolov <= 0 Then Exit Function
    
    If KolOsmotrov <= 0 Then KolOsmotrov = OKRUGLVVERKH(kolSimvolov / 43)
    
    Dim dlinaOsmotra As Double
    dlinaOsmotra = DlinaNadpisi(Shrift, kolSimvolov) / KolOsmotrov
    
    Dim tBase As Double
    tBase = VisualnyyKontrol(Shrift, dlinaOsmotra, Slozhnost) * KolOsmotrov
    
    VisualnyyKontrol_Zatirka = IIf(tBase < 0.1, 0.1, tBase)
End Function

Function VisualnyyKontrolPoPloshadi(ByVal Ploschad_mm2 As Double, Optional Slozhnost As Integer = 1, Optional isLupa As Boolean = False, Optional isSravnenieSEtalonom As Boolean = False) As Double
    
    If Ploschad_mm2 <= 0 Then Exit Function
    
    ' ‘ËÍÒËÓ‚‡ÌÌÓÂ ÒÓÓÚÌÓ¯ÂÌËÂ ÒÚÓÓÌ ËÒıÓ‰ˇ ËÁ ÏËÌËÏ‡Î¸Ì˚ı „‡·‡ËÚÓ‚ (65 / 25)
    Const K_FORM As Double = 2.6
    
    Dim p As Double, L As Double
    p = Sqr(Ploschad_mm2 / K_FORM)
    L = p * K_FORM
    
    VisualnyyKontrolPoPloshadi = VisualnyyKontrol(p, L, Slozhnost, isLupa, isSravnenieSEtalonom)
    
End Function

Function IzmerShtangenCircul_CPL(ByVal Size_1 As Double, ByVal Size_2 As Double, Optional isHighPrecision As Boolean = False, Optional KolVo As Double = 1) As Double
    KolVo = MinimalnoeZnachenie(KolVo, 1)
    IzmerShtangenCircul_CPL = IzmerShtangenCircul(Size_1, Size_2, isHighPrecision, KolVo) + IzmerShtangenCircul(Size_2, Size_1, isHighPrecision, KolVo)
End Function

