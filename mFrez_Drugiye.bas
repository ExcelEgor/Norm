Attribute VB_Name = "mFrez_Drugiye"
'@Folder Frezerovaniye

Option Explicit

'Material - Материал:
'1 Ал-2
'2 Сталь 35
'3 25Х13Н2
'4 12Х18Н9Т
'5 ВНМ3


Function FrezerovaniyeGabaritovKontsevymiFrezami(Material As EnumMaterialy, ByVal DlinaZagotovki As Double, ByVal ShirinaZagotovki As Double, ByVal TolshchinaZagotovki As Double, _
    Optional ByVal DlinaDetali As Double = 0, Optional ByVal ShirinaDetali As Double = 0, Optional ByVal TolshchinaDetali As Double = 0, _
    Optional ByVal Ra As Double = 6.3, Optional ByVal IT As Integer = 11)
    
    Dim RaschotRezhimov As New clsRezhimy_FrezKonts, Rezhimy As clsMillingRegimes, Dfrezy As Double
    
    Dim ArrZagotovka(1 To 3) As Double
    ArrZagotovka(1) = DlinaZagotovki
    ArrZagotovka(2) = ShirinaZagotovki
    ArrZagotovka(3) = TolshchinaZagotovki
    
    Dim ArrDetal(1 To 3) As Double
    ArrDetal(1) = DlinaDetali
    ArrDetal(2) = ShirinaDetali
    ArrDetal(3) = TolshchinaDetali
    
    With WorksheetFunction
        DlinaZagotovki = .Large(ArrZagotovka, 1)
        ShirinaZagotovki = .Large(ArrZagotovka, 2)
        TolshchinaZagotovki = .Large(ArrZagotovka, 3)
        
        DlinaDetali = .Large(ArrDetal, 1)
        ShirinaDetali = .Large(ArrDetal, 2)
        TolshchinaDetali = .Large(ArrDetal, 3)
    End With
    
    Dim KolVoUstanovok As Integer
    
    Dim ctx As New clsMillingContext
    
    With ctx
        .AvtoRezhimy = True
        .CHPU = False
        .IT = IT
        .Korka = False
        .Material = MEDNYYE_SPLAVY
        .PoRezhimam = True
        .Ra = Ra
        .TipStanka = VertikalnoGorizontalnoFrezernyy
    End With
    
    Dim tFrezTolshchina As Double, hTolshchina As Double
    If TolshchinaDetali > 0 Then
        hTolshchina = (TolshchinaZagotovki - TolshchinaDetali) / 2
        
        With ctx
            .DlinaPoverkhnosti = DlinaZagotovki
            .Pripusk = hTolshchina
            .ShirinaPoverkhnosti = ShirinaZagotovki
        End With
        
        tFrezTolshchina = FrezKontsevymiFrezami_Kompleks(ctx)(1)
        KolVoUstanovok = KolVoUstanovok + 2
        
    End If
    
    Dim tFrezDlina As Double, hDlina As Double
    If DlinaDetali > 0 Then
        hDlina = (DlinaZagotovki - DlinaDetali) / 2
        
        With ctx
            .DlinaPoverkhnosti = ShirinaZagotovki
            .Pripusk = hDlina
            .ShirinaPoverkhnosti = TolshchinaZagotovki
        End With
        
        tFrezDlina = FrezKontsevymiFrezami_Kompleks(ctx)
        KolVoUstanovok = KolVoUstanovok + 2
    End If
    
    Dim tFrezShirina As Double, hShirina As Double
    If ShirinaDetali > 0 Then
        hShirina = (ShirinaZagotovki - ShirinaDetali) / 2
        
        With ctx
            .DlinaPoverkhnosti = DlinaZagotovki
            .Pripusk = hDlina
            .ShirinaPoverkhnosti = TolshchinaZagotovki
        End With
        
        tFrezShirina = FrezKontsevymiFrezami_Kompleks(ctx)(1)
        KolVoUstanovok = KolVoUstanovok + 2
    End If
    
    Dim tFrezGabarity As Double
    tFrezGabarity = (tFrezTolshchina + tFrezDlina + tFrezShirina) * 2
    
    Dim MassaZagotovki As Double
    MassaZagotovki = MassaLista(DlinaZagotovki, ShirinaZagotovki, TolshchinaZagotovki, CInt(Material))
    
    Dim tUstanovki As Double
    If MassaZagotovki <= 20 Then
        tUstanovki = Ustanov_Vtiskah(MassaZagotovki)
    Else
        tUstanovki = Ustanov_NaStoleKreplBoltPlank(MassaZagotovki)
    End If
    tUstanovki = tUstanovki * KolVoUstanovok
    
    FrezerovaniyeGabaritovKontsevymiFrezami = tFrezGabarity + tUstanovki

End Function

Function FrezOtverstiyPodKvadratnyieKnopki(RazmerKvadrata As Double, Tolshchina As Double, Material As EnumMaterialy)

    If PolozhitelnyyeChisla(RazmerKvadrata, Tolshchina) = False Then Exit Function

    'Material: 1-Эбонит; Ал -2; 2-Сталь 35; 3-Стекловолокнит; 4-25Х13Н2; 5-12Х18Н9Т; 6-Титан ВТ-1; 7-ВНМ3

    Dim tVspom As Double:       tVspom = 0.1
    Dim Lnedobeg As Integer:    Lnedobeg = 5
    
    Dim kMaterial  As Double
    If Material = ALUMINIYEVYYE_SPLAVY Then
        kMaterial = 0.85
    ElseIf Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1.25
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = 1.5
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = 2
    Else
        Exit Function
    End If
    
    Dim Svrezaniye As Double
    Svrezaniye = 300 / kMaterial
    
    Dim Lperebeg As Integer
    Dim tSverleniye As Double
    
    Lperebeg = 3
    tSverleniye = (Lnedobeg + Lperebeg + Tolshchina) / Svrezaniye + tVspom
    
    Dim Sfrezerovaniye As Double
    Dim Dfrezy As Double
    
    Dim tFrezerovaniye
    Dim i As Integer
    For i = 1 To 3
        Select Case i
            Case 1
                Dfrezy = 10
                Sfrezerovaniye = 500
            Case 2
                Dfrezy = 6
                Sfrezerovaniye = 600
            Case 3
                Dfrezy = 3
                Sfrezerovaniye = 400
        End Select
        Sfrezerovaniye = Sfrezerovaniye / kMaterial
        Lperebeg = 0
        tFrezerovaniye = tFrezerovaniye + ((Lnedobeg + Lperebeg + Tolshchina) / Svrezaniye) + ((RazmerKvadrata - Dfrezy) * 4) / Sfrezerovaniye + tVspom
    Next
    
    Dim tFrez As Double:    tFrez = (tSverleniye + tFrezerovaniye) * 1.1
    Dim tSles As Double:    tSles = VypilivaniyeVnutrennegoRadiusa(Tolshchina, 1.5, CInt(Material)) * 4
    
    'tKontrol делим на 2 т.к. фрезеруя две поверхности мы проводим только один размер
    Dim tKontrol As Double: tKontrol = IzmerShtangenCircul_CPL(RazmerKvadrata, RazmerKvadrata)

    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezOtverstiyPodKvadratnyieKnopki = Normy
    
End Function

Function FrezOkon(Material As EnumMaterialy, ByVal DlinaOkna As Double, ByVal ShirinaOkna As Double, ByVal GlubinaOkna As Double, Optional Glukhoye As Boolean = False, _
    Optional DlinaObnizki As Double = 0, Optional ShirinaObnizki As Double = 0, Optional GlubinaObnizki As Double = 0, _
    Optional IT As Integer = 14, Optional Ra As Double, Optional Tonkosten As Boolean = False, Optional TipStanka As EnumTipStankaRastFrez = 2) As Variant
    
    If PolozhitelnyyeChisla(DlinaOkna, ShirinaOkna, GlubinaOkna) = False Then Exit Function
    
    Dim Parametry
    Parametry = Array(DlinaOkna, ShirinaOkna, GlubinaOkna)
    
    With WorksheetFunction
        DlinaOkna = .Large(Parametry, 1)
        ShirinaOkna = .Large(Parametry, 2)
        GlubinaOkna = .Large(Parametry, 3)
    End With
    
    Dim tTorcevayaFreza As Double, tKoncevayaFreza As Double
    If Glukhoye Then
        tTorcevayaFreza = FrezTortsovymyFrezami_Kompleks(TipStanka, Material, GlubinaOkna, ShirinaOkna, DlinaOkna, IT, Ra, , , Tonkosten)(1)
        tKoncevayaFreza = FrezUstupov_Kompleks(TipStanka, CInt(Material), (DlinaOkna + ShirinaOkna) * 2, GlubinaOkna, 0, IT, Ra, , , Tonkosten)(1)
    Else
        tTorcevayaFreza = FrezTortsovymyFrezami_Kompleks(TipStanka, Material, GlubinaOkna, ShirinaOkna, DlinaOkna, IT, 20, , , Tonkosten)(1)
        
        Dim RaschotRezhimov As New clsRezhimy_FrezKonts, Rezhimy As clsMillingRegimes, Dfrezy As Double
        Dfrezy = RaschotRezhimov.OpredelitDfrezy_FrezPloskKonts(GlubinaOkna)
        Rezhimy = RaschotRezhimov.RaschotRezhimov(Material, MIN_PRIPUSK_RASTFREZ, Dfrezy, GlubinaOkna, IT, Ra, False)
        
        Dim ctx As clsMillingContext
        With ctx
            .AvtoRezhimy = True
            .CHPU = True
            .DlinaPoverkhnosti = 2 * (DlinaOkna + ShirinaOkna)
            .IT = IT
            .Korka = False
            .Material = MEDNYYE_SPLAVY
            .NaUdar = False
            .PoRezhimam = True
            .Pripusk = MIN_PRIPUSK_RASTFREZ
            .Ra = Ra
            .ShirinaPoverkhnosti = GlubinaOkna
            .TipStanka = TipStanka
        End With
        
        tKoncevayaFreza = FrezKontsevymiFrezami_Kompleks(ctx)(1)
    
    End If
    
    Dim tObnizka As Double
    Dim GlubinaIliShirina As Double
    If DlinaObnizki > 0 And ShirinaObnizki > 0 And GlubinaObnizki > 0 Then
        Parametry = Array(DlinaObnizki, ShirinaObnizki, GlubinaObnizki)
        With WorksheetFunction
            DlinaObnizki = .Large(Parametry, 1)
            ShirinaObnizki = .Large(Parametry, 2)
            GlubinaObnizki = .Large(Parametry, 3)
        End With
        
        GlubinaIliShirina = WorksheetFunction.Max(DlinaOkna - DlinaObnizki, ShirinaOkna - ShirinaObnizki) / 2
        If GlubinaIliShirina > 20 Then  'Максимальная глубина резания при фрезеровании уступов = 20 мм
            tObnizka = FrezTortsovymyFrezami_Kompleks(TipStanka, Material, GlubinaObnizki, GlubinaIliShirina, (DlinaObnizki + DlinaOkna) * 2, 14, Ra, , , Tonkosten)(1)
            tObnizka = tObnizka + FrezUstupov_Kompleks(TipStanka, CInt(Material), (DlinaObnizki + DlinaOkna) * 2, GlubinaObnizki, 0, IT, Ra, , , Tonkosten)(1)
        Else
            tObnizka = tObnizka + FrezUstupov_Kompleks(TipStanka, CInt(Material), (DlinaObnizki + DlinaOkna) * 2, GlubinaObnizki, GlubinaIliShirina, IT, Ra, , , Tonkosten)(1)
        End If
    End If

    Dim tFrez As Double
    tFrez = tTorcevayaFreza + tKoncevayaFreza + tObnizka
    
    Dim tSles As Double
    tSles = ZachistkaZausencev_PoKonturu_Napilnikom(Material, IIf(Glukhoye, 2, 4) * (DlinaOkna + ShirinaOkna), False, False)
    If tObnizka > 0 Then
        tSles = tSles + ZachistkaZausencev_PoKonturu_Napilnikom(Material, 2 * (DlinaObnizki + ShirinaObnizki), False, False)
    End If
    
    Dim tKontrol As Double
    tKontrol = KontrolOkna(DlinaOkna, ShirinaOkna, GlubinaOkna, DlinaObnizki, ShirinaObnizki, GlubinaObnizki, IT, Glukhoye)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezOkon = Normy

End Function

Private Function KontrolOkna(DlinaOkna As Double, ShirinaOkna As Double, GlubinaOkna As Double, _
    Optional DlinaObnizki As Double = 0, Optional ShirinaObnizki As Double = 0, Optional GlubinaObnizki As Double = 0, _
    Optional IT As Integer = 14, Optional Glukhoye As Boolean = False)
    
    Dim tKontrol_1, tKontrol_2
    
    tKontrol_1 = IzmerShtangenCircul_CPL(DlinaOkna, ShirinaOkna, IT < 11)
    If Glukhoye Then
        tKontrol_1 = tKontrol_1 + IzmerShtangenCircul(GlubinaOkna, (DlinaOkna + ShirinaOkna) * 2, IT < 11)
    End If
    If DlinaObnizki > 0 And ShirinaObnizki > 0 And GlubinaObnizki > 0 Then
        tKontrol_2 = IzmerShtangenCircul_CPL(DlinaObnizki, ShirinaObnizki, IT < 11)
        tKontrol_2 = tKontrol_2 + IzmerShtangenCircul(GlubinaObnizki, (DlinaObnizki + ShirinaObnizki) * 2, IT < 11)
    End If

    KontrolOkna = tKontrol_1 + tKontrol_2

End Function

Function FrezShestigrannikaIzKrugaKoncevoyFrezoy(Material As EnumMaterialy, S As Double, Vysota As Double, Massa As Double, _
    Optional IT As Integer = 14, Optional Ra As Double = 10, Optional Dzag As Double, Optional Ustup As Boolean = False, Optional PoRezhimam As Boolean = False, Optional CHPU As Boolean = False) As Variant
    
    If PolozhitelnyyeChisla(S, Vysota) = False Then Exit Function
    
    Dim DlinaGrani As Double
    DlinaGrani = DlinaGraniShestigrannika(S)
    
    Dim Pripusk As Double
    If Dzag > 0 And Dzag > DlinaGrani * 2 Then
        Pripusk = (Dzag - S) / 2
    Else
        Pripusk = (DlinaGrani * 2 - S) / 2
    End If
    
    Dim tFrez As Double
    If Ustup Then
        tFrez = 6 * FrezUstupov_Kompleks(2, CInt(Material), DlinaGrani, Vysota, Pripusk, IT, Ra)(1)
    Else
        Dim RaschotRezhimov As New clsRezhimy_FrezKonts, Rezhimy As clsMillingRegimes, Dfrezy As Double
    
        Dim ctx As New clsMillingContext
        With ctx
            .AvtoRezhimy = True
            .CHPU = CHPU
            .DlinaPoverkhnosti = DlinaGrani
            .IT = IT
            .Korka = False
            .Material = Material
            .NaUdar = False
            .PoRezhimam = PoRezhimam
            .Pripusk = Pripusk
            .Ra = Ra
            .ShirinaPoverkhnosti = Vysota
            .TipStanka = VertikalnoGorizontalnoFrezernyy
            .TonkostennayaKonfiguraciya = False
        End With
        
        tFrez = 6 * FrezKontsevymiFrezami_Kompleks(ctx)(1)
    End If
    
    Dim tUst As Double
    tUst = Ustanov_Vtiskah(Massa) * 6

    Dim tSles As Double, DlinaZachistki As Double
    tSles = ZachistkaZausencevNaShestigrannike(Material, S, IIf(Ustup, 1, 2), Vysota)
    
    Dim tKontrol As Double: tKontrol = IzmerShtangenCircul(S, DlinaGrani, IT < 11) + IIf(Ustup, IzmerShtangenCircul(Vysota, DlinaGrani, IT < 11), 0)

    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez + tUst
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezShestigrannikaIzKrugaKoncevoyFrezoy = Normy
    
End Function








