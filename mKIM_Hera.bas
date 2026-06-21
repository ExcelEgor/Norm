Attribute VB_Name = "mKIM_Hera"
Option Explicit
Option Private Module

Private Const T_TOCHKI As Double = 0.045

Function KIM_Hera_Tpz(GruppaSloznosti As Integer, VremyaNapisaniyaUP As Double, _
    KolVoKalibrovokSchupa_15 As Integer, KolVoKalibrovokSchupa_40 As Integer, KolVoKalibrovokSchupa_99 As Integer) As Double
    
    Dim TpzOsnovnoe As Double
    TpzOsnovnoe = KIM_TpzOsnovnoe(GruppaSloznosti)
    
    Dim TpzKalibrovka As Double
    TpzKalibrovka = KolVoKalibrovokSchupa_15 * TpzKalibrovka_Hera(15) + _
        KolVoKalibrovokSchupa_40 * TpzKalibrovka_Hera(40) + _
        KolVoKalibrovokSchupa_99 * TpzKalibrovka_Hera(99)
        
    KIM_Hera_Tpz = VremyaNapisaniyaUP + TpzOsnovnoe + TpzKalibrovka

End Function

Function KIM_Hera_VremyaNapisaniyaUPOdnogoPerekhoda(SlozhnostPoverkhnosti As Integer, ByVal KolVoPoverkhnostey As Double) As Double
    
    If KolVoPoverkhnostey < 0 Then Exit Function
    
    Dim tUP As Integer
    Select Case SlozhnostPoverkhnosti
        Case 1:     tUP = 10
        Case 2:     tUP = 15
        Case 3:     tUP = 20
        Case Else:  Exit Function
    End Select
    
    If KolVoPoverkhnostey = 0 Then KolVoPoverkhnostey = 1
    KIM_Hera_VremyaNapisaniyaUPOdnogoPerekhoda = SlozhnostPoverkhnosti * KolVoPoverkhnostey
    
End Function

Private Function TpzKalibrovka_Hera(DlinaSchupa As Double) As Double
    
    Select Case DlinaSchupa
        Case Is <= 15:  TpzKalibrovka_Hera = 7
        Case Is <= 40:  TpzKalibrovka_Hera = 14
        Case Else:      TpzKalibrovka_Hera = 21
    End Select
    
End Function

Function KIM_Hera_Ustanovka(Massa As Double)
    
    If Massa <= 0 Or Massa > 3000 Then Exit Function
    
    Dim a As Double, X As Double
    
    Dim tUst As Double
    Select Case Massa
        Case Is <= 0.1
            a = 0.4
            X = 0.09
        Case Is <= 20
            a = 0.96
            X = 0.2
        Case Else
            a = 2
            X = 0.22
    End Select
    
    If Massa > 0.1 Then
        tUst = a * Massa ^ X
    Else
        tUst = a / Massa ^ X
    End If
    
    Dim tBazirovaniya As Double
    tBazirovaniya = 4
    
    KIM_Hera_Ustanovka = tUst + tBazirovaniya
    
End Function

Function KIM_Hera_RaschotPloskosti(b As Double, L As Double, TrigernyyRezhim As Boolean, _
    ProverkaDopuskFormy As Boolean, ProverkaDopuskRaspolozheniya As Boolean, ProverkaLineynogoRazmera As Boolean) As Double

    If b <= 0 Or L <= 0 Then Exit Function
    
    Dim TipDopuska As Integer
    TipDopuska = OpredelitTipDopuska(ProverkaDopuskFormy, ProverkaDopuskRaspolozheniya)

    Dim tPostroeniya As Double, KolVoProhodov As Double
    If TrigernyyRezhim Then
        KolVoProhodov = KIM_KolVoSecheniy(b, TipDopuska) * KIM_KolVoSecheniy(L, TipDopuska)
        tPostroeniya = T_TOCHKI * KolVoProhodov
    Else
        KolVoProhodov = KIM_KolVoSecheniy(b, TipDopuska)
        tPostroeniya = Top_Hera(L, KolVoProhodov)
    End If
    
    Dim tAnaliz_Forma As Double
    tAnaliz_Forma = IIf(ProverkaDopuskFormy, AnalizFormiRaspolozeniya(b, L), 0)
    
    Dim tAnaliz_Raspolozeniye
    tAnaliz_Raspolozeniye = IIf(ProverkaDopuskRaspolozheniya, AnalizFormiRaspolozeniya(b, L), 0)
    
    Dim tAnaliz_Razmer As Double
    tAnaliz_Razmer = IIf(ProverkaLineynogoRazmera, 0.5 * KIM_T_ANALIZA_RAZMERA, 0)
    
    KIM_Hera_RaschotPloskosti = tPostroeniya + tAnaliz_Forma + tAnaliz_Raspolozeniye + tAnaliz_Razmer
    
End Function

Function KIM_Hera_RaschotTsilindra(d As Double, L As Double, ProverkaDopuskaDiametra As Boolean, _
    ProverkaDopuskFormy As Boolean, ProverkaDopuskRaspolozheniya As Boolean, ProverkaOtOsi As Boolean) As Double

    If d <= 0 Or L <= 0 Then Exit Function

    Dim NomerShchupa As Integer
    Select Case L
        Case Is <= 15:  NomerShchupa = 1
        Case Is <= 40:  NomerShchupa = 2
        Case Else:      NomerShchupa = 3
    End Select
    
    Dim TipDopuska As Integer
    TipDopuska = OpredelitTipDopuska(ProverkaDopuskFormy, ProverkaDopuskRaspolozheniya)
    
    Dim tPostroeniya As Double
    tPostroeniya = Top_Hera(PI * d, KIM_KolVoSecheniy(L, TipDopuska), NomerShchupa)
    
    Dim tAnaliz_Diametr As Double
    tAnaliz_Diametr = IIf(ProverkaDopuskaDiametra, KIM_T_ANALIZA_RAZMERA, 0)
    
    Dim tAnaliz_Forma As Double
    tAnaliz_Forma = IIf(ProverkaDopuskFormy, AnalizFormiRaspolozeniya(d, L), 0)
    
    Dim tAnaliz_Raspolozeniye As Double
    tAnaliz_Raspolozeniye = IIf(ProverkaDopuskRaspolozheniya, AnalizFormiRaspolozeniya(d, L), 0)
    
    Dim tAnaliz_Razmer As Double
    tAnaliz_Razmer = IIf(ProverkaOtOsi, 0.5 * KIM_T_ANALIZA_RAZMERA, 0)
    
    KIM_Hera_RaschotTsilindra = tPostroeniya + tAnaliz_Diametr + tAnaliz_Forma + tAnaliz_Raspolozeniye + tAnaliz_Razmer

End Function

Function KIM_Hera_RaschotKrugloyPloskosti(D1 As Double, d2 As Double, TrigernyyRezhim As Boolean, _
    ProverkaDopuskFormy As Boolean, ProverkaDopuskRaspolozheniya As Boolean, ProverkaLineynogoRazmera As Boolean) As Double

    If (D1 + d2) <= 0 Or d2 >= D1 Then Exit Function
    
    Dim TipDopuska As Integer
    TipDopuska = OpredelitTipDopuska(ProverkaDopuskFormy, ProverkaDopuskRaspolozheniya)
   
    Dim tPostroeniya As Double, KolVoProhodov As Double
    Dim L As Double, b As Double
    If TrigernyyRezhim Then
        L = (D1 * PI + d2 * PI) / 2
        b = (D1 - d2) / 2
        KolVoProhodov = KIM_KolVoSecheniy(b, TipDopuska) * KIM_KolVoSecheniy(L, TipDopuska)
        tPostroeniya = T_TOCHKI * KolVoProhodov
    Else
        KolVoProhodov = KIM_KolVoSecheniy((D1 - d2) / 2, TipDopuska)
        tPostroeniya = Top_Hera(L, KolVoProhodov)
    End If
    
    Dim tAnaliz_Forma As Double
    tAnaliz_Forma = IIf(ProverkaDopuskFormy, AnalizFormiRaspolozeniya(D1, (D1 - d2) / 2), 0)
    
    Dim tAnaliz_Raspolozeniye As Double
    tAnaliz_Raspolozeniye = IIf(ProverkaDopuskRaspolozheniya, AnalizFormiRaspolozeniya(D1, (D1 - d2) / 2), 0)
    
    Dim tAnaliz_Razmer As Double
    tAnaliz_Razmer = IIf(ProverkaLineynogoRazmera, 0.5 * KIM_T_ANALIZA_RAZMERA, 0)
    
    KIM_Hera_RaschotKrugloyPloskosti = tPostroeniya + tAnaliz_Forma + tAnaliz_Raspolozeniye + tAnaliz_Razmer
    
End Function

Private Function AnalizFormiRaspolozeniya(ShirinaIliDiametr As Double, Dlina As Double) As Double
    
    AnalizFormiRaspolozeniya = 0.175 * ShirinaIliDiametr ^ 0.36 * Dlina ^ 0.3
    
    Dim tMin As Integer
    tMin = 1
    
    If AnalizFormiRaspolozeniya < tMin Then AnalizFormiRaspolozeniya = tMin
    
End Function

Private Function Top_Hera(L As Double, KolVoSecheniy As Double, Optional NomerShchupa As Integer = 1) As Double

    Dim Smm_min As Double
    Smm_min = Smm_min_Hera(NomerShchupa)
    
    Top_Hera = KolVoSecheniy * (L / Smm_min + KIM_T_VSPOMOGATELNOYE)
    
End Function

Private Function Smm_min_Hera(NomerShchupa As Integer) As Double
    Select Case NomerShchupa
        Case 1: Smm_min_Hera = 250
        Case 2: Smm_min_Hera = 170
        Case 3: Smm_min_Hera = 110
    End Select
End Function

Private Function OpredelitTipDopuska(ProverkaDopuskFormy As Boolean, ProverkaDopuskRaspolozheniya As Boolean) As Integer
    
    Dim TipDopuska As Integer
    If ProverkaDopuskFormy Then
        TipDopuska = 3
    ElseIf ProverkaDopuskRaspolozheniya Then
        TipDopuska = 2
    Else
        TipDopuska = 1
    End If
    
    OpredelitTipDopuska = TipDopuska

End Function

