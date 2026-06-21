Attribute VB_Name = "mKonservatciya"
Option Explicit
Const K_tPz_tObs_Kramatorsk As Double = 1.16
Const K_tPz_tObs_Moskva As Double = 1.12

Private Function ObvyazyvaniyeProvolokoy(DlinaOdnogoVitka As Double, Optional DiametrProvoloki As Double, Optional ChisloVitkov As Integer) As Double
    
    If DiametrProvoloki = 0 Then DiametrProvoloki = 3
    If ChisloVitkov = 0 Then ChisloVitkov = 3
    
    Dim X As Double: X = 0.0579
    Dim Y As Double: Y = 0.29
    Dim z As Double: z = 0.92
    Dim W As Double: W = 0.3
    
    ObvyazyvaniyeProvolokoy = (X * DiametrProvoloki ^ Y * ChisloVitkov ^ z * DlinaOdnogoVitka ^ W) * K_tPz_tObs_Moskva
    
End Function


Function UpakovyvaniyeUgolnika(ByVal Dlina As Double, ByVal Shirina As Double, ByVal Vysota As Double)

    Dim Gabarity
    Gabarity = DlinaShirina(Dlina, Shirina, Vysota)
    Dlina = Gabarity(1)
    Shirina = Gabarity(2)
    Vysota = Gabarity(3)
    
    Dim Gipotenuza As Double
    Gipotenuza = Sqr(Shirina ^ 2 + Vysota ^ 2)
    Dim DlinaOdnogoVitka As Double: DlinaOdnogoVitka = Shirina + Vysota + Gipotenuza
    
    Dim tObvyaz As Double
    tObvyaz = ObvyazyvaniyeProvolokoy(DlinaOdnogoVitka)
    
    Dim tObert As Double
    tObert = ObertyvaniyeBumagoy_PloskieDetalie(Shirina, Dlina)
    
    UpakovyvaniyeUgolnika = tObvyaz + tObert
    
End Function

Function UpakovyvaniyePramougolnika(ByVal Dlina As Double, ByVal Shirina As Double, ByVal Tolshchina As Double)
    
    Dim Gabarity
    Gabarity = DlinaShirina(Dlina, Shirina, Tolshchina)
    Dlina = Gabarity(1)
    Shirina = Gabarity(2)
    Tolshchina = Gabarity(3)
    
    Dim tObvyaz As Double, tObert As Double
    tObert = ObertyvaniyeBumagoy_PloskieDetalie(Shirina, Dlina)
    tObvyaz = ObvyazyvaniyeProvolokoyPramougolnika(Shirina, Tolshchina)
    UpakovyvaniyePramougolnika = tObert + tObvyaz
    
End Function
Function UpakovyvaniyeShestigrannika(S As Double, L As Double)
    
    Dim DlinaOdnogoVitka As Double
    DlinaOdnogoVitka = DlinaGraniShestigrannika(S) * 6
    
    Dim tObvyaz As Double
    tObvyaz = ObvyazyvaniyeProvolokoy(DlinaOdnogoVitka)
    
    Dim tObert As Double
    tObert = ObertyvaniyeBumagoy_CilindicheskiyeDetalie(DiametrShestigrannika(S), L)
    
    UpakovyvaniyeShestigrannika = tObvyaz + tObert
    
End Function

Function UpakovyvaniyeKruglogoPrutka(d As Double, L As Double)

    Dim tObvyaz As Double
    tObvyaz = ObvyazyvaniyeProvolokoy(PI * d)
    
    Dim tObert As Double
    tObert = ObertyvaniyeBumagoy_CilindicheskiyeDetalie(d, L)
    
    UpakovyvaniyeKruglogoPrutka = tObvyaz + tObert
    
End Function

Function ObvyazyvaniyeProvolokoyPramougolnika(a As Double, b As Double) As Double
    
    Dim DlinaOdnogoVitka As Double: DlinaOdnogoVitka = (a + b) * 2
    
    ObvyazyvaniyeProvolokoyPramougolnika = ObvyazyvaniyeProvolokoy(DlinaOdnogoVitka)

End Function

Function Komplektovaniye(ObshchayaMassa As Double, KolVoIzdeliy As Double, Optional Transportirovaniye As Boolean = False) As Double
    'Укрупненные нормативы времени на консервацию и укаповку готовой продукции. Краматорск 1988
    'Карта 86
    
    If KolVoIzdeliy > 0 Then
    
        Dim MassOdnogoIzdeliya As Double
        MassOdnogoIzdeliya = ObshchayaMassa / KolVoIzdeliy
        
        Dim tKomplekt As Double
        If MassOdnogoIzdeliya > 20 Then
            tKomplekt = 0.18 * MassOdnogoIzdeliya ^ 0.45 * KolVoIzdeliy ^ 0.8
        Else
            tKomplekt = 0.16 * ObshchayaMassa ^ 0.5 * KolVoIzdeliy ^ 0.5
        End If
        'С учётом разгрузки тары
        tKomplekt = tKomplekt + tKomplekt * 0.7
        
        Dim Rasstoyaniye As Double
        Rasstoyaniye = IIf(Transportirovaniye, 100, 3)
        
        Dim tTransport As Double
        tTransport = (0.03 * Rasstoyaniye ^ 0.93) + (0.0135 * Rasstoyaniye ^ 0.92 * ObshchayaMassa ^ 0.27)
        
        Komplektovaniye = (tKomplekt + tTransport) * K_tPz_tObs_Kramatorsk
        
    End If
    
End Function


Private Function ObertyvaniyeBumagoy(TipDetali As Integer, ShirinaIliDiametr As Double, Dlina As Double) As Double
    'Общемашиностроительные нормативы времени на консервацию и упаковочные работы. Москва 1988
    'Карта 32
    
    'TipDetali:
    '1-Плоские
    '2-Цилиндрические
    
    'TipBumagi:
    '1-Бумага
    '2-Битумная бумага, пергамин, рубероид, толь
    '3-Ткань
    
    ObertyvaniyeBumagoy = (IIf(TipDetali = 1, 0.00019, 0.00041) * ShirinaIliDiametr ^ 0.67 * Dlina ^ 0.67) * K_tPz_tObs_Moskva

End Function

Function ObertyvaniyeBumagoy_PloskieDetalie(Shirina As Double, Dlina As Double) As Double
    
    ObertyvaniyeBumagoy_PloskieDetalie = _
        ObertyvaniyeBumagoy(1, WorksheetFunction.Min(Shirina, Dlina), WorksheetFunction.Max(Shirina, Dlina))

End Function

Function ObertyvaniyeBumagoy_CilindicheskiyeDetalie(Diametr As Double, Dlina As Double) As Double
    
    ObertyvaniyeBumagoy_CilindicheskiyeDetalie = ObertyvaniyeBumagoy(2, Diametr, Dlina)

End Function
