Attribute VB_Name = "mTermoObrabotka"
Option Explicit

Public Enum TipSecheniya
    Krugloe = 1
    Kvadratnoye = 2
    Pryamougolnoye = 3
End Enum

Function Zakalka(TipStali As Integer, TipSecheniya As TipSecheniya, Secheniye As Double, MassaOdnoyDetaly As Double, _
    Optional KolVoPechey As Integer = 3, Optional KolVoDetaley As Double = 3) As Double
        
    Dim ќсновное¬рем€ As Double
    ќсновное¬рем€ = VremyaNagreva(TipStali, CInt(TipSecheniya), Secheniye)
    
    Dim ¬спомогательное¬рем€ As Double
    ¬спомогательное¬рем€ = ZagruzkaNaPodPechi(MassaOdnoyDetaly, KolVoDetaley) + VygruzkaIzPechiVOkhlazhdSredu(MassaOdnoyDetaly, KolVoDetaley)
    
    Zakalka = TermoObrabotka(ќсновное¬рем€, ¬спомогательное¬рем€, KolVoPechey, KolVoDetaley)
        
End Function
Function VremyaNagreva(“ип—тали As Integer, “ип—ечени€ As Integer, —ечение As Double) As Double
    Dim ¬рем€Ќаћћ—ек As Double
    Select Case “ип—ечени€
        Case 1  ' руглое
            ¬рем€Ќаћћ—ек = IIf(“ип—тали = 1, (50 + 60) / 2, (75 + 90) / 2)
        Case 2  ' вадратное
            ¬рем€Ќаћћ—ек = IIf(“ип—тали = 1, (60 + 70) / 2, (80 + 90) / 2)
        Case 3  'ѕр€моугольное
            ¬рем€Ќаћћ—ек = IIf(“ип—тали = 1, (75 + 80) / 2, (90 + 100) / 2)
    End Select
    VremyaNagreva = (¬рем€Ќаћћ—ек * —ечение) / 60
End Function

Function VremyaVyderzhki(“ип—тали As Integer) As Double
    VremyaVyderzhki = IIf(“ип—тали = 1, 127.5, 135)
End Function

Function Otpusk(“ип—тали As Integer, ћассаќднойƒетали As Double, _
    Optional KolVoPechey As Integer = 3, Optional  ол¬оƒеталей As Double = 3) As Double
    
    If ћассаќднойƒетали < 0 Then Exit Function

    Dim ќсновное¬рем€ As Double
    ќсновное¬рем€ = VremyaVyderzhki(“ип—тали)
    
    Dim ¬спомогательное¬рем€ As Double
    ¬спомогательное¬рем€ = ZagruzkaNaPodPechi(ћассаќднойƒетали,  ол¬оƒеталей) + VygruzkaIzPechiVTaru(ћассаќднойƒетали,  ол¬оƒеталей)
    
    Otpusk = TermoObrabotka(ќсновное¬рем€, ¬спомогательное¬рем€, KolVoPechey,  ол¬оƒеталей)

End Function

Function ZagruzkaNaPodPechi(ћассаќднойƒетали As Double, Optional  олвоƒеталей—в€зок As Double = 3) As Double
    If ћассаќднойƒетали < 0 Then Exit Function
    If  олвоƒеталей—в€зок = 0 Then  олвоƒеталей—в€зок = 3
    ZagruzkaNaPodPechi = 0.1326 * ћассаќднойƒетали ^ 0.3 *  олвоƒеталей—в€зок ^ 0.77
End Function

Function VygruzkaIzPechiVOkhlazhdSredu(ћассаќднойƒетали As Double, Optional  олвоƒеталей—в€зок As Double = 3) As Double
    If ћассаќднойƒетали < 0 Then Exit Function
    If  олвоƒеталей—в€зок = 0 Then  олвоƒеталей—в€зок = 3
    VygruzkaIzPechiVOkhlazhdSredu = 0.143 * ћассаќднойƒетали ^ 0.3 *  олвоƒеталей—в€зок ^ 0.77
End Function
Function VygruzkaIzPechiVTaru(ћассаќднойƒетали As Double, Optional  олвоƒеталей—в€зок As Double = 3) As Double
    If  олвоƒеталей—в€зок = 0 Then  олвоƒеталей—в€зок = 3
    VygruzkaIzPechiVTaru = 0.0865 * ћассаќднойƒетали ^ 0.3 *  олвоƒеталей—в€зок ^ 0.77
End Function

Function TermoObrabotka(ќсновное¬рем€ As Double, ¬спомогательное¬рем€ As Double, Optional KolVoPechey As Integer = 3, Optional  ол¬оƒеталей As Double = 3) As Double
    
    If Not ќсновное¬рем€ > 0 Then Exit Function
    
    Dim ќперативное¬рем€ As Double
    Dim ¬рем€Ќаблюдени€ As Double
    Dim ¬рем€«ан€тости As Double
    Dim KoeffitsientZanyatosti As Double
    
    ќперативное¬рем€ = ќсновное¬рем€ + ¬спомогательное¬рем€
    ¬рем€Ќаблюдени€ = ќсновное¬рем€ * 0.06
    ¬рем€«ан€тости = ¬спомогательное¬рем€ + ¬рем€Ќаблюдени€
    KoeffitsientZanyatosti = ¬рем€«ан€тости / ќперативное¬рем€
    
    Dim  оэффициент—овпадени€ As Double
     оэффициент—овпадени€ = KoeffitsientSovpadeniya(KoeffitsientZanyatosti, KolVoPechey)
    
    Dim ќбслуживаниеќтдых As Double
    ќбслуживаниеќтдых = IIf(KoeffitsientZanyatosti < 0.5, 1.05, 1.13)
    
    TermoObrabotka = ((ќперативное¬рем€ *  оэффициент—овпадени€) / (KolVoPechey *  ол¬оƒеталей)) * ќбслуживаниеќтдых

    
End Function

Private Function KoeffitsientSovpadeniya(KoeffitsientZanyatosti As Double, Optional KolVoPechey As Integer = 3) As Double
    
    Dim KzArray  'ћассив коэффициентов зан€тости
    Dim KsArray  'ћассив коэффициентов совпадени€
    
    Select Case KolVoPechey
    
        Case Is <= 2
            KzArray = Array(0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45)
            KsArray = Array(1.01, 1.02, 1.04, 1.06, 1.09, 1.12, 1.16, 1.2, 1.26)
            
        Case Is <= 3
            KzArray = Array(0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45)
            KsArray = Array(1.01, 1.02, 1.05, 1.1, 1.15, 1.2, 1.3, 1.4, 1.5, 1.6)

        Case Is <= 4
            KzArray = Array(0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45)
            KsArray = Array(1.02, 1.04, 1.1, 1.16, 1.26, 1.4, 1.5, 1.7, 1.85, 2.05)

        Case Is <= 5
            KzArray = Array(0.05, 0.1, 0.15, 0.2, 0.25)
            KsArray = Array(1.03, 1.04, 1.15, 1.23, 1.4, 1.65)
            
        Case Is <= 6
            KzArray = Array(0.05, 0.1, 0.15)
            KsArray = Array(1.04, 1.07, 1.2, 1.35)
            
        Case Is <= 7
            KzArray = Array(0.05, 0.1)
            KsArray = Array(1.07, 1.12, 1.25)
            
        Case Is <= 8
            KzArray = Array(0.05)
            KsArray = Array(1.09, 1.21)
            
    End Select
    
    If KoeffitsientZanyatosti > KzArray(UBound(KzArray)) Then
        KsArray = KsArray(UBound(KsArray))
    Else
        KsArray = KsArray(BlizhBolshRavn_Pozic_Array(KoeffitsientZanyatosti, KzArray) - 1)
    End If
    
    KoeffitsientSovpadeniya = KsArray
    
End Function

