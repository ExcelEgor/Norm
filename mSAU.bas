Attribute VB_Name = "mSAU"
Option Explicit
'ÑÁÎĞÎ×ÍÛÅ ĞÀÁÎÒÛ ÏĞÈ ÈÇÃÎÒÎÂËÅÍÈÈ ÈÇÄÅËÈÉ ÑÀÓ. ÍÎĞÌÀÒÈÂÛ ÂĞÅÌÅÍÈ
'ÌÅËÊÎÑÅĞÈÉÍÎÅ ÏĞÎÈÇÂÎÄÑÒÂÎ. 1979 ã.

Public Enum eTipKleya
    BF_AK20_88N = 1     'ÁÔ, ÀÊ-20, 88-Í;
    Germetik_EPA_LN = 2 'Ãåğìåòèê, İÏÀ, ËÍ
End Enum

Function UstanovkaDetaleyNaKley(TipDetali As EnumTipPoverhnosti, Dlina As Double, ShirinaIliDiametr As Double, _
    Optional Slozhnost As Integer = 2, Optional TipKleya As eTipKleya = 2, Optional KolVoSloev As Integer = 2, Optional UstanovkaVKanavkuPaz As Boolean = False) As Double
     'Êàğòà 8. Óñòàíîâêà äåòàëåé íà êëåé
     
    If TipDetali <> Ploskaya And TipDetali <> Tsilindricheskaya Or _
        Dlina <= 0 Or ShirinaIliDiametr <= 0 Or _
        Slozhnost < 1 Or Slozhnost > 2 Or _
        TipKleya <> BF_AK20_88N And TipKleya <> Germetik_EPA_LN Or _
        KolVoSloev < 1 Or KolVoSloev > 2 Then Exit Function

    Dim tSkleivaniye As Double
    If TipDetali = Ploskaya Then
        tSkleivaniye = UstanovkaDetaleyNaKley_Ploskiye(Dlina, ShirinaIliDiametr, Slozhnost, TipKleya, KolVoSloev)
    Else
        tSkleivaniye = UstanovkaDetaleyNaKley_Tsilindricheskiye(Dlina, ShirinaIliDiametr, Slozhnost, TipKleya, KolVoSloev)
    End If
    
    If tSkleivaniye <= 0 Then Exit Function

    Dim kKanavka As Double
    kKanavka = IIf(UstanovkaVKanavkuPaz, 1.2, 1)

    UstanovkaDetaleyNaKley = kKanavka * tSkleivaniye

End Function

Private Function UstanovkaDetaleyNaKley_Ploskiye(Dlina As Double, Shirina As Double, _
    Optional Slozhnost As Integer = 2, Optional TipKleya As eTipKleya = 2, Optional KolVoSloev As Integer = 2) As Double

    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina)

    Dim MaxDlina As Integer
    MaxDlina = 2000

    If Gabarity.Dlina > MaxDlina Then Exit Function

    Static Pokazateli()
    Pokazateli = ThisWorkbook.Worksheets("ÑÀÓ_8").ListObjects("SAU_Skleivanie_8").DataBodyRange.Value

    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    For i = LBound(Pokazateli) To UBound(Pokazateli)
        If Pokazateli(i, 1) >= Gabarity.Dlina And Pokazateli(i, 2) = Slozhnost And Pokazateli(i, 3) = TipKleya And Pokazateli(i, 4) = KolVoSloev Then
            a = Pokazateli(i, 5)
            b = Pokazateli(i, 6)
            c = Pokazateli(i, 7)
            Exit For
        End If
    Next

    UstanovkaDetaleyNaKley_Ploskiye = a * Gabarity.Shirina ^ 2 + b * Gabarity.Shirina + c

End Function

Private Function UstanovkaDetaleyNaKley_Tsilindricheskiye(Dlina As Double, Diametr As Double, _
    Optional Slozhnost As Integer = 2, Optional TipKleya As eTipKleya = 2, Optional KolVoSloev As Integer = 2, Optional UstanovkaVKanavkuPaz As Boolean = False)

    Dim MaxDlina As Integer
    MaxDlina = 100
    
    If Dlina > MaxDlina Then Exit Function
    
    Dim Pokazateli()
    Pokazateli = ThisWorkbook.Worksheets("ÑÀÓ_8").ListObjects("tbUstNaKley_Tsilindr").DataBodyRange.Value
    
    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    For i = LBound(Pokazateli) To UBound(Pokazateli)
        If Pokazateli(i, 1) >= Dlina And Pokazateli(i, 2) = Slozhnost And Pokazateli(i, 3) = TipKleya And Pokazateli(i, 4) = KolVoSloev Then
            a = Pokazateli(i, 5)
            b = Pokazateli(i, 6)
            Exit For
        End If
    Next
    
    UstanovkaDetaleyNaKley_Tsilindricheskiye = a * Diametr + b
        
End Function



