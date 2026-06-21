Attribute VB_Name = "mDatron_ObrabotkaOtverstiy"
Option Explicit

Private Type ParametryFrezerovaniya
    Dfrezy As Double
    AP_CHERN As Double
    ap_chist As Double
    Podacha As Long
End Type


Function Datron_ObrabotkaOtverstiy(Material As EnumMaterialy, Dotv As Double, Lotv As Double, Glukhoye As Boolean, VysotaFaski As Double, _
    Optional Dobnizheniya As Double, Optional Lobnizheniya As Double) As Variant
    
    If PolozhitelnyyeChisla(Dotv, Lotv) = False Then Exit Function
    
    Dim NormyOtverstiya As NormyVremeni
    NormyOtverstiya = RaschotOtverstiya(Material, Dotv, Lotv, Glukhoye)
    
    Dim NormyObnizheniya As NormyVremeni
    NormyObnizheniya = RaschotObnizheniya(Material, Dobnizheniya, Lobnizheniya, Dotv)
    
    Dim tFaska As Double
    If VysotaFaski > 0 Then tFaska = RaschotTmash_FrezFaski(Material, Dotv, VysotaFaski)
    
    Dim tOtv As Double
    tOtv = RaschotTsht_Datron(NormyOtverstiya.tMekhanika + NormyObnizheniya.tMekhanika + tFaska)

    Dim Normy(1 To 3) As Double
    Normy(1) = tOtv
    Normy(2) = NormyOtverstiya.tSlesar + NormyObnizheniya.tSlesar
    Normy(3) = NormyOtverstiya.tKontrol + NormyObnizheniya.tKontrol
    
    Datron_ObrabotkaOtverstiy = Normy
    
End Function

Private Function RaschotOtverstiya(Material As EnumMaterialy, Dotv As Double, Lotv As Double, Glukhoye As Boolean) As NormyVremeni

    Const MAKS_D_ZA_ODNU_OKR As Integer = 23
    Const MIN_D_DLYA_VYREZANIYA As Integer = 50

    With RaschotOtverstiya
                
        If Dotv >= MIN_D_DLYA_VYREZANIYA And Not Glukhoye Then
            .tMekhanika = Datron_VyrezaniyeOtverstiya(Material, Dotv, Lotv)
        Else
            If Dotv <= MAKS_D_ZA_ODNU_OKR Then
                .tMekhanika = RaschotTmash_ZaOdnuOkruzhnost(Material, Dotv, Lotv)
            Else
                .tMekhanika = RaschotTmash_ZaNeskolkoOkruzhnostey(Material, Dotv, Lotv)
            End If
        End If
        
        .tSlesar = ZachistkaZausencevSOtverstiiVruchnuyu(Material, Dotv, 2, IIf(Glukhoye, 1, 2), False)
        .tKontrol = KontrolOtverstiy(Dotv, Lotv, 11, Glukhoye)
        
    End With

End Function

Private Function RaschotTmash_ZaOdnuOkruzhnost(Material As EnumMaterialy, Dotv As Double, Lotv As Double) As Double

    Dim Parametry As ParametryFrezerovaniya
    Parametry = OpredelitParametryFrezerovaniya_ObrabotkaZaOdnuOkruzhnost(Material, Dotv)
    
    With Parametry
    
        Dim KolVoVrezaniy As Long
        KolVoVrezaniy = 1 + OKRUGLVVERKH(Lotv / .AP_CHERN) + OKRUGLVVERKH(Lotv / .ap_chist)
        
        Dim lRezaniya As Double
        lRezaniya = KolVoVrezaniy * PI * (Dotv - .Dfrezy) + Lotv
    
        RaschotTmash_ZaOdnuOkruzhnost = lRezaniya / .Podacha
    
    End With
    
End Function

Private Function OpredelitParametryFrezerovaniya_ObrabotkaZaOdnuOkruzhnost(Material As EnumMaterialy, Dotv As Double) As ParametryFrezerovaniya
    
    With OpredelitParametryFrezerovaniya_ObrabotkaZaOdnuOkruzhnost

        Select Case Dotv
            Case Is <= 1.5: .Dfrezy = 1
            Case Is <= 2:   .Dfrezy = 1.5
            Case Is <= 3:   .Dfrezy = 2
            Case Is <= 5:   .Dfrezy = 3
            Case Is <= 7:   .Dfrezy = 4
            Case Is <= 9:   .Dfrezy = 5
            Case Is <= 11:  .Dfrezy = 6
            Case Is <= 15:  .Dfrezy = 8
            Case Is <= 19:  .Dfrezy = 10
            Case Else:      .Dfrezy = 12
        End Select

        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                .AP_CHERN = IIf(.Dfrezy <= 6, 0.2, 0.5)
                .Podacha = IIf(.Dfrezy <= 4, 500, 1000)
            Case TEKSTOLIT, STEKLOTEKSTOLIT
                .AP_CHERN = IIf(.Dfrezy <= 6, 0.2, 0.5)
                .Podacha = 400
            Case ORGSTEKLO
                .AP_CHERN = IIf(.Dfrezy <= 6, 0.2, 0.5)
                .Podacha = 800
            Case POLIAMID
                .AP_CHERN = IIf(.Dfrezy <= 6, 0.5, 1)
                .Podacha = IIf(.Dfrezy <= 4, 1000, 2000)
            Case Else
                Exit Function
        End Select
        
        .ap_chist = .Dfrezy

    End With
    
End Function

Private Function RaschotTmash_ZaNeskolkoOkruzhnostey(Material As EnumMaterialy, Dotv As Double, Lotv As Double) As Double

    Dim Parametry As ParametryFrezerovaniya
    Parametry = OpredelitParametryFrezerovaniya_ObrabotkaZaNeskolkoOkruzhnostey(Material, Dotv)
    
    With Parametry
    
        Dim DlinaRezaniya As Double
        DlinaRezaniya = RaschotDlinyNeskolkikhOkruzhnostey(Dotv, 0, .Dfrezy, Lotv, .AP_CHERN, .ap_chist)
    
        RaschotTmash_ZaNeskolkoOkruzhnostey = DlinaRezaniya / .Podacha
    
    End With
    
End Function

Private Function OpredelitParametryFrezerovaniya_ObrabotkaZaNeskolkoOkruzhnostey(Material As EnumMaterialy, Dotv As Double) As ParametryFrezerovaniya

    Const D_OTV_DLYA_FREZY_20 As Long = 50

    With OpredelitParametryFrezerovaniya_ObrabotkaZaNeskolkoOkruzhnostey
    
        .AP_CHERN = 1
    
        If Dotv > D_OTV_DLYA_FREZY_20 Then
            .Dfrezy = 20
        Else
            .Dfrezy = 8
        End If
        .ap_chist = 8
        
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                .Podacha = 3000
            Case POLIAMID
                .Podacha = 2000
            Case ORGSTEKLO
                .Podacha = 800
            Case TEKSTOLIT, STEKLOTEKSTOLIT
                .Podacha = 400
            Case Else
                Exit Function
        End Select
        
    End With

End Function

Private Function RaschotObnizheniya(Material As EnumMaterialy, Dobnizheniya As Double, Lobnizheniya As Double, Dotv As Double) As NormyVremeni

    If PolozhitelnyyeChisla(Dobnizheniya, Lobnizheniya, Dotv) = False Or Dotv >= Dobnizheniya Then Exit Function

    Dim Parametry As ParametryFrezerovaniya
    Parametry = OpredelitParametryFrezerovaniya_ObrabotkaZaNeskolkoOkruzhnostey(Material, Dobnizheniya)
    
    With Parametry
     
        Dim DlinaRezaniya As Double
        DlinaRezaniya = RaschotDlinyNeskolkikhOkruzhnostey(Dobnizheniya, Dotv, .Dfrezy, Lobnizheniya, .AP_CHERN, .ap_chist)
    
        Dim tOtv As Double
        tOtv = DlinaRezaniya / .Podacha
    
    End With
    
    With RaschotObnizheniya
        .tMekhanika = tOtv
        .tSlesar = ZachistkaZausencevSOtverstiiVruchnuyu(1, Dobnizheniya, 2, 1, False)
        .tKontrol = KontrolOtverstiy(Dobnizheniya, Lobnizheniya, 11, True)
    End With
    
End Function
 
Private Function RaschotDlinyNeskolkikhOkruzhnostey(Dotv As Double, Dvnutr As Double, Dfrezy As Double, Lotv As Double, AP_CHERN As Double, ap_chist As Double) As Double
    
    Const KOEFFITSIENT_PEREKRYTIYA As Double = 0.5
    
    Dim RadialnyyShag As Double
    RadialnyyShag = Dfrezy * KOEFFITSIENT_PEREKRYTIYA
    
    Dim DlinaSloya As Double
    DlinaSloya = RaschotDlinySloya(Dotv, Dvnutr, Dfrezy, RadialnyyShag)
    
    Dim KolVoVrezaniy As Long
    KolVoVrezaniy = 1 + OKRUGLVVERKH(Lotv / AP_CHERN) + OKRUGLVVERKH(Lotv / ap_chist)

    RaschotDlinyNeskolkikhOkruzhnostey = Round(DlinaSloya * KolVoVrezaniy + Lotv, 2)
     
End Function

Private Function RaschotDlinySloya(Dotv As Double, Dvnutr As Double, Dfrezy As Double, RadialnyyShag As Double) As Double
    Dim d2 As Double
    d2 = IIf(Dvnutr > Dfrezy, Dvnutr, Dfrezy)
    RaschotDlinySloya = (PI * (Dotv ^ 2 - d2 ^ 2)) / (4 * RadialnyyShag)
    RaschotDlinySloya = Round(RaschotDlinySloya, 2)
End Function

Function Datron_ObrabotkaRezbovykhOtverstiy(Material As EnumMaterialy, Diametr As Double, Lotv As Double, Lrezby As Double, Glukhoye As Boolean, _
    Optional Dobnizheniya As Double, Optional Lobnizheniya As Double) As Variant
    
    If PolozhitelnyyeChisla(Diametr, Lotv, Lrezby) = False Or Lrezby > Lotv Then Exit Function
       
    Dim ShagRezby As Double
    ShagRezby = KrupShagRezb(Diametr)

    Dim NormyOtv As Variant
    NormyOtv = Datron_ObrabotkaOtverstiy(Material, Diametr, Lotv, Glukhoye, ShagRezby, Dobnizheniya, Lobnizheniya)
    
    Dim tTshtRezb As Double
    tTshtRezb = Datron_NarezaniyeRezby(Material, Diametr, Lrezby, ShagRezby)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = NormyOtv(1) + tTshtRezb
    Normy(2) = ProgonkaRezby(ALUMINIYEVYYE_SPLAVY, Diametr, Lrezby) + NormyOtv(2)
    Normy(3) = Kontrol_RezbovayaProbka(Diametr, ShagRezby, 7)
    
    Datron_ObrabotkaRezbovykhOtverstiy = Normy
    
End Function

Public Function Datron_NarezaniyeRezby(Material As EnumMaterialy, Diametr As Double, Lrezby As Double, ShagRezby As Double) As Double

    Dim DlinaRezaniya As Double
    DlinaRezaniya = RaschotDlinaRezaniya_NarezRezby(Diametr, Lrezby, ShagRezby)
    
    If Not DlinaRezaniya > 0 Then Exit Function
    
    Dim Fchern As Double, Fchist As Double
    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, POLIAMID
            Fchern = 250
            Fchist = 200
        Case ORGSTEKLO
            Fchern = 150
            Fchist = 120
        Case TEKSTOLIT, STEKLOTEKSTOLIT
            Fchern = 100
            Fchist = 80
        Case Else
            Exit Function
    End Select

    Datron_NarezaniyeRezby = RaschotTsht_Datron(DlinaRezaniya / Fchern + DlinaRezaniya / Fchist)
    
End Function

Private Function RaschotDlinaRezaniya_NarezRezby(Diametr As Double, Lrezby As Double, ShagRezby As Double) As Double

    Dim KolVoVitkov As Long
    KolVoVitkov = OKRUGLVVERKH(2 + Lrezby / ShagRezby)
    
    Dim Dfrezy As Double
    Select Case Diametr
        Case Is <= 1.2: Dfrezy = 0.8
        Case 1.4:       Dfrezy = 0.95
        Case Is <= 2.5: Dfrezy = 1.4
        Case Is <= 4:   Dfrezy = 2
        Case Is <= 10:  Dfrezy = 4
        Case Else:      Dfrezy = 8
    End Select
    
    RaschotDlinaRezaniya_NarezRezby = KolVoVitkov * PI * (Diametr - Dfrezy)
    
End Function

Private Function RaschotTmash_FrezFaski(Material As EnumMaterialy, Dotv As Double, VysotaFaski As Double) As Double

    If PolozhitelnyyeChisla(Dotv, VysotaFaski) = False Then Exit Function
    
    Const GLUBINA_REZANIYA As Double = 0.3
    
    Dim KolVoProkhodov As Long
    KolVoProkhodov = OKRUGLVVERKH(VysotaFaski / GLUBINA_REZANIYA)
    
    Dim DlinaRezaniya As Double
    DlinaRezaniya = KolVoProkhodov * PI * (Dotv + 2 * VysotaFaski)
    
    If DlinaRezaniya <= 0 Then Exit Function

    Dim Podacha As Double
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Podacha = 4000
    ElseIf Material = POLIAMID Then
        Podacha = 1900
    ElseIf Material = ORGSTEKLO Then
        Podacha = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Podacha = 400
    End If
    
    RaschotTmash_FrezFaski = DlinaRezaniya / Podacha

End Function


