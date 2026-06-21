Attribute VB_Name = "mFrezPazy"
'@Folder Frezerovaniye

Option Explicit

Function FrezPazov_Kompleks(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal DlinaPaza As Double, ByVal ShirinaPaza As Double, GlubinaPaza As Double, _
    Optional ByVal IT As Integer = 12, Optional ByVal Ra As Double = 10, Optional Gluhoy As Boolean = False, _
    Optional Udar As Boolean = False, Optional Korka As Boolean = False, Optional Tonkosten As Boolean = False, Optional ZachistkaBokovykhStoron As Boolean = True)
    
    If DlinaPaza <= 0 Or ShirinaPaza <= 0 Or GlubinaPaza <= 0 Then Exit Function
    
    Dim tFrez As Double
    If Not ZachistkaBokovykhStoron Then
        tFrez = FrezPazovBezZachistkiBokovykhStoron(TipStanka, Material, DlinaPaza, ShirinaPaza, GlubinaPaza, Udar, Korka, Tonkosten)
    
    Else
        tFrez = FrezPazovSZachistkoyBokovykhStoron(TipStanka, Material, DlinaPaza, ShirinaPaza, GlubinaPaza, IT, Ra, Gluhoy, Udar, Korka, Tonkosten)
    End If
    
    Dim tSles As Double, lZachistki As Double
    If Gluhoy Then
        lZachistki = 4 * GlubinaPaza + 2 * ShirinaPaza + 2 * DlinaPaza
    Else
        lZachistki = 4 * (DlinaPaza + ShirinaPaza)
    End If
    tSles = ZachistkaZausencev_PoKonturu_Napilnikom(Material, lZachistki, IT < 9, False)
    
    Dim tKontrol As Double
    tKontrol = IzmerShtangenCircul_CPL(DlinaPaza, ShirinaPaza, IT < 11) + IIf(Gluhoy, IzmerShtangenCircul(GlubinaPaza, DlinaPaza, IT < 11), 0)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezPazov_Kompleks = Normy

End Function

Private Function FrezPazovBezZachistkiBokovykhStoron(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal DlinaPaza As Double, ByVal ShirinaPaza As Double, GlubinaPaza As Double, _
    Optional Udar As Boolean = False, Optional Korka As Boolean = False, Optional Tonkosten As Boolean = False) As Double
    
    If ShirinaPaza > MAX_SHIRINA_KONTS_USTUP_PAZ Or GlubinaPaza <= 0 Then Exit Function
    
    Dim i_chern As Long
    i_chern = WorksheetFunction.RoundUp(GlubinaPaza / MaksGlubinaRezaniya_Paz(ShirinaPaza), 0)
    
    Dim GlubinaRez As Double
    GlubinaRez = GlubinaPaza / i_chern
        
    Dim tFrez As Double
    If Korka Then
        tFrez = FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza, GlubinaRez, Udar, Korka, Tonkosten)
        tFrez = tFrez + (i_chern - 1) * FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza, GlubinaRez, Udar, False, Tonkosten)
    Else
        tFrez = i_chern * FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza, GlubinaRez, Udar, False, Tonkosten)
    End If
    
    FrezPazovBezZachistkiBokovykhStoron = tFrez

End Function

Private Function FrezPazovSZachistkoyBokovykhStoron(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal DlinaPaza As Double, ByVal ShirinaPaza As Double, GlubinaPaza As Double, _
    Optional ByVal IT As Integer = 12, Optional ByVal Ra As Double = 10, Optional Gluhoy As Boolean = False, _
    Optional Udar As Boolean = False, Optional Korka As Boolean = False, Optional Tonkosten As Boolean = False, _
    Optional PoRezhimam As Boolean = False, Optional CHPU As Boolean = False) As Double
    
    If Not DlinaPaza > 0 Or Not ShirinaPaza > 0 Then Exit Function
    
    Dim tFrez As Double
    
    Dim ShirinaPaza_Chern As Double
    If ShirinaPaza > MAX_SHIRINA_KONTS_USTUP_PAZ Then
        ShirinaPaza_Chern = MAX_SHIRINA_KONTS_USTUP_PAZ
    Else
        ShirinaPaza_Chern = ShirinaPaza
    End If

    Dim i_chern As Long
    i_chern = WorksheetFunction.RoundUp(GlubinaPaza / MaksGlubinaRezaniya_Paz(ShirinaPaza_Chern), 0)
    
    Dim GlubinaRez As Double
    GlubinaRez = GlubinaPaza / i_chern
        
    Dim t_chern As Double
    If Korka Then
        t_chern = FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza_Chern, GlubinaRez, Udar, Korka, Tonkosten)
        t_chern = t_chern + (i_chern - 1) * FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza_Chern, GlubinaRez, Udar, False, Tonkosten)
    Else
        t_chern = i_chern * FrezPazov_Chernovoye(TipStanka, Material, DlinaPaza, ShirinaPaza_Chern, GlubinaRez, Udar, False, Tonkosten)
    End If
    
    Dim Pripusk As Double
    Pripusk = WorksheetFunction.Max(ShirinaPaza - ShirinaPaza_Chern, MIN_PRIPUSK_RASTFREZ)
    
    Dim t_chist
    If Gluhoy Then
        t_chist = FrezUstupov_Kompleks(TipStanka, Material, DlinaPaza, GlubinaPaza, Pripusk, IT, Ra, Udar, False, Tonkosten)
    Else
    
        Dim ctx As New clsMillingContext

        With ctx
            .TipStanka = TipStanka
            .Material = Material
            .DlinaPoverkhnosti = DlinaPaza
            .ShirinaPoverkhnosti = GlubinaPaza
            
            ' Даже если здесь передашь 0, класс сам запишет MIN_PRIPUSK_RASTFREZ
            .Pripusk = Pripusk
            .IT = IT
            .Ra = Ra
            
            .NaUdar = Udar
            .Korka = Korka
            .TonkostennayaKonfiguraciya = Tonkosten
            
            .CHPU = CHPU
            .PoRezhimam = PoRezhimam
            .AvtoRezhimy = True

        End With
        
        t_chist = FrezKontsevymiFrezami_Kompleks(ctx)(1)
        
    End If

    tFrez = (t_chern + 2 * IIf(IT <= 11, 2, 1) * t_chist(1))
    
    FrezPazovSZachistkoyBokovykhStoron = tFrez
    
End Function

Private Function FrezPazov_Chernovoye(ByVal TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal DlinaPaza As Double, ShirinaPaza As Double, GlubinaRez As Double, _
    Optional FrezNaUdar As Boolean = False, Optional FrezPoKorke As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Double
    'Фрезерование сквозных окон, пазов концевыми и шпоночными фрезами (карты 32, 33)
    
    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        ShirinaPaza > MAX_SHIRINA_KONTS_USTUP_PAZ Or GlubinaRez > MaksGlubinaRezaniya_Paz(ShirinaPaza) Then Exit Function
        
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)

    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_pazy")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaPaza And Normativ(i, 4) = 20 And Normativ(i, 5) >= GlubinaRez Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, FrezNaUdar, TonkostennayaKonfiguraciya) * IIf(FrezPoKorke, K_KORKA_FREZ_PLOSK, 1)
    
    FrezPazov_Chernovoye = PoprKoef * (a * DlinaPaza + b)

End Function

Private Function MaksGlubinaRezaniya_Paz(ShirinaPaza As Double) As Double

    Select Case ShirinaPaza
        Case Is <= 2:                               MaksGlubinaRezaniya_Paz = 3
        Case Is <= 5:                               MaksGlubinaRezaniya_Paz = 5
        Case Is <= 8:                               MaksGlubinaRezaniya_Paz = 7
        Case Is <= 12:                              MaksGlubinaRezaniya_Paz = 10
        Case Is <= 16:                              MaksGlubinaRezaniya_Paz = 15
        Case Is <= 20:                              MaksGlubinaRezaniya_Paz = 20
        Case Is <= MAX_SHIRINA_KONTS_USTUP_PAZ:     MaksGlubinaRezaniya_Paz = 30
    End Select
    
End Function
