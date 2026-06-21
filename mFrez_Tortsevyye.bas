Attribute VB_Name = "mFrez_Tortsevyye"
'@Folder Frezerovaniye

Option Explicit

Private Const MAX_SHIRINA_TORTS As Integer = 90

Function FrezTortsovymyFrezami_Kompleks(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, Pripusk As Double, _
    Shirina As Double, Dlina As Double, Optional IT As Integer = 11, Optional Ra As Double = 10, Optional NaUdar As Boolean = False, _
    Optional Korka As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Variant
     
    If Not Shirina > 0 Or Not Dlina > 0 Then Exit Function
           
    If IT = 0 Then IT = 11
    Ra = MinRa_RastFrez(Ra, IT)

    If Shirina = 0 Then Shirina = 6
    If Pripusk = 0 Then Pripusk = MIN_PRIPUSK_RASTFREZ
    
    Dim Bfrez As Double
    Bfrez = ShirinaFrezerovaniya_TortsovyeFrezy(Shirina)
    
    Dim h_Chist As Double, t_chist As Double
    If Ra <= 10 Then
    
        If Ra = 5 Then
            h_Chist = MIN_PRIPUSK_RASTFREZ
        Else
            h_Chist = 2
        End If
        
        t_chist = FrezTortsovymyFrezami_Chistovoye(TipStanka, Material, Bfrez, Dlina, Ra, NaUdar, TonkostennayaKonfiguraciya) * IIf(IT <= 11, 2, 1)

    Else
        t_chist = 0
        h_Chist = 0
    End If
    
    Dim h_Chern As Double, h_ChernKork As Double, i_chern As Long
    Dim t_ChernKork As Double, t_chern As Double
    Dim GlubinaRez As Double

    h_Chern = Pripusk - h_Chist
    If Korka And Not h_Chern > 0 Then h_Chern = MIN_PRIPUSK_RASTFREZ
    
    If h_Chern > 0 Then
        GlubinaRez = GlubinaRez_TortsovyyeFrezy(h_Chern)
        If Korka Then
            h_ChernKork = GlubinaRez
            t_ChernKork = FrezTortsovymyFrezami_Chernovoye(TipStanka, CInt(Material), Bfrez, Dlina, GlubinaRez, NaUdar, Korka, TonkostennayaKonfiguraciya)
        Else
            h_ChernKork = 0
            t_ChernKork = 0
        End If
        h_Chern = h_Chern - h_ChernKork
        If h_Chern > 0 Then
            i_chern = WorksheetFunction.RoundUp(h_Chern / GlubinaRez, 0)
            t_chern = FrezTortsovymyFrezami_Chernovoye(TipStanka, CInt(Material), Bfrez, Dlina, GlubinaRez, NaUdar, False, TonkostennayaKonfiguraciya) * i_chern
        End If
    Else
        t_ChernKork = 0
        t_chern = 0
    End If
    
    Dim iB As Integer:      iB = WorksheetFunction.RoundUp(Shirina / Bfrez, 0)
    
    Dim tFrez As Double:    tFrez = (t_ChernKork + t_chern + t_chist) * iB
    Dim tSles As Double:    tSles = ZachistkaZausencev_PoKonturu_Napilnikom(Material, 2 * (Shirina + Dlina), False, False)
    
    'tKontrol делим на 2 т.к. фрезеру€ две поверхности мы проводим только один размер
    Dim tKontrol As Double: tKontrol = IzmereniyePosleFrezerovaniya(Shirina, Dlina, Pripusk, IT) / 2

    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezTortsovymyFrezami_Kompleks = Normy
    
End Function

Private Function FrezTortsovymyFrezami_Chernovoye(ByVal TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ShirinaFrezerovaniya As Double, Dlina As Double, GlubinaRez As Double, _
    Optional FrezNaUdar As Boolean = False, Optional FrezPoKorke As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Double
    '‘резерование плоскостей торцевыми фрезами (карты 22, 23)
    
    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        ShirinaFrezerovaniya > MAX_SHIRINA_TORTS Or GlubinaRez > MAX_GLUBINA_REZANIYA_KONTS_TORTS Then Exit Function
    
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)
    
    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, True)
    
    Dim kMaterial As Double
    If Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.25
    Else
        kMaterial = 1
    End If
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_tortsovyye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaFrezerovaniya And Normativ(i, 4) = 20 And Normativ(i, 5) >= GlubinaRez Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, FrezNaUdar, TonkostennayaKonfiguraciya) * IIf(FrezPoKorke, K_KORKA_FREZ_PLOSK, 1)
    
    FrezTortsovymyFrezami_Chernovoye = kMaterial * PoprKoef * (a * Dlina + b)
    
End Function

Private Function FrezTortsovymyFrezami_Chistovoye(ByVal TipStanka As EnumTipStankaRastFrez, ByVal Material As EnumMaterialy, ShirinaFrezerovaniya As Double, Dlina As Double, _
    Optional Ra As Double = 10, Optional FrezNaUdar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Double
    '‘резерование плоскостей торцевыми фрезами (карты 22, 23)
    
    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        ShirinaFrezerovaniya > MAX_SHIRINA_TORTS Or Ra < MIN_RA_FREZ Then Exit Function
    
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)
    
    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, True)

    Dim kMaterial As Double
    If Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.2
    Else
        kMaterial = 1
    End If
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_tortsovyye")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaFrezerovaniya And Normativ(i, 4) <= Ra Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, FrezNaUdar, TonkostennayaKonfiguraciya)
    
    FrezTortsovymyFrezami_Chistovoye = kMaterial * PoprKoef * (a * Dlina + b)
    
    
End Function

Private Function ShirinaFrezerovaniya_TortsovyeFrezy(ShirinaPoverkhnosti As Double) As Integer
    Select Case ShirinaPoverkhnosti
        Case Is <= 60: ShirinaFrezerovaniya_TortsovyeFrezy = 60
        Case Is <= 90: ShirinaFrezerovaniya_TortsovyeFrezy = 90
        Case Else:
            If WorksheetFunction.RoundUp(ShirinaPoverkhnosti / 60, 0) = _
                WorksheetFunction.RoundUp(ShirinaPoverkhnosti / 90, 0) And _
                ShirinaPoverkhnosti / 60 <> Int(ShirinaPoverkhnosti / 60) Then
                ShirinaFrezerovaniya_TortsovyeFrezy = 60
            Else
                ShirinaFrezerovaniya_TortsovyeFrezy = 90
            End If
    End Select
End Function

Private Function GlubinaRez_TortsovyyeFrezy(Pripusk As Double)

    Select Case Pripusk
        Case Is <= 3:   GlubinaRez_TortsovyyeFrezy = 3
        Case Is <= 5:   GlubinaRez_TortsovyyeFrezy = 5
        Case Else:
            If WorksheetFunction.RoundUp(Pripusk / 3, 0) = _
                WorksheetFunction.RoundUp(Pripusk / 5, 0) And _
                Pripusk / 7 <> Int(Pripusk / 3) Then
                GlubinaRez_TortsovyyeFrezy = 3
            Else
                GlubinaRez_TortsovyyeFrezy = 5
            End If
    End Select

End Function


