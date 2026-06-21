Attribute VB_Name = "mFrez_Shlits"
'@Folder Frezerovaniye

Option Explicit

Private Const MAX_SHIRINA_SHLITS As Integer = 4
Private Const MAX_GLUBINA_SHLITS As Integer = 15
Private Const MIN_RA_SHLITS As Double = 3.2

Function FrezShlitsevymiFrezami_Kompleks(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal Shirina As Double, Dlina As Double, Glubina As Double, Ra As Double, _
    Optional FrezNaUdar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Variant
    'Фрезерование канавок, пазов шлицевыми фрезами (Карты 46, 47)

    If Not Shirina > 0 Or Not Dlina > 0 Or Not Glubina > 0 Then Exit Function
    
    Dim tFrez As Double
    tFrez = FrezShlitsevymiFrezami(TipStanka, Material, Shirina, Dlina, Glubina, Ra, FrezNaUdar, TonkostennayaKonfiguraciya)
    
    Dim tSles As Double, DlinaZachistki As Double
    DlinaZachistki = 2 * Shirina + 4 * Glubina + 2 * Dlina
    tSles = ZachistkaZausencevPosleFrezPazovShlicev(CInt(Material), DlinaZachistki, 1)
    
    Dim tKontrol As Double
    tKontrol = IzmereniyeShlitsevKalibrProbkoy(Shirina, Dlina)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezShlitsevymiFrezami_Kompleks = Normy
    
End Function

Private Function FrezShlitsevymiFrezami(TipStanka As EnumTipStankaRastFrez, ByVal Material As EnumMaterialy, ByVal Shirina As Double, Dlina As Double, Glubina As Double, Ra As Double, _
    Optional NaUdar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False) As Double
    'Фрезерование канавок, пазов шлицевыми фрезами (Карты 46, 47)

    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        Shirina > MAX_SHIRINA_SHLITS Or Glubina > MAX_GLUBINA_SHLITS Or Ra < MIN_RA_SHLITS Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("frez_shlitsevyye")
    
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)
    
    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)

    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) <= Shirina And Normativ(i, 4) <= Ra And Normativ(i, 5) >= Glubina Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next
    
    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, NaUdar, TonkostennayaKonfiguraciya)
    
    FrezShlitsevymiFrezami = PoprKoef * (a * Dlina + b)

End Function


