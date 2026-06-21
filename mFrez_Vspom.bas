Attribute VB_Name = "mFrez_Vspom"
'@Folder Frezerovaniye

Option Explicit
Option Private Module

Public Const MAX_GLUBINA_REZANIYA_KONTS_TORTS As Integer = 10
Public Const MIN_RA_FREZ As Double = 1.25
Public Const MIN_PRIPUSK_RASTFREZ As Integer = 1
Public Const MAX_SHIRINA_KONTS_USTUP_PAZ As Integer = 40
Public Const MAKS_D_SVERL_FREZ_STANOK As Integer = 30
Public Const MAKS_D_REZBY_FREZERNYY As Integer = 20
Public Const MIN_D_FREZ As Integer = 2

Public Const K_KORKA_FREZ_PLOSK As Double = 1.3

Public Enum EnumTipStankaRastFrez
    GorizontalnoRastochnoy = 1
    VertikalnoGorizontalnoFrezernyy = 2
End Enum

Function MaterialKorrektnyy_RastFrez(Material As EnumMaterialy) As Boolean

    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, STAL_UGLERODISTAYA, STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY, MEDNYYE_SPLAVY
            MaterialKorrektnyy_RastFrez = True
        Case Else
            MaterialKorrektnyy_RastFrez = False
    End Select

End Function

Function MinRa_RastFrez(Ra As Double, IT As Integer) As Double
    
    If Ra > 10 And IT <= 11 Or Ra = 0 Then
        MinRa_RastFrez = 10
    Else
        MinRa_RastFrez = Ra
    End If
    
End Function

Function IzmereniyePosleFrezerovaniya(Shirina As Double, Dlina As Double, Pripusk As Double, IT As Integer)
 
    Dim tKontrol As Double
    Select Case IT
        Case Is <= 7: tKontrol = IzmereniyeMikrometrom(Shirina, Dlina, False) + IzmereniyeMikrometrom(Pripusk, Dlina, False)
        Case Else: tKontrol = IzmerShtangenCircul(Shirina, Dlina, IT < 11) + IzmerShtangenCircul(Pripusk, Dlina, IT < 11)
    End Select
    
    IzmereniyePosleFrezerovaniya = tKontrol
    
End Function

Function IzmenitTipStanka(TipStanka As EnumTipStankaRastFrez, ByVal Material As EnumMaterialy) As EnumTipStankaRastFrez
    
    IzmenitTipStanka = TipStanka
    If TipStanka = GorizontalnoRastochnoy Then
        If Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
            IzmenitTipStanka = VertikalnoGorizontalnoFrezernyy
        End If
    End If
    
End Function

Public Function RaschotObshchegoKoeffitsienta_RastFrez(Material As EnumMaterialy, TipStanka As EnumTipStankaRastFrez, FrezNaUdar As Boolean, TonkostennayaKonfiguraciya As Boolean)
    
    Dim kStanok As Double, kUdar As Double, kKorka As Double, kTonkosten As Double
    
    kStanok = 1
    If TipStanka = GorizontalnoRastochnoy Then
        If Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
            kStanok = 1.2
        End If
    End If

    kUdar = IIf(FrezNaUdar = True, 1.2, 1)
    kTonkosten = IIf(TonkostennayaKonfiguraciya = True, 1.3, 1)

    RaschotObshchegoKoeffitsienta_RastFrez = WorksheetFunction.Product(kStanok, kUdar, kTonkosten)
                                                    
End Function

Public Function IzmenitMaterialDlyaZaprosa_RastFrez(Material As EnumMaterialy, Tortevoye As Boolean) As EnumMaterialy
    
    Dim newMaterial As EnumMaterialy
    
    If Material = MEDNYYE_SPLAVY Then
        newMaterial = ALUMINIYEVYYE_SPLAVY
    ElseIf Material = TITANOVYYE_SPLAVY And Tortevoye = True Then
        newMaterial = STAL_NERZHAVEYUSHCHAYA
    Else
        newMaterial = Material
    End If
    
    IzmenitMaterialDlyaZaprosa_RastFrez = newMaterial
    
End Function
