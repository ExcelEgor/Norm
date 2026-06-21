Attribute VB_Name = "mElektroeroziya"
Option Explicit

Private Type ParameryIzTablitsy
    t10Bolshe As Double
    t10Menshe As Double
    TolshchinaBolshe As Double
    TolshchinaMenshe As Double
    TolshchinaMin As Double
End Type

Function ElektroEroz_Dorabotka(Material As Integer, Tolshchina As Double, DlinaRezaniya As Double, Optional TipKonfiguratsii As Integer = 1) As Double
    
    If Tolshchina > ElektroEroz_Dorabotka_Smax(Material, TipKonfiguratsii) Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("elektroeroz_dorabotka")

    Dim i As Integer, NachalnayaStroka As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) = TipKonfiguratsii Then
            NachalnayaStroka = i
            Exit For
        End If
    Next
    
    Dim Parametry As ParameryIzTablitsy
    Parametry = VzyatParametryIzTablitsy(Normativ, NachalnayaStroka, Tolshchina, 3, 4)
    
    ElektroEroz_Dorabotka = RaschotVremeni(DlinaRezaniya, Tolshchina, Parametry)

End Function

Public Function ElektroEroz_Dorabotka_Smax(Material As Integer, Optional Konfig As Integer = 1) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("elektroeroz_dorabotka_s_max")

    Dim i As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) = Konfig Then
            ElektroEroz_Dorabotka_Smax = Normativ(i, 3)
            Exit For
        End If
    Next

End Function

Function ElektroEroz_Kontur(Material As Integer, IT As Integer, Tolshchina As Double, DlinaRezaniya As Double, Optional TipKonfiguratsii As Integer = 1) As Double
    
    If Tolshchina > ElektroEroz_Kontur_Smax(Material, IT, TipKonfiguratsii) Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("elektroeroz_kontur")

    Dim i As Integer, NachalnayaStroka As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) <= IT And Normativ(i, 3) = TipKonfiguratsii Then
            NachalnayaStroka = i
            Exit For
        End If
    Next
    
    If NachalnayaStroka = 0 Then Exit Function
    
    Dim Parametry As ParameryIzTablitsy
    Parametry = VzyatParametryIzTablitsy(Normativ, NachalnayaStroka, Tolshchina, 4, 5)
    
    ElektroEroz_Kontur = RaschotVremeni(DlinaRezaniya, Tolshchina, Parametry)
    
End Function

Function ElektroEroz_Kontur_Smax(Material As Integer, IT As Integer, Optional Konfig As Integer = 1) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("elektroeroz_kontur_s_max")

    Dim i As Integer
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) <= IT And Normativ(i, 3) = Konfig Then
            ElektroEroz_Kontur_Smax = Normativ(i, 4)
            Exit For
        End If
    Next
    
End Function

Private Function VzyatParametryIzTablitsy(Normy As Variant, NachalnayaStroka As Integer, _
    Tolshchina As Double, StolbetsTolshchina As Integer, StolbetsVremya As Integer) As ParameryIzTablitsy

    Dim i As Integer
    With VzyatParametryIzTablitsy
        .TolshchinaMin = Normy(NachalnayaStroka, StolbetsTolshchina)
        For i = NachalnayaStroka To UBound(Normy)
            If Normy(i, StolbetsTolshchina) >= Tolshchina Then
                .TolshchinaBolshe = Normy(i, StolbetsTolshchina)
                .t10Bolshe = Normy(i, StolbetsVremya)
                If Tolshchina > .TolshchinaMin Then
                    .TolshchinaMenshe = Normy(i - 1, StolbetsTolshchina)
                    .t10Menshe = Normy(i - 1, StolbetsVremya)
                End If
                Exit For
            End If
        Next
    End With
    
End Function

Function ElektroEroz_Ustanov(Sloznost As Integer) As Double
    Select Case Sloznost
        Case 1: ElektroEroz_Ustanov = 5
        Case 2: ElektroEroz_Ustanov = 10
        Case 3: ElektroEroz_Ustanov = 15
    End Select
End Function

Private Function RaschotVremeni(Dlina As Double, Tolshchina As Double, Parametry As ParameryIzTablitsy) As Double
    
    Dim t10 As Double
    With Parametry
        If Tolshchina > .TolshchinaMin Then
            t10 = mPrivate.LineynayaInterpolyatisiya(.t10Bolshe, .t10Menshe, .TolshchinaBolshe, .TolshchinaMenshe, Tolshchina)
        Else
            t10 = .t10Bolshe
        End If
    End With
    
    RaschotVremeni = ((Dlina + 1) / 10) * t10
    
End Function

