Attribute VB_Name = "mRaschotPloshchadiMassy"
Option Explicit
Public Const PlotnostStali As Integer = 7920

Private Function RaschotPloshchadi(Figura As ITechnologicalShape, Dlina As Double) As Double

    If Dlina > 0 Then
        RaschotPloshchadi = Figura.PloshchadPolnaya
    Else
        RaschotPloshchadi = Figura.PloshchadPoperSech
    End If
    
End Function

Function MassaLista(Dlina As Double, Shirina As Double, Tolshchina As Double, Material As EnumMaterialy) As Double

    Dim List As New clsFigura_List
    List.Init Material, Shirina, Tolshchina, Dlina
    MassaLista = List.Massa
    
End Function

Function MassaKruga(d As Double, Dlina As Double, Material As EnumMaterialy) As Double

    Dim Krug As New clsFigura_Krug
    Krug.Init Material, d, Dlina
    MassaKruga = Krug.Massa
    
End Function

Function MassaShestigrannika(Material As EnumMaterialy, S As Double, Dlina As Double) As Double

    Dim Shestigrannik As New clsFigura_Shestigrnnik
    Shestigrannik.Init Material, S, Dlina
    MassaShestigrannika = Shestigrannik.Massa
    
End Function

Function MassaUgolnika(Material As EnumMaterialy, Shirina As Double, Vysota As Double, TolshchinaStenki_a As Double, TolshchinaStenki_b As Double, Dlina As Double) As Double
    
    Dim Ugolnik As New clsFigura_Ugolnik
    Ugolnik.Init Material, Shirina, Vysota, TolshchinaStenki_a, TolshchinaStenki_b, Dlina
    MassaUgolnika = Ugolnik.Massa
    
End Function

Function MassaPryamougolnoyTruby(Shirina As Double, Vysota As Double, TolshchinaStenki As Double, Dlina As Double) As Double

    Dim Truba As New clsFigura_PryamougolnayaTruba
    Truba.Init ALUMINIYEVYYE_SPLAVY, Shirina, Vysota, TolshchinaStenki, Dlina
    MassaPryamougolnoyTruby = Truba.Massa
    
End Function

Function MassaShvellera_U(NumShveller As String, Dlina As Double) As Double

    Dim Shveller As New clsFigura_Shveller_U
    Shveller.Init ALUMINIYEVYYE_SPLAVY, NumShveller, Dlina
    MassaShvellera_U = Shveller.Massa

End Function

Function MassaShvellera_P(Shirina As Double, Vysota As Double, TolshchinaStenki As Double, Dlina As Double) As Double

    Dim Shveller As New clsFigura_Shveller_P
    Shveller.Init ALUMINIYEVYYE_SPLAVY, Shirina, Vysota, TolshchinaStenki, Dlina
    MassaShvellera_P = Shveller.Massa

End Function

Function MassaTruby(Diametr As Double, TolshchinaStenki As Double, Dlina As Double, Material As EnumMaterialy) As Double

    Dim Truba As New clsFigura_Truba
    Truba.Init Material, Diametr, TolshchinaStenki, Dlina
    MassaTruby = Truba.Massa
    
End Function

Function PloshchadShvellera(Nomer As String, Dlina As Double) As Double
    
    Dim Shveller As New clsFigura_Shveller_U
    Shveller.Init ALUMINIYEVYYE_SPLAVY, Nomer, Dlina
    PloshchadShvellera = RaschotPloshchadi(Shveller, Dlina)
    
End Function

Function PloshchadUgolnika(Shirina As Double, Vysota As Double, TolshchinaStenki_a As Double, TolshchinaStenki_b As Double, Optional Dlina As Double = 0) As Double
    
    Dim Ugolnik As New clsFigura_Ugolnik
    Ugolnik.Init ALUMINIYEVYYE_SPLAVY, Shirina, Vysota, TolshchinaStenki_a, TolshchinaStenki_b, Dlina
    PloshchadUgolnika = RaschotPloshchadi(Ugolnik, Dlina)
    
End Function

Function PloshchadKruga(d As Double, Optional Dlina As Double = 0) As Double

    Dim Krug As New clsFigura_Krug
    Krug.Init ALUMINIYEVYYE_SPLAVY, d, Dlina
    PloshchadKruga = RaschotPloshchadi(Krug, Dlina)

End Function
Function PloshchadPramougolnika(Shirina As Double, Vysota As Double, Optional Dlina As Double = 0) As Double

    Dim List As New clsFigura_List
    List.Init ALUMINIYEVYYE_SPLAVY, Shirina, Vysota, Dlina
    PloshchadPramougolnika = RaschotPloshchadi(List, Dlina)
    
End Function

Function PloshchadTruby(d As Double, S As Double, Optional Dlina As Double = 0) As Double
    
    Dim Truba As New clsFigura_Truba
    Truba.Init ALUMINIYEVYYE_SPLAVY, d, S, Dlina
    PloshchadTruby = RaschotPloshchadi(Truba, Dlina)
 
End Function

Function PloshchadPryamougolnoyTruby(Shirina As Double, Vysota As Double, TolshchinaStenki As Double, Optional Dlina As Double = 0) As Double

    Dim Truba As New clsFigura_PryamougolnayaTruba
    Truba.Init ALUMINIYEVYYE_SPLAVY, Shirina, Vysota, TolshchinaStenki, Dlina
    PloshchadPryamougolnoyTruby = RaschotPloshchadi(Truba, Dlina)
 
End Function

Function PloshchadShestigrannika(S As Double, Optional Dlina As Double = 0) As Double

    Dim Shestigrannik As New clsFigura_Shestigrnnik
    Shestigrannik.Init ALUMINIYEVYYE_SPLAVY, S, Dlina
    PloshchadShestigrannika = RaschotPloshchadi(Shestigrannik, Dlina)
    
End Function

Function DiametrShestigrannika(S As Double) As Double

    Dim Shestigrannik As New clsFigura_Shestigrnnik
    Shestigrannik.Init ALUMINIYEVYYE_SPLAVY, S, 0
    DiametrShestigrannika = Shestigrannik.DiametrOpisannoyOkruzhnosti
    
End Function

Function DiametrKvadrata(a As Double, b As Double) As Double
    DiametrKvadrata = Sqr(a ^ 2 + b ^ 2)
End Function

Function DlinaLiski(d As Double, razmer As Double)

    Dim R As Double
    R = d / 2
    
    Dim H As Double
    H = razmer / 2
    H = R - H
    
    DlinaLiski = Sqr(8 * R * H - 4 * H ^ 2)

End Function

