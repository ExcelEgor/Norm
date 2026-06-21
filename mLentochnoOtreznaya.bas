Attribute VB_Name = "mLentochnoOtreznaya"
Option Explicit

' Единый внутренний диспетчер расчетов. За счет Static объект класса создается
' всего один раз в памяти при первом вычислении, а не пересоздается на каждой ячейке.
Private Function RaschotFigury(ByVal Figura As ITechnologicalShape) As Double
    Static Raschot As clsLentOtrez_Raschot
    If Raschot Is Nothing Then Set Raschot = New clsLentOtrez_Raschot
    RaschotFigury = Raschot.tLentOtrez(Figura)
End Function

' =========================================================================
' ПОЛЬЗОВАТЕЛЬСКИЕ ФУНКЦИИ (UDF) ДЛЯ РАБОЧИХ ЛИСТОВ EXCEL
' =========================================================================

Function LentOtrez_Krug(ByVal Material As EnumMaterialy, ByVal d As Double, ByVal Dlina As Double) As Double
    Dim Krug As New clsFigura_Krug
    Krug.Init Material, d, Dlina
    LentOtrez_Krug = RaschotFigury(Krug)
End Function

Function LentOtrez_Truba(ByVal Material As EnumMaterialy, ByVal d As Double, ByVal S As Double, ByVal Dlina As Double) As Double
    If d < S * 2 Then Exit Function

    Dim Truba As New clsFigura_Truba
    Truba.Init Material, d, S, Dlina
    Dim tTruba As Double
    tTruba = RaschotFigury(Truba)
    
    Dim Krug As New clsFigura_Krug
    Krug.Init Material, d, Dlina
    Dim tKrug As Double
    tKrug = RaschotFigury(Krug)

    LentOtrez_Truba = IIf(tTruba > tKrug, tKrug, tTruba)
End Function

Function LentOtrez_Ugolnik(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal a As Double, ByVal b As Double, ByVal Dlina As Double) As Double
    Dim Ugolnik As New clsFigura_Ugolnik
    Ugolnik.Init Material, Shirina, Vysota, a, b, Dlina
    Dim tUglonik As Double
    tUglonik = RaschotFigury(Ugolnik)
    
    Dim List As New clsFigura_List
    List.Init Material, Shirina, Vysota, Dlina
    Dim tList As Double
    tList = RaschotFigury(List)

    LentOtrez_Ugolnik = IIf(tUglonik > tList, tList, tUglonik)
End Function

Function LentOtrez_Shveller_P(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal TolshchinaStenki As Double, ByVal Dlina As Double) As Double
    If Shirina < 2 * TolshchinaStenki Or Vysota < 2 * TolshchinaStenki Then Exit Function
    
    Dim ShvellerP As New clsFigura_Shveller_P
    ShvellerP.Init Material, Shirina, Vysota, TolshchinaStenki, Dlina
    Dim tShvellerP As Double
    tShvellerP = RaschotFigury(ShvellerP)
    
    Dim List As New clsFigura_List
    List.Init Material, Shirina, Vysota, Dlina
    Dim tList As Double
    tList = RaschotFigury(List)
    
    LentOtrez_Shveller_P = IIf(tShvellerP > tList, tList, tShvellerP)
End Function

Function LentOtrez_PryamougolnayaTruba(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal TolshchinaStenki As Double, ByVal Dlina As Double) As Double
    If Shirina < 2 * TolshchinaStenki Or Vysota < 2 * TolshchinaStenki Then Exit Function
    
    Dim Truba As New clsFigura_PryamougolnayaTruba
    Truba.Init Material, Shirina, Vysota, TolshchinaStenki, Dlina
    Dim tTruba As Double
    tTruba = RaschotFigury(Truba)
    
    Dim List As New clsFigura_List
    List.Init Material, Shirina, Vysota, Dlina
    Dim tList As Double
    tList = RaschotFigury(List)

    LentOtrez_PryamougolnayaTruba = IIf(tTruba > tList, tList, tTruba)
End Function

Function LentOtrez_Shveller_U(ByVal Material As EnumMaterialy, ByVal Oboznacheniye As String, ByVal Dlina As Double) As Double
    Dim Shveller As New clsFigura_Shveller_U
    Shveller.Init Material, Oboznacheniye, Dlina
    Dim tShveller As Double
    tShveller = RaschotFigury(Shveller)
    
    Dim List As New clsFigura_List
    List.Init Material, Shveller.Shirina, Shveller.Vysota, Dlina
    Dim tList As Double
    tList = RaschotFigury(List)
    
    LentOtrez_Shveller_U = IIf(tShveller > tList, tList, tShveller)
End Function

Function LentOtrez_Pryamougolnik(ByVal Material As EnumMaterialy, ByVal a As Double, ByVal b As Double, ByVal Dlina As Double) As Double
    Dim List As New clsFigura_List
    List.Init Material, a, b, Dlina
    LentOtrez_Pryamougolnik = RaschotFigury(List)
End Function

Function LentOtrez_Shestigrannik(ByVal Material As EnumMaterialy, ByVal S As Double, ByVal Dlina As Double) As Double
    Dim Shestigrannik As New clsFigura_Shestigrnnik
    Shestigrannik.Init Material, S, Dlina
    LentOtrez_Shestigrannik = RaschotFigury(Shestigrannik)
End Function

' =========================================================================
' МЕТОДЫ СБОРКИ ПОЛНЫХ МАРШРУТОВ ДЛЯ ВНЕШНИХ ВЫЗОВОВ И ФОРМУЛ
' =========================================================================

Function LentOtrez_Krug_Marshrut(ByVal Material As EnumMaterialy, ByVal Diametr As Double, ByVal Dlina As Double) As Variant
    Dim Krug As New clsFigura_Krug
    Krug.Init Material, Diametr, Dlina

    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Krug_Marshrut = Marshrut.RaschotMarshrutaFigury(Krug, False)
End Function

Function LentOtrez_Truba_Marshrut(ByVal Material As EnumMaterialy, ByVal d As Double, ByVal S As Double, ByVal Dlina As Double) As Variant
    Dim Truba As New clsFigura_Truba
    Truba.Init Material, d, S, Dlina
    
    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Truba_Marshrut = Marshrut.RaschotMarshrutaFigury(Truba, False)
End Function

Function LentOtrez_Ugolnik_Marshrut(ByVal Material As EnumMaterialy, ByVal H As Double, ByVal W As Double, ByVal a As Double, ByVal b As Double, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim Ugolnik As New clsFigura_Ugolnik
    Ugolnik.Init Material, W, H, a, b, Dlina
    
    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Ugolnik_Marshrut = Marshrut.RaschotMarshrutaFigury(Ugolnik, Rikhtovka)
End Function

Function LentOtrez_Shveller_P_Marshrut(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal TolshchinaStenki As Double, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim Shveller As New clsFigura_Shveller_P
    Shveller.Init Material, Shirina, Vysota, TolshchinaStenki, Dlina
    
    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Shveller_P_Marshrut = Marshrut.RaschotMarshrutaFigury(Shveller, Rikhtovka)
End Function

Function LentOtrez_PryamougolnayaTruba_Marshrut(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal TolshchinaStenki As Double, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim PryamougolnayaTruba As New clsFigura_PryamougolnayaTruba
    PryamougolnayaTruba.Init Material, Shirina, Vysota, TolshchinaStenki, Dlina
    
    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_PryamougolnayaTruba_Marshrut = Marshrut.RaschotMarshrutaFigury(PryamougolnayaTruba, Rikhtovka)
End Function

Function LentOtrez_Shveller_U_Marshrut(ByVal Material As EnumMaterialy, ByVal OboznacheniyeShvellera As String, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim Shveller As New clsFigura_Shveller_U
    Shveller.Init Material, OboznacheniyeShvellera, Dlina
    
    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Shveller_U_Marshrut = Marshrut.RaschotMarshrutaFigury(Shveller, Rikhtovka)
End Function

Function LentOtrez_Pryamougolnik_Marshrut(ByVal Material As EnumMaterialy, ByVal Shirina As Double, ByVal Vysota As Double, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim Pryamougolnik As New clsFigura_List
    Pryamougolnik.Init Material, Shirina, Vysota, Dlina

    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Pryamougolnik_Marshrut = Marshrut.RaschotMarshrutaFigury(Pryamougolnik, Rikhtovka)
End Function

Function LentOtrez_Shestigrannik_Marshrut(ByVal Material As EnumMaterialy, ByVal S As Double, ByVal Dlina As Double, ByVal Rikhtovka As Boolean) As Variant
    Dim Shestigrannik As New clsFigura_Shestigrnnik
    Shestigrannik.Init Material, S, Dlina

    Dim Marshrut As New clsLentOtrez_Raschot
    LentOtrez_Shestigrannik_Marshrut = Marshrut.RaschotMarshrutaFigury(Shestigrannik, Rikhtovka)
End Function
