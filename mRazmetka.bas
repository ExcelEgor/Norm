Attribute VB_Name = "mRazmetka"
Option Explicit
'ĞÀÇÌÅÒÎ×ÍÛÅ ĞÀÁÎÒÛ. ÍÎĞÌÀÒÈÂÛ ÂĞÅÌÅÍÈ. ÌÅËÊÎÑÅĞÈÉÍÎÅ ÏĞÎÈÇÂÎÄÑÒÂÎ. 1988

Function RazmKontPoShab(Material As Integer, Mass As Double, p As Double, Optional Krivolinein As Boolean = False)
    Dim a As Double, X As Double, Y As Double
    Y = 0.28
    If Material = 1 Then
        X = 0.39
        a = IIf(Mass <= 3, 0.078, 0.0684)
    Else
        X = IIf(Mass <= 3, 0.39, 0.43)
        a = IIf(Mass <= 3, 0.065, 0.057)
    End If
    
    If p > 6000 Then
        Dim t1000 As Double
        t1000 = ((p - 6000) / 1000) * 0.25
        RazmKontPoShab = (a * Mass ^ X * 6000 ^ Y) + 0.25
        
    Else
        RazmKontPoShab = a * Mass ^ X * p ^ Y
    End If
    If Krivolinein Then RazmKontPoShab = RazmKontPoShab * 1.2
    
End Function

Function RazmOtvPoShab(Material As Integer, d As Double, KolVoOtv As Double)
    Dim a As Double, X As Double, Y As Double
    X = 0.7
    Y = 0.78
    a = IIf(Material = 1, 0.0471, 0.0393)
    RazmOtvPoShab = a * d ^ X * KolVoOtv ^ Y
End Function

Function RazmetkaLinii_Lineikoi(DlinaLinii As Double, Optional Material As Integer = 1, Optional RazmetkaMelom As Boolean = False) As Double

    'Material - Ìàòåğèàë:
    '1 - Ñòàëü
    '2 - Àëşìèíèåâûå è ìåäíûå ñïëàâû
    
    If Material < 1 Or Material > 2 Then Exit Function
    
    Dim MinDlinaLinii As Integer, MaksDlinaLinii As Integer
    MinDlinaLinii = 5
    MaksDlinaLinii = 5000
    
    If DlinaLinii > MaksDlinaLinii Then Exit Function
    
    If DlinaLinii < MinDlinaLinii Then DlinaLinii = MinDlinaLinii
    
    Dim MalenikayaDlinaLinii As Integer
    MalenikayaDlinaLinii = 10
    
    Dim a As Double, b As Double
    If DlinaLinii > MalenikayaDlinaLinii Then
        a = 0.00016
        b = 0.08
    Else
        a = 0.0002
        b = 0.1
    End If
    
    Dim K1 As Double, K2 As Double
    K1 = IIf(Material = 2, 0.8, 1)
    K2 = IIf(RazmetkaMelom, 0.9, 1)
    
    RazmetkaLinii_Lineikoi = (a * DlinaLinii + b) * K1 * K2
    
End Function

Function RazmetkaKonturaBezShablona_Lineykoy(Material As EnumMaterialy, ByVal Massa As Double, ByVal Perimetr As Double) As Double
    'ÊÀĞÒÀ 12. ĞÀÇÌÅÒÊÀ ÊÎÍÒÓĞÀ ÁÅÇ ØÀÁËÎÍÀ (ËÈÍÅÉÊÎÉ)
    If Not Perimetr > 0 Then Exit Function

    Dim X As Double, Y As Double, z As Double
    Select Case Material
        Case EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.REZINA, EnumMaterialy.TEKSTOLIT, EnumMaterialy.STEKLOTEKSTOLIT, EnumMaterialy.ORGSTEKLO, EnumMaterialy.GETINAKS, EnumMaterialy.POLIAMID, EnumMaterialy.KARTON
            X = IIf(Massa <= 3, 0.083, 0.0634)
        Case EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.TITANOVYYE_SPLAVY, EnumMaterialy.CHUGUN
            X = IIf(Massa <= 3, 0.106, 0.0762)
        Case Else
            MsgBox "RazmetkaKonturaBezShablona_Lineykoy" & vbNewLine & "Íåäîïóñòèìûé ìàòåğèàë", , "Îøèáêà ìàòåğèàëà"
            Exit Function
    End Select
    Y = 0.32
    z = 0.32
    
    Dim tRazmetka As Double
    
    Massa = MinimalnoeZnachenie(Massa, 1)
    Perimetr = MinimalnoeZnachenie(Perimetr, 50)
    
    Dim MaxPerimetr As Integer
    MaxPerimetr = 6000
    
    If Perimetr > MaxPerimetr Then
        tRazmetka = X * Massa ^ Y * MaxPerimetr ^ z
        tRazmetka = tRazmetka + 0.25 * ((Perimetr - MaxPerimetr) / 1000)
    Else
        tRazmetka = X * Massa ^ Y * Perimetr ^ z
    End If
    
    RazmetkaKonturaBezShablona_Lineykoy = tRazmetka
    
End Function













