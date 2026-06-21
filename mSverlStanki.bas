Attribute VB_Name = "mSverlStanki"
Option Explicit

'леуюмхвеяйюъ напюанрйю дерюкеи мю ябепкхкэмшу ярюмйюу
'сйпсомеммше мнплюрхбш бпелемх
'едхмхввмне х лекйняепхимне опнхгбндярбн
'дехб.212.100-88
'йнд 2723
'1988

Public Const D_MAKS_METCHIK_RADIALSVERL As Integer = 48

Function NarezaniyeRezbyMetchikom_RadialSverlStanok(Diametr As Double, DlinaNarezki As Double, _
    Optional ByVal ShagRezby As Double = 0, Optional Kvalitet As Integer = 7, Optional Glukhoye As Boolean = False, _
    Optional Kalibrovaniye As Boolean = False, Optional DvaMetchika As Boolean = False) As Double
    
    'йЮПР 22. мЮПЕГЮМЭЕ ПЕГЭАШ ОН 9 ЙБЮКХРЕРС. лЕРВХЙХ ЛЮЬХММШЕ.

    If Diametr <= 0 Or DlinaNarezki <= 0 Or Diametr > D_MAKS_METCHIK_RADIALSVERL Then Exit Function
    If ShagRezby = 0 Then ShagRezby = KrupShagRezb(Diametr)
    
    Dim kKvalitet As Double, kGlukhoye As Double, kKalibrovaniye As Double, kDvaMetchika As Double
    kKvalitet = IIf(Kvalitet < 9, 1.3, 1)
    kGlukhoye = IIf(Glukhoye, 1.1, 1)
    kKalibrovaniye = IIf(Kalibrovaniye, 0.25, 1)
    kDvaMetchika = IIf(DvaMetchika, 2, 1)
    
    Dim lVrezPerebeg As Double
    lVrezPerebeg = 4 * ShagRezby
    
    Dim Oboroty As Integer
    Oboroty = RaschotOborotov(Diametr, DlinaNarezki)
    
    Dim tNarez As Double
    tNarez = ((DlinaNarezki + lVrezPerebeg) / (ShagRezby * Oboroty) * 2 + 0.18) * 1.08

End Function

Private Function RaschotOborotov(Diametr As Double, DlinaNarezki As Double) As Integer

    Dim kL As Double
    Select Case DlinaNarezki / Diametr
        Case Is <= 3: kL = 1
        Case Is <= 5: kL = 0.85
        Case Is <= 7: kL = 0.7
        Case Else: kL = 0.6
    End Select
    
    Dim N As Double
    Select Case Diametr
        Case Is <= 3: N = 308
        Case Is <= 4: N = 280
        Case Is <= 5: N = 230
        Case Is <= 6: N = 244
        Case Is <= 8: N = 187
        Case Is <= 10: N = 171
        Case Is <= 12: N = 156
        Case Is <= 14: N = 116
        Case Is <= 16: N = 116
        Case Is <= 20: N = 99
        Case Is <= 30: N = 90
        Case Is <= 36: N = 81
        Case Is <= 39: N = 68
        Case Is <= 42: N = 57
        Case Is <= 48: N = 49
    End Select
    
    RaschotOborotov = Round(kL * N, 0)
    
End Function
