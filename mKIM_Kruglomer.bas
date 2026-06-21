Attribute VB_Name = "mKIM_Kruglomer"
Option Explicit
Option Private Module

Private Const KOLVO_LINEYNYKH_PROKHODOV As Integer = 10
Private Const T_OBOROTA As Double = 0.2 'Время одного оборота при скорости 5 об/мин

Function KIM_Kruglomer_Tpz(GruppaSloznosti As Integer, VremyaNapisaniyaUP As Double) As Double

    Dim TpzOsnovnoe As Double
    TpzOsnovnoe = KIM_TpzOsnovnoe(GruppaSloznosti)
    
    KIM_Kruglomer_Tpz = TpzOsnovnoe + VremyaNapisaniyaUP
    
End Function

Function KIM_Kruglomer_VremyaNapisaniyaUPOdnogoPerekhoda(Dlina As Double, Slozhnost As Integer, KolVoPoverkhnostey As Double, _
    RadialnoyeIzmereniye As Boolean, LineynoyeIzmereniye As Boolean) As Double

    Dim KolVoSecheniy As Double
    If RadialnoyeIzmereniye = True Then
        KolVoSecheniy = KIM_KolVoSecheniy(Dlina, 3)
    End If
    
    If LineynoyeIzmereniye = True Then
        KolVoSecheniy = KOLVO_LINEYNYKH_PROKHODOV
    End If
    
    Dim KolVoStrokUP As Integer
    KolVoStrokUP = KolVoSecheniy * KolVoPoverkhnostey
    
    If KolVoStrokUP <= 0 Then Exit Function

    Dim tUP As Integer
    Select Case Slozhnost
        Case 1: tUP = 5
        Case 2: tUP = 10
        Case 3: tUP = 15
    End Select
    
    Dim tStrokaUP As Integer    'Время написания одной строки УП
    tStrokaUP = 2
    
    KIM_Kruglomer_VremyaNapisaniyaUPOdnogoPerekhoda = tUP + (KolVoStrokUP * tStrokaUP)
    
End Function

Function KIM_Kruglomer_Izmereniye(Dlina As Double, RadialnoyeIzmereniye As Boolean, LineynoyeIzmereniye As Boolean) As Double

    If Dlina <= 0 Then Exit Function

    If RadialnoyeIzmereniye Then
        Dim tRadialnoye As Double: tRadialnoye = Kruglomer_RadialnoyeIzmereniye(Dlina)
    End If

    If LineynoyeIzmereniye Then
        Dim tLineynoye As Double: tLineynoye = Kruglomer_LineynoyeIzmereniye(Dlina)
    End If

    KIM_Kruglomer_Izmereniye = tRadialnoye + tLineynoye

End Function

Private Function Kruglomer_LineynoyeIzmereniye(L As Double) As Double
    Dim Skorost As Integer
    Skorost = 180
    Kruglomer_LineynoyeIzmereniye = (L / Skorost + KIM_T_VSPOMOGATELNOYE) * KOLVO_LINEYNYKH_PROKHODOV + KIM_T_ANALIZA_RAZMERA
End Function

Private Function Kruglomer_RadialnoyeIzmereniye(L As Double) As Double
    Kruglomer_RadialnoyeIzmereniye = (T_OBOROTA + KIM_T_VSPOMOGATELNOYE) * KIM_KolVoSecheniy(L, 3) + KIM_T_ANALIZA_RAZMERA
End Function

Function KIM_Kruglomer_Vyravnivaniye(Dlina As Double, Optional KolVoTsiklov As Integer = 10) As Double

    If Dlina <= 0 Then Exit Function

    Dim UskorennyyKhod As Integer
    UskorennyyKhod = 900
    
    KIM_Kruglomer_Vyravnivaniye = (Dlina / UskorennyyKhod + T_OBOROTA) * KolVoTsiklov * 2
    
End Function

Function KIM_Kruglomer_Ustanovka(Massa As Double) As Double

    If Massa <= 0 Or Massa > 20 Then Exit Function
    
    Dim tUst As Double
    Select Case Massa
        Case Is <= 0.01:    tUst = 0.76
        Case Is <= 0.05:    tUst = 0.62
        Case Is <= 0.08:    tUst = 0.58
        Case Is <= 0.3:     tUst = 1.15
        Case Is <= 1:       tUst = 1.53
        Case Is <= 3:       tUst = 2
        Case Is <= 5:       tUst = 2.15
        Case Is <= 10:      tUst = 2.65
        Case Else:          tUst = 3.15
    End Select
    
    KIM_Kruglomer_Ustanovka = tUst
    
End Function


