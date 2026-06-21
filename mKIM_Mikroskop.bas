Attribute VB_Name = "mKIM_Mikroskop"
Option Explicit
Option Private Module

Function KIM_Mikroskop_Tpz(Slozhnost As Integer) As Double
    
    Select Case Slozhnost
        Case 1: KIM_Mikroskop_Tpz = 5
        Case 2: KIM_Mikroskop_Tpz = 10
        Case 3: KIM_Mikroskop_Tpz = 15
    End Select
    
End Function

Function KIM_Mikroskop_Ustanovka(Massa As Double)
    
    If Massa <= 0 Or Massa > 20 Then Exit Function
    
    Dim tUst As Double
    Select Case Massa
        Case Is <= 0.01:    tUst = 0.3
        Case Is <= 0.05:    tUst = 0.27
        Case Is <= 0.08:    tUst = 0.25
        Case Is <= 0.3:     tUst = 0.3
        Case Is <= 1:       tUst = 0.4
        Case Is <= 3:       tUst = 0.52
        Case Is <= 5:       tUst = 0.59
        Case Is <= 10:      tUst = 0.7
        Case Else:          tUst = 0.82
    End Select

    KIM_Mikroskop_Ustanovka = tUst
    
End Function

Function KIM_Mikroskop_IzmereniyeDiametra(Diametr As Double) As Double
    KIM_Mikroskop_IzmereniyeDiametra = KIM_Mikroskop_IzmereniyeKontura(PI * Diametr)
End Function

Function KIM_Mikroskop_IzmereniyeKontura(DlinaKontura As Double) As Double

    If DlinaKontura <= 0 Then Exit Function

    Dim Smm_min_Mikroskop As Integer
    Smm_min_Mikroskop = 70
    
    Dim tVspom_Mikroskop As Integer
    tVspom_Mikroskop = 1
    
    KIM_Mikroskop_IzmereniyeKontura = DlinaKontura / Smm_min_Mikroskop + tVspom_Mikroskop
    
End Function

