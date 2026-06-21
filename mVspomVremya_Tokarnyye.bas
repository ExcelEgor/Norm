Attribute VB_Name = "mVspomVremya_Tokarnyye"
Option Explicit

Private Const L_MIN_POP As Integer = 25
Private Const D_NAD_STANINOY = 420
Private Const L_MIN_PROD As Integer = 50
Private Const D_MIN As Integer = 10


 Function Tvsp_PoperechnoyeTocheniye_Chernovoye(ByVal DlinaTocheniya As Double) As Double
    
    Dim a As Double, Y As Double, K As Double
    Dim t_vsp As Double

    a = 0.0021
    Y = 0.29
    K = 0.59
    DlinaTocheniya = WorksheetFunction.Max(DlinaTocheniya, L_MIN_POP)
    t_vsp = a * DlinaTocheniya ^ Y * D_NAD_STANINOY ^ K

    Tvsp_PoperechnoyeTocheniye_Chernovoye = t_vsp
    
End Function

 Function Tvsp_PoperechnoyeTocheniye_Chistovoye(ByVal izmeryaemyyRazmer As Double, ByVal DlinaTocheniya As Double, Kvalitet As Integer) As Double
    
    Dim a As Double, X As Double, Y As Double, K As Double
    Dim t_vsp As Double

    DlinaTocheniya = WorksheetFunction.Max(DlinaTocheniya, L_MIN_POP)
    izmeryaemyyRazmer = WorksheetFunction.Max(izmeryaemyyRazmer, RAZMER_MIN_POP)
    
    Select Case Kvalitet
        Case Is > 11
            a = 0.00786
            X = 0.16
            Y = 0.24
            K = 0.37
        Case Is > 9
            a = 0.018
            X = 0.2
            Y = 0.24
            K = 0.32
        Case Is > 7
            a = 0.0257
            X = 0.17
            Y = 0.27
            K = 0.31
        Case Else
            a = 0.0241
            X = 0.16
            Y = 0.26
            K = 0.36
    End Select
    
    t_vsp = a * izmeryaemyyRazmer ^ X * DlinaTocheniya ^ Y * D_NAD_STANINOY ^ K
    
    Tvsp_PoperechnoyeTocheniye_Chistovoye = t_vsp
    
End Function

Function Tvsp_ProdolnoyeTocheniye_Chernovoye(ByVal DlinaTocheniya As Double) As Double
    
    Dim a As Double
    
    Dim Y As Double
    Dim K As Double
    Dim t_vsp As Double

    a = 0.00162
    Y = 0.35
    K = 0.51
    DlinaTocheniya = WorksheetFunction.Max(DlinaTocheniya, L_MIN_PROD)
    t_vsp = a * DlinaTocheniya ^ Y * D_NAD_STANINOY ^ K

    Tvsp_ProdolnoyeTocheniye_Chernovoye = t_vsp
    
End Function

Function Tvsp_Prodolnoye_Tocheniye_Chistovoye(ByVal DiametrDetali As Double, ByVal DlinaTocheniya As Double, Kvalitet As Integer) As Double
    
    Dim a As Double, X As Double, Y As Double, K As Double
    Dim t_vsp As Double

    DlinaTocheniya = WorksheetFunction.Max(DlinaTocheniya, L_MIN_PROD)
    DiametrDetali = WorksheetFunction.Max(DiametrDetali, D_MIN)
    
    Select Case Kvalitet
        Case Is > 11
            a = 0.00416
            X = 0.2
            Y = 0.29
            K = 0.39
        Case Is > 9
            a = 0.0065
            X = 0.18
            Y = 0.23
            K = 0.53
        Case Is > 7
            a = 0.00975
            X = 0.17
            Y = 0.26
            K = 0.46
        Case Else
            a = 0.00846
            X = 0.19
            Y = 0.24
            K = 0.52
    End Select
    
    t_vsp = a * DiametrDetali ^ X * DlinaTocheniya ^ Y * D_NAD_STANINOY ^ K
        
    
    Tvsp_Prodolnoye_Tocheniye_Chistovoye = t_vsp
    
End Function
