Attribute VB_Name = "mGidroabraziv"
Option Explicit
Private Const YS As Double = -1.15

Function GidrAbrazivSmm_min(Material As EnumMaterialy, H As Double, Optional VysokoeKachestvo As Boolean = False)
    
    Dim X As Double
    
    If H > 0 Then

        If Material = ALUMINIYEVYYE_SPLAVY Then
            If VysokoeKachestvo Then
                X = 2238.078952
            Else
                X = 2826.95
            End If
        
        ElseIf Material = STAL_UGLERODISTAYA Then
            If VysokoeKachestvo Then
                X = 727.80581369764
            Else
                X = 921.4738049
            End If
        
        ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
            If VysokoeKachestvo Then
                X = 763.784298706124
            Else
                X = 967.385838062869
            End If
        
        ElseIf Material = TITANOVYYE_SPLAVY Then
            If VysokoeKachestvo Then
                X = 888.244208735275
            Else
                X = 1121.82844155998
            End If
        
        ElseIf Material = REZINA Or Material = TEKSTOLIT Or Material = POLIAMID Then
            If VysokoeKachestvo Then
                X = 5625.19466975413
            Else
                X = 7110.5022690512
            End If
            
        End If

        GidrAbrazivSmm_min = X * H ^ YS
        
    End If

End Function

Function GidrAbraziv_Vremya(Material As EnumMaterialy, S As Double, L As Double, KolVoUglov As Double, Optional VysokoeKachestvo As Boolean = False) As Double

    Dim Podacha As Double
    Podacha = GidrAbrazivSmm_min(Material, S, VysokoeKachestvo)
    
    Dim tReza As Double
    tReza = L / Podacha
    
    Dim tUgol As Double
    If KolVoUglov > 0 Then
        tUgol = GidrAbraziv_VremyaNaUgol(Material, S) * KolVoUglov
    End If
    
    GidrAbraziv_Vremya = OkruglenieTsht((tReza + tUgol) * 1.1)

End Function

Function GidrAbraziv_VremyaNaUgol(Material As EnumMaterialy, S As Double)

    '    Dim PodachaNa100 As Double
    '    PodachaNa100 = GidrAbrazivSmm_min(Material, 100)
    '
    '    Dim t100 As Double
    '    t100 = 4.5
    '
    '    Dim PodachaFakt As Double
    '    PodachaFakt = GidrAbrazivSmm_min(Material, s)
    '
    '    GidrAbraziv_VremyaNaUgol = t100 / (PodachaFakt / PodachaNa100)

    Dim PodachaBazovaya As Double
    PodachaBazovaya = 10.2445728676679
    
    Dim t50 As Double
    t50 = 4.5
    
    Dim PodachaFakt As Double
    PodachaFakt = GidrAbrazivSmm_min(Material, S)
    
    GidrAbraziv_VremyaNaUgol = (t50 / PodachaFakt) * PodachaBazovaya
    
End Function
