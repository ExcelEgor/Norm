VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShlifovaniye_Koeff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function RaschotKoeff_Ploskoe(ShirinaShlifKruga As Double, TolshchinaDetali As Double, DopuskPloskostnostiParallelnosti As Double) As Double

    Dim K1 As Double, K2 As Double, K3 As Double

    Select Case ShirinaShlifKruga
        Case Is <= 50:  K1 = 1
        Case Is <= 63:  K1 = 0.85
        Case Else:      K1 = 0.7
    End Select

    Select Case TolshchinaDetali
        Case Is <= 2:   K2 = 1.3
        Case Is <= 5:   K2 = 1.15
        Case Else:      K2 = 1
    End Select

    Select Case DopuskPloskostnostiParallelnosti
        Case 0:             K3 = 1
        Case Is >= 0.01:    K3 = 1.15
        Case Is >= 0.005:   K3 = 1.2
        Case Else:          K3 = 1.3
    End Select
    
    RaschotKoeff_Ploskoe = K1 * K2 * K3
    
End Function

Public Function RaschotKoeff_Vnutrenneye(IT As Integer, Ra As Double, KonicheskayaPoverhnost As Boolean, _
    DopuskBieniyaSoosnostiKruglosti As Double, ShirinaShlifKruga As Double) As Double

    Dim K1 As Double
    K1 = Koeffitsient_IT_Ra(IT, Ra, 2)
    
    Dim K2 As Double
    K2 = RaschotKoeff_ProdNaruzhVnutrn(KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, ShirinaShlifKruga)
    
    RaschotKoeff_Vnutrenneye = K1 * K2
    
End Function

Public Function RaschotKoeff_NaruzhRadial(IT As Integer, Ra As Double, KolVoGalteley As Integer, DopuskBieniyaSoosnostiKruglosti As Double) As Double

    Dim K1 As Double
    K1 = Koeffitsient_IT_Ra(IT, Ra, 1)
    
    Dim K2 As Double
    K2 = RaschotKoeffitsienta_KolVoGalteley(KolVoGalteley)
    
    Dim K3 As Double
    K3 = Koeffitsient_DopuskBieniyaSoosnostiKruglosti(DopuskBieniyaSoosnostiKruglosti)
    
    RaschotKoeff_NaruzhRadial = K1 * K2 * K3
    
End Function

Public Function RaschotKoeff_NauzhProd(IT As Integer, Ra As Double, KolVoGalteley As Integer, _
    DopuskBieniyaSoosnostiKruglosti As Double, KonicheskayaPoverhnost As Boolean, ShirinaKruga As Double) As Double

    Dim K1 As Double
    K1 = Koeffitsient_IT_Ra(IT, Ra, 1)
    
    Dim K2 As Double
    K2 = RaschotKoeffitsienta_KolVoGalteley(KolVoGalteley)
    
    Dim K3 As Double
    K3 = RaschotKoeff_ProdNaruzhVnutrn(KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, ShirinaKruga)
    
    RaschotKoeff_NauzhProd = K1 * K2 * K3

End Function

Private Function RaschotKoeffitsienta_KolVoGalteley(KolVoGalteley As Integer) As Double
    Select Case KolVoGalteley
        Case 0: RaschotKoeffitsienta_KolVoGalteley = 1
        Case 1: RaschotKoeffitsienta_KolVoGalteley = 1.1
        Case 2: RaschotKoeffitsienta_KolVoGalteley = 1.2
    End Select
End Function

Private Function RaschotKoeff_ProdNaruzhVnutrn(KonicheskayaPoverhnost As Boolean, DopuskBieniyaSoosnostiKruglosti As Double, ShirinaKruga As Double) As Double

    Dim K1 As Double
    K1 = IIf(KonicheskayaPoverhnost, 1.3, 1)
    
    Dim K2 As Double
    K2 = Koeffitsient_DopuskBieniyaSoosnostiKruglosti(DopuskBieniyaSoosnostiKruglosti)
    
    Dim K3 As Double
    Select Case ShirinaKruga
        Case Is <= 50: K3 = 1
        Case Is <= 63: K3 = 0.85
        Case Is <= 80: K3 = 0.7
    End Select

    RaschotKoeff_ProdNaruzhVnutrn = K1 * K2 * K3
    
End Function

Private Function Koeffitsient_DopuskBieniyaSoosnostiKruglosti(DopuskBieniyaSoosnostiKruglosti As Double) As Double
    Select Case DopuskBieniyaSoosnostiKruglosti
        Case 0:             Koeffitsient_DopuskBieniyaSoosnostiKruglosti = 1
        Case Is > 0.005:    Koeffitsient_DopuskBieniyaSoosnostiKruglosti = 1.2
        Case Is > 0.002:    Koeffitsient_DopuskBieniyaSoosnostiKruglosti = 1.25
        Case Else:          Koeffitsient_DopuskBieniyaSoosnostiKruglosti = 1.35
    End Select
End Function

Private Function Koeffitsient_IT_Ra(ByVal IT As Integer, ByVal Ra As Double, TipPoverkhnosti As Integer) As Double
    'TipPoverkhnosti: 1-вал, 2-отверстие
     
    Select Case Ra
        Case Is <= 0.32
            If IT > IIf(TipPoverkhnosti = 1, 7, 8) Then
                IT = IIf(TipPoverkhnosti = 1, 7, 8)
            End If
        Case Is <= 0.63
            If IT > 9 Then
                IT = 9
            End If
    End Select
    
    Select Case IT
        Case Is > 9
            Koeffitsient_IT_Ra = IIf(Ra >= 2.5, 0.43, 0.52)
        Case Is > IIf(TipPoverkhnosti = 1, 7, 8)
            Select Case Ra
                Case Is > 1.25:     Koeffitsient_IT_Ra = 0.53
                Case Is > 0.63:     Koeffitsient_IT_Ra = 0.64
                Case Else:          Koeffitsient_IT_Ra = 0.77
            End Select
        Case Is > IIf(TipPoverkhnosti = 1, 5, 6)
            Select Case Ra
                Case Is > 0.63:     Koeffitsient_IT_Ra = 1
                Case Is > 0.32:     Koeffitsient_IT_Ra = 1.2
                Case Else:          Koeffitsient_IT_Ra = 1.55
            End Select
        Case Else
            Select Case Ra
                Case Is > 0.63:     Koeffitsient_IT_Ra = 1.25
                Case Is > 0.32:     Koeffitsient_IT_Ra = 1.5
                Case Else:          Koeffitsient_IT_Ra = 1.9
            End Select
    End Select
    
End Function

