VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShlifovaniye_Vnutrenneye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Normativ_Vnutrenneye

Public Function RaschotTsht(Material As EnumMaterialy, ByVal Diametr As Double, Pripusk As Double, Dlina As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, Optional ShirinaShlifKruga As Double = 40) As Double
    
    If ProverkaParametrov(Material, Diametr, Pripusk, Dlina, IT, Ra, DopuskBieniyaSoosnostiKruglosti, ShirinaShlifKruga) = False Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("shlif_vnutrenneye")
    
    Dim Shlifovaniye_Koeff As New clsShlifovaniye_Koeff, K As Double
    K = Shlifovaniye_Koeff.RaschotKoeff_Vnutrenneye(IT, Ra, KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, ShirinaShlifKruga)
    
    Dim i As Integer, Stx As Double, Sm As Double
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) >= Diametr And Normativ(i, 3) >= Pripusk Then
            Stx = Normativ(i, 4)
            Sm = Normativ(i, 5)
            Exit For
        End If
    Next
    
    Dim a As Double
    Select Case Diametr
        Case Is <= 45:      a = 0.54
        Case Is <= 135:     a = 0.66
        Case Else:          a = 1.15
    End Select

    RaschotTsht = 1.1 * K * ((Dlina * 2 * Pripusk) / (Stx * Sm) + a * Dlina ^ 0.1)

End Function

Private Function ProverkaParametrov(Material As EnumMaterialy, ByVal Diametr As Double, Pripusk As Double, Dlina As Double, IT As Integer, Ra As Double, _
    DopuskBieniyaSoosnostiKruglosti As Double, ShirinaShlifKruga As Double) As Boolean
    
    Const MIN_IT As Integer = 5
    Const MIN_RA As Double = 0.32
    Const MAKS_SHIRINA_SHLIF_KRUGA As Integer = 80
    
    ProverkaParametrov = False
    
    If IT < MIN_IT Or Ra < MIN_RA Or Diametr > D_MAKS_VNUTR_SHLIF _
        Or ShirinaShlifKruga > MAKS_SHIRINA_SHLIF_KRUGA Or DopuskBieniyaSoosnostiKruglosti < 0 Then Exit Function
    
    If MaterialKorrektnyy_Shlifovaniye(Material) = False Then Exit Function
    
    If PolozhitelnyyeChisla(Diametr, Dlina, Pripusk, IT, Ra, ShirinaShlifKruga) = False Then Exit Function
    
    If Pripusk > OpredelitMaksimalnyyPripusk(Material, Diametr) Then Exit Function
    
    ProverkaParametrov = True

End Function

Private Function OpredelitMaksimalnyyPripusk(Material As EnumMaterialy, Diametr As Double) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("shlif_vnutrenneye")
    
    Dim i As Long
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) = Material And Normativ(i, 2) >= Diametr Then
            OpredelitMaksimalnyyPripusk = Normativ(i, 3)
            If PoslednyayaStrokaVMassive(Normativ, i, 2) Then Exit For
        End If
    Next
    
End Function


