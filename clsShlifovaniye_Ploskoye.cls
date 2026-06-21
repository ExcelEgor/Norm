VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShlifovaniye_Ploskoye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Normativ_Ploskoye

Function RaschotTsht(Material As EnumMaterialy, ByVal Shirina As Double, ByVal Dlina As Double, _
    ByVal IT As Integer, Ra As Double, Pripusk As Double, _
    Optional ShirinaShlifKruga As Double, Optional TolshchinaDetali As Double, Optional DopuskPloskostnostiParallelnosti As Double = 0) As Double
    
    Const ShirinaKrugaPoUmolchaniyu As Integer = 50
    Const TolshchinaDetaliPoUmolchaniyu As Integer = 6
    
    If ShirinaShlifKruga = 0 Then ShirinaShlifKruga = ShirinaKrugaPoUmolchaniyu
    If TolshchinaDetali = 0 Then TolshchinaDetali = TolshchinaDetaliPoUmolchaniyu

    If ProverkaParametrov_Ploskoye(Material, Shirina, Dlina, Pripusk, IT, Ra, ShirinaShlifKruga, TolshchinaDetali, DopuskPloskostnostiParallelnosti) = False Then Exit Function
    
    'Ўирина шлифовани€ должна быть меньше длины шлифовани€
    Dlina = DlinaShirina(Dlina, Shirina)(1)
    Shirina = DlinaShirina(Dlina, Shirina)(2)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("shlif_ploskoye")
    
    Dim tShlif As Double
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Material And _
            Normativ(i, 2) >= Shirina And Normativ(i, 3) >= Dlina And _
            Normativ(i, 4) >= Pripusk And Normativ(i, 5) <= IT And Normativ(i, 6) <= Ra Then
            tShlif = Normativ(i, 7)
            Exit For
        End If
    Next
    
    Dim Shlifovaniye_Koeff As New clsShlifovaniye_Koeff, K As Double
    K = Shlifovaniye_Koeff.RaschotKoeff_Ploskoe(ShirinaShlifKruga, TolshchinaDetali, DopuskPloskostnostiParallelnosti)
    
    RaschotTsht = K * tShlif

End Function

Private Function ProverkaParametrov_Ploskoye(Material As EnumMaterialy, Shirina As Double, Dlina As Double, Pripusk As Double, IT As Integer, Ra As Double, _
    ShirinaShlifKruga As Double, TolshchinaDetali As Double, DopuskPloskostnostiParallelnosti As Double) As Boolean
    
    Const MAKS_SHIRINA_SHLIF_KRUGA As Integer = 80
    Const MIN_IT As Integer = 5
    Const MIN_RA As Double = 0.32
    
    ProverkaParametrov_Ploskoye = False

    If PolozhitelnyyeChisla(Shirina, Dlina, Pripusk, IT, Ra, ShirinaShlifKruga, TolshchinaDetali) = False Then Exit Function
    
    If MaterialKorrektnyy_Shlifovaniye(Material) = False Then Exit Function
    
    If Shirina > MAKS_SHIRINA_SHLIFOVANIYA Or IT < MIN_IT Or Ra < MIN_RA _
        Or DopuskPloskostnostiParallelnosti < 0 Or ShirinaShlifKruga > MAKS_SHIRINA_SHLIF_KRUGA Then Exit Function

    If Dlina > RaschotLmax(Material, Shirina) Then Exit Function
    
    ProverkaParametrov_Ploskoye = True

End Function

Public Function RaschotLmax(Material As EnumMaterialy, ByVal Shirina As Double) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("shlif_ploskoye")
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = Material Then
            If Normativ(i, 2) >= Shirina Then
                Lmax = Normativ(i, 3)
                If PoslednyayaStrokaVMassive(Normativ, i, 3) = True Then Exit For
            End If
        End If
    Next
    
    RaschotLmax = Lmax

End Function




