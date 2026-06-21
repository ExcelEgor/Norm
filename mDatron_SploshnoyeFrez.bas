Attribute VB_Name = "mDatron_SploshnoyeFrez"
Option Explicit

Private Type ParamChernFrez
    Dfrezy As Double
    ap As Double
End Type

Public Function Datron_FrezerovaniyeObnizheniy(ByVal Material As EnumMaterialy, _
                                               ByVal Dlina As Double, _
                                               ByVal Shirina As Double, _
                                               ByVal Glubina As Double, _
                                               ByVal Radius As Double, _
                                               Optional BezZachistki As Boolean = False) As Variant
    
    If PolozhitelnyyeChisla(Dlina, Shirina, Glubina) = False Or MaterialKorrektnyy_Datron(Material) = False Then Exit Function
    
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina)
    Dlina = Gabarity.Dlina
    Shirina = Gabarity.Shirina
    
    Dim tMash As Double
    tMash = RaschotTmash(Material, Dlina, Shirina, Glubina, Radius, BezZachistki)

    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(tMash)
    Normy(2) = SnyatiyeZausentsevPoKonturu(2 * (Dlina + Shirina), Material, False, NAPILNIK_SHABER, False)
    Normy(3) = IzmerShtangenCircul_CPL(Dlina, Shirina) + IzmereniyeShtangenGlubinomerom(Glubina)
    
    Datron_FrezerovaniyeObnizheniy = Normy
    
End Function

Private Function RaschotTmash(ByVal Material As EnumMaterialy, _
                                               ByVal Dlina As Double, _
                                               ByVal Shirina As Double, _
                                               ByVal Glubina As Double, _
                                               ByVal Radius As Double, BezZachistki As Boolean) As Double
    
    If PolozhitelnyyeChisla(Dlina, Shirina, Glubina) = False Then Exit Function
    
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina)
    Dlina = Gabarity.Dlina
    Shirina = Gabarity.Shirina
    
    ' 1. ќпределение параметров инструмента и режимов
    Dim ChernParam As ParamChernFrez
    ChernParam = RaschotChernParam(Shirina)
    
    Dim DfrezyChist As Double
    DfrezyChist = VyborDfrezyPodRadius(Radius)
    
    Dim Fchern As Double: Fchern = RaschotPodachi(Material, ChernParam.Dfrezy, True)
    Dim Fchist1 As Double: Fchist1 = RaschotPodachi(Material, ChernParam.Dfrezy, False)
    
    ' 2. ¬ычисление машинного времени по этапам
    Dim tMash As Double
    
    tMash = CalcTime_Chern(Dlina, Shirina, Glubina, ChernParam.Dfrezy, ChernParam.ap, Fchern)
    tMash = tMash + Raschot_Dno(Dlina, Shirina, ChernParam.Dfrezy, Fchist1)
    
    If BezZachistki = False Then tMash = tMash + Raschot_Perimetr(Dlina, Shirina, Glubina, ChernParam.Dfrezy, Fchist1)
    
    ' ƒоработка углов (только если требуетс€)
    If DfrezyChist <> ChernParam.Dfrezy Then
        Dim Fchist2 As Double: Fchist2 = RaschotPodachi(Material, DfrezyChist, False)
        tMash = tMash + Raschot_Ugly(Glubina, ChernParam.Dfrezy, DfrezyChist, Fchist2)
    End If
    
    RaschotTmash = tMash
End Function

Private Function CalcTime_Chern(ByVal L As Double, ByVal b As Double, ByVal H As Double, _
                                ByVal Dfrezy As Double, ByVal ap As Double, ByVal f As Double) As Double
    Const KOEF_PEREKRYTIYA As Double = 0.5
    Dim iB As Long: iB = OKRUGLVVERKH(b / (Dfrezy * KOEF_PEREKRYTIYA))
    Dim iH As Long: iH = OKRUGLVVERKH(H / ap)
    CalcTime_Chern = (L * iB * iH) / f
End Function

Private Function Raschot_Dno(ByVal L As Double, ByVal b As Double, _
                              ByVal Dfrezy As Double, ByVal f As Double) As Double
    Const KOEF_PEREKRYTIYA As Double = 0.5
    Dim iB As Long: iB = OKRUGLVVERKH(b / (Dfrezy * KOEF_PEREKRYTIYA))
    Raschot_Dno = (L * iB) / f
End Function

Private Function Raschot_Perimetr(ByVal L As Double, ByVal b As Double, ByVal H As Double, _
                                 ByVal Dfrezy As Double, ByVal f As Double) As Double
    Dim p As Double: p = 2 * (L + b)
    Dim iH As Long: iH = OKRUGLVVERKH(H / Dfrezy)
    Raschot_Perimetr = (p * iH) / f
End Function

Private Function Raschot_Ugly(ByVal H As Double, ByVal DfrezyChern As Double, _
                               ByVal DfrezyChist As Double, ByVal f As Double) As Double
    Dim L_ugly As Double: L_ugly = PI * DfrezyChern
    Dim iH As Long: iH = OKRUGLVVERKH(H / DfrezyChist)
    Raschot_Ugly = (L_ugly * iH) / f
End Function

Private Function RaschotChernParam(ByVal Shirina As Double) As ParamChernFrez
    With RaschotChernParam
        Select Case Shirina
            Case Is <= 2:   .Dfrezy = 1: .ap = 0.3
            Case Is <= 3.5: .Dfrezy = 2: .ap = 0.3
            Case Is <= 4.5: .Dfrezy = 3: .ap = 0.5
            Case Is <= 6.5: .Dfrezy = 4: .ap = 0.5
            Case Is <= 8.5: .Dfrezy = 6: .ap = 0.7
            Case Is <= 12:  .Dfrezy = 8: .ap = 1
            Case Else:      .Dfrezy = 10: .ap = 1
        End Select
    End With
End Function

Private Function VyborDfrezyPodRadius(ByVal Radius As Double) As Double
    Dim Dfrezy As Double
    
    If Radius = 0 Then
        Dfrezy = 8
    Else
        Select Case Radius
            Case Is >= 5:   Dfrezy = 10
            Case Is >= 4:   Dfrezy = 8
            Case Is >= 3:   Dfrezy = 6
            Case Is >= 2:   Dfrezy = 4
            Case Is >= 1.5: Dfrezy = 3
            Case Is >= 1:   Dfrezy = 2
            Case Else:      Dfrezy = 1
        End Select
    End If
    
    VyborDfrezyPodRadius = Dfrezy
End Function

Private Function RaschotPodachi(ByVal Material As EnumMaterialy, ByVal Dfrezy As Double, ByVal Chern As Boolean) As Integer
    Dim Podacha As Integer
    
    If Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Podacha = 400
    ElseIf Material = ORGSTEKLO Then
        Podacha = 800
    ElseIf Material = POLIAMID Then
        If Chern Then Podacha = 1900 Else Podacha = 1000
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
        If Dfrezy > 6 Then
            If Chern Then Podacha = 1600 Else Podacha = 800
        Else
            If Chern Then Podacha = 1000 Else Podacha = 500
        End If
    End If
    
    RaschotPodachi = Podacha
End Function




