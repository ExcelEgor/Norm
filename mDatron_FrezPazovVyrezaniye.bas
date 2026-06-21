Attribute VB_Name = "mDatron_FrezPazovVyrezaniye"
Option Explicit

Private Type RezhimyRezaniya
    ap As Double
    f As Double
End Type

Private Type ParametryPaza
    Shirina As Double
    Glubina As Double
    Dlina As Double
End Type

Private Const D_FREZY As Integer = 6

Public Function Datron_VyrezaniyePoKonturu(ByVal Material As EnumMaterialy, ByVal Dlina As Double, ByVal Glubina As Double) As Variant

    If PolozhitelnyyeChisla(Dlina, Glubina) = False Then Exit Function
    
    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(RaschotTmash_Vyrezaniye(Material, Dlina, Glubina))
    Normy(2) = SnyatiyeZausentsevPoKonturu(2 * Dlina, Material, False, NAPILNIK_SHABER, False)
    Normy(3) = IzmerShtangenCircul(Dlina / 4, Dlina / 4)
    
    Datron_VyrezaniyePoKonturu = Normy
    
End Function

Public Function Datron_VyrezaniyeOtverstiya(Material As EnumMaterialy, Dotv As Double, Lotv As Double) As Double
    Dim Dlina As Double
    Dlina = PI * (Dotv - D_FREZY)
    Datron_VyrezaniyeOtverstiya = Datron_VyrezaniyePoKonturu(Material, Dlina, Lotv)(1)
End Function

Private Function RaschotTmash_Vyrezaniye(ByVal Material As EnumMaterialy, ByVal Dlina As Double, ByVal Glubina As Double) As Double

    If PolozhitelnyyeChisla(Dlina, Glubina) = False Then Exit Function
    
    Dim tChern As Double
    tChern = RaschotTmash_Paz(Material, D_FREZY, Glubina, Dlina, True)
    
    Dim Fchist As Double
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Fchist = 1500
    ElseIf Material = POLIAMID Then
        Fchist = 1000
    ElseIf Material = ORGSTEKLO Then
        Fchist = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Fchist = 400
    Else
        Exit Function
    End If
    
    Dim iH_chist As Long
    iH_chist = OKRUGLVVERKH(Glubina / D_FREZY)
    
    Dim tChist As Double
    tChist = (iH_chist * Dlina) / Fchist
    
    RaschotTmash_Vyrezaniye = tChern + tChist
    
End Function

Public Function Datron_FrezerovaniyePazov(Material As EnumMaterialy, ShirinaPaza As Double, GlubinaPaza As Double, DlinaPaza As Double, Optional ByVal BezZachistki As Boolean = False) As Variant

    If MaterialKorrektnyy_Datron(Material) = False Or PolozhitelnyyeChisla(ShirinaPaza, GlubinaPaza, DlinaPaza) = False Then Exit Function

    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(RaschotTmash_Paz(Material, ShirinaPaza, GlubinaPaza, DlinaPaza, BezZachistki))
    Normy(2) = SnyatiyeZausentsevPoKonturu(4 * GlubinaPaza, Material, False, NADFIL, False)
    Normy(3) = IzmerShtangenCircul(ShirinaPaza, DlinaPaza)

    Datron_FrezerovaniyePazov = Normy
    
End Function

Private Function RaschotTmash_Paz(Material As EnumMaterialy, ShirinaPaza As Double, GlubinaPaza As Double, DlinaPaza As Double, Optional ByVal BezZachistki As Boolean = False) As Double

    If MaterialKorrektnyy_Datron(Material) = False Or PolozhitelnyyeChisla(ShirinaPaza, GlubinaPaza, DlinaPaza) = False Then Exit Function
    
    Dim Paz As ParametryPaza
    With Paz
        .Dlina = DlinaPaza
        .Glubina = GlubinaPaza
        .Shirina = ShirinaPaza
    End With
        
    Dim Dfrezy As Double
    Dfrezy = VyborDiametraFrezy(Paz)

    Dim tChern1 As Double
    tChern1 = Raschot_tChern_Do3D(Dfrezy, Paz, Material)
    
    Dim tChern2 As Double
    tChern2 = Raschot_tChern_Posle3D(Dfrezy, Paz, Material)
    
    If BezZachistki = False Then
    
        Dim tChist As Double
        tChist = Raschot_tChist(Dfrezy, Paz, Material)
        
        Dim tPritup As Double
        tPritup = Raschot_tPritup(Paz, Material)
        
    End If

    RaschotTmash_Paz = tChern1 + tChern2 + tChist + tPritup
   
End Function

Private Function VyborDiametraFrezy(Paz As ParametryPaza) As Double

    Dim Dfrezy As Double
    Select Case Paz.Shirina
        Case Is < 1.5:  Dfrezy = 1
        Case Is < 2:    Dfrezy = 1.5
        Case Is < 3:    Dfrezy = 2
        Case Is < 4:    Dfrezy = 3
        Case Is < 5:    Dfrezy = 4
        Case Is < 6:    Dfrezy = 5
        Case Is < 8:    Dfrezy = 6
        Case Is < 10:   Dfrezy = 8
        Case Is < 12:   Dfrezy = 10
        Case Is < 14:   Dfrezy = 12
        Case Is < 20:   Dfrezy = 14
        Case Else:      Dfrezy = 20
    End Select
    
    VyborDiametraFrezy = Dfrezy
    
End Function

Private Function Raschot_tChern_Do3D(Dfrezy As Double, Paz As ParametryPaza, Material As EnumMaterialy) As Double

    Dim Rezhymy As RezhimyRezaniya
    Rezhymy = RezhimyRezaniya_Do3D(Dfrezy, Material)
    
    Dim GlubinaFrezerovaniya As Double
    GlubinaFrezerovaniya = MathMin(Paz.Glubina, 3 * Dfrezy)
        
    Raschot_tChern_Do3D = Raschot_tChern(Dfrezy, GlubinaFrezerovaniya, Rezhymy, Paz)
    
End Function

Private Function Raschot_tChern_Posle3D(Dfrezy As Double, Paz As ParametryPaza, Material As EnumMaterialy) As Double
    
    If Paz.Glubina > 3 * Dfrezy Then
        Dim Rezhymy As RezhimyRezaniya
        Rezhymy = RezhimyRezaniya_Posle3D(Dfrezy, Material)
        
        Dim GlubinaFrezerovaniya As Double
        GlubinaFrezerovaniya = Paz.Glubina - 3 * Dfrezy
        Raschot_tChern_Posle3D = Raschot_tChern(Dfrezy, GlubinaFrezerovaniya, Rezhymy, Paz)
    End If
    
End Function

Private Function Raschot_tChern(Dfrezy As Double, GlubinaFrezerovaniya As Double, Rezhymy As RezhimyRezaniya, Paz As ParametryPaza) As Double

    Dim iH As Long
    iH = OKRUGLVVERKH(GlubinaFrezerovaniya / Rezhymy.ap)
    
    Dim iB As Long
    iB = OKRUGLVNIZ(Paz.Shirina / Dfrezy)
    If iB < 1 Then iB = 1
    
    Dim tChern As Double
    tChern = iH * iB * Paz.Dlina / Rezhymy.f

    Raschot_tChern = tChern
    
End Function

Private Function Raschot_tChist(Dfrezy As Double, Paz As ParametryPaza, Material As EnumMaterialy) As Double
    
    Dim Rezhymy As RezhimyRezaniya
    Rezhymy = RezhimyRezaniya_ZachistkaBojovykhStoron(Dfrezy, Material)
    
    Dim iH As Long
    iH = OKRUGLVVERKH(Paz.Glubina / Rezhymy.ap)
    
    Raschot_tChist = 2 * iH * Paz.Dlina / Rezhymy.f
    
End Function

Private Function Raschot_tPritup(Paz As ParametryPaza, Material As EnumMaterialy) As Double
    Raschot_tPritup = 2 * Paz.Dlina / PodachaPriPritupleniiOstrykhKromok(Material)
End Function

Private Function RezhimyRezaniya_Do3D(Dfrezy As Double, Material As EnumMaterialy) As RezhimyRezaniya
    
    Dim Rezhimy As RezhimyRezaniya
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Rezhimy.ap = IIf(Dfrezy < 5, 0.5, 1)
        Rezhimy.f = 3000
    ElseIf Material = POLIAMID Then
        Rezhimy.ap = 1
        Rezhimy.f = 2000
    ElseIf Material = ORGSTEKLO Then
        Rezhimy.ap = 1
        Rezhimy.f = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Rezhimy.ap = 1
        Rezhimy.f = 400
    Else
        Exit Function
    End If
    
    RezhimyRezaniya_Do3D = Rezhimy
    
End Function

Private Function RezhimyRezaniya_Posle3D(Dfrezy As Double, Material As EnumMaterialy) As RezhimyRezaniya

    Dim Rezhimy As RezhimyRezaniya
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Rezhimy.ap = IIf(Dfrezy <= 5, 0.25, 0.5)
        Rezhimy.f = 1500
    ElseIf Material = POLIAMID Then
        Rezhimy.ap = 0.5
        Rezhimy.f = 1000
    ElseIf Material = ORGSTEKLO Then
        Rezhimy.ap = 0.5
        Rezhimy.f = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Rezhimy.ap = 0.5
        Rezhimy.f = 400
    Else
        Exit Function
    End If
    
    RezhimyRezaniya_Posle3D = Rezhimy
    
End Function

Private Function RezhimyRezaniya_ZachistkaBojovykhStoron(Dfrezy As Double, Material As EnumMaterialy) As RezhimyRezaniya

    Dim Rezhimy As RezhimyRezaniya
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Rezhimy.f = 1500
    ElseIf Material = POLIAMID Then
        Rezhimy.f = 1000
    ElseIf Material = ORGSTEKLO Then
        Rezhimy.f = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Rezhimy.f = 400
    Else
        Exit Function
    End If
    
    Rezhimy.ap = IIf(Dfrezy < 20, Dfrezy, 10)
    
    RezhimyRezaniya_ZachistkaBojovykhStoron = Rezhimy
    
End Function

Private Function PodachaPriPritupleniiOstrykhKromok(Material As EnumMaterialy) As Double

    Dim f As Double
    If Material = ALUMINIYEVYYE_SPLAVY Then
        f = 3000
    ElseIf Material = POLIAMID Then
        f = 2000
    ElseIf Material = ORGSTEKLO Then
        f = 800
    ElseIf Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        f = 400
    Else
        Exit Function
    End If
    
    PodachaPriPritupleniiOstrykhKromok = f
    
End Function

Private Function MathMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MathMin = a Else MathMin = b
End Function
