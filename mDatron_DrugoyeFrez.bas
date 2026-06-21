Attribute VB_Name = "mDatron_DrugoyeFrez"
Option Explicit

Public Function Datron_FrezerovaniyeUstupov(ByVal Material As EnumMaterialy, ByVal Dlina As Double, ByVal Shirina As Double, ByVal Vysota As Double) As Variant
    
    If PolozhitelnyyeChisla(Dlina, Shirina, Vysota) = False Then Exit Function
    
    Dim tMash As Double
    tMash = RaschotTmash_Ustup(Material, Dlina, Shirina, Vysota)
        
    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(tMash)
    Normy(2) = SnyatiyeZausentsevPoKonturu(2 * Dlina + 2 * Shirina + 2 * Vysota, Material, False, NAPILNIK_SHABER, False)
    Normy(3) = IzmereniyeNutromerom(Vysota, Dlina) + IzmerShtangenCircul(Shirina, Dlina)

    Datron_FrezerovaniyeUstupov = Normy
    
End Function

 Function RaschotTmash_Ustup(ByVal Material As EnumMaterialy, ByVal Dlina As Double, ByVal Shirina As Double, ByVal Vysota As Double) As Double

    Const K_PEREKRYTIYA As Double = 0.7
    Const AP_CHERN As Double = 1
    Const Dfrezy  As Integer = 10
    
    Dim Fchern As Double, Fchist As Double
    Call RasschitatPodachi(Material, Fchern, Fchist)
    
    Dim iB As Long
    iB = OKRUGLVVERKH(Shirina / (Dfrezy * K_PEREKRYTIYA))
    
    Dim iH_chern As Long
    iH_chern = OKRUGLVVERKH(Vysota / AP_CHERN)
    
    Dim tChern As Double
    tChern = (Dlina * iB * iH_chern) / Fchern
    
    Dim iH_chist As Long
    iH_chist = OKRUGLVVERKH(Vysota / Dfrezy)
    
    Dim tChist As Double
    tChist = (Dlina * iB * iH_chist) / Fchist
    
    RaschotTmash_Ustup = tChern + tChist

End Function

Private Sub RasschitatPodachi(Material As EnumMaterialy, ByRef Fchern As Double, ByRef Fchist As Double)

    If Material = TEKSTOLIT Or Material = STEKLOTEKSTOLIT Then
        Fchern = 400
        Fchist = 400
    ElseIf Material = ORGSTEKLO Then
        Fchern = 800
        Fchist = 800
    ElseIf Material = POLIAMID Then
        Fchern = 1900
        Fchist = 1000
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
        Fchern = 1600
        Fchist = 800
    End If
    
End Sub

Public Function Datron_PostrochnoyeFrezerovaniye(ByVal Material As EnumMaterialy, ByVal DlinaFrezerovaniya As Double, ByVal DlinaKrivoy As Double) As Variant

    If PolozhitelnyyeChisla(DlinaFrezerovaniya, DlinaKrivoy) = False Then Exit Function

    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(RaschotTmash_Stupenchatoye(Material, DlinaFrezerovaniya, DlinaKrivoy))
    Normy(2) = ZachistkaShlifovalnoyShkurkoy(CInt(Material), DlinaFrezerovaniya, DlinaKrivoy)
    Normy(3) = IzmerShtangenCircul(DlinaFrezerovaniya / 4, DlinaFrezerovaniya / 4)
    
    Datron_PostrochnoyeFrezerovaniye = Normy
    
End Function

Private Function RaschotTmash_Stupenchatoye(ByVal Material As EnumMaterialy, ByVal DlinaFrezerovaniya As Double, ByVal DlinaKrivoy As Double) As Double
    
    Const Shag As Double = 0.1
    
    Dim KolVoProkhodov As Long
    KolVoProkhodov = OKRUGLVVERKH(DlinaKrivoy / Shag)
    
    Dim Podacha As Double
    If Material = STEKLOTEKSTOLIT Or Material = TEKSTOLIT Then
        Podacha = 400
    ElseIf Material = STEKLOTEKSTOLIT Then
        Podacha = 800
    ElseIf Material = POLIAMID Then
        Podacha = 1900
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
        Podacha = 3000
    Else
        Exit Function
    End If
    
    RaschotTmash_Stupenchatoye = (KolVoProkhodov * DlinaFrezerovaniya) / Podacha
    
End Function


