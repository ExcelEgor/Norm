Attribute VB_Name = "mDatron_FrezTorstFrez"
Option Explicit

Private Type ParametryFrezerovaniya
    Dfrezy As Integer
    Fchern As Integer
    Fchist As Integer
End Type

Public Function Datron_FrezerovaniyeTortsevoyFrezoy(Material As EnumMaterialy, ByVal Dlina As Double, ByVal Shirina As Double, Pripusk As Double) As Variant
                                         
    Dim Normy(1 To 3) As Double
    Normy(1) = RaschotTsht_Datron(RaschotTmash(Material, Dlina, Shirina, Pripusk))
    Normy(2) = SnyatiyeZausentsevPoKonturu(2 * (Dlina + Shirina), Material, False, NAPILNIK_SHABER, False)
    Normy(3) = IzmerShtangenCircul(MIN_IZMER_RAZMER, Dlina)
    
    Datron_FrezerovaniyeTortsevoyFrezoy = Normy

End Function

Private Function RaschotTmash(Material As EnumMaterialy, ByVal Dlina As Double, ByVal Shirina As Double, Pripusk As Double) As Double
                                         
    Const GLUBINA_REZANIYA As Double = 1
    Const K_PEREKRYTIYA As Double = 0.7

    Dim Parametry As ParametryFrezerovaniya
    Parametry = OpredelitParametryFrezerovaniya(Material)
    
    With Parametry
        If .Dfrezy <= 0 Or .Fchern <= 0 Or .Fchist <= 0 Then Exit Function
    End With
    
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina)
    Shirina = Gabarity.Shirina
    Dlina = Gabarity.Dlina
    
    Dim KolVoProkhodovPoGlubine As Long, KolVoProkhodovPoShirine As Long
    KolVoProkhodovPoGlubine = OKRUGLVVERKH(Pripusk / GLUBINA_REZANIYA)
    KolVoProkhodovPoShirine = OKRUGLVVERKH(Shirina / (Parametry.Dfrezy * K_PEREKRYTIYA))
    
    Dim tChern As Double, tChist As Double
    tChern = (Dlina * KolVoProkhodovPoGlubine * KolVoProkhodovPoShirine) / Parametry.Fchern
    tChist = (Dlina * KolVoProkhodovPoShirine) / Parametry.Fchist

    RaschotTmash = tChern + tChist
    
End Function

Private Function OpredelitParametryFrezerovaniya(Material As EnumMaterialy) As ParametryFrezerovaniya
    
    With OpredelitParametryFrezerovaniya
        Select Case Material
            Case ALUMINIYEVYYE_SPLAVY
                .Fchern = 4000
                .Fchist = 2000
                .Dfrezy = 20
                
            Case POLIAMID
                .Fchern = 1900
                .Fchist = 1000
                .Dfrezy = 10
            
            Case ORGSTEKLO
                .Fchern = 800
                .Fchist = 800
                .Dfrezy = 10
                
            Case TEKSTOLIT, STEKLOTEKSTOLIT
                .Fchern = 400
                .Fchist = 400
                .Dfrezy = 10
                
            Case Else
                Exit Function
        End Select
    End With
    
End Function

