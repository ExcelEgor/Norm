Attribute VB_Name = "mShtampovka"
Option Explicit

Private Type ParamtryVyrubki
    MaterialKorrektnyy As Boolean
    Material As EnumMaterialy
    MaksTolshchina As Integer
End Type

Function VyrubkaDetaleyIzPolosy(ByVal Material As EnumMaterialy, ShagProdvizheniya As Double, TolshchinaPolosy As Double, DlinaPolosy As Double, ShirinaPolosy As Double) As Double

    Dim Vyrubka As ParamtryVyrubki
    Vyrubka = OpredelitOsnovnyyeParametryVyrubki(Material)
    
    Const MAKS_SHAG_PRODVIZHENIYA As Integer = 400
    Const MAKS_SHIRINA_POLOSY As Integer = 400
    Const MAKS_DLINA_POLOSY As Integer = 5100
    
    If Vyrubka.MaterialKorrektnyy = False Or TolshchinaPolosy > Vyrubka.MaksTolshchina Or DlinaPolosy > MAKS_DLINA_POLOSY _
        Or ShagProdvizheniya > MAKS_SHAG_PRODVIZHENIYA Or ShirinaPolosy > MAKS_SHIRINA_POLOSY Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("vyrubka_iz_polosy")

    Dim i As Integer, tVyrubka As Double
    For i = LBound(Normativ) To UBound(Normativ)
        If Normativ(i, 1) >= ShagProdvizheniya _
            And Normativ(i, 2) = Vyrubka.Material _
            And Normativ(i, 3) >= TolshchinaPolosy _
            And Normativ(i, 4) >= ShirinaPolosy Then
            tVyrubka = Normativ(i, 5)
            Exit For
        End If
    Next
    
    Dim kDlina As Double
    kDlina = RaschotKoeffitsienta_VyrubkaDetaleyIzPolosy(DlinaPolosy)
    
    VyrubkaDetaleyIzPolosy = WorksheetFunction.RoundUp(kDlina * tVyrubka, 1)
    
End Function

Private Function OpredelitOsnovnyyeParametryVyrubki(Material As EnumMaterialy) As ParamtryVyrubki

    Dim Vyrubka As ParamtryVyrubki

    If Material = MEDNYYE_SPLAVY Or Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = STAL_UGLERODISTAYA Then
        
        Vyrubka.MaterialKorrektnyy = True
        Vyrubka.Material = STAL_UGLERODISTAYA
        Vyrubka.MaksTolshchina = 10
        
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
    
        Vyrubka.MaterialKorrektnyy = True
        Vyrubka.Material = Material
        Vyrubka.MaksTolshchina = 13
    
    Else
    
        Vyrubka.MaterialKorrektnyy = False
        
    End If
    
    OpredelitOsnovnyyeParametryVyrubki = Vyrubka

End Function

Private Function RaschotKoeffitsienta_VyrubkaDetaleyIzPolosy(DlinaPolosy As Double) As Double

    Dim kDlina As Double
    Select Case DlinaPolosy
        Case Is <= 700:     kDlina = 1.02
        Case Is <= 1000:    kDlina = 1
        Case Is <= 1500:    kDlina = 0.97
        Case Is <= 2300:    kDlina = 0.96
        Case Is <= 3400:    kDlina = 0.95
        Case Is <= 5100:    kDlina = 0.94
    End Select
    
    RaschotKoeffitsienta_VyrubkaDetaleyIzPolosy = kDlina
    
End Function
