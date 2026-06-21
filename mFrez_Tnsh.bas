Attribute VB_Name = "mFrez_Tnsh"
'@Folder Frezerovaniye

Option Explicit

Private Const DIAMETR_SHPINDELYA As Integer = 105    'Станок WHQ 105 CNC
Private Const DLINA_STOLA As Integer = 1500
Private Const MIN_DLINA_RAZMER As Integer = 100
Private Const K_SLEDUYUSHCHIY_PROKHOD_PO_GLUBINE As Double = 0.6
Private Const K_OBSRABMESTA_OTH_LICHNNADOBN As Double = 1.1

Public Function RaschotTnshRastFrez_Chist(TipStanka As EnumTipStankaRastFrez, DlinaFrezerovaniya As Double, DlinaNedobegaPerebega As Double, Podacha As Double, _
    KolVoProkhodovPoShirine As Double, Kvalitet As Integer, CHPU As Boolean)
    
    If DlinaFrezerovaniya <= 0 Or Podacha <= 0 Or KolVoProkhodovPoShirine <= 0 Or Kvalitet < 6 Then Exit Function
    
    Dim tVspom As Double
    If CHPU Then
        tVspom = 0.3 / K_OBSRABMESTA_OTH_LICHNNADOBN
    Else
        tVspom = RaschotTvspRastFrez_Chistovoye(TipStanka, DlinaFrezerovaniya, KolVoProkhodovPoShirine, Kvalitet)
    End If
    
    Dim tMash As Double
    tMash = IIf(Kvalitet <= 11, 2, 1) * KolVoProkhodovPoShirine * (DlinaFrezerovaniya + DlinaNedobegaPerebega) / Podacha
    
    RaschotTnshRastFrez_Chist = K_OBSRABMESTA_OTH_LICHNNADOBN * (tVspom + tMash)
    
End Function

Public Function RaschotTnshRastFrez_Chern(TipStanka As EnumTipStankaRastFrez, DlinaFrezerovaniya As Double, DlinaNedobegaPerebega As Double, Podacha As Double, _
    KolVoProkhodovPoShirine As Double, KolVoChernProkhodovPoGlubine As Double, CHPU As Boolean)
    
    Dim tVspom As Double
    If CHPU Then
        tVspom = 0.3 / K_OBSRABMESTA_OTH_LICHNNADOBN
    Else
        tVspom = RaschotTvspRastFrez_Cnernovoye(TipStanka, DlinaFrezerovaniya, KolVoProkhodovPoShirine, KolVoChernProkhodovPoGlubine)
    End If
    
    Dim tMash As Double
    tMash = KolVoProkhodovPoShirine * KolVoChernProkhodovPoGlubine * (DlinaFrezerovaniya + DlinaNedobegaPerebega) / Podacha
    
    RaschotTnshRastFrez_Chern = K_OBSRABMESTA_OTH_LICHNNADOBN * (tVspom + tMash)
    
End Function

Public Function RaschotTvspRastFrez_Cnernovoye(TipStanka As EnumTipStankaRastFrez, ByVal Dlina As Double, KolVoProkhodovPoShirine As Double, KolVoChernProkhodovPoGlubine As Double)
   
    If TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy _
        Or Dlina <= 0 Or KolVoProkhodovPoShirine <= 0 Or KolVoChernProkhodovPoGlubine <= 0 Then Exit Function
        
    Dlina = IIf(Dlina < MIN_DLINA_RAZMER, MIN_DLINA_RAZMER, Dlina)
    
    Dim tVsp As Double
    tVsp = tVspChernovoyeNaProkhod(TipStanka, Dlina)
    
    Dim KolVoProkhodov As Double
    KolVoProkhodov = KolVoProkhodovPoShirine * KolVoChernProkhodovPoGlubine
    
    Dim tPervyyProkhod As Double, tSledProkhod As Double
    If KolVoProkhodov > 1 Then
        tPervyyProkhod = tVsp
        tSledProkhod = K_SLEDUYUSHCHIY_PROKHOD_PO_GLUBINE * tVsp
        RaschotTvspRastFrez_Cnernovoye = tPervyyProkhod + tSledProkhod * (KolVoProkhodov - 1)
    Else
        RaschotTvspRastFrez_Cnernovoye = tVsp
    End If

End Function

Private Function tVspChernovoyeNaProkhod(TipStanka As EnumTipStankaRastFrez, Dlina As Double) As Double

    If TipStanka = GorizontalnoRastochnoy Then
        tVspChernovoyeNaProkhod = 0.042 * Dlina ^ 0.16 * DIAMETR_SHPINDELYA ^ 0.44
    Else
        tVspChernovoyeNaProkhod = 0.0016 * Dlina ^ 0.41 * DLINA_STOLA ^ 0.42
    End If
    
End Function

Public Function RaschotTvspRastFrez_Chistovoye(TipStanka As EnumTipStankaRastFrez, ByVal Dlina As Double, KolVoProkhodovPoShirine As Double, Kvalitet As Integer) As Double

    If TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy _
        Or Dlina <= 0 Or KolVoProkhodovPoShirine <= 0 Or Kvalitet < 6 Then Exit Function
    
    Dlina = IIf(Dlina < MIN_DLINA_RAZMER, MIN_DLINA_RAZMER, Dlina)

    Dim tVsp As Double
    If TipStanka = GorizontalnoRastochnoy Then
        tVsp = tVspChistovoyeNaProkhod_GRS(Dlina, Kvalitet)
    Else
        tVsp = tVspChistovoyeNaProkhod_FRV(Dlina, Kvalitet)
    End If
    
    Dim kShirina As Double, tSledProkhod As Double
    If KolVoProkhodovPoShirine > 1 Then
        If Kvalitet > 11 Then
            kShirina = 0.6
        Else
            kShirina = 0.7
        End If
        tSledProkhod = kShirina * tVsp
    End If
    
    RaschotTvspRastFrez_Chistovoye = tVsp + tSledProkhod * (KolVoProkhodovPoShirine - 1)
    
End Function

Private Function tVspChistovoyeNaProkhod_GRS(Dlina As Double, Kvalitet As Integer) As Double
    
    Dim Pokazateli As Variant
    Select Case Kvalitet
        Case Is > 11:   Pokazateli = Array(0.07, 0.11, 0.1, 0.38)
        Case Is > 9:    Pokazateli = Array(0.079, 0.13, 0.15, 0.38)
        Case Is > 7:    Pokazateli = Array(0.125, 0.14, 0.14, 0.36)
        Case Else:      Pokazateli = Array(0.163, 0.14, 0.14, 0.36)
    End Select
    
    Dim a As Double, X As Double, Y As Double, K As Double
    a = Pokazateli(0)
    X = Pokazateli(1)
    Y = Pokazateli(2)
    K = Pokazateli(3)
    
    tVspChistovoyeNaProkhod_GRS = a * MIN_DLINA_RAZMER ^ X * Dlina ^ Y * DIAMETR_SHPINDELYA ^ K
    
End Function

Private Function tVspChistovoyeNaProkhod_FRV(Dlina As Double, Kvalitet As Integer) As Double
    
    Dim Pokazateli As Variant
    Select Case Kvalitet
        Case Is > 11:   Pokazateli = Array(0.016, 0.16, 0.2, 0.31)
        Case Is > 9:    Pokazateli = Array(0.048, 0.13, 0.2, 0.28)
        Case Is > 7:    Pokazateli = Array(0.224, 0.09, 0.2, 0.13)
        Case Else:      Pokazateli = Array(0.265, 0.09, 0.23, 0.12)
    End Select
    
    Dim a As Double, X As Double, Y As Double, K As Double
    a = Pokazateli(0)
    X = Pokazateli(1)
    Y = Pokazateli(2)
    K = Pokazateli(3)

    tVspChistovoyeNaProkhod_FRV = a * MIN_DLINA_RAZMER ^ X * Dlina ^ Y * DLINA_STOLA ^ K

End Function





