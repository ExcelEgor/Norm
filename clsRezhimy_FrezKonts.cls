VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRezhimy_FrezKonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

Public Function RaschotVremeni_FrezPloskKonts(Rezhimy As clsMillingRegimes, TipStanka As EnumTipStankaRastFrez, Dfrezy As Double, _
    DlinaPoverkhnosti As Double, Shirinaoverkhnosti As Double, _
    IT As Integer, Ra As Double, CHPU As Boolean) As Double
    
    Dim tChern_Kork As Double, tChern As Double, tChist As Double
 
    With Rezhimy.ChernKork
        If .iH > 0 Then
            tChern_Kork = RaschotTnshRastFrez_Chern(TipStanka, DlinaPoverkhnosti, Dfrezy, .Pokazateli.S, .iB, .iH, CHPU)
        End If
    End With
 
    With Rezhimy.Chern
        If .iH > 0 Then
            tChern = RaschotTnshRastFrez_Chern(TipStanka, DlinaPoverkhnosti, Dfrezy, .Pokazateli.S, .iB, .iH, CHPU)
        End If
    End With
 
    With Rezhimy.Chist
        If .iH > 0 Then
            tChist = RaschotTnshRastFrez_Chist(TipStanka, DlinaPoverkhnosti, Dfrezy, .Pokazateli.S, .iB, IT, CHPU)
        End If
    End With
 
    RaschotVremeni_FrezPloskKonts = tChern_Kork + tChern + tChist

End Function

Public Function OpredelitDfrezy_FrezPloskKonts(Shirinaoverkhnosti As Double) As Double

    Select Case Shirinaoverkhnosti
        Case Is <= 10
            OpredelitDfrezy_FrezPloskKonts = 10
        Case Is <= 20
            OpredelitDfrezy_FrezPloskKonts = 32
        Case Else
            OpredelitDfrezy_FrezPloskKonts = 40
    End Select
    
End Function

Public Function RaschotVremeni_FrezOtv(Material As EnumMaterialy, TipStanka As EnumTipStankaRastFrez, Dfrezy As Double, Dotv As Double, dSverl As Double, Lotv As Double, ByVal Ra As Double, ByVal IT As Integer) As Double

    If MaterialKorrektnyy_RastFrez(Material) = False Or TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy _
        Or Dotv <= MIN_D_FREZ Or Lotv <= 0 Then Exit Function

    If Ra <= 0 Then Ra = 6.3
    If IT <= 0 Then IT = 11
    
    Dim Pripusk As Double
    Pripusk = 0.5 * (Dotv - dSverl)

    Dim Rezhimy As clsMillingRegimes
    Set Rezhimy = RaschotRezhimov(Material, Pripusk, Dfrezy, Lotv, IT, Ra, False) ' Добавил Set

    Dim tChern As Double, tChist As Double
    Dim DotvChern As Double

    With Rezhimy.Chern
        If .iH > 0 Then
            DotvChern = Dotv - 2 * .GlubRez * .iH
            
            Dim LrezChern As Double
            LrezChern = DlinaRezaniyaPriFrezOtverstiy(DotvChern, dSverl, Dfrezy, .GlubRez)
            
            tChern = RaschotTnshRastFrez_Chern(TipStanka, LrezChern, 0, .Pokazateli.S, .iB, .iH, True)
        Else
            DotvChern = dSverl
        End If
    End With
 
    With Rezhimy.Chist
        If .iH > 0 Then
            Dim LrezChist As Double
            LrezChist = DlinaRezaniyaPriFrezOtverstiy(Dotv, DotvChern, Dfrezy, .GlubRez)
            
            tChist = RaschotTnshRastFrez_Chist(TipStanka, LrezChist, 0, .Pokazateli.S, .iB, IT, True)
        End If
    End With
    
    RaschotVremeni_FrezOtv = tChern + tChist
    
End Function
 
Public Function OpredelitDfrezy_FrezOtv(dSverl As Double) As Double
    Dim Dfrezy As Double
    Select Case dSverl
        Case Is <= 3:  Dfrezy = 2
        Case Is <= 4:  Dfrezy = 3
        Case Is <= 5:  Dfrezy = 4
        Case Is <= 6:  Dfrezy = 5
        Case Is <= 8:  Dfrezy = 6
        Case Is <= 10: Dfrezy = 8
        Case Is <= 12: Dfrezy = 10
        Case Is <= 16: Dfrezy = 12
        Case Is <= 20: Dfrezy = 16
        Case Is <= 25: Dfrezy = 20
        Case Is <= 32: Dfrezy = 25
        Case Is <= 40: Dfrezy = 32
        Case Else:     Dfrezy = 40
    End Select
    
    OpredelitDfrezy_FrezOtv = Dfrezy
End Function

Public Function RaschotRezhimov(Material As EnumMaterialy, Pripusk As Double, DiametrFrezy As Double, ShirinaPoverkhnosti As Double, IT As Integer, Ra As Double, Korka As Boolean) As clsMillingRegimes
  
    If MaterialKorrektnyy_RastFrez(Material) = False Or Ra < MIN_RA_FREZ Then Exit Function

    If PolozhitelnyyeChisla(Pripusk, ShirinaPoverkhnosti, DiametrFrezy) = False Then Exit Function
    
    IT = IIf(IT = 0, 11, IT)
    
    Dim hChist As Double, hChern As Double
    RaschitatPripuski Ra, Pripusk, hChist, hChern

    Dim Rezhimy As New clsMillingRegimes
    Dim Bfrez As Double
    Bfrez = OpredelitShirinuFrezerovaniya(DiametrFrezy)

    Dim iB As Long
    iB = WorksheetFunction.RoundUp(ShirinaPoverkhnosti / Bfrez, 0)
    Bfrez = ShirinaPoverkhnosti / iB
    
    RaschitatRezhimy_Chist Material, Bfrez, hChist, iB, DiametrFrezy, IT, Rezhimy.Chist
    RaschitatRezhimy_Chern Material, hChern, DiametrFrezy, Rezhimy, Bfrez, iB, Korka
    
    Set RaschotRezhimov = Rezhimy

End Function

Private Sub RaschitatPripuski(ByVal Ra As Double, ByVal Pripusk As Double, ByRef hChist As Double, ByRef hChern As Double)

    hChist = OpredelitChistovoyPripusk(Ra, Pripusk)
    
    If Pripusk > hChist Then
        hChern = Pripusk - hChist
    End If
    
End Sub

Private Function OpredelitChistovoyPripusk(ByVal Ra As Double, ByVal Pripusk As Double) As Double

    Dim hChistNorm As Integer
    hChistNorm = IIf(Ra > 2.5, 2, 1)
    
    Dim hChist As Double
    If Ra <= 10 Then
        hChist = WorksheetFunction.Min(hChistNorm, Pripusk)
    Else
        hChist = 0
    End If
    
    OpredelitChistovoyPripusk = hChist
    
End Function

Private Function RaschitatGlubinyRezaniya(Dfrezy As Double, hChern As Double) As Double
    
    Dim GlubRez As Double

    Select Case Dfrezy
        Case 6:     GlubRez = VybratMaksGlubRez(hChern, 1.5, 3, 5)
        Case 8:     GlubRez = VybratMaksGlubRez(hChern, 1.5, 4, 6)
        Case 10:    GlubRez = VybratMaksGlubRez(hChern, 2, 5, 8)
        Case 12:    GlubRez = VybratMaksGlubRez(hChern, 2, 5, 10)
        Case 16:    GlubRez = VybratMaksGlubRez(hChern, 3, 6, 10, 14)
        Case 20:    GlubRez = VybratMaksGlubRez(hChern, 4, 8, 12, 18)
        Case 25:    GlubRez = VybratMaksGlubRez(hChern, 5, 10, 15, 20)
        Case 32:    GlubRez = VybratMaksGlubRez(hChern, 6, 13, 20, 28)
        Case Else:  GlubRez = VybratMaksGlubRez(hChern, 7, 15, 25, 36)
    End Select
    
    RaschitatGlubinyRezaniya = GlubRez
    If RaschitatGlubinyRezaniya > hChern Then RaschitatGlubinyRezaniya = hChern
    
End Function

Private Function VybratMaksGlubRez(Val As Double, ParamArray Steps()) As Double

    Dim i As Long
    For i = LBound(Steps) To UBound(Steps) - 1
        If Val <= Steps(i) Then
            VybratMaksGlubRez = Steps(i)
            Exit Function
        End If
    Next i
    VybratMaksGlubRez = Steps(UBound(Steps))
    
End Function


Private Function OpredelitShirinuFrezerovaniya(Dfrezy As Double) As Double

    Select Case Dfrezy
        Case Is <= 5:   OpredelitShirinuFrezerovaniya = 3
        Case Is <= 8:   OpredelitShirinuFrezerovaniya = 5
        Case Is <= 16:  OpredelitShirinuFrezerovaniya = 10
        Case Is <= 32:  OpredelitShirinuFrezerovaniya = 20
        Case Else:      OpredelitShirinuFrezerovaniya = 30
    End Select
    
End Function

Private Sub RaschitatRezhimy_Chern(Material As EnumMaterialy, hChern As Double, Dfrezy As Double, ByRef Rezhimy As clsMillingRegimes, Bfrez As Double, iB As Long, Korka As Boolean)

    Dim RezhimyOsnova As New clsRezhimy_FrezKonts_Osnova

    If Not hChern > 0 Then Exit Sub
    
    Dim GlubinaRez As Double
    GlubinaRez = RaschitatGlubinyRezaniya(Dfrezy, hChern)
    
    Dim iH As Long
    iH = WorksheetFunction.RoundUp(hChern / GlubinaRez, 0)
    
    GlubinaRez = hChern / iH
    
    If Korka Then
        With Rezhimy.ChernKork
            .Bfrez = Bfrez
            .GlubRez = GlubinaRez
            .iB = iB
            .iH = 1
            
            Set .Pokazateli = RezhimyOsnova.RaschotMinutnoyPodachi(1, Material, Dfrezy, .Bfrez, .GlubRez, True)
        End With
        iH = iH - 1
    End If
    
    If iH > 0 Then
        With Rezhimy.Chern
            .Bfrez = Bfrez
            .GlubRez = GlubinaRez
            .iB = iB
            .iH = iH
            
            Set .Pokazateli = RezhimyOsnova.RaschotMinutnoyPodachi(1, Material, Dfrezy, .Bfrez, .GlubRez, False)
        End With
    End If

End Sub

Private Sub RaschitatRezhimy_Chist(Material As EnumMaterialy, Bfrez As Double, hChist As Double, iB As Long, Dfrezy As Double, IT As Integer, ByRef ChistStage As clsMillingStage)
        
    If Not hChist > 0 Then Exit Sub
    
    Dim Osnova As New clsRezhimy_FrezKonts_Osnova
    
    With ChistStage
        .Bfrez = Bfrez
        .iB = iB
        .iH = IIf(IT <= 11, 2, 1)
        .GlubRez = hChist / .iH
        
        Set .Pokazateli = Osnova.RaschotMinutnoyPodachi(2, Material, Dfrezy, .Bfrez, .GlubRez, False)
    End With
    
End Sub

