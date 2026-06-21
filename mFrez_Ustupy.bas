Attribute VB_Name = "mFrez_Ustupy"
'@Folder Frezerovaniye

Option Explicit
Private Const MAX_GLUBINA_REZANIYA_USTUP As Integer = 20

Function FrezUstupov_Kompleks(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, Dlina As Double, VysotaUstupa As Double, Pripusk As Double, _
    Optional ByVal IT As Integer = 11, Optional ByVal Ra As Double = 10, _
    Optional Udar As Boolean = False, Optional Korka As Boolean = False, Optional Tonkosten As Boolean = False)
    
    If Not Dlina > 0 Or Not VysotaUstupa > 0 Then Exit Function
        
    If IT = 0 Then IT = 11
    Ra = MinRa_RastFrez(Ra, IT)
        
    Dim MinPripusk As Integer:  MinPripusk = 1
    Dim MinVysotaUstupa As Integer:  MinVysotaUstupa = 6
    
    If VysotaUstupa < MinVysotaUstupa Then VysotaUstupa = MinVysotaUstupa
    If Pripusk < MinPripusk Then Pripusk = MinPripusk
    
    Dim Bfrez As Double
    Bfrez = ShirinaFrezerovaniya_Ustup(VysotaUstupa)
    
    Dim h_Chist As Double, t_chist As Double
    If Ra <= 10 Then
    
        If Ra = 5 Then
            h_Chist = MinPripusk
        Else
            h_Chist = 2
        End If
        
        t_chist = FrezUstupov_Chistovoye(TipStanka, CInt(Material), Dlina, Bfrez, Ra, Udar, Tonkosten) * IIf(IT <= 11, 2, 1)

    Else
        t_chist = 0
        h_Chist = 0
    End If
    
    Dim h_Chern As Double, h_ChernKork As Double, i_chern As Long
    Dim t_ChernKork As Double, t_chern As Double
    Dim GlubinaRez As Double

    h_Chern = Pripusk - h_Chist
    If Korka And Not h_Chern > 0 Then h_Chern = MinPripusk
    
    If h_Chern > 0 Then
        GlubinaRez = GlubinaRez_Ustup(h_Chern)
        If Korka Then
            h_ChernKork = GlubinaRez
            t_ChernKork = FrezUstupov_Chernovoye(TipStanka, CInt(Material), Dlina, Bfrez, GlubinaRez, Udar, Korka, Tonkosten)
        Else
            h_ChernKork = 0
            t_ChernKork = 0
        End If
        h_Chern = h_Chern - h_ChernKork
        If h_Chern > 0 Then
            i_chern = WorksheetFunction.RoundUp(h_Chern / GlubinaRez, 0)
            t_chern = FrezUstupov_Chernovoye(TipStanka, CInt(Material), Dlina, Bfrez, GlubinaRez, Udar, False, Tonkosten) * i_chern
        End If
    Else
        t_ChernKork = 0
        t_chern = 0
    End If
    
    Dim iB As Integer:      iB = WorksheetFunction.RoundUp(VysotaUstupa / Bfrez, 0)
    
    Dim tFrez As Double:    tFrez = (t_ChernKork + t_chern + t_chist) * iB
    Dim tSles As Double:    tSles = ZachistkaZausencev_PoKonturu_Napilnikom(Material, 2 * (Pripusk + VysotaUstupa + Dlina), False, False)
    Dim tKontrol As Double: tKontrol = IzmereniyePosleFrezerovaniya(VysotaUstupa, Dlina, Pripusk, IT)
    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    FrezUstupov_Kompleks = Normy
    
End Function


Function FrezUstupov_Chernovoye(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, Dlina As Double, ShirinaFrezerovaniya As Double, GlubinaRez As Double, _
    Optional FrezNaUdar As Boolean = False, Optional FrezPoKorke As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False)
    'Фрезерование уступов, выемок концевыми фрезами (карты 30, 31)
    
    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        ShirinaFrezerovaniya > MAX_SHIRINA_KONTS_USTUP_PAZ Or GlubinaRez > MAX_GLUBINA_REZANIYA_USTUP Then Exit Function
        
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)

    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_ustupy")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaFrezerovaniya And Normativ(i, 4) = 20 And Normativ(i, 5) >= GlubinaRez Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, FrezNaUdar, TonkostennayaKonfiguraciya) * IIf(FrezPoKorke, K_KORKA_FREZ_PLOSK, 1)
    
    FrezUstupov_Chernovoye = PoprKoef * (a * Dlina + b)
    
End Function

Function FrezUstupov_Chistovoye(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, Dlina As Double, ShirinaFrezerovaniya As Double, _
    Optional Ra As Double = 10, Optional FrezNaUdar As Boolean = False, Optional TonkostennayaKonfiguraciya As Boolean = False)
    'Фрезерование уступов, выемок концевыми фрезами (карты 30, 31)
    
    If MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        ShirinaFrezerovaniya > MAX_SHIRINA_KONTS_USTUP_PAZ Or Ra < MIN_RA_FREZ Then Exit Function
        
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)

    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_ustupy")
    
    Dim i As Long
    Dim a As Double, b As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = Material_DlyaPoiska And _
            Normativ(i, 3) >= ShirinaFrezerovaniya And Normativ(i, 4) <= Ra Then
            a = Normativ(i, 6)
            b = Normativ(i, 7)
            Exit For
        End If
    Next

    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, FrezNaUdar, TonkostennayaKonfiguraciya)
    
    FrezUstupov_Chistovoye = PoprKoef * (a * Dlina + b)
    
End Function

Private Function ShirinaFrezerovaniya_Ustup(Shirina As Double) As Integer
    Dim Bfrez As Integer
    Select Case Shirina
        Case Is <= 6:   Bfrez = 6
        Case Is <= 12:  Bfrez = 12
        Case Is <= 25:  Bfrez = 25
        Case Else
            If WorksheetFunction.RoundUp(Shirina / 25, 0) = WorksheetFunction.RoundUp(Shirina / 40, 0) And Shirina / 25 <> Int(Shirina / 25) Then
                Bfrez = 25
            Else
                Bfrez = 40
            End If
    End Select
    ShirinaFrezerovaniya_Ustup = Bfrez
End Function

Private Function GlubinaRez_Ustup(PripuskChernovoy As Double)
    
    Select Case PripuskChernovoy
        Case Is <= 3:   GlubinaRez_Ustup = 3
        Case Is <= 5:   GlubinaRez_Ustup = 5
        Case Is <= 7:   GlubinaRez_Ustup = 7
        Case Is <= 10:  GlubinaRez_Ustup = 10
        Case Is <= 15:  GlubinaRez_Ustup = 15
        Case Else:
            If WorksheetFunction.RoundUp(PripuskChernovoy / 15, 0) = _
                WorksheetFunction.RoundUp(PripuskChernovoy / 20, 0) And _
                PripuskChernovoy / 15 <> Int(PripuskChernovoy / 15) Then
                GlubinaRez_Ustup = 15
            Else
                GlubinaRez_Ustup = 20
            End If
    End Select

End Function

