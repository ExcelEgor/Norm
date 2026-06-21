Attribute VB_Name = "mFrez_ObrOtv"
'@Folder Frezerovaniye

Option Explicit

Private Function Sverleniye_Frezernyy(ByVal Material As EnumMaterialy, d As Double, L As Double, Optional SovmeshcheniyeOsi As Boolean = False, Optional ByVal TipStanka As EnumTipStankaRastFrez = 2, _
    Optional Korka As Boolean = False, Optional NaUdar As Boolean = False, Optional NaOdnoyOsiSvyshePyati As Boolean = False) As Double

    If d <= 0 Or L <= 0 Or _
        MaterialKorrektnyy_RastFrez(Material) = False Or _
        TipStanka <> GorizontalnoRastochnoy And TipStanka <> VertikalnoGorizontalnoFrezernyy Or _
        d > Dmax_Sverleniye_Frezernyy(TipStanka) Then Exit Function
        
    If L > Lmax_Sverleniye_Frezernyy(d, TipStanka) Then Exit Function
    
    Dim TipStanka_DlyaPoiska As EnumTipStankaRastFrez
    TipStanka_DlyaPoiska = IzmenitTipStanka(TipStanka, Material)
    
    Dim Material_DlyaPoiska As EnumMaterialy
    Material_DlyaPoiska = IzmenitMaterialDlyaZaprosa_RastFrez(Material, False)
    
    Dim PoprKoef As Double
    PoprKoef = RaschotObshchegoKoeffitsienta_RastFrez(Material, TipStanka, NaUdar, False) * IIf(NaOdnoyOsiSvyshePyati, 0.8, 1) * IIf(Korka, K_KORKA_FREZ_PLOSK, 1)
    
    Dim Normativ
    Normativ = ZagruzitNormativ("frez_sverleniye")
    
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = TipStanka_DlyaPoiska And Normativ(i, 2) = 1 And Normativ(i, 3) = Material_DlyaPoiska And _
            Normativ(i, 4) >= d And Normativ(i, 5) >= L And Normativ(i, 6) = SovmeshcheniyeOsi Then
            Sverleniye_Frezernyy = PoprKoef * Normativ(i, 7)
            Exit For
        End If
    Next
    
End Function

Private Function Dmax_Sverleniye_Frezernyy(TipStanka As EnumTipStankaRastFrez) As Integer
    Dmax_Sverleniye_Frezernyy = IIf(TipStanka = 2, 55, 30)
End Function

Function Lmax_Sverleniye_Frezernyy(d As Double, TipStanka As EnumTipStankaRastFrez) As Integer

    If d > Dmax_Sverleniye_Frezernyy(TipStanka) Then Exit Function

    Dim Lmax As Integer
    
    If TipStanka = 2 And d > 20 Then    'Фрезерный
        Lmax = 160
    Else
        Select Case d
            Case Is <= 6:   Lmax = 30
            Case Is <= 10:  Lmax = 55
            Case Is <= 15:  Lmax = 100
            Case Is <= 20:  Lmax = 125
            Case Is <= 30:  Lmax = 160
            Case Is <= 40:  Lmax = 200
            Case Else:      Lmax = 160
        End Select
    End If
    
    Lmax_Sverleniye_Frezernyy = Lmax
    
End Function

Function ObrabotkaOtverstiiNaFrezernomStanke(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, d As Double, L As Double, _
    Optional Glukhoye As Boolean = False, Optional Dobnizki As Double = 0, Optional ByVal GlubinaObnizki As Double = 0, _
    Optional IT As Integer = 14, Optional Ra As Double = 10, _
    Optional Tonkosten As Boolean = False, _
    Optional Frezerovat As Boolean = True, _
    Optional Faska As Boolean)
    
    Dim Normy(1 To 3) As Double

    If Frezerovat And d <= MIN_D_FREZ Then
        Normy(1) = 0
        Normy(2) = 0
        Normy(3) = 0
        ObrabotkaOtverstiiNaFrezernomStanke = Normy
        Exit Function
    End If
    
    If d > 0 And L > 0 Then
        Dim dSverleniya As Double
        If d > 20 Then
            Select Case L
                Case Is <= 125: dSverleniya = 20
                Case Else:      dSverleniya = 30
            End Select
            dSverleniya = 20
        Else
            dSverleniya = d
        End If
        
        If L > Lmax_Sverleniye_Frezernyy(dSverleniya, TipStanka) Then Exit Function
    
        Dim tSverleniya As Double
        tSverleniya = Sverleniye_Frezernyy(Material, CDbl(dSverleniya), L, , TipStanka, , , False)
        
        Dim tFrez As Double
        If Frezerovat Or d > MAKS_D_SVERL_FREZ_STANOK Then
            Dim FrezOtv As New clsRezhimy_FrezKonts
            tFrez = FrezOtv.RaschotVremeni_FrezOtv(Material, TipStanka, FrezOtv.OpredelitDfrezy_FrezOtv(dSverleniya), d, dSverleniya, L, Ra, IT)
        Else
            tFrez = 0
        End If
        
        Dim tFaska As Double
        If Faska Then tFaska = ZenkovaniyeFasokNaSverlilnovStanke(d, True)
        
        Dim tSles As Double
        tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, IIf(Glukhoye, 1, 2), False)
    
        Dim tKontrol As Double
        tKontrol = KontrolOtverstiy(d, L, IT)
    End If

    Dim tObnizki As Double
    If Dobnizki > 0 And GlubinaObnizki > 0 Then
        tObnizki = FrezUstupov_Kompleks(TipStanka, CInt(Material), PI * Dobnizki, GlubinaObnizki, 0.5 * (Dobnizki - d), IT, Ra, , False, Tonkosten)(1)
        tSles = SnyatiyeZausentsevPoKonturu(Dobnizki * PI, Material, True, NADFIL, False)
    Else
        tObnizki = 0
    End If
    
    Normy(1) = tFrez + tSverleniya + tObnizki + tFaska
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    ObrabotkaOtverstiiNaFrezernomStanke = Normy
    
End Function

Function ObrabotkaRezbovihOtverstiiNaFrezernimStanke(TipStanka As EnumTipStankaRastFrez, Material As EnumMaterialy, ByVal d As Double, Lotv As Double, Lrezb As Double, _
    Optional Glukhoye As Boolean = False, Optional Dobnizki As Double = 0, Optional GlubinaObnizki As Double = 0, _
    Optional IT As Integer = 7, Optional ByVal Shag As Double = 0, _
    Optional Tonkosten As Boolean = False, Optional KalibrRezbSlesOper As Boolean = False, Optional Frezerovat As Boolean = True)
    
    If d > MAKS_D_REZBY_FREZERNYY Or Lotv > Lmax_Sverleniye_Frezernyy(d, TipStanka) Or Lrezb > Lmax_Metchik(d) _
        Or Not d > 0 Or Not Lrezb > 0 Then Exit Function
        
    If Shag = 0 Then Shag = KrupShagRezb(d)
    
    Dim tSverleniya As Double
    tSverleniya = ObrabotkaOtverstiiNaFrezernomStanke(TipStanka, Material, d, Lotv, Glukhoye, Dobnizki, GlubinaObnizki, 11, 5, Tonkosten, Frezerovat, True)(1)
    
    If tSverleniya <= 0 And Lotv > 0 Then Exit Function
    
    Dim tNarez As Double
    If d <= 6 And Material = ALUMINIYEVYYE_SPLAVY Then
        tNarez = Datron_NarezaniyeRezby(Material, d, Lrezb, Shag)
    Else
        tNarez = NarezaniyeRezbyMetchikom(CInt(Material), d, Lrezb, IT, Glukhoye, Shag)(1)
    End If
    
    Dim tFrez As Double
    tFrez = tSverleniya + tNarez
    
    Dim tSles As Double
    tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, IIf(Glukhoye, 1, 2))
    If KalibrRezbSlesOper Then
        tSles = tSles + ProgonkaRezby(CInt(Material), d, Lrezb)
    End If
    
    Dim tKontrol As Double
    tKontrol = Kontrol_RezbovayaProbka(d, Shag, IT)
    
    Dim Normy(1 To 3) As Double
    Normy(1) = tFrez
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    If Dobnizki > 0 And GlubinaObnizki > 0 Then
        tSles = tSles + SnyatiyeZausentsevPoKonturu(Dobnizki * PI, Material, True, NADFIL, False)
        tKontrol = tKontrol + KontrolOtverstiy(Dobnizki, GlubinaObnizki, IT)
    End If
    
    ObrabotkaRezbovihOtverstiiNaFrezernimStanke = Normy
        
End Function


