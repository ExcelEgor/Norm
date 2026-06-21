Attribute VB_Name = "mShlifovaniye_RabotaSFormoy"
Option Explicit
Option Private Module

Dim idMaterial As Integer
Dim Dlina_Ploskoye As Double, Shirina_Ploskoye As Double, KolVo_Ploskoye As Double, Lmax_Ploskoe As Integer
Dim KolVo_Naruzhnoye As Double
Dim GlavnyyKoeffitsient As Double
Dim RezhimRedaktirovaniya As Boolean
Dim ShirinaBolshe As Boolean, DlinaBolshe As Boolean
Dim MaterialSlesar As Integer

Sub DobavitVTablitsu(lstShlifovaniye As Control, OpisaniyePerekhoda As String, tKontrol As String, Optional tSles As String, Optional ArrParametry)

    'Столбцы:
    '0-Переход; 1-tшлиф; 2-tконтр; 3-слес; 4-tшлиф*; 5-tконтр*; 6-tслес*; 7-Параметры

    Dim StrokaDobavanleniya As Integer
    
    With lstShlifovaniye
        If RezhimRedaktirovaniya Then
            StrokaDobavanleniya = .ListIndex
        Else
            .AddItem
            StrokaDobavanleniya = .ListCount - 1
        End If
        .List(StrokaDobavanleniya, 0) = OpisaniyePerekhoda
        .List(StrokaDobavanleniya, 4) = frmShlifonaiye.txtVremya_Material_1
        .List(StrokaDobavanleniya, 5) = tKontrol
        .List(StrokaDobavanleniya, 6) = tSles
        .List(StrokaDobavanleniya, 7) = Join(ArrParametry)
    End With
    
    Call RaschetTsht
            
    Call IzmenitRezhimRedaktirovaniya(False)
            
End Sub

Sub DobavitVTablitsu_Naruzhnoye()

    With frmShlifonaiye

        If IsNumeric(.txtVremya) Then
        
            Dim tKontrol As String
            tKontrol = Format(IzmereniyeMikrometrom(.txtDiametr_Naruzhnoye, .txtDlina_Naruzhnoye) * KolVo_Naruzhnoye, "0.00")
            
            Dim ArrParametry(0 To 10) As String
            ArrParametry(0) = ZnacheniyeParametra(.optCilindricheskaya_Naruzhnoye)
            ArrParametry(1) = ZnacheniyeParametra(.txtDiametr_Naruzhnoye)
            ArrParametry(2) = ZnacheniyeParametra(.txtDlina_Naruzhnoye)
            ArrParametry(3) = ZnacheniyeParametra(.cboPripusk_Naruzhnoye)
            ArrParametry(4) = ZnacheniyeParametra(.optProdolnaya_Naruzh)
            ArrParametry(5) = ZnacheniyeParametra(.cboIT_Naruzhnoye)
            ArrParametry(6) = ZnacheniyeParametra(.cboRa_Naruzhnoye)
            ArrParametry(7) = ZnacheniyeParametra(.cboDupusk_Naruzhnoye)
            ArrParametry(8) = ZnacheniyeParametra(.cboGaltel_Naruzhnoye)
            ArrParametry(9) = ZnacheniyeParametra(.cboShirinaKruga_Naruzhnoye)
            ArrParametry(10) = ZnacheniyeParametra(.txtKolVo_Naruzhnoye)
            
            Dim OpisaniyePerekhoda As String
            OpisaniyePerekhoda = "Шлифовать D" & frmShlifonaiye.txtDiametr_Naruzhnoye
            Call DobavitVTablitsu(.lstShlifovaniye_Naruzhnoye, OpisaniyePerekhoda, tKontrol, "", ArrParametry)

        End If
    
    End With

End Sub

Sub DobavitVTablitsu_Ploskoye()


    With frmShlifonaiye

        If IsNumeric(.txtVremya) Then
        
            Dim razmer As Double
            razmer = 10
            
            Dim KolVo As Double
            KolVo = KolVoPoverhnosteiProhodov(.txtKolVo_Ploskoye)
            
            Dim tKontrol As String
            tKontrol = Format( _
                (IzmereniyeMikrometrom(razmer, .txtDlina_Ploskoye) + IzmereniyeMikrometrom(razmer, .txtShirina_Ploskoye)) * KolVo, "0.00")
            
            Dim tSles As String
            tSles = Format(ZachistkaZausencev(CInt(MaterialSlesar), (DblFromCtrl(.txtDlina_Ploskoye) + DblFromCtrl(.txtShirina_Ploskoye)) * 2) * KolVo, "0.00")
        
            Dim ArrParametry(0 To 10) As String
            ArrParametry(0) = ZnacheniyeParametra(.optStolPryamougol)
            ArrParametry(1) = ZnacheniyeParametra(.optPereferiyaKruga)
            ArrParametry(2) = ZnacheniyeParametra(.txtShirina_Ploskoye)
            ArrParametry(3) = ZnacheniyeParametra(.txtDlina_Ploskoye)
            ArrParametry(4) = ZnacheniyeParametra(.txtKolVo_Ploskoye)
            ArrParametry(5) = ZnacheniyeParametra(.cboIT_Ploskoye)
            ArrParametry(6) = ZnacheniyeParametra(.cboRa_Ploskoye)
            ArrParametry(7) = ZnacheniyeParametra(.cboPripusk_Ploskoye)
            ArrParametry(8) = ZnacheniyeParametra(.cboDopuskPloskParallel)
            ArrParametry(9) = ZnacheniyeParametra(.cboTolschinaDetali_Ploskoye)
            ArrParametry(10) = ZnacheniyeParametra(.cboShirinaKruga_Ploskoye)
            
            Dim OpisaniyePerekhoda As String
            OpisaniyePerekhoda = "Шлифовать пл-ть"
            
            Call DobavitVTablitsu(.lstShlifovaniye_Ploskoye, OpisaniyePerekhoda, tKontrol, tSles, ArrParametry)

        End If
    
    End With
    
End Sub

Sub DobavitVTablitsu_Vnutrenneye()

    With frmShlifonaiye
            
        If IsNumeric(.txtVremya) Then
        
            Dim tKontrol As String
            tKontrol = Format(IzmereniyeNutromerom(.txtDiametr_Vnutrenneye, .txtDlina_Vnutrenneye) * KolVo_Naruzhnoye, "0.00")
        
            Dim ArrParametry(0 To 8) As String
            ArrParametry(0) = ZnacheniyeParametra(.optCilindricheskaya_Vnutrenneye)
            ArrParametry(1) = ZnacheniyeParametra(.txtDiametr_Vnutrenneye)
            ArrParametry(2) = ZnacheniyeParametra(.txtDlina_Vnutrenneye)
            ArrParametry(3) = ZnacheniyeParametra(.cboPripusk_Vnutrenneye)
            ArrParametry(4) = ZnacheniyeParametra(.cboIT_Vnutrenneye)
            ArrParametry(5) = ZnacheniyeParametra(.cboRa_Vnutrenneye)
            ArrParametry(6) = ZnacheniyeParametra(.cboDopusk_Vnutrenneye)
            ArrParametry(7) = ZnacheniyeParametra(.cboShirinaKruga_Vnutrenneye)
            ArrParametry(8) = ZnacheniyeParametra(.txtKolVo_Vnutrenneye)

            Dim OpisaniyePerekhoda As String
            OpisaniyePerekhoda = "Шлифовать D" & frmShlifonaiye.txtDiametr_Vnutrenneye
            Call DobavitVTablitsu(.lstShlifovaniye_Vnutrenneye, OpisaniyePerekhoda, tKontrol, , ArrParametry)
            
        End If
    
    End With

End Sub

Sub IzmeneniyeMateriala()

'    With frmShlifonaiye
'
'        If .cboMaterial.ListCount > 0 Then
'            idMaterial = .cboMaterial.List(.cboMaterial.ListIndex, .cboMaterial.ColumnCount - 1)
'        End If
'        MaterialSlesar = KonvertMaterial_Shlifovaniye_v_SlesarnoKarkasnyye(idMaterial)
'
'        Call RaschetGlavnogoKoeffitsienta
'
'    End With
    
End Sub

Sub IzmeneniyeTipaProizvodstva()

    With frmShlifonaiye
    
        If .cboTipProizvodstva.ListIndex = 3 Then
            .lblTipProizvodstva_KolVo.Caption = "более 100 деталей в партии"
        Else
            Dim KolVo As Variant
            Select Case .cboTipProizvodstva.ListIndex
                Case 0: KolVo = 3
                Case 1: KolVo = 20
                Case 2: KolVo = 100
            End Select
            .lblTipProizvodstva_KolVo.Caption = "до " & KolVo & " деталей в партии"
        End If
        
        Call RaschetGlavnogoKoeffitsienta
        
    End With
    
End Sub

Sub IzmenitRezhimRedaktirovaniya(Sostoyaniye As Boolean)

    RezhimRedaktirovaniya = Sostoyaniye

    With frmShlifonaiye

        .lstShlifovaniye_Naruzhnoye.Enabled = Not RezhimRedaktirovaniya
        Dim i As Integer
        For i = 0 To .mltShlifonaiye.Pages.Count - 1
            If i <> .mltShlifonaiye.Value Then
                .mltShlifonaiye.Pages(i).Enabled = Not RezhimRedaktirovaniya
            End If
        Next
        If RezhimRedaktirovaniya Then
            .cmdAdd_Ploskoye.Caption = " Редактировать"
            .cmdAdd_Ploskoye.Picture = .cmdPictureEdit.Picture
        Else
            .cmdAdd_Ploskoye.Caption = " Добавить"
            .cmdAdd_Ploskoye.Picture = .cmdPictureAdd.Picture
        End If
        
        .cmdOtmenitRedaktirovaniye.Visible = RezhimRedaktirovaniya
    
    End With
    
End Sub

Sub RaschetBmax_Ploskoe()

    ShirinaBolshe = False
    
    With frmShlifonaiye
    
        .lblBmax_Ploskoye.Caption = "до 300"
        .txtShirina_Ploskoye.BorderColor = vbActiveBorder
        
        Shirina_Ploskoye = DblFromCtrl(.txtShirina_Ploskoye)
        
        If Shirina_Ploskoye > MAKS_SHIRINA_SHLIFOVANIYA Then
            ShirinaBolshe = True
            .txtShirina_Ploskoye.BorderColor = vbRed
        End If
    
        .Repaint
        
    End With
    
End Sub

Function RaschetGlavnogoKoeffitsienta()

    With frmShlifonaiye
    
        If .cboTipProizvodstva.ListIndex <> -1 And .cboVozrast.ListIndex <> -1 Then
        
            Dim KolVoDetaleyVPartii As Double
            Select Case .cboTipProizvodstva.ListIndex
                Case 0: KolVoDetaleyVPartii = 3
                Case 1: KolVoDetaleyVPartii = 20
                Case 2: KolVoDetaleyVPartii = 100
                Case 3: KolVoDetaleyVPartii = 999
            End Select
            
            Dim VozrastStanka As Double
            Select Case .cboVozrast.ListIndex
                Case 0: VozrastStanka = 10
                Case 1: VozrastStanka = 20
                Case 2: VozrastStanka = 999
            End Select
            
            GlavnyyKoeffitsient = GlavnyyKoeffitsient_Shlifovaniye(KolVoDetaleyVPartii, .chkNezhestkaya.Value, VozrastStanka, .chkShlifovaniyeDvukhPoverhnostey.Value)
            
        Else
            
            GlavnyyKoeffitsient = 1
        
        End If
        
        If Not .tglIdotZagruzkaFormy Then
            Select Case .mltShlifonaiye.SelectedItem.Name
                Case "pNaruzhnoye"
                    Call RaschetVremeni_Naruzhnoye
                Case "pPloskoye"
                    Call RaschetVremeni_Ploskoe
                Case "pVnutrenneye"
                    Call RaschetVremeni_Vnutrenneye
            End Select
            Call RaschetTsht
        End If
    
    End With
    
End Function

Sub RaschetLmax_Ploskoe()

    Dim Shlifovaniye_Ploskoye As New clsShlifovaniye_Ploskoye

    DlinaBolshe = False

    With frmShlifonaiye
    
        .lblLmax_Ploskoye.Caption = "до 1500"
        .txtDlina_Ploskoye.BorderColor = vbActiveBorder
        
        Dlina_Ploskoye = DblFromCtrl(.txtDlina_Ploskoye)
        
        If Shirina_Ploskoye > 0 And Shirina_Ploskoye <= MAKS_SHIRINA_SHLIFOVANIYA Then
        
            Lmax_Ploskoe = Shlifovaniye_Ploskoye.RaschotLmax(CInt(idMaterial), Shirina_Ploskoye)
            .lblLmax_Ploskoye.Caption = "до " & Lmax_Ploskoe
            
            If Dlina_Ploskoye > Lmax_Ploskoe Then
                DlinaBolshe = True
                .txtDlina_Ploskoye.BorderColor = vbRed
            End If
            
        End If
        
        .Repaint
    
    End With
    
End Sub

Private Sub RaschetTsht()
    
    With frmShlifonaiye
    
        Dim Tsht As Double, Tsht_Kontrol As Double
        
        Dim KoefTipProizvKontrol As Double
        Select Case .cboTipProizvodstva.ListIndex
            Case 0:     KoefTipProizvKontrol = 1.3
            Case 1:     KoefTipProizvKontrol = 1
            Case Else:  KoefTipProizvKontrol = 0.7
        End Select
        
        Dim lstShlifovaniye As Control, txtTsht As Control, txtTsht_Kontrol As Control

        Dim i As Integer
        Select Case .mltShlifonaiye.SelectedItem.Name
            'Столбцы:
            '0-Переход; 1-tшлиф; 2-tконтр; 3-слес; 4-tшлиф*; 5-tконтр*; 6-tслес*; 7-Параметры
            Case "pPloskoye"
                Set lstShlifovaniye = .lstShlifovaniye_Ploskoye
                Set txtTsht = .txtTsht_Ploskoye
                Set txtTsht_Kontrol = .txtTsht_Kontrol_Ploskoye
                
            Case "pNaruzhnoye"
                Set lstShlifovaniye = .lstShlifovaniye_Naruzhnoye
                Set txtTsht = .txtTsht_Naruzhnoye
                Set txtTsht_Kontrol = .txtTsht_Kontrol_Naruzhnoye

            Case "pVnutrenneye"
                Set lstShlifovaniye = .lstShlifovaniye_Vnutrenneye
                Set txtTsht = .txtTsht_Vnutrenneye
                Set txtTsht_Kontrol = .txtTsht_Kontrol_Vnutrenneye
        End Select
        
    End With
    
    For i = 0 To lstShlifovaniye.ListCount - 1
        lstShlifovaniye.List(i, 1) = Format(Split(lstShlifovaniye.List(i, 4))(idMaterial - 1) * GlavnyyKoeffitsient, "0.00")
        Tsht = Tsht + CDbl(lstShlifovaniye.List(i, 1))
        lstShlifovaniye.List(i, 2) = Format(lstShlifovaniye.List(i, 5) * KoefTipProizvKontrol, "0.00")  'Контрольная операция
        If IsNumeric(lstShlifovaniye.List(i, 6)) Then
            lstShlifovaniye.List(i, 3) = Format(lstShlifovaniye.List(i, 6) * KoefTipProizvKontrol, "0.00")  'Слесарная операция
        End If
        Tsht_Kontrol = Tsht_Kontrol + CDbl(lstShlifovaniye.List(i, 2))
    Next
    
    txtTsht.text = IIf(Tsht > 0, WorksheetFunction.RoundUp(Tsht, 0), Empty)
    txtTsht_Kontrol = IIf(Tsht_Kontrol > 0, CStr(OkruglenieTsht(Tsht_Kontrol)), Empty)
    
End Sub

Sub RaschetVremeni_Naruzhnoye()

'    With frmShlifonaiye
'
'        If .tglIdotZagruzkaFormy Then Exit Sub
'
'        .txtVremya.text = Empty
'        .optRadialnaya_Naruzh.Enabled = True
'        .optProdolnaya_Naruzh.Enabled = True
'        Call ProverkaNaMaksimalnoyeZnacheniye(.txtDiametr_Naruzhnoye, D_MAKS_SHLIF_NARUZH_PROD, frmShlifonaiye)
'
'        Dim D As Double:        D = DblFromCtrl(.txtDiametr_Naruzhnoye)
'        Dim L As Double:        L = DblFromCtrl(.txtDlina_Naruzhnoye)
'        Dim Pripusk As Double:  Pripusk = DblFromCtrl(.cboPripusk_Naruzhnoye)
'
'        If D > 0 Then
'
'            Dim PripuskMaxRadialnoye As Double: PripuskMaxRadialnoye = PripuskMax_Shlifovanie_NaruzhnoyeRadialnoye(idMaterial, D)
'            Dim LmaxRadialnoe As Integer:       LmaxRadialnoe = Lmax_Shlifovanie_NaruzhnoyeRadialnoye(idMaterial, D)
'            Dim LmaxProdolnoye As Integer:      LmaxProdolnoye = Lmax_NaruzhnoyeShlifovaniye(DblFromCtrl(.txtDiametr_Naruzhnoye))
'
'            If L > LmaxRadialnoe Or Pripusk > PripuskMaxRadialnoye Or D > D_MAKS_SHLIF_NARUZH_RADIAL Then
'                .optProdolnaya_Naruzh.Value = True
'                .optRadialnaya_Naruzh.Enabled = False
'            End If
'
'            Dim Lmax As Double: Lmax = IIf(.optProdolnaya_Naruzh, LmaxProdolnoye, LmaxRadialnoe)
'            .lblLmax_Naruzhnoye.Caption = "до " & Lmax & " мм"
'
'            Call ProverkaNaMaksimalnoyeZnacheniye(.txtDlina_Naruzhnoye, Lmax, frmShlifonaiye)
'
'            If D <= D_MAKS_SHLIF_NARUZH_PROD And L > 0 And L <= Lmax And Pripusk > 0 And .cboRa_Naruzhnoye.ListIndex <> -1 Then
'
'                Dim IT As Integer:      IT = .cboIT_Naruzhnoye.List(.cboIT_Naruzhnoye.ListIndex, 1)
'                Dim Dopusk As Double:   Dopusk = IIf(.cboDupusk_Naruzhnoye.ListIndex = 0, 0, DblFromCtrl(.cboDupusk_Naruzhnoye))
'
'                Dim ShirinaKruga As Double
'                If .cboShirinaKruga_Naruzhnoye.ListIndex = 0 Then
'                    ShirinaKruga = 50
'                Else
'                    ShirinaKruga = CDbl(.cboShirinaKruga_Naruzhnoye)
'                End If
'
'                KolVo_Naruzhnoye = KolVoPoverhnosteiProhodov(.txtKolVo_Naruzhnoye)
'
'                Dim Vremya()
'                Dim i As Integer
'                Dim idMaterial_i As Integer
'
'                With .cboMaterial
'                    For i = 0 To .ListCount - 1
'                        ReDim Preserve Vremya(i)
'                        idMaterial_i = .List(i, .ColumnCount - 1)
'                        If frmShlifonaiye.optProdolnaya_Naruzh Then
'                            Vremya(i) = Shlifovanie_NaruzhnoyeProdolnoye(idMaterial_i, D, Pripusk, L, IT, frmShlifonaiye.cboRa_Naruzhnoye, _
'                                frmShlifonaiye.optKonicheskaya_Naruzhnoye.Value, Dopusk, frmShlifonaiye.cboGaltel_Naruzhnoye.ListIndex, ShirinaKruga)
'                        Else
'                            Vremya(i) = Shlifovanie_NaruzhnoyeRadialnoye(idMaterial_i, D, Pripusk, L, IT, CDbl(frmShlifonaiye.cboRa_Naruzhnoye), _
'                                frmShlifonaiye.optKonicheskaya_Naruzhnoye.Value, Dopusk, frmShlifonaiye.cboGaltel_Naruzhnoye.ListIndex)
'                        End If
'                        Vremya(i) = Vremya(i) * KolVo_Naruzhnoye
'                        Vremya(i) = WorksheetFunction.Round(Vremya(i), 2)
'                    Next
'                End With
'
'
'                .txtVremya_Material_1 = Join(Vremya)
'                .txtVremya = Format(Vremya(idMaterial - 1) * GlavnyyKoeffitsient, "0.00")
'
'            End If
'
'        End If
'
'    End With
'

End Sub

Sub RaschetVremeni_Ploskoe()

    With frmShlifonaiye
        
        .txtVremya.text = Empty
        
        If .tglIdotZagruzkaFormy Then Exit Sub
        
        Call RaschetBmax_Ploskoe
        Call RaschetLmax_Ploskoe
        
        If DlinaBolshe = True Or ShirinaBolshe = True Or .cboRa_Ploskoye.ListIndex = -1 Or .cboPripusk_Ploskoye.ListIndex = -1 Then Exit Sub
        
        Dim IT As Integer
        If .cboIT_Ploskoye.ListIndex = 1 Then
            IT = 6
        Else
            IT = CInt(.cboIT_Ploskoye)
        End If
            
        Dim ShirinaKruga As Double
        If .cboShirinaKruga_Ploskoye.ListIndex = 0 Then
            ShirinaKruga = 50
        Else
            ShirinaKruga = CDbl(.cboShirinaKruga_Ploskoye)
        End If
            
        Dim TolschinaDetali As Double
        Select Case .cboTolschinaDetali_Ploskoye.ListIndex
            Case 0: TolschinaDetali = 2
            Case 1: TolschinaDetali = 5
            Case 2: TolschinaDetali = 6
        End Select
            
        Dim DopuskPloskParallel As Double
        If .cboDopuskPloskParallel.ListIndex = 0 Then
            DopuskPloskParallel = 0
        Else
            DopuskPloskParallel = CDbl(.cboDopuskPloskParallel)
        End If
            
        KolVo_Ploskoye = KolVoPoverhnosteiProhodov(.txtKolVo_Ploskoye)
            
        Dim Vremya()
        Dim i As Integer
        Dim idMaterial_i As Integer
                
        With .cboMaterial
            For i = 0 To .ListCount - 1
                ReDim Preserve Vremya(i)
                idMaterial_i = .List(i, .ColumnCount - 1)
                Vremya(i) = Shlifovanie_Ploskoe(CInt(idMaterial_i), Shirina_Ploskoye, _
                    Dlina_Ploskoye, IT, CDbl(frmShlifonaiye.cboRa_Ploskoye), CDbl(frmShlifonaiye.cboPripusk_Ploskoye), ShirinaKruga, TolschinaDetali, DopuskPloskParallel) * KolVo_Ploskoye
                Vremya(i) = WorksheetFunction.Round(Vremya(i), 2)
            Next
        End With
        
        .txtVremya_Material_1 = Join(Vremya)
        .txtVremya = Format(Vremya(idMaterial - 1) * GlavnyyKoeffitsient, "0.00")

        Call RaschetTsht
    
    End With
    
End Sub

Sub RaschetVremeni_Vnutrenneye()

    With frmShlifonaiye
    
        If .tglIdotZagruzkaFormy Then Exit Sub
        
        .txtVremya.text = Empty
        
        Dim d As Double:        d = DblFromCtrl(.txtDiametr_Vnutrenneye)
        Call ProverkaNaMaksimalnoyeZnacheniye(.txtDiametr_Vnutrenneye, D_MAKS_VNUTR_SHLIF, frmShlifonaiye)
        If d > D_MAKS_VNUTR_SHLIF Then Exit Sub
        
        Dim L As Double:        L = DblFromCtrl(.txtDlina_Vnutrenneye)
        Dim Pripusk As Double:  Pripusk = DblFromCtrl(.cboPripusk_Vnutrenneye)
        Dim Ra As Double:       Ra = DblFromCtrl(.cboRa_Vnutrenneye)
        Dim IT As Integer:      IT = .cboIT_Naruzhnoye.List(.cboIT_Naruzhnoye.ListIndex, 1)
        Dim dopusk As Double:   dopusk = IIf(.cboDopusk_Vnutrenneye.ListIndex = 0, 0, DblFromCtrl(.cboDopusk_Vnutrenneye))
        Dim KolVo As Double:    KolVo = KolVoPoverhnosteiProhodov(.txtKolVo_Vnutrenneye)
        
        Dim ShirinaKruga As Double
        If .cboShirinaKruga_Vnutrenneye.ListIndex = 0 Then
            ShirinaKruga = 50
        Else
            ShirinaKruga = CDbl(.cboShirinaKruga_Vnutrenneye)
        End If
    
        If d > 0 And L > 0 And Pripusk > 0 And Ra > 0 And IT > 0 Then

            KolVo_Naruzhnoye = KolVoPoverhnosteiProhodov(.txtKolVo_Naruzhnoye)
            
            Dim Vremya()
            Dim i As Integer
            Dim idMaterial_i As Integer

            With .cboMaterial
                For i = 0 To .ListCount - 1
                    ReDim Preserve Vremya(i)
                    idMaterial_i = .List(i, .ColumnCount - 1)
                    Vremya(i) = Shlifovanie_Vnutrenneye(CInt(idMaterial_i), d, Pripusk, L, IT, Ra, _
                        frmShlifonaiye.optKonicheskaya_Vnutrenneye.Value, dopusk, ShirinaKruga)
                    Vremya(i) = Vremya(i) * KolVo
                    Vremya(i) = WorksheetFunction.Round(Vremya(i), 2)
                Next
            End With
            
    
            .txtVremya_Material_1 = Join(Vremya)
            .txtVremya = Format(Vremya(idMaterial - 1) * GlavnyyKoeffitsient, "0.00")
                
        End If
    
    End With
    

End Sub

Sub RedaktirovaniyePerekhoda()

    Call IzmenitRezhimRedaktirovaniya(True)

    Dim Parametry
    
    With frmShlifonaiye
        Select Case .mltShlifonaiye.SelectedItem.Name
            Case "pNaruzhnoye"
                Parametry = Split(.lstShlifovaniye_Naruzhnoye.List(.lstShlifovaniye_Naruzhnoye.ListIndex, 7))
            Case "pVnutrenneye"
                Parametry = Split(.lstShlifovaniye_Vnutrenneye.List(.lstShlifovaniye_Vnutrenneye.ListIndex, 7))
            Case "pBestsentrovoye"
            Case "pPloskoye"
                Parametry = Split(.lstShlifovaniye_Ploskoye.List(.lstShlifovaniye_Ploskoye.ListIndex, 7))
        End Select
    
        Dim i As Integer
        Dim ctrl As Control
        Dim ctrlName As String
        Dim ctrlValue
        For i = 0 To UBound(Parametry)
            ctrlName = Split(Parametry(i), ":")(0)
            ctrlValue = Split(Parametry(i), ":")(1)
            Set ctrl = .Controls(ctrlName)
            Select Case TypeName(ctrl)
                Case "TextBox"
                    ctrl.text = ctrlValue
                Case "ComboBox"
                    ctrl.ListIndex = CInt(ctrlValue)
                Case "OptionButton"
                    ctrl.Value = ctrlValue
            End Select
        Next
    
    End With
    
End Sub

Sub SpisokPripuskov_Naruzhnoye()

    With frmShlifonaiye

        .cboPripusk_Naruzhnoye.Clear
        .cboPripusk_Naruzhnoye.Enabled = False
        
        Dim d As Double
        d = DblFromCtrl(.txtDiametr_Naruzhnoye)
        
        If d > 0 Then
            If d <= 10 Then
                .cboPripusk_Naruzhnoye.List = Split("0,1 0,3")
            Else
                If idMaterial = 1 Then
                    .cboPripusk_Naruzhnoye.List = Split("0,3 0,4 0,5 0,6")
                Else
                    .cboPripusk_Naruzhnoye.List = Split("0,4 0,5 0,6")
                End If
            End If
            .cboPripusk_Naruzhnoye.Enabled = True
            .cboPripusk_Naruzhnoye.ListIndex = .cboPripusk_Naruzhnoye.ListCount - 1
        End If
    
    End With
    
End Sub

Sub SpisokPripuskov_Vnutrenneyee()

    With frmShlifonaiye
    
        .cboPripusk_Vnutrenneye.Clear
        .cboPripusk_Vnutrenneye.Enabled = False
        Dim d As Double
        d = DblFromCtrl(.txtDiametr_Vnutrenneye)
        If d > 0 And d <= D_MAKS_VNUTR_SHLIF Then
            With .cboPripusk_Vnutrenneye
                If d <= 25 Then
                    .List = Split("0,3 0,5")
                Else
                    .List = Split("0,5 0,6")
                End If
                .Enabled = True
                .ListIndex = 0
            End With
        End If
    
    End With
    
End Sub

Sub UdaleniyeIzTablitsy()
    
    Dim ActTable As Control
    
    With frmShlifonaiye
        Select Case .mltShlifonaiye.SelectedItem.Name
            Case "pNaruzhnoye"
                Set ActTable = .lstShlifovaniye_Naruzhnoye
            Case "pVnutrenneye"
                Set ActTable = .lstShlifovaniye_Vnutrenneye
            Case "pBestsentrovoye"
                Set ActTable = .lstShlifovaniye_Bestsentrovoye
            Case "pPloskoye"
                Set ActTable = .lstShlifovaniye_Ploskoye
        End Select
    End With
    
    Dim i As Integer

    With ActTable
        If .ListIndex <> -1 Then
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    .RemoveItem (i)
                End If
            Next
        End If
    End With
    
    Call RaschetTsht
End Sub

Function ZnacheniyeParametra(ctrl As Control) As String
    Select Case TypeName(ctrl)
        Case "OptionButton"
            ZnacheniyeParametra = ctrl.Name & ":" & ctrl.Value
        Case "TextBox"
            ZnacheniyeParametra = ctrl.Name & ":" & ctrl.text
        Case "ComboBox"
            ZnacheniyeParametra = ctrl.Name & ":" & ctrl.ListIndex
    End Select
End Function

