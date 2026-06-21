VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKoordinat 
   Caption         =   "Координатно-расточне станки"
   ClientHeight    =   11460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   21030
   OleObjectBlob   =   "frmKoordinat.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKoordinat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IdotZagruzka As Boolean
Dim Parametry_Koordinat As Koordinat
Dim ParametryUstanovki As ParameryUstanovki_Koordinat
Dim Material As Integer
Dim Vremya As Variant, tKoordinat As Double

Dim DmaxVytochki As Integer
Dim KolVo As Double

Dim newWB As Workbook
Dim tbKoordinat As ListObject
Dim rMaterial As Range, rTipProizvodstva As Range, rTipStanka As Range, rVozrast As Range, rMaterialInstrumenta As Range

Dim RezhimRedaktirovaniya As Boolean

Dim ArrTxt() As New clsmZapuskRaschetov
Dim ArrCbo() As New clsmZapuskRaschetov
Dim ArrOpt() As New clsmZapuskRaschetov
Dim ArrChk() As New clsmZapuskRaschetov
Dim ArrMlt() As New clsmZapuskRaschetov

Dim Stolbets_idPerekhoda As Integer
Dim Stolbets_Perekhod As Integer
Dim Stolbets_D As Integer
Dim Stolbets_L As Integer
Dim Stolbets_SovmeshcheniyeOsi As Integer
Dim Stolbets_PredvaritelnayaTsentrovka As Integer
Dim Stolbets_Glukhoye As Integer
Dim Stolbets_NaklonnayaPloskost As Integer
Dim Stolbets_TonkostennyaDetal As Integer
Dim Stolbets_NaUdar As Integer
Dim Stolbets_IT
Dim Stolbets_Ra
Dim Stolbets_BieniyeaSoosnostKruglost
Dim Stolbets_ZakalonnayaStal
Dim Stolbets_KolVoDopProhodov
Dim Stolbets_KolVo As Integer
Dim Stolbets_tKoordinat As Integer
Dim Stolbets_tSlesar As Integer
Dim Stolbets_tKontrol As Integer
Dim Stolbets_Massa As Integer
Dim Stolbets_Sposob As Integer
Dim Stolbets_HarakterViverki As Integer
Dim Stolbets_Nezestk As Integer
Dim Stolbets_TochnostViverki As Integer
Dim Stolbets_KolVoDopBoltov As Integer
Dim Stolbets_KolVoDopDomkratov As Integer
Dim Stolbets_KolVoUstDetaley As Integer

Dim Tablitsa As clsTablitsa

Private Sub UserForm_Initialize()

    IdotZagruzka = True
    
    Application.ScreenUpdating = False
    wsKoordinat_Tablitsa.Copy
    Set newWB = ActiveWorkbook
    newWB.Windows(1).Visible = False
    Application.ScreenUpdating = True
    
    Call PrisvoeniyeDiapazonov
   
    Call PrisvoitNomeraStolbtsov
    
    Set Tablitsa = New clsTablitsa
    With Tablitsa
        Set .lstRaschot = lstKoordinat
        Set .txtTsht_Kontrol = txtTsht_Kontrol
        Set .txtTsht_Mekhanika = txtTsht_Koordinat
        Set .txtTsht_Slesar = txtTsht_Slesar
        .Stolbets_Perekhod = 3
        .Stolbets_tKontrol = tbKoordinat.ListColumns("tконтроль").Index
        .Stolbets_tMekhanika = tbKoordinat.ListColumns("tкоординат").Index
        .Stolbets_tSlesar = tbKoordinat.ListColumns("tслесар").Index
    End With
    
    Call ZapolnitElementyUpravleniya_Koordinat
    
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok(mltKoordinat)
    Call ZagruzkaElementovVKlass
    
    Call VyravnitElementyFormy_Koordinat(Me)
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    lblDmax_Rastach.Caption = "до " & DMAX_RASTACHIVANIYE_KOORDINAT & " мм"

    
    IdotZagruzka = False
    
End Sub

Private Sub PrisvoitNomeraStolbtsov()

    With tbKoordinat
    
        Stolbets_Perekhod = .ListColumns("Переход").Index
        Stolbets_idPerekhoda = .ListColumns("id перехода").Index
        Stolbets_D = .ListColumns("D").Index
        Stolbets_L = .ListColumns("L").Index
        Stolbets_SovmeshcheniyeOsi = .ListColumns("SovmeshcheniyeOsi").Index
        Stolbets_PredvaritelnayaTsentrovka = .ListColumns("PredvaritelnayaTsentrovka").Index
        Stolbets_Glukhoye = .ListColumns("Gluhoe").Index
        Stolbets_NaklonnayaPloskost = .ListColumns("NaklonnayaPloskost").Index
        Stolbets_TonkostennyaDetal = .ListColumns("TonkostennyaDetal").Index
        Stolbets_NaUdar = .ListColumns("NaUdar").Index
        Stolbets_IT = .ListColumns("IT").Index
        Stolbets_Ra = .ListColumns("Ra").Index
        Stolbets_BieniyeaSoosnostKruglost = .ListColumns("BieniyeaSoosnostKruglost").Index
        Stolbets_ZakalonnayaStal = .ListColumns("ZakalonnayaStal").Index
        Stolbets_KolVoDopProhodov = .ListColumns("KolVoDopProhodov").Index
        Stolbets_KolVo = .ListColumns("Кол-во").Index
        Stolbets_tKoordinat = .ListColumns("tкоординат").Index
        Stolbets_tSlesar = .ListColumns("tслесар").Index
        Stolbets_tKontrol = .ListColumns("tконтроль").Index
        Stolbets_Massa = .ListColumns("Massa").Index
        Stolbets_Sposob = .ListColumns("Sposob").Index
        Stolbets_HarakterViverki = .ListColumns("HarakterViverki").Index
        Stolbets_Nezestk = .ListColumns("Nezestk").Index
        Stolbets_TochnostViverki = .ListColumns("TochnostViverki").Index
        Stolbets_KolVoDopBoltov = .ListColumns("KolVoDopBoltov").Index
        Stolbets_KolVoDopDomkratov = .ListColumns("KolVoDopDomkratov").Index
        Stolbets_KolVoUstDetaley = .ListColumns("KolVoUstDetaley").Index
        
    End With
    
End Sub

Private Sub ZagruzkaElementovVKlass()

    Dim ctrl As Control
    Dim iTxt As Integer, iCbo As Integer, iOpt As Integer, iChk As Integer, iMlt As Integer
    
    For Each ctrl In Me.Controls
        If LCase(ctrl.Tag) <> "notclsm" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    ReDim Preserve ArrTxt(iTxt)
                    Set ArrTxt(iTxt).txtKoordinat = ctrl
                    iTxt = iTxt + 1
                Case "OptionButton"
                    ReDim Preserve ArrOpt(iOpt)
                    Set ArrOpt(iOpt).optKoordinat = ctrl
                    iOpt = iOpt + 1
                Case "ComboBox"
                    ReDim Preserve ArrCbo(iCbo)
                    Set ArrCbo(iCbo).cboKoordinat = ctrl
                    iCbo = iCbo + 1
                Case "CheckBox"
                    ReDim Preserve ArrChk(iChk)
                    Set ArrChk(iChk).chkKoordinat = ctrl
                    iChk = iChk + 1
                Case "MultiPage"
                    ReDim Preserve ArrMlt(iMlt)
                    Set ArrMlt(iMlt).mltKoordinat = ctrl
                    iMlt = iMlt + 1
            End Select
        End If
    Next
    
End Sub

Private Sub cboMaterialInstrument_Koordinat_Change()
    lblKoefMaterialInstriment_Koordinat.Caption = Format(PoprKoef_Koordinat_MatrialInstrument(cboMaterialInstrument_Koordinat.ListIndex + 1), "0.0")
    Call ZagruzkaVListBox
End Sub

Private Sub cboTipProizv_Koordinat_Change()
    lblKoefTipProizv_Koordinat.Caption = Format(PoprKoef_Koordinat_TipProizv(cboTipProizv_Koordinat.ListIndex + 1), "0.0")
    Call ZagruzkaVListBox
End Sub

Private Sub cboTipStanka_Koordinat_Change()
    lblKoefTipStanka_Koordinat.Caption = Format(PoprKoef_Koordinat_TipStanka(cboTipStanka_Koordinat.ListIndex + 1), "0.0")
    Call ZagruzkaVListBox
End Sub

Private Sub cboVozrast_Koordinat_Change()
    lblKoefVozrast_Koordinat.Caption = Format(PoprKoef_Koordinat_Vozrast(cboVozrast_Koordinat.ListIndex + 1), "0.0")
    Call ZagruzkaVListBox
End Sub

Private Sub cmdCopy_Click()

    If lstKoordinat.ListIndex = -1 Then Exit Sub
        
    Dim i As Integer
    Dim Stroka As Integer
    
    For i = 0 To lstKoordinat.ListCount - 1
        If lstKoordinat.Selected(i) Then
            Stroka = i + 1
            Exit For
        End If
    Next
    
    With tbKoordinat
        .ListRows.Add
        .ListRows(Stroka).Range.Copy
        .ListRows(.ListRows.Count).Range.PasteSpecial (xlPasteAll)
    End With
    
    Call ZagruzkaVListBox
    
End Sub

Private Sub cmdDelete_Click()
    Call Tablitsa.UdalitStroki(tbKoordinat)
End Sub

Private Sub cmdDown_Click()
    Call Tablitsa.PeremestitStrokuVniz(tbKoordinat)
End Sub

Private Sub cmdUp_Click()
    Call Tablitsa.PeremestitStrokuVverkh(tbKoordinat)
End Sub

Private Sub spnDopBolt_Koordinat_SpinDown()
    If txtDopBolt_Koordinat > 0 Then txtDopBolt_Koordinat = txtDopBolt_Koordinat - 1
End Sub

Private Sub spnDopBolt_Koordinat_SpinUp()
    txtDopBolt_Koordinat = txtDopBolt_Koordinat + 1
End Sub

Private Sub spnDopDomkrat_Koordinat_SpinDown()
    If txtDopDomkrat_Koordinat > 0 Then txtDopDomkrat_Koordinat = txtDopDomkrat_Koordinat - 1
End Sub

Private Sub spnDopDomkrat_Koordinat_SpinUp()
    txtDopDomkrat_Koordinat = txtDopDomkrat_Koordinat + 1
End Sub

Private Sub spnKolDet_Koordinat_SpinDown()
    If txtKolDet_Koordinat > 1 Then txtKolDet_Koordinat = txtKolDet_Koordinat - 1
End Sub

Private Sub spnKolDet_Koordinat_SpinUp()
    If txtKolDet_Koordinat < 20 Then txtKolDet_Koordinat = txtKolDet_Koordinat + 1
End Sub

Private Sub cmdDobavit_Click()
    
    If tKoordinat > 0 Then
        
        With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
        End With
        
        Call DobavitVTablitsu
        
        With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
        End With
        
        Call ZagruzkaVListBox
        
        RezhimRedaktirovaniya = False
        
        Call IzmenitRezhimFormy
        
    End If
    
End Sub
Private Sub ZagruzkaVListBox()

    If IdotZagruzka Then Exit Sub
    
    If tbKoordinat.ListRows.Count = 0 Then
        lstKoordinat.Clear
    
    Else
        
        rMaterial = Material
        rTipProizvodstva = cboTipProizv_Koordinat.ListIndex + 1
        rTipStanka = cboTipStanka_Koordinat.ListIndex + 1
        rVozrast = cboVozrast_Koordinat.ListIndex + 1
        rMaterialInstrumenta = cboMaterialInstrument_Koordinat.ListIndex + 1
        
        Application.DisplayAlerts = False
        newWB.Worksheets(1).Cells.Replace What:=newWB.Name & "!", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
        Application.DisplayAlerts = True
        tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count).DataBodyRange.Formula = tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count).DataBodyRange.Formula
        tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count - 1).DataBodyRange.Formula = tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count - 1).DataBodyRange.Formula
        tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count - 2).DataBodyRange.Formula = tbKoordinat.ListColumns(tbKoordinat.ListColumns.Count - 2).DataBodyRange.Formula

        Call Tablitsa.ZagruzkaVListBox(tbKoordinat)
    
    End If

End Sub
Private Sub ZagruzkaRaschota()

    cboMaterial.ListIndex = rMaterial - 1
    cboTipProizv_Koordinat.ListIndex = rTipProizvodstva - 1
    cboTipStanka_Koordinat.ListIndex = rTipStanka - 1
    cboVozrast_Koordinat.ListIndex = rVozrast - 1
    cboMaterialInstrument_Koordinat.ListIndex = rMaterialInstrumenta - 1

    Call ZagruzkaVListBox
    
End Sub

Private Sub DobavitVTablitsu()

    Dim NomerStroki As Integer
    Dim DiapazonParametrov As Range
    
    With tbKoordinat
        If RezhimRedaktirovaniya Then
            NomerStroki = lstKoordinat.ListIndex + 1
        Else
            .ListRows.Add
            NomerStroki = .ListRows.Count
        End If
        Set DiapazonParametrov = .ListRows(NomerStroki).Range
        .DataBodyRange.Orientation = 0
    End With
    
    Dim Perekhod As String
    Dim id As Integer
    
    With Parametry_Koordinat
    
        Select Case mltKoordinat.SelectedItem.Name
            Case "pUstanovka"
                id = 1
                Perekhod = "Установить " & "(x" & KolVo & ")"
            Case "pNametka"
                id = 2
                Perekhod = "Наметить " & KolVo & " отв."
            Case "pSverleniye"
                id = 3
                Perekhod = "Сверлить " & KolVo & " отв. Ф" & .d
            Case "pRatachivaniye"
                id = 5
                Perekhod = "Расточить " & KolVo & " отв. Ф" & .d
            Case "pVytochki"
                id = 6
                Perekhod = "Расточить " & KolVo & " выточ. Ф" & .d
        End Select
        
        DiapazonParametrov.Columns(Stolbets_idPerekhoda) = id
        DiapazonParametrov.Columns(Stolbets_Perekhod) = Perekhod
        DiapazonParametrov.Columns(Stolbets_D) = .d
        DiapazonParametrov.Columns(Stolbets_L) = .L
        DiapazonParametrov.Columns(Stolbets_SovmeshcheniyeOsi) = .SposobSovmeshcheniyaOsi
        DiapazonParametrov.Columns(Stolbets_PredvaritelnayaTsentrovka) = .PredvaritelnayaTsentrovka
        DiapazonParametrov.Columns(Stolbets_Glukhoye) = .Glukhoye
        DiapazonParametrov.Columns(Stolbets_NaklonnayaPloskost) = .NaklonnayaPloskost
        DiapazonParametrov.Columns(Stolbets_TonkostennyaDetal) = .TonkostennyaDetal
        DiapazonParametrov.Columns(Stolbets_NaUdar) = .NaUdar
        DiapazonParametrov.Columns(Stolbets_IT) = .IT
        DiapazonParametrov.Columns(Stolbets_Ra) = .Ra
        DiapazonParametrov.Columns(Stolbets_BieniyeaSoosnostKruglost) = .BieniyeaSoosnostKruglost
        DiapazonParametrov.Columns(Stolbets_ZakalonnayaStal) = .ZakalonnayaStal
        DiapazonParametrov.Columns(Stolbets_KolVoDopProhodov) = .KolVoDopProhodov
        DiapazonParametrov.Columns(Stolbets_KolVo) = KolVo
    
    End With
    
    With ParametryUstanovki
        DiapazonParametrov.Columns(Stolbets_Massa) = .Massa
        DiapazonParametrov.Columns(Stolbets_Sposob) = .Sposob
        DiapazonParametrov.Columns(Stolbets_HarakterViverki) = .HarakterViverki
        DiapazonParametrov.Columns(Stolbets_Nezestk) = .Nezestk
        DiapazonParametrov.Columns(Stolbets_TochnostViverki) = .TochnostViverki
        DiapazonParametrov.Columns(Stolbets_KolVoDopBoltov) = .KolVoDopBoltov
        DiapazonParametrov.Columns(Stolbets_KolVoDopDomkratov) = .KolVoDopDomkratov
        DiapazonParametrov.Columns(Stolbets_KolVoUstDetaley) = .KolVoUstDetaley
    End With
    
End Sub

Private Sub Redaktirovaniye()

    If lstKoordinat.ListIndex = -1 Then Exit Sub
    
    RezhimRedaktirovaniya = True
    
    Dim StrokaRedaktirovaniya As Integer
    StrokaRedaktirovaniya = lstKoordinat.ListIndex + 1
    
    Dim DiapazonRedaktirovaniya As Range
    Set DiapazonRedaktirovaniya = tbKoordinat.ListRows(StrokaRedaktirovaniya).Range
    
    Dim optSovmOsi_1 As MSForms.OptionButton, optSovmOsi_2 As MSForms.OptionButton, optSovmOsi_3 As MSForms.OptionButton
    
    IdotZagruzka = True
    
    Dim i As Integer
    If DiapazonRedaktirovaniya(Stolbets_idPerekhoda) = 1 Then
        mltKoordinat.Value = mltKoordinat.Pages("pUstanovka").Index
        txtMass_Koordinat = DiapazonRedaktirovaniya(Stolbets_Massa)
        txtKolVoUstanovok = DiapazonRedaktirovaniya(Stolbets_KolVo)
        
        With lstSposobUst_Koordinat
            For i = 0 To .ListCount - 1
                If .List(i, 1) = DiapazonRedaktirovaniya(Stolbets_Sposob) - 1 Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End With
        
        With cboHarakterViverki_Koordinat
            For i = 0 To .ListCount - 1
                If .List(i, 1) = DiapazonRedaktirovaniya(Stolbets_HarakterViverki) - 1 Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End With
        
        cboTochnostViverki_Koordinat.ListIndex = DiapazonRedaktirovaniya(Stolbets_TochnostViverki) - 1
        chkNezestk_Koordinat = DiapazonRedaktirovaniya(Stolbets_Nezestk)
        txtDopBolt_Koordinat = DiapazonRedaktirovaniya(Stolbets_KolVoDopBoltov)
        txtDopDomkrat_Koordinat = DiapazonRedaktirovaniya(Stolbets_KolVoDopDomkratov)
        txtKolDet_Koordinat = DiapazonRedaktirovaniya(Stolbets_KolVoUstDetaley)
    Else
    
        Select Case DiapazonRedaktirovaniya(Stolbets_idPerekhoda)
            Case 2
                mltKoordinat.Value = mltKoordinat.Pages("pNametka").Index
            
                Set optSovmOsi_1 = optSovmOsi_Nametka_1
                Set optSovmOsi_2 = optSovmOsi_Nametka_2
                Set optSovmOsi_3 = optSovmOsi_Nametka_3
            Case 3
                mltKoordinat.Value = mltKoordinat.Pages("pSverleniye").Index
            
                Set optSovmOsi_1 = optSovmOsi_Sverl_1
                Set optSovmOsi_2 = optSovmOsi_Sverl_2
                Set optSovmOsi_3 = optSovmOsi_Sverl_3
            
            Case 5
                mltKoordinat.Value = mltKoordinat.Pages("pRatachivaniye").Index
            
                Set optSovmOsi_1 = optSovmOsi_Rastachivaniye_1
                Set optSovmOsi_2 = optSovmOsi_Rastachivaniye_2
                Set optSovmOsi_3 = optSovmOsi_Rastachivaniye_3
            
                Select Case DiapazonRedaktirovaniya(Stolbets_IT)
                    Case Is >= 11:  cboIT_Rastachivaniye.ListIndex = 0
                    Case Is >= 9:   cboIT_Rastachivaniye.ListIndex = 1
                    Case Is >= 7:   cboIT_Rastachivaniye.ListIndex = 2
                    Case Is <= 6:   cboIT_Rastachivaniye.ListIndex = 3
                End Select
            
                Select Case DiapazonRedaktirovaniya(Stolbets_Ra)
                    Case Is >= 5:       cboRa_Rastachivaniye.text = "5,0"
                    Case Is >= 2.5:     cboRa_Rastachivaniye.text = "2,5"
                    Case Is >= 1.25:    cboRa_Rastachivaniye.text = "1,25"
                    Case Is <= 0.63:    cboRa_Rastachivaniye.text = "0,63"
                End Select

                txtKolVoDopProhodov_Rastachivaniye = IIf(DiapazonRedaktirovaniya(Stolbets_KolVoDopProhodov) > 0, DiapazonRedaktirovaniya(Stolbets_KolVoDopProhodov), Empty)
            
                Select Case DiapazonRedaktirovaniya(Stolbets_BieniyeaSoosnostKruglost)
                    Case 0:             cboDopusk_Rastachivaniye.ListIndex = 0
                    Case Is <= 0.002:   cboDopusk_Rastachivaniye.ListIndex = 1
                    Case Is <= 0.005:   cboDopusk_Rastachivaniye.ListIndex = 2
                    Case Is <= 0.01:    cboDopusk_Rastachivaniye.ListIndex = 3
                End Select
                
            Case 6
                mltKoordinat.Value = mltKoordinat.Pages("pVytochki").Index
                Set optSovmOsi_1 = optSovmOsi_Vytochki_1
                Set optSovmOsi_2 = optSovmOsi_Vytochki_2
            
        End Select
    
        Dim ctrl As Control
        For Each ctrl In mltKoordinat.SelectedItem.Controls
            If ctrl.Name Like "txtD_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_D)
            ElseIf ctrl.Name Like "txtL_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_L)
            ElseIf ctrl.Name Like "txtKolVo_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_KolVo)
            ElseIf ctrl.Name Like "chkGlukhoye_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_Glukhoye)
            ElseIf ctrl.Name Like "chkNaklon_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_NaklonnayaPloskost)
            ElseIf ctrl.Name Like "chkTonkosten_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_TonkostennyaDetal)
            ElseIf ctrl.Name Like "chkUdar_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_NaUdar)
            ElseIf ctrl.Name Like "chkZakalStal_*" Then
                ctrl = DiapazonRedaktirovaniya(Stolbets_ZakalonnayaStal)
            End If
        Next
        
        If DiapazonRedaktirovaniya(Stolbets_SovmeshcheniyeOsi) = 1 Then
            optSovmOsi_1.Value = True
        ElseIf DiapazonRedaktirovaniya(Stolbets_SovmeshcheniyeOsi) = 2 Then
            optSovmOsi_2.Value = True
        ElseIf DiapazonRedaktirovaniya(Stolbets_SovmeshcheniyeOsi) = 3 Then
            optSovmOsi_3.Value = True
        End If
        
    End If
    
    Call IzmenitRezhimFormy
            
    IdotZagruzka = False
    
    Call RaschotVremeni
    
End Sub

Private Sub cboIT_Rastachivaniye_Change()
    
    cboRa_Rastachivaniye.Clear
    If cboIT_Rastachivaniye.ListIndex <> -1 Then
        Select Case cboIT_Rastachivaniye.ListIndex
            Case 0
                cboRa_Rastachivaniye.AddItem "5,0"
            Case 1
                cboRa_Rastachivaniye.AddItem "2,5"
            Case 2
                With cboRa_Rastachivaniye
                    .AddItem "2,5"
                    .AddItem "1,25"
                    .AddItem "0,63"
                End With
            Case 3
                With cboRa_Rastachivaniye
                    .AddItem "1,25"
                    .AddItem "0,63"
                End With
        End Select
        cboRa_Rastachivaniye.ListIndex = 0
    End If
    
End Sub

Private Sub cmdEdit_Click()
    Call Redaktirovaniye
End Sub

Private Sub cmdOtkryt_Click()

    Dim ImyaKnigi As String
    ImyaKnigi = Application.GetOpenFilename
    
    If ImyaKnigi <> "False" Then
        newWB.Close (False)
        Set newWB = Workbooks.Open(ImyaKnigi)
        newWB.Windows(1).Visible = False
        Call PrisvoeniyeDiapazonov
        Call ZagruzkaRaschota
    End If
    
End Sub

Private Sub cmdSohranit_Click()
    
    Application.ScreenUpdating = False
    
    newWB.Windows(1).Visible = True
    
    Dim ImyaFayla As String
    ImyaFayla = Application.GetSaveAsFilename
    
    Application.DisplayAlerts = False
    If ImyaFayla <> "False" Then newWB.SaveAs FileFormat:=51
    Application.DisplayAlerts = True
    
    newWB.Windows(1).Visible = False
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub IzmenitRezhimFormy()
    
    If RezhimRedaktirovaniya Then
        cmdDobavit.Caption = "Редактировать"
    Else
        cmdDobavit.Caption = "Добавить переход в расчет"
    End If
    
    tglEdit.Visible = RezhimRedaktirovaniya
    
    Dim i As Integer
    With mltKoordinat
        For i = 0 To .Pages.Count - 1
            If i <> .SelectedItem.Index Then .Pages(i).Enabled = Not RezhimRedaktirovaniya
        Next
    End With
    
    Dim ctrl As Control
    For Each ctrl In fraRaschot.Controls
        ctrl.Enabled = Not RezhimRedaktirovaniya
    Next
    
End Sub

Private Sub tglEdit_Click()
    RezhimRedaktirovaniya = False
    Call IzmenitRezhimFormy
End Sub

Private Sub lstSposobUst_Koordinat_Change()
    Call ZapolnitKharakterVyverki_Koordinat(Me)
End Sub

Private Sub PrisvoeniyeDiapazonov()

    With newWB.ActiveSheet
        Set tbKoordinat = .ListObjects("tbKoordinat")
        Set rMaterial = .Range("Material")
        Set rTipProizvodstva = .Range("TipProizvodstva")
        Set rTipStanka = .Range("TipStanka")
        Set rVozrast = .Range("Vozrast")
        Set rMaterialInstrumenta = .Range("MaterialInstrumenta")
    End With
    
End Sub

Private Sub cboMaterial_Change()

    Material = MaterialFromCboMaterial(cboMaterial)

    Dim MaterialColor As Double
    MaterialColor = TsvetMateriala(CInt(Material))

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        With ctrl
            If .Tag = "color" Then
                .BackColor = MaterialColor
            End If
        End With
    Next
    
    DmaxVytochki = Dmax_Vytocheki(CInt(Material))
    lblDmax_Vytochka.Caption = "до " & DmaxVytochki & " мм"
    
    Call ZagruzkaVListBox
    
End Sub

Sub RaschotVremeni()

    Vremya = 0
    txtVremya.text = Empty
    KolVo = 0
    
    Select Case mltKoordinat.SelectedItem.Name
        Case "pSverleniye"
            Call RaschotVremeni_Sverleniye
        Case "pNametka"
            Call RaschotVremeni_Nametka
        Case "pRatachivaniye"
            Call RaschotVremeni_Rastachivaniye
        Case "pUstanovka"
            Call RaschotVremeni_Ustanovka
        Case "pVytochki"
            Call RaschotVremeni_Vytochki
    End Select
    
    Dim GlavnyyKoeffitsient As Double
    
    If IsArray(Vremya) Then
        tKoordinat = Vremya(1) * KolVo
    Else
        tKoordinat = Vremya * KolVo
    End If
    
    If tKoordinat > 0 Then
        If mltKoordinat.SelectedItem.Name = "pUstanovka" Then
            GlavnyyKoeffitsient = PoprKoef_Koordinat_TipProizv(cboTipProizv_Koordinat.ListIndex + 1)
        Else
            GlavnyyKoeffitsient = PoprKoef_Koordinat_MatrialInstrument(cboMaterialInstrument_Koordinat.ListIndex + 1) * _
                PoprKoef_Koordinat_TipProizv(cboTipProizv_Koordinat.ListIndex + 1) * _
                PoprKoef_Koordinat_TipStanka(cboTipStanka_Koordinat.ListIndex + 1) * _
                PoprKoef_Koordinat_Vozrast(cboVozrast_Koordinat.ListIndex + 1)
        End If
        txtVremya.text = Format(tKoordinat * GlavnyyKoeffitsient, "0.0")
    End If
    
End Sub

Private Sub RaschotVremeni_Sverleniye()

    txtD_Sverl.BackColor = vbWhite
    txtL_Sverl.BackColor = vbWhite
    txtD_Sverl_Lmax_Kd.Caption = "-"
    
    Dim Parametry As Koordinat
    
    With Parametry
        
        .d = DblFromCtrl(txtD_Sverl)
        .L = DblFromCtrl(txtL_Sverl)
        
        If .d > DMAX_SVERLENIYE_KOORDINAT Then txtD_Sverl.BackColor = vbRed
        
        Dim Lmax As Double: Lmax = Lmax_Sverlenie_Koordinat(.d)
        If Lmax > 0 Then txtD_Sverl_Lmax_Kd.Caption = "до " & Lmax & " мм"
        If .L > Lmax Then txtL_Sverl.BackColor = vbRed
        
        If .d > 0 And .L > 0 And .d <= DMAX_SVERLENIYE_KOORDINAT And .L <= Lmax Then
        
            .Glukhoye = chkGlukhoye_Sverl.Value
            .NaklonnayaPloskost = chkNaklon_Sverl.Value
            .NaUdar = chkNaUdar_Sverl.Value
            .PredvaritelnayaTsentrovka = chkTsentrovka_Sverl.Value
            
            If optSovmOsi_Sverl_1 Then
                .SposobSovmeshcheniyaOsi = PoNoniusnoyIliMasshtabnoyLineyke
            ElseIf optSovmOsi_Sverl_2 Then
                .SposobSovmeshcheniyaOsi = PoShtikhmasuIliIzmeritelnymPlitkam
            ElseIf optSovmOsi_Sverl_3 Then
                .SposobSovmeshcheniyaOsi = BezSovmeshcheniyaOsey
            End If

            .TonkostennyaDetal = chkTonkosten_Sverl.Value
        
            KolVo = KolVoPoverhnosteiProhodov(txtKolVo_Sverl)
            Vremya = Sverlenie_Koordinat(CInt(Material), .d, .L, .SposobSovmeshcheniyaOsi, .PredvaritelnayaTsentrovka, .Glukhoye, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar)
            
        End If
    
    End With
    
    Parametry_Koordinat = Parametry
    
End Sub

Private Sub RaschotVremeni_Nametka()

    Dim Parametry As Koordinat
    
    With Parametry

        .NaklonnayaPloskost = chkNaklon_Nametka.Value
        
        If optSovmOsi_Nametka_1 Then
            .SposobSovmeshcheniyaOsi = PoNoniusnoyIliMasshtabnoyLineyke
        ElseIf optSovmOsi_Nametka_2 Then
            .SposobSovmeshcheniyaOsi = PoShtikhmasuIliIzmeritelnymPlitkam
        ElseIf optSovmOsi_Nametka_3 Then
            .SposobSovmeshcheniyaOsi = BezSovmeshcheniyaOsey
        End If
        
        KolVo = DblFromCtrl(txtKolVo_Nametka)
        If KolVo > 0 Then
            Vremya = NametkaOtverstiy_Koordinat(.SposobSovmeshcheniyaOsi, .NaklonnayaPloskost)
        End If
    
    End With
    
    Parametry_Koordinat = Parametry
    
End Sub
Private Sub RaschotVremeni_Ustanovka()

    txtMass_Koordinat.BackColor = vbWhite
    
    With ParametryUstanovki
    
        .Massa = DblFromCtrl(txtMass_Koordinat)
    
        If .Massa > 0 And lstSposobUst_Koordinat.ListIndex <> -1 And cboHarakterViverki_Koordinat.ListIndex <> -1 _
            And cboTochnostViverki_Koordinat.ListIndex <> -1 Then
    
            .Sposob = lstSposobUst_Koordinat.List(lstSposobUst_Koordinat.ListIndex, 1)
            .HarakterViverki = cboHarakterViverki_Koordinat.List(cboHarakterViverki_Koordinat.ListIndex, 1)
            .TochnostViverki = cboTochnostViverki_Koordinat.ListIndex + 1
            .KolVoUstDetaley = KolVoPoverhnosteiProhodov(txtKolDet_Koordinat)
            .Nezestk = chkNezestk_Koordinat.Value
            .KolVoDopBoltov = CInt(txtDopBolt_Koordinat.text)
            .KolVoDopDomkratov = CInt(txtDopDomkrat_Koordinat.text)
            .KolVoUstDetaley = CInt(txtKolDet_Koordinat.text)
            KolVo = KolVoPoverhnosteiProhodov(txtKolVoUstanovok)
            
            Dim MaksMassa As Integer
            MaksMassa = MaksMassaUstanovki_Koordinat(.Sposob, .HarakterViverki)
            If .Massa > MaksMassa Then txtMass_Koordinat.BackColor = vbRed
            lblMmax.Caption = "до " & MaksMassa & " кг"

            Vremya = Ustanovka_Koordinat(.Massa, .Sposob, .HarakterViverki, .Nezestk, .TochnostViverki, .KolVoDopBoltov, .KolVoDopDomkratov, .KolVoUstDetaley)
        
        End If
    
    End With
    
    
End Sub

Private Sub RaschotVremeni_Rastachivaniye()

    txtD_Rastachivaniye.BackColor = vbWhite
    txtL_Rastachivaniye.BackColor = vbWhite
    lblLmax_Rastach_Koordinat.Caption = "-"
    txtPripusk_Rastachivaniye.text = Empty
    
    Dim Parametry As Koordinat
    With Parametry
    
        .d = DblFromCtrl(txtD_Rastachivaniye)
        .L = DblFromCtrl(txtL_Rastachivaniye)
        .Ra = DblFromCtrl(cboRa_Rastachivaniye)
            
        If .d > DMAX_RASTACHIVANIYE_KOORDINAT Then txtD_Rastachivaniye.BackColor = vbRed
        
        Dim Lmax As Double: Lmax = Lmax_Rastachivaniye_Koordinat(.d)
        If Lmax > 0 Then lblLmax_Rastach_Koordinat.Caption = "до " & Lmax & " мм"
        If .L > Lmax Then txtL_Rastachivaniye.BackColor = vbRed
        
        If .d > 0 Then
            txtPripusk_Rastachivaniye = "1,0...4,0"
        Else
            txtPripusk_Rastachivaniye = "1,0...3,0"
        End If
        
        If .d > 0 And .L > 0 And .d <= DMAX_RASTACHIVANIYE_KOORDINAT And .L <= Lmax And .Ra > 0 Then
    
            If cboDopusk_Rastachivaniye.ListIndex = 0 Then
                .BieniyeaSoosnostKruglost = 0
            Else
                .BieniyeaSoosnostKruglost = DblFromCtrl(cboDopusk_Rastachivaniye)
            End If
            
            .Glukhoye = chkGlukhoye_Rastachivaniye.Value
            
            Select Case cboIT_Rastachivaniye.ListIndex
                Case 0: .IT = 11
                Case 1: .IT = 9
                Case 2: .IT = 7
                Case 3: .IT = 6
            End Select
            
            .KolVoDopProhodov = DblFromCtrl(txtKolVoDopProhodov_Rastachivaniye)
            
            .NaklonnayaPloskost = chkNaklon_Rastachivaniye.Value
            .NaUdar = chkUdar_Rastachivaniye.Value
            
            If optSovmOsi_Rastachivaniye_1 Then
                .SposobSovmeshcheniyaOsi = PoNoniusnoyIliMasshtabnoyLineyke
            ElseIf optSovmOsi_Rastachivaniye_2 Then
                .SposobSovmeshcheniyaOsi = PoShtikhmasuIliIzmeritelnymPlitkam
            ElseIf optSovmOsi_Rastachivaniye_3 Then
                .SposobSovmeshcheniyaOsi = BezSovmeshcheniyaOsey
            End If
            
            .TonkostennyaDetal = chkTonkosten_Rastachivaniye.Value
            .ZakalonnayaStal = chkZakalStal_Rastachivaniye.Value
        
            KolVo = KolVoPoverhnosteiProhodov(txtKolVo_Rastachivaniye)
            Vremya = Rastachivaniye_Koordinat(CInt(Material), .d, .L, .SposobSovmeshcheniyaOsi, .IT, .Ra, .Glukhoye, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar, .BieniyeaSoosnostKruglost, .ZakalonnayaStal, .KolVoDopProhodov)
        
        End If
    
    End With
        
    Parametry_Koordinat = Parametry
    
End Sub

Private Sub RaschotVremeni_Vytochki()

    txtD_Vytochki.BackColor = vbWhite
    txtL_Vytochki.BackColor = vbWhite
    lblLmax_Vytochka.Caption = "-"
    txtPripusk_Vytochki.text = Empty
    txtRa_Vytochki.text = Empty
    
    Dim Parametry As Koordinat
    With Parametry
    
        .d = DblFromCtrl(txtD_Vytochki)
        .L = DblFromCtrl(txtL_Vytochki)
        
        If .d > DmaxVytochki Then txtD_Vytochki.BackColor = vbRed
        
        If .d > 0 Then
            Dim Lmax As Double: Lmax = Lmax_Vytocheki(CInt(Material), .d)
            If Lmax > 0 Then lblLmax_Vytochka.Caption = "до " & Lmax & " мм"
            If .L > Lmax Then txtL_Vytochki.BackColor = vbRed
        End If
        
        If .d > 0 And .L > 0 And .d <= DmaxVytochki And .L <= Lmax Then
    
            If cboDopusk_Vytochki.ListIndex = 0 Then
                .BieniyeaSoosnostKruglost = 0
            Else
                .BieniyeaSoosnostKruglost = DblFromCtrl(cboDopusk_Vytochki)
            End If
            
            Select Case cboIT_Vytochki.ListIndex
                Case 0
                    txtRa_Vytochki.text = "-"
                    txtPripusk_Vytochki.text = "2,0...4,0"
                    .IT = 14
                Case 1
                    txtRa_Vytochki.text = "5,0"
                    txtPripusk_Vytochki.text = "1,2...1,7"
                    .IT = 11
                Case 2
                    txtRa_Vytochki.text = "2,5"
                    txtPripusk_Vytochki.text = "1,2...1,7"
                    .IT = 9
                Case 3
                    txtRa_Vytochki.text = "1,25"
                    txtPripusk_Vytochki.text = "0,5...0,9"
                    .IT = 7
                Case 4
                    txtRa_Vytochki.text = "1,25"
                    txtPripusk_Vytochki.text = "0,5...0,9"
                    .IT = 6
            End Select
            
            .NaklonnayaPloskost = chkNaklon_Vytochki.Value
            .NaUdar = chkUdar_Vytochki.Value
            
            If optSovmOsi_Vytochki_1 Then
                .SposobSovmeshcheniyaOsi = 1
            ElseIf optSovmOsi_Vytochki_2 Then
                .SposobSovmeshcheniyaOsi = 0
            End If
            
            .TonkostennyaDetal = chkTonkosten_Vytochki.Value
            .ZakalonnayaStal = chkZakal_Vytochki.Value
        
            KolVo = KolVoPoverhnosteiProhodov(txtKolVo_Vytochki)
            Vremya = RastachivaniyeVytochek(CInt(Material), .d, .L, .IT, .SposobSovmeshcheniyaOsi <> BezSovmeshcheniyaOsey, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar, .BieniyeaSoosnostKruglost, .ZakalonnayaStal)
        
        End If
    
    End With
        
    Parametry_Koordinat = Parametry
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    newWB.Close (False)
End Sub



