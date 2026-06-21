VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPeskostruy 
   Caption         =   "Пескоструйная обработка"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmPeskostruy.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPeskostruy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tPeskostruy As Double, tKontrol As Double

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Sub UserForm_Initialize()

    With cboTipProizvodstva
        .AddItem "1-2"
        .AddItem "3-5"
        .AddItem "6-10"
        .AddItem "11-15"
        .AddItem "16-24"
        .AddItem "25-39"
        .AddItem "40 и более"
        .ListIndex = 0
    End With
    
    With cboUsloviya
        .AddItem "Удобно"
        .AddItem "Неудобно"
        .AddItem "Очень неудобно"
        .ListIndex = 0
    End With
    
    With cboMaterial
        .AddItem "Сталь"
        .AddItem "Медные и алюминиевые сплавы"
        .AddItem "Гетинакс"
        .AddItem "Текстолит"
        .AddItem "Оргстекло"
        .ListIndex = 0
    End With
    
    optSloznost_1.Value = True
    optTipZatogovki_1.Value = True
    
    Set ControlHandlers = New Collection
    Call DobavitVKlass_GlavnyyRaschot(Me, ControlHandlers, Handler)
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
End Sub

Public Sub GlavnyyRaschot()

    tPeskostruy = 0
    tKontrol = 0
    
    Dim Massa As Double
    Massa = DblFromCtrl(txtMassa)
    
    Dim kTipProizv As Double
    kTipProizv = Koeffitsient_KolVoDetaleyVPartii_IzgotVhodDetaley(KolVoDetaleyVPartii)
    lblKoeffitsientNaTipProizvodstva.Caption = "К=" & Format(kTipProizv, "0.0")
    
    Dim kGlavnyy As Double
    kGlavnyy = GlavnyyKoefftsient_IzgotVhodDetaley(KolVoDetaleyVPartii, cboUsloviya.ListIndex + 1, Massa > 20)
    lblKoefUsloviya.Caption = "К=" & Format(RaschotKoefUsloviya_IzgotVhodDetaley(cboUsloviya.ListIndex + 1), "0.0")
    
    If mltPeskostruy.SelectedItem.Name = "pZagotovka" Then
        Call RaschetVremeni_Zagotovka(Massa, kTipProizv, kGlavnyy)
    Else
        Call RaschetVremeni_Karkas(Massa, kTipProizv, kGlavnyy)
    End If
    
    txtVremya.text = CStr(OkruglenieTsht(tPeskostruy))
    txtKontrol.text = CStr(OkruglenieTsht(tKontrol))
    
End Sub

Private Sub RaschetVremeni_Zagotovka(Massa As Double, kTipProizv As Double, kGlavnyy As Double)

    txtVremya.text = Empty
    txtKontrol.text = Empty
    
    Dim TipZagotovki As Integer
    optSloznost_2.Enabled = True
    If optTipZatogovki_1.Value = True Then
        TipZagotovki = 1
        optSloznost_2.Enabled = False
    ElseIf optTipZatogovki_2.Value = True Then
        TipZagotovki = 2
    Else
        TipZagotovki = 3
    End If

    Dim DiametrIliVysota As Double, Dlina As Double
    
    DiametrIliVysota = DblFromCtrl(txtDB_Zagotovka)
    Dlina = DblFromCtrl(txtL_Zagotovka)
    
    Call SdelatFonKrasnym(txtL_Zagotovka, MAX_DLINA_PESKOSTRUY)
    Call SdelatFonKrasnym(txtDB_Zagotovka, MAX_SHIRINA_PESKOSTRUY)
    
    If DiametrIliVysota > MAX_SHIRINA_PESKOSTRUY Or Dlina > MAX_DLINA_PESKOSTRUY Then Exit Sub
    
    If DiametrIliVysota > 0 And Dlina > 0 Then

        tPeskostruy = PeskostruynayaOchistka(DiametrIliVysota, Dlina, TipZagotovki, IIf(optSloznost_1, 1, 2), Massa, cboMaterial.ListIndex + 1) * kGlavnyy
        tKontrol = VisualnyyKontrol(WorksheetFunction.Min(DiametrIliVysota, Dlina), WorksheetFunction.Max(DiametrIliVysota, Dlina)) * kTipProizv

    End If

End Sub

Private Sub RaschetVremeni_Karkas(Massa As Double, kTipProizv As Double, kGlavnyy As Double)

    txtVremya.text = Empty
    txtKontrol.text = Empty

    Dim Dlina As Double, Shirina As Double, Vysota As Double, KolVoDetaley As Double
    
    Dlina = DblFromCtrl(txtDlina_Karkas)
    Shirina = DblFromCtrl(txtShirina_Karkas)
    Vysota = DblFromCtrl(txtVysota_Karkas)
    KolVoDetaley = DblFromCtrl(txtKolVoDetaley_Karkas)
    
    Call SdelatFonKrasnym(txtDlina_Karkas, MAX_DLINA_PESKOSTRUY)
    Call SdelatFonKrasnym(txtShirina_Karkas, MAX_SHIRINA_PESKOSTRUY)
    Call SdelatFonKrasnym(txtVysota_Karkas, MAX_VYSOTA_PESKOSTRUY)
    Call SdelatFonKrasnym(txtKolVoDetaley_Karkas, MAX_KOLVO_DETALEY_PESKOSTRUY)
    
    If Dlina > MAX_DLINA_PESKOSTRUY Or Shirina > MAX_SHIRINA_PESKOSTRUY Or Vysota > MAX_VYSOTA_PESKOSTRUY Or KolVoDetaley > MAX_KOLVO_DETALEY_PESKOSTRUY Then Exit Sub
    
    If Dlina > 0 And Shirina > 0 And Vysota > 0 Then
    
        tPeskostruy = PeskostruyKarkasov(Dlina, Shirina, Vysota, KolVoDetaley, Massa) * kGlavnyy
        tKontrol = VisualnyyKontrol(WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 2), WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 1)) * kTipProizv
        
    End If

End Sub

Private Function KolVoDetaleyVPartii() As Double

    Select Case cboTipProizvodstva.ListIndex
        Case 0:     KolVoDetaleyVPartii = 2
        Case 1:     KolVoDetaleyVPartii = 5
        Case 2:     KolVoDetaleyVPartii = 10
        Case 3:     KolVoDetaleyVPartii = 15
        Case 4:     KolVoDetaleyVPartii = 24
        Case 5:     KolVoDetaleyVPartii = 39
        Case Else:  KolVoDetaleyVPartii = 40
    End Select
    
End Function
