VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtrezkaNaCircularnihPilah 
   Caption         =   "Разметка и отрезка заготовок на циркулярных пилах"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   OleObjectBlob   =   "frmOtrezkaNaCircularnihPilah.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtrezkaNaCircularnihPilah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaksTolshchina As Integer
Dim tRazmetki As Double, tOtrezki As Double, tSlesar As Double
Dim Dlina As Double, Shirina As Double, Tolshchina As Double
Dim Dlina_Detali As Double, Shirina_Detali As Double
Dim Massa As Double
Dim Material As EnumMaterialy

Private Sub cboMaterial_Change()

    Material = MaterialFromCboMaterial(cboMaterial)
    
    MaksTolshchina = OtrezkaTsirkulyarnoyPiloy_List_MaksTolshchina(CInt(Material))
    
    lblTolschina_Lmax.Caption = "до " & MaksTolshchina
    
    Call BolsheMaksimalnogo(MaksTolshchina, txtTolschina)
    
    Call RaschetVremeni
    
End Sub

Private Sub optPesk_Click()
    Call RaschetVremeni
End Sub

Private Sub optShkurka_Click()
    Call RaschetVremeni
End Sub

Private Sub txtDlina_Change()
    Call BolsheMaksimalnogo(MAX_DLINA_CIRCULYARKA, txtDlina)
    Call RaschetVremeni
End Sub
Private Sub BolsheMaksimalnogo(MaksZnachenie, txt As MSForms.TextBox)
    
    Dim Znachenie As Double
    Znachenie = DblFromCtrl(txt)
    
    If Znachenie > MaksZnachenie Then
        txt.BorderColor = vbRed
    Else
        txt.BorderColor = &H8000000D
    End If
    
    Me.Repaint
    
End Sub

Private Sub txtDlina_Detal_Change()
    Call RaschetVremeni
End Sub

Private Sub txtShirina_Change()
    Call BolsheMaksimalnogo(MAX_DLINA_CIRCULYARKA, txtShirina)
    Call RaschetVremeni
End Sub


Private Sub txtShirina_Detal_Change()
    Call RaschetVremeni
End Sub

Private Sub txtTolschina_Change()
    Call BolsheMaksimalnogo(MaksTolshchina, txtTolschina)
    Call RaschetVremeni
End Sub

Private Sub UserForm_Initialize()

    Dim MATERIALY
    MATERIALY = Array(EnumMaterialy.STEKLOTEKSTOLIT, EnumMaterialy.TEKSTOLIT)
    Call DobavitMaterialyVListBox(MATERIALY, cboMaterial)
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ReDim Preserve ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel)
                Set ArrTxtKontrolVvodaChisel(iTxt_KontrolVvodaChisel).txt = ctrl
                iTxt_KontrolVvodaChisel = iTxt_KontrolVvodaChisel + 1
        End Select
    Next
    
End Sub

Private Sub RaschetVremeni()

    lstOtrez.Clear
        
    Dlina = WorksheetFunction.Max(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina))
    Shirina = WorksheetFunction.Min(DblFromCtrl(txtDlina), DblFromCtrl(txtShirina))
    Dlina_Detali = WorksheetFunction.Max(DblFromCtrl(txtDlina_Detal), DblFromCtrl(txtShirina_Detal))
    Shirina_Detali = WorksheetFunction.Min(DblFromCtrl(txtDlina_Detal), DblFromCtrl(txtShirina_Detal))
    Tolshchina = DblFromCtrl(txtTolschina)
        
    If Dlina > 0 And Shirina > 0 And Tolshchina > 0 And _
        Dlina <= MAX_DLINA_CIRCULYARKA And Shirina <= MAX_DLINA_CIRCULYARKA And Tolshchina <= MaksTolshchina Then
            
        'Расчет разметки
        Massa = MassaLista(Dlina, Shirina, Tolshchina, CInt(Material))
        tRazmetki = 1.15 * 1.2 * RazmetkaKonturaBezShablona_Lineykoy(CInt(Material), Massa, 2 * (Dlina + Shirina))
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        'Расчет отрезки
        tOtrezki = 1.3 * (OtrezkaTsirkulyarnoyPiloy_List(CInt(Material), Dlina, Tolshchina) + OtrezkaTsirkulyarnoyPiloy_List(CInt(Material), Shirina, Tolshchina))
            
        'Расчет слесарной операции
        tSlesar = ZachistkaShlifovalnoyShkurkoy(CInt(Material), Dlina, Shirina, 1) * 1.3 * 2
        
        Call ZapolneniyeTablitsy
     
    End If

End Sub

Private Sub ZapolneniyeTablitsy()
    
    Call DobavleniyeStrokVListBox(lstOtrez, "Отрезная", 10, tRazmetki + tOtrezki)
    Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, 1.3 * IzmerShtangenCircul_CPL(Dlina, Shirina))
    Call DobavleniyeStrokVListBox(lstOtrez, "Маркирование (этикеткой)", 0, 3)
    Call DobavleniyeStrokVListBox(lstOtrez, "Упаковывание", 0, 1.3 * UpakovyvaniyePramougolnika(Dlina, Shirina, Tolshchina))
    Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, 1.3 * VisualnyyKontrol(Shirina, Dlina))
    
    Dim MaterialPeskostuy As Integer
    If optShkurka Then
        Call DobavleniyeStrokVListBox(lstOtrez, "Слесарная", 5, tSlesar)
    Else
        MaterialPeskostuy = IIf(Material = 1, 4, 5)
        Call DobavleniyeStrokVListBox(lstOtrez, "Пескоструйная", 5, 1.3 * PeskostruynayaOchistka(Shirina, Dlina, 2, 2, Massa, Material))
    End If
        
    Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, VisualnyyKontrol(Shirina, Dlina) * 1.3)
    
    Dim Sdm As Double
    If cboMaterial.ListIndex = 0 Then
        Sdm = PloshchadPramougolnika(Dlina_Detali, Shirina_Detali, Tolshchina) / 10000
        Call DobavleniyeStrokVListBox(lstOtrez, "Лакирование", 5, Lakirovanie(Sdm, , 1) * 1.3)
    Else
        Call DobavleniyeStrokVListBox(lstOtrez, "Полирование", 5, Polirovanie(Dlina_Detali, Shirina_Detali, 2) * 1.3)
    End If
    
    Call DobavleniyeStrokVListBox(lstOtrez, "Контрольная", 0, VisualnyyKontrol(Shirina_Detali, Dlina_Detali) * 1.3)
        
End Sub
