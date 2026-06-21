VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPromyvka 
   Caption         =   "Промывка"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715.001
   OleObjectBlob   =   "frmPromyvka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPromyvka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboKolVoDetaleyVPartii_Change()
    Call RaschetVremeniProverki
End Sub

Private Sub cboPromyvka_Change()
    Call RaschetVremeniProverki
End Sub

Private Sub chkObduvka_Click()
    Call RaschetVremeniProverki
End Sub

Private Sub optCilindricheskaya_Click()
    lblShirina.Caption = "Диаметр промываемой поверхности"
    Call RaschetVremeniProverki
End Sub

Private Sub optPloskaya_Click()
    lblShirina.Caption = "Ширина промываемой поверхности"
    Call RaschetVremeniProverki
End Sub

Private Sub optProstaya_Click()
    Call RaschetVremeniProverki
End Sub

Private Sub optSlozhnaya_Click()
    Call RaschetVremeniProverki
End Sub

Private Sub txtDlina_Change()
    Call RaschetVremeniProverki
End Sub

Private Sub txtShirina_Change()
    Call RaschetVremeniProverki
End Sub

Private Sub UserForm_Initialize()

    With cboPromyvka
        .AddItem "От пыли и стружки"
        .AddItem "От масла"
        .AddItem "От консистентной смазки"
        .ListIndex = 2
    End With
    
    cboKolVoDetaleyVPartii.List = wsIzgorVhodDet.ListObjects("tbKoefPartiynosti_Slesar1990").DataBodyRange.Value
    cboKolVoDetaleyVPartii.ListIndex = 0
    
    Call DobavitVKlass_KontrolVvodaChisel_Ctrl(txtDlina)
    Call DobavitVKlass_KontrolVvodaChisel_Ctrl(txtShirina)
    
End Sub

Private Sub RaschetVremeniProverki()
    
    txtTsht.text = Empty
    
    Dim tPromyvki As Double, tObduvki As Double, Tsht As Double, K As Double
    If IsNumeric(txtShirina) And IsNumeric(txtDlina) Then
        tPromyvki = Promivka_m64(IIf(optPloskaya, 1, 2), cboPromyvka.ListIndex + 1, txtShirina, txtDlina, IIf(optProstaya, 1, 2))
        tObduvki = IIf(chkObduvka, Obduvka_m65(IIf(optPloskaya, 1, 2), txtShirina, txtDlina, IIf(optProstaya, 1, 2)), 0)
        K = CDbl(cboKolVoDetaleyVPartii.List(cboKolVoDetaleyVPartii.ListIndex, 1))
        Tsht = (tPromyvki + tObduvki) * K
        txtTsht = FormatTime(Tsht, 1, , 1)
    End If
    
End Sub
