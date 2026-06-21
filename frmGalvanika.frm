VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGalvanika 
   Caption         =   "Нанесение гальванических покрытий"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6330
   OleObjectBlob   =   "frmGalvanika.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGalvanika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboGalvanika_Change(): GlavnyyRaschot: End Sub

Private Sub chkGalvanikaPartiya_Click(): GlavnyyRaschot: End Sub

Private Sub txtGalvanika_S_Change(): GlavnyyRaschot: End Sub

Private Sub txtKolVo_Change(): GlavnyyRaschot: End Sub

Private Sub UserForm_Initialize()
    
    Call VyravnitElementy
    
    DobavitVKlass_KontrolVvodaChisel_Ctrl txtGalvanika_S
    DobavitVKlass_KontrolVvodaChisel_Ctrl txtKolVo
    
    With cboGalvanika
        .List = Split("Обезжиривание;Травление;Цинкование;Ан. Окс. хр;Ан. Окс. тв;Ан. Окс. Черн;Хим. Пас;Н12Х;Н18;О-Ви", ";")
        .ListIndex = 0
    End With
    
End Sub

Private Sub VyravnitElementy()

    fraVvod.Top = 6
    fraNormy.Top = fraVvod.Top + fraVvod.Height + 6
    
    fraVvod.Left = 6
    fraNormy.Left = 6
    
    Me.Width = fraNormy.Left + fraNormy.Width + PLUS_SHIRINA
    Me.Height = fraNormy.Top + fraNormy.Height + PLUS_VYSOTA
    
End Sub

Private Sub GlavnyyRaschot()

    OchistitFormu
    
    Dim Smax As Double, S As Double
    S = DblFromCtrl(txtGalvanika_S)
    Smax = Galvanika_MaksPloshchadM2(cboGalvanika.text)
    
    lblSmax.Caption = "до " & CStr(Smax) & " мм" & ChrW(178)
    If S > Smax Then txtGalvanika_S.BackColor = vbRed
    
    If Not S > 0 Or S > Smax Then Exit Sub
    
    Dim Normy As Variant
    Normy = GalvanikaNovayaFunktsiya(cboGalvanika.text, S, KolVoPoverhnosteiProhodov(txtKolVo), chkGalvanikaPartiya)
    
    ZapisatRezyltaty Normy
     
End Sub

Private Sub OchistitFormu()
    
    txtGalvanika_S.BackColor = vbWhite
    lblSmax.Caption = "-"

    Dim ctrl As Control
    For Each ctrl In fraNormy.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.text = "-"
        End If
    Next
    
End Sub

Private Sub ZapisatRezyltaty(Normy As Variant)

    If Not IsArray(Normy) Then Exit Sub

    txtRazryad_Galvanika.text = OpredelitRazryad_Galvanika(cboGalvanika.text)
    txtTpz_Galvanika.text = 5
    txtTsht_Galvanika.text = CStr(OkruglenieTsht(CDbl(Normy(1))))
    
    txtRazryad_Kontrol.text = 3
    txtTpz_Kontrol.text = 0
    txtTsht_Kontrol.text = CStr(OKRUGLVVERKH(CDbl(Normy(2)), 1))
    
End Sub

