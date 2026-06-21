VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarkiravanie 
   Caption         =   "Маркировка кистью"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   OleObjectBlob   =   "frmMarkiravanie.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMarkiravanie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboVisotaShrifta_Change(): Call GlavnyyRaschot: End Sub

Private Sub chkLakirovanie_Click(): Call GlavnyyRaschot: End Sub

Private Sub optKolvoCvetov_1_Click(): Call GlavnyyRaschot: End Sub

Private Sub optKolvoCvetov_2_Click(): Call GlavnyyRaschot: End Sub

Private Sub optSloznost_1_Click(): Call GlavnyyRaschot: End Sub

Private Sub optSloznost_2_Click(): Call GlavnyyRaschot: End Sub

Private Sub txtKolvoNadpisei_Change(): Call GlavnyyRaschot: End Sub

Private Sub txtKolvoSimvolov_Change(): Call GlavnyyRaschot: End Sub

Private Sub UserForm_Initialize()

    VyravnitElementy

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    With cboVisotaShrifta
        .List = Split("3 5 10")
        .ListIndex = 0
    End With
    
End Sub

Private Sub VyravnitElementy()

    Const Shirina As Integer = 327

    With fraParametry
        .Left = 3
        .Top = 6
        .Width = Shirina
        .Height = chkLakirovanie.Top + chkLakirovanie.Height + 18
    End With
    
    With fraNormy
        .Top = fraParametry.Top + fraParametry.Height + 6
        .Left = 3
        .Width = Shirina
    End With
    
    With mltMarkirovaniye
        .Left = 3
        .Top = 6
        .Width = Shirina + 9
        .Height = fraNormy.Top + fraNormy.Height + 3 + .TabFixedHeight + .Top + 21
    End With
    
    With Me
        .Width = mltMarkirovaniye.Left + mltMarkirovaniye.Width + 15
        .Height = mltMarkirovaniye.Top + mltMarkirovaniye.Height + 30
    End With
    
    
End Sub

Private Sub GlavnyyRaschot()

    txtMarkirovanie.text = Empty
    txtKontrol.text = Empty

    Dim VisotaShrifta As Double, KolVoSimvolov As Double, KolvoNadpisei As Double
    Dim Lakirovanie As Boolean
    
    VisotaShrifta = DblFromCtrl(cboVisotaShrifta)
    KolVoSimvolov = DblFromCtrl(txtKolvoSimvolov)
    KolvoNadpisei = KolVoPoverhnosteiProhodov(txtKolvoNadpisei)
    
    If VisotaShrifta * KolVoSimvolov * KolvoNadpisei > 0 Then
    
        Dim Sloznost As Integer
        Sloznost = IIf(optSloznost_1, 1, 2)

        txtMarkirovanie = OKRUGLVVERKH(MarkirovkaKistyu(VisotaShrifta, KolVoSimvolov, Sloznost, IIf(optKolvoCvetov_1, 1, 2), KolvoNadpisei, chkLakirovanie.Value) * 1.3)
        
        txtKontrol = FormatTime(KolVoSimvolov / 60, 3)
        
        txtRazryad_Markirovanie = IIf(Sloznost = 1, 2, 3)
    
    End If
        
End Sub










