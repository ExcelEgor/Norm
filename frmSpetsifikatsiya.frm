VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpetsifikatsiya 
   Caption         =   "Расчёт из спецификации"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4185
   OleObjectBlob   =   "frmSpetsifikatsiya.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpetsifikatsiya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wb_spetsifikatsiya As Workbook

Private Sub cmdBlok_Click()
    Call mIzSpetsifikatsii.RaschotBlokovStoyekModuley(wb_spetsifikatsiya, cmdBlok)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdModul_Click()
    Call mIzSpetsifikatsii.RaschotBlokovStoyekModuley(wb_spetsifikatsiya, cmdModul)
End Sub

Private Sub cmdPlata_Click()
    Application.Run "montazh.xlam!RaschotKomplektovaniya", wb_spetsifikatsiya
End Sub

Private Sub cmdSborka_Click()
    Call mIzSpetsifikatsii.RaschotKomplektovaniyaSborki(wb_spetsifikatsiya)
End Sub


Private Sub cmdStoyka_Click()
    Call mIzSpetsifikatsii.RaschotBlokovStoyekModuley(wb_spetsifikatsiya, cmdStoyka)
End Sub

Private Sub lblBlok_1_Click()
    Call cmdBlok_Click
End Sub

Private Sub lblBlok_2_Click()
    Call cmdBlok_Click
End Sub

Private Sub lblModul_1_Click()
    Call cmdModul_Click
End Sub

Private Sub lblModul_2_Click()
    Call cmdModul_Click
End Sub

Private Sub lblStoyka_1_Click()
    Call cmdStoyka_Click
End Sub

Private Sub lblStoyka_2_Click()
    Call cmdStoyka_Click
End Sub

Private Sub UserForm_Initialize()
    Set wb_spetsifikatsiya = ActiveWorkbook
End Sub

