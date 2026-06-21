VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGeometrRascheti 
   Caption         =   "Геометрические расчеты"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4035
   OleObjectBlob   =   "frmGeometrRascheti.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGeometrRascheti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RaschetDliniLiski(txtD As MSForms.TextBox, txtH As MSForms.TextBox, txtL As MSForms.TextBox)

    Dim R As Double
    Dim H As Double
    Dim L As Double
    
    R = DblFromCtrl(txtD) / 2
    H = DblFromCtrl(txtH)
    
    If R > 0 And H > 0 And R > H Then
        H = R - H
        L = Sqr(8 * R * H - 4 * H ^ 2)
        txtL.text = Format(L, "0.0")
    Else
        txtL.text = Empty
    End If
    
End Sub

Private Sub RaschetOpisannogoDiametra()
    txtDopis_D.text = Empty
    Dim a As Double, b As Double, d As Double
    a = DblFromCtrl(txtDopis_A)
    b = DblFromCtrl(txtDopis_B)
    If a > 0 And b > 0 Then
        d = Sqr(a ^ 2 + b ^ 2)
        txtDopis_D.text = Format(d, "0.0")
    End If
End Sub

Private Sub txtDopis_A_Change()
    Call RaschetOpisannogoDiametra
End Sub

Private Sub txtDopis_B_Change()
    Call RaschetOpisannogoDiametra
End Sub

Private Sub txtD_Liska_Change()
    Call RaschetDliniLiski(txtD_Liska, txth_Liska, txtL_Liska)
End Sub

Private Sub txth_Liska_Change()
    Call RaschetDliniLiski(txtD_Liska, txth_Liska, txtL_Liska)
End Sub

Private Sub txtS_RaschetShestigran_Change()
    Dim S As Double, L As Double, H As Double
    If IsNumeric(txtS_RaschetShestigran) Then
        S = CDbl(txtS_RaschetShestigran.text)
        L = DlinaGraniShestigrannika(S)
        H = (L * 2 - S) / 2
        txtL_RaschetShestigran.text = Format(L, "0.0")
        txth_RaschetShestigran.text = Format(H, "0.0")
    Else
        txtL_RaschetShestigran.text = Empty
        txth_RaschetShestigran.text = Empty
    End If
End Sub
