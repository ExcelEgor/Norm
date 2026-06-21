VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputBox 
   Caption         =   "Номер операции"
   ClientHeight    =   810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4080
   OleObjectBlob   =   "frmInputBox.frx":0000
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ws As Worksheet

Private Sub cmdCancel_Click()
    frmTokFrez.OpName = Empty
    txtNumOp = Empty
    txtNumOp.SetFocus
    Me.Hide
End Sub

Private Sub cmdOK_Click()

    frmTokFrez.OpName = txtNumOp
    txtNumOp = Empty
    txtNumOp.SetFocus

    If frmTokFrez.OpName <> Empty And TakoyNomerOperatsiiUzheEst = False Then
        frmTokFrez.tabOperatsii.SelectedItem.Caption = frmTokFrez.OpName
        ws.Name = frmTokFrez.OpName
    End If
    
    Me.Hide
    
End Sub

Private Function TakoyNomerOperatsiiUzheEst() As Boolean
    
    TakoyNomerOperatsiiUzheEst = False
    
    Dim i As Integer
    For i = 0 To frmTokFrez.tabOperatsii.tabs.Count - 1
        If frmTokFrez.tabOperatsii.tabs(i).Caption = frmTokFrez.OpName Then
            TakoyNomerOperatsiiUzheEst = True
            Exit For
        End If
    Next
    
End Function

Sub Init(wsTokarnoFrezer As Worksheet)
    Set ws = wsTokarnoFrezer
End Sub

