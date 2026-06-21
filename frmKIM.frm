VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKIM 
   Caption         =   "Нормирование КИМ"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16605
   OleObjectBlob   =   "frmKIM.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmKIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Raschot As New clsKIM_Ffm_Raschoty

Public ParameryKIM As clsKIM
Private Tablitsa As New clsKIM_Tablitsa

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Const LBL_TOP As Integer = 24

Private Sub cmdDelete_Click(): Tablitsa.UdalitStroki: End Sub

Private Sub cmdDobavit_Click(): Call Tablitsa.DobavitNovuyuStroku(ParameryKIM): End Sub

Private Sub cmdNovyyRaschet_Click(): Tablitsa.NovyyRaschot: End Sub

Private Sub lblHera_Click(): mltKIM.Value = 0: End Sub

Private Sub lblMikroskop_Click(): mltKIM.Value = 2: End Sub

Private Sub lblRoundScan_Click(): mltKIM.Value = 1: End Sub

Private Sub mltKIM_Change(): IzmenitStilVkladok: End Sub

Private Sub optSloznost_1_Click(): Tablitsa.ZapisatParametryOperatsii: End Sub

Private Sub optSloznost_2_Click(): Tablitsa.ZapisatParametryOperatsii: End Sub

Private Sub optSloznost_3_Click(): Tablitsa.ZapisatParametryOperatsii: End Sub

Private Sub spnKalibrovka_1_SpinUp(): txtKalibrovka_1 = txtKalibrovka_1 + 1: End Sub

Private Sub spnKalibrovka_2_SpinUp(): txtKalibrovka_2 = txtKalibrovka_2 + 1: End Sub

Private Sub spnKalibrovka_3_SpinUp(): txtKalibrovka_3 = txtKalibrovka_3 + 1: End Sub

Private Sub spnKalibrovka_1_SpinDown()
    If CInt(txtKalibrovka_1) > 0 Then txtKalibrovka_1 = txtKalibrovka_1 - 1
End Sub

Private Sub spnKalibrovka_2_SpinDown()
    If CInt(txtKalibrovka_2) > 0 Then txtKalibrovka_2 = txtKalibrovka_2 - 1
End Sub

Private Sub spnKalibrovka_3_SpinDown()
    If CInt(txtKalibrovka_3) > 0 Then txtKalibrovka_3 = txtKalibrovka_3 - 1
End Sub

Private Sub UserForm_Initialize()
    
    mltHera.Value = 0
    Tablitsa.NovyyRaschot
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Name Like "cboSloznost_*" Then
            ctrl.List = Split("Простая Средняя Сложная")
            ctrl.ListIndex = 0
        End If
    Next
    
    VyravnitElementy
    
    IzmenitStilVkladok

End Sub

Private Sub VyravnitElementy()

    With lblGlavnyyZagolovok
        .Left = 0
        .Top = 0
    End With
    
    With franVvod
        .Top = lblGlavnyyZagolovok.Top + lblGlavnyyZagolovok.Height + 6
        .Width = 468
        .Height = 276
    End With
    
    lblHera.Top = LBL_TOP
    lblHera.Left = 3
    
    lblLine.Left = lblHera.Left
    lblLine.Top = lblHera.Top + lblHera.Height
    
    With mltKIM
        .Style = fmTabStyleNone
        .Top = lblLine.Top + lblLine.Height
        .Left = 2
        .Width = franVvod.Width - 9
        .TabFixedWidth = .Width / 3 - 2.5
        .Height = franVvod.Height - .Top - 12
    End With
    
    lblLine.Width = mltKIM.Width - 1
    lblHera.Width = mltKIM.Width / 3
    
    lblRoundScan.Left = lblHera.Width + lblHera.Left
    lblRoundScan.Width = mltKIM.Width / 3
    
    lblMikroskop.Left = lblRoundScan.Width + lblRoundScan.Left
    lblMikroskop.Width = mltKIM.Width / 3
    
    With fraVremya
        .Left = franVvod.Left
        .Top = franVvod.Top + franVvod.Height + 3
        .Width = franVvod.Width
    End With
    
    With fraTablitsa
        .Left = franVvod.Left + franVvod.Width + 3
        .Top = franVvod.Top
        .Height = fraVremya.Top + fraVremya.Height - .Top
    End With
       
    With txtTsht
        .Top = fraTablitsa.Height - txtTsht.Height - 12
        .Left = 54
    End With
    
    With txtTpz
        .Top = txtTsht.Top - txtTsht.Height - 3
        .Left = 54
    End With
    
    With lblTsht
        .Left = 3
        .Top = txtTsht.Top + txtTsht.Height / 8
    End With
    
    With lblTpz
        .Left = 3
        .Top = txtTpz.Top + txtTpz.Height / 8
    End With
    
    lstHera.Height = txtTpz.Top - lstHera.Top - 3

    Call VyravnitElementyVnitriVkladok
    
    cmdNovyyRaschet.Left = fraVremya.Left
    cmdNovyyRaschet.Top = fraVremya.Top + fraVremya.Height + 6
    
    Me.Width = fraTablitsa.Left + fraTablitsa.Width + 9
    lblGlavnyyZagolovok.Width = Me.Width
    Me.Height = cmdNovyyRaschet.Top + cmdNovyyRaschet.Height + 30
    
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok(mltHera)
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok(mltMikroskop)
    Call VyravnitPoLevomuKrayuNazvaniyaVkladok(mltRoundScan)
    
End Sub

Private Sub VyravnitElementyVnitriVkladok()

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        
        If ctrl.Name Like "lblZagolovok*" Then
            ctrl.Left = 3
            ctrl.Top = 6
            ctrl.Width = mltKIM.Width
        ElseIf TypeName(ctrl) = "MultiPage" And ctrl.Name <> "mltKIM" Then
            ctrl.Top = 21
            ctrl.Left = 3
            ctrl.Width = mltKIM.Width - 9
            ctrl.Height = mltKIM.Height - ctrl.Top - 6
        End If

    Next
    
End Sub

Private Sub IzmenitStilVkladok()
    
    Dim lblNeaktiv
    Dim lblAktiv As Control

    If mltKIM.Value = 0 Then
        Set lblAktiv = lblHera
        lblNeaktiv = Array(lblRoundScan, lblMikroskop)
    ElseIf mltKIM.Value = 1 Then
        Set lblAktiv = lblRoundScan
        lblNeaktiv = Array(lblMikroskop, lblHera)
    Else
        Set lblAktiv = lblMikroskop
        lblNeaktiv = Array(lblRoundScan, lblHera)
    End If
    
    With lblAktiv
        .BackColor = &HFFFFFF
        .ForeColor = &H8000000D
        .Font.Name = "Segoe UI Semibold"
        .Top = LBL_TOP
        .Height = 21
    End With
    
    Dim i As Integer
    For i = 0 To UBound(lblNeaktiv)
        With lblNeaktiv(i)
            .BackColor = &H8000000F
            .ForeColor = &H80000011
            .Font.Name = "Segoe UI"
            .Top = LBL_TOP + 3
            .Height = 18
        End With
     Next

    
End Sub

Public Sub GlavnyyRaschot()
    Raschot.VremyaKIM
End Sub







