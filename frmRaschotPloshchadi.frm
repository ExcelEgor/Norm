VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRaschotPloshchadi 
   Caption         =   "Расчет площадей"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9180.001
   OleObjectBlob   =   "frmRaschotPloshchadi.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRaschotPloshchadi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ZapuskRaschotaTxt() As New clsmZapuskRaschetov

Private ptxtS_mm As MSForms.TextBox, ptxtS_EdIzm As MSForms.TextBox, ptxtKolVo As MSForms.TextBox


Private Sub cmdKrug_Minus_Click(): DobavitVTablitsu (True): End Sub

Private Sub cmdKrug_Plus_Click(): DobavitVTablitsu (False): End Sub

Private Sub cmdKvadrat_Plus_Click(): DobavitVTablitsu (False): End Sub

Private Sub cmdShestigran_Minus_Click(): DobavitVTablitsu (True): End Sub

Private Sub cmdShestigran_Plus_Click(): DobavitVTablitsu (False): End Sub

Private Sub cmdTruba_Minus_Click(): DobavitVTablitsu (True): End Sub

Private Sub cmdTruba_Plus_Click(): DobavitVTablitsu (False): End Sub

Private Sub cmdUgol_Minus_Click(): DobavitVTablitsu (True): End Sub

Private Sub cmdUgol_Plus_Click(): DobavitVTablitsu (False): End Sub

Private Sub UserForm_Initialize()

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Dim ctrl As Control, iTxt As Integer
    For Each ctrl In Me.Controls
        If Not LCase(ctrl.Name) Like "*area*" And Not ctrl.Name = "txtItog" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    ReDim Preserve ZapuskRaschotaTxt(iTxt)
                    Set ZapuskRaschotaTxt(iTxt).txtRaschotPloshchadi = ctrl
                    iTxt = iTxt + 1
            End Select
        End If
    Next
    
    With cboEdIzmer
        .List = Split("мм см дм м")
        .ListIndex = 3
    End With
    
End Sub

Public Sub GlavnyyRaschot()

    Call PrisvoitTekstBoksy
    
    Dim Ploshchad As Double
    Select Case mltArea.SelectedItem.Caption
        Case "Круг"
            Ploshchad = PloshchadKruga(DblFromCtrl(txtKrug_D), DblFromCtrl(txtKrug_L))
        
        Case "Прямоугольник"
            Ploshchad = PloshchadPramougolnika(DblFromCtrl(txtKvadrat_A), DblFromCtrl(txtKvadrat_B), DblFromCtrl(txtKvadrat_L))
        
        Case "Шестигранник"
            Ploshchad = PloshchadShestigrannika(DblFromCtrl(txtShestigran_S), DblFromCtrl(txtShestigran_L))

        Case "Труба"
            Ploshchad = PloshchadTruby(DblFromCtrl(txtKrug_D), DblFromCtrl(txtTruba_s), DblFromCtrl(txtTruba_L))
        
        Case "Угольник"
            Ploshchad = PloshchadUgolnika(DblFromCtrl(txtUgol_A), DblFromCtrl(txtUgol_B), DblFromCtrl(txtUgol_c), DblFromCtrl(txtUgol_d), DblFromCtrl(txtUgol_L))
    
    End Select
    
    Ploshchad = KolVoPoverhnosteiProhodov(ptxtKolVo) * Ploshchad
    ptxtS_mm = Ploshchad
    ptxtS_EdIzm = FormatirovatPloschad(KonvertEdIzmer(Ploshchad))
    
End Sub

Private Sub PrisvoitTekstBoksy()
    
    Dim ctrl As Control
    For Each ctrl In mltArea.SelectedItem.Controls
        If TypeName(ctrl) = "TextBox" Then
            If ctrl.Name Like "*_Area_mm" Then
                Set ptxtS_mm = ctrl
            ElseIf ctrl.Name Like "*_Area" Then
                Set ptxtS_EdIzm = ctrl
            ElseIf ctrl.Name Like "*_i" Then
                Set ptxtKolVo = ctrl
            End If
        End If
    Next
    
End Sub

Public Sub RaschotVTablitse()

    Dim i As Long
    Dim Ploshchad As Double
    
    With frmRaschotPloshchadi
    
        With .lstPloshchad
            For i = 0 To .ListCount - 1
                .List(i, 1) = FormatirovatPloschad(KonvertEdIzmer(CDbl(Replace(.List(i, 2), ".", ","))))
                If .List(i, 0) = "-" Then
                    Ploshchad = Ploshchad - CDbl(Replace(.List(i, 2), ".", ","))
                Else
                    Ploshchad = Ploshchad + CDbl(Replace(.List(i, 2), ".", ","))
                End If
            Next
        End With
        
        Ploshchad = KonvertEdIzmer(Ploshchad)
        .txtItog = FormatirovatPloschad(Ploshchad)
    
    End With
    
End Sub

Private Function KonvertEdIzmer(Ploshchad As Double) As Double

    Dim Delitel As Long
    Select Case frmRaschotPloshchadi.cboEdIzmer.ListIndex
        Case 0: Delitel = 1
        Case 1: Delitel = 100
        Case 2: Delitel = 10000
        Case 3: Delitel = 1000000
    End Select
    
    KonvertEdIzmer = Ploshchad / Delitel
    
End Function

Private Function FormatirovatPloschad(Ploshchad As Double) As String

    If InStr(1, CStr(Ploshchad), ",") = 0 Then
        FormatirovatPloschad = Format(Ploshchad, "#,##0")
    Else
    
        Dim NuliPosleZapyatoy As String
        Dim KolVoNuley As Integer
        
        KolVoNuley = Len(CStr(Ploshchad)) - InStr(1, CStr(Ploshchad), ",")
        If KolVoNuley > 6 Then KolVoNuley = 6
        
        NuliPosleZapyatoy = String(KolVoNuley, "0")
        FormatirovatPloschad = Format(Ploshchad, "#,##0." & NuliPosleZapyatoy)
        
    End If
    
End Function

Private Sub cboEdIzmer_Change()

    Call GlavnyyRaschot
        
    Dim str As String
    str = Me.Controls("lbllAreaEdIzmer_" & cboEdIzmer.ListIndex + 1).Caption
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Name Like "lbl*_Area_mm" Then ctrl.Caption = str
    Next

    Call RaschotVTablitse
    
End Sub

Private Sub cmdDelete_Click()

    Dim i As Long
    With lstPloshchad
        If .ListIndex <> -1 Then
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then .RemoveItem (i)
            Next
            Call RaschotVTablitse
        End If
    End With
    
End Sub

Private Sub DobavitVTablitsu(Minus As Boolean)
    
    Dim ctrl As Control
    For Each ctrl In frmRaschotPloshchadi.mltArea.SelectedItem.Controls
        If ctrl.Name Like "txt*_Area_mm" Then
            Set ctrl = ctrl
            Exit For
        End If
    Next
    
    If Minus Then Call cmdMinus(ctrl) Else Call cmdPlus(ctrl)

    Call frmRaschotPloshchadi.RaschotVTablitse
    
End Sub

Private Sub cmdPlus(Area_mm As Control)

    With lstPloshchad
        .AddItem
        .List(.ListCount - 1, 0) = "+"
        .List(.ListCount - 1, 2) = Area_mm
    End With

End Sub

Private Sub cmdMinus(Area_mm As Control)

    With lstPloshchad
        .AddItem
        .List(.ListCount - 1, 0) = "-"
        .List(.ListCount - 1, 2) = Area_mm
    End With

End Sub



