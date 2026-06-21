VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGravirovka 
   Caption         =   "Гравировка Datron"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17145
   OleObjectBlob   =   "frmGravirovka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGravirovka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type NormyGravirovka
    tMash As Double
    tZatirka As Double
    tKontrolGravirovki As Double
    tKontrolZatirki As Double
End Type

Private Const KOLVO_STROK As Integer = 5

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Sub cmdNew_Click(): Call NovyyRaschot: End Sub

Private Sub UserForm_Initialize()

    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    With Me
        .Width = fraNormy.Left + fraNormy.Width + PLUS_SHIRINA
        .Height = cmdNew.Top + cmdNew.Height + PLUS_VYSOTA
    End With
    
    Dim ctrl As Control
    For Each ctrl In fraSpetsSimvoly.Controls
        If ctrl.Name Like "cboSimvol_*" Then
            ctrl.List = Split("Земля Молния")
        End If
    Next
    
    Call GlavnyyRaschot

End Sub

Public Sub GlavnyyRaschot()

    Dim Normy As NormyGravirovka

    Call OchistitNormy

    Dim SumDlinaLinii As Double
    Dim Instrumenty As New Collection

    RaschitatBukvy Normy, Instrumenty
    RaschitatLinii Normy, SumDlinaLinii
    RaschitatSpetsSimvoly Normy, Instrumenty

    If SumDlinaLinii > 0 Then
        With Normy
            .tKontrolGravirovki = .tKontrolGravirovki + VisualKontrolZatirkiLiniy(SumDlinaLinii, 2)
            .tKontrolZatirki = .tKontrolZatirki + VisualKontrolZatirkiLiniy(SumDlinaLinii, 1)
        End With
    End If

    With Normy
        .tMash = .tMash
        .tZatirka = 1.3 * .tZatirka
        .tKontrolGravirovki = 1.3 * .tKontrolGravirovki
        .tKontrolZatirki = 1.3 * .tKontrolZatirki
    End With
    
    ZapisatRezultaty Instrumenty, SumDlinaLinii, Normy
    
End Sub

Private Sub RaschitatBukvy(ByRef Normy As NormyGravirovka, ByRef Instrumenty As Collection)

    Dim i As Integer, Shrift As Double, KolVo As Double
    
    For i = 1 To KOLVO_STROK

        Shrift = DblFromCtrl(Me.Controls("txtShrift_" & i))
        KolVo = KolVoPoverhnosteiProhodov(Me.Controls("txtKolVo_Bukvy_" & i))
        
        If Shrift > 0 And KolVo > 0 Then
        
            With Normy
            
                .tMash = .tMash + GravirovkaDatron_Sivmoly(Shrift, KolVo)
                .tZatirka = .tZatirka + ZapolGravEmal_Znak_45(Shrift, KolVo)
                .tKontrolGravirovki = .tKontrolGravirovki + VisualnyyKontrol_Zatirka(Shrift, KolVo, , 3)
                .tKontrolZatirki = .tKontrolZatirki + VisualnyyKontrol_Zatirka(Shrift, KolVo, , 2)
            
            End With
            
            On Error Resume Next
            Instrumenty.Add Me.Controls("txtShrift_" & i), Me.Controls("txtShrift_" & i)
            On Error GoTo 0
            
        End If
            
    Next
    
End Sub

Private Sub RaschitatLinii(ByRef Normy As NormyGravirovka, ByRef SumDlinaLinii As Double)

    Dim i As Integer, DlinaLinii As Double, KolVo As Double
    
    For i = 1 To KOLVO_STROK
        
        DlinaLinii = DblFromCtrl(Me.Controls("txtDlina_" & i))
        KolVo = KolVoPoverhnosteiProhodov(Me.Controls("txtKolVo_Linii_" & i))
    
        If DlinaLinii > 0 Then
            With Normy
                .tMash = .tMash + GravirovkaDatron_Linii(KolVo * DlinaLinii)
                .tZatirka = .tZatirka + ZapolGravEmal_Riska_45(DlinaLinii, 1, 2, KolVo)
                SumDlinaLinii = SumDlinaLinii + DlinaLinii * KolVo
            End With
        End If
            
    Next
    
End Sub

Private Sub RaschitatSpetsSimvoly(ByRef Normy As NormyGravirovka, ByRef Instrumenty As Collection)

    Dim i As Integer, Vysota As Double, KolVo As Double
    
    For i = 1 To KOLVO_STROK
    
        Vysota = DblFromCtrl(Me.Controls("txtVysota_" & i))
        KolVo = DblFromCtrl(Me.Controls("txtKolVo_Simvol_" & i))
        
        If Vysota > 0 Then
            With Normy
                If Me.Controls("cboSimvol_" & i).text = "Земля" Then
                    .tMash = .tMash + KolVo * GravirovkaDatron_Zemlya(Vysota)
                ElseIf Me.Controls("cboSimvol_" & i).text = "Молния" Then
                    .tMash = .tMash + KolVo * GravirovkaDatron_Molniya(Vysota)
                End If
                .tZatirka = .tZatirka + ZapolGravEmal_Znak_45(Vysota, KolVo)
                .tKontrolGravirovki = .tKontrolGravirovki + VisualnyyKontrol_Zatirka(Vysota, KolVo, , 3)
                .tKontrolZatirki = .tKontrolZatirki + VisualnyyKontrol_Zatirka(Vysota, KolVo, , 2)
            End With

            On Error Resume Next
            Instrumenty.Add Me.Controls("cboSimvol_" & i).text & Vysota, Me.Controls("cboSimvol_" & i).text & Vysota
            On Error GoTo 0
            
        End If
        
    Next
End Sub

Private Sub ZapisatRezultaty(Instrumenty As Collection, SumDlinaLinii As Double, Normy As NormyGravirovka)

    Dim KolVoInstrumentov As Integer
    Dim Tsht_Gravirovka As Double
    
    If Normy.tMash > 0 Then
    
        KolVoInstrumentov = Instrumenty.Count + IIf(SumDlinaLinii > 0, 1, 0)
        
        Tsht_Gravirovka = Tsht_GravirovkaDatron(Normy.tMash, KolVoInstrumentov, KolVoPoverhnosteiProhodov(txtKolVoUstanovok))
        
        txtR = 3
        txtTpz = 20
        txtTsht = CStr(OkruglenieTsht(Tsht_Gravirovka))
        
        txtR_Kontrol_Gravirovka = 3
        txtTpz_Kontrol_Gravirovka = 0
        txtTsht_Kontrol_Gravirovka = CStr(OkruglenieTsht(Normy.tKontrolGravirovki))
        
        txtR_Zatirka = 2
        txtTpz_Zatirka = 0
        txtTsht_Zatirka = CStr(OkruglenieTsht(Normy.tZatirka))
        
        txtR_Kontrol_Zatirka = 3
        txtTpz_Kontrol_Zatirka = 0
        txtTsht_Kontrol_Zatirka = CStr(OkruglenieTsht(Normy.tKontrolZatirki))
    
    End If
    
End Sub

Private Sub OchistitNormy()
    Dim ctrl As Control
    
    For Each ctrl In fraNormy.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.text = "-"
        End If
    Next
    
End Sub

Private Sub NovyyRaschot()
    
    Dim ctrl As Control
    
    For Each ctrl In fraGravirovka.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.text = Empty
        ElseIf TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.ListIndex = -1
        End If
    Next
    
    Call GlavnyyRaschot
    
End Sub



