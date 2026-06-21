VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZatirka 
   Caption         =   "Затирка"
   ClientHeight    =   8595.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9240.001
   OleObjectBlob   =   "frmZatirka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZatirka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Const KolVoStrok As Integer = 5

Private Type NormyZatirka
    tZatirka As Double
    tKontrol As Double
End Type

Private Sub cboKolVoTonov_Change(): GlavnyyRaschot: End Sub

Private Sub optSlozhnost_1_Click(): GlavnyyRaschot: End Sub

Private Sub optSlozhnost_2_Click(): GlavnyyRaschot: End Sub

Private Sub UserForm_Initialize()
    
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If ctrl.Name Like "cbpTip_*" Then
            With ctrl
                .AddItem "Буква, цифра, спец. знак"
                .AddItem "Линия"
            End With
        End If
    Next
    
    cbpTip_1.ListIndex = 0
    
    With cboKolVoTonov
        .List = Array(1, 2, 3, 5, 8)
        .ListIndex = 1
    End With
    
    Call VyravnitElementy
    
    mltZatirka.Value = 0
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    Call GlavnyyRaschot

End Sub

Private Sub VyravnitElementy()

    mltZatirka.Top = 3
    mltZatirka.Left = 3
    mltZatirka.Width = fraNormy.Left + fraNormy.Width + 9
    mltZatirka.Height = fraNormy.Top + fraNormy.Height + mltZatirka.TabFixedHeight + 9
    
    Me.Height = mltZatirka.Top + mltZatirka.Height + mltZatirka.TabFixedHeight + 6
    Me.Width = mltZatirka.Left + mltZatirka.Width + 9

    
End Sub

Public Sub GlavnyyRaschot()

    Dim KolVoTonov As Integer
    If cboKolVoTonov.ListIndex = -1 Then Exit Sub
    KolVoTonov = CInt(cboKolVoTonov)

    Dim i As Integer
    Dim Normy As NormyZatirka, tKontrol As Double
    Dim cboTip As MSForms.ComboBox, txtRazmer As MSForms.TextBox, txtKolVo As MSForms.TextBox, txtVremya As MSForms.TextBox
    
    For i = 1 To KolVoStrok
        
        Set cboTip = Me.Controls("cbpTip_" & i)
        Set txtRazmer = Me.Controls("txtRazmer_" & i)
        Set txtKolVo = Me.Controls("txtKolvo_" & i)
        Set txtVremya = Me.Controls("txtVremya_" & i)
        
        Normy = RaschitatNormy(cboTip, txtRazmer, txtKolVo, KolVoTonov)
          
        If Normy.tZatirka > 0 Then
            txtVremya.text = OkruglenieTsht(Normy.tZatirka)
        Else
            txtVremya.text = "-"
        End If
        
        tKontrol = tKontrol + Normy.tKontrol
        
    Next
    
    Call ZapisatItogi(KolVoTonov, tKontrol)
    
End Sub

Private Function RaschitatNormy(cboTip As MSForms.ComboBox, txtRazmer As MSForms.TextBox, txtKolVo As MSForms.TextBox, KolVoTonov As Integer) As NormyZatirka
    
    Dim Slozhnost As Integer
    Slozhnost = IIf(optSlozhnost_1, 1, 2)

    Dim Vremya As Double, razmer As Double, KolVo As Double
    If cboTip.ListIndex <> -1 Then
        razmer = DblFromCtrl(txtRazmer)
        KolVo = DblFromCtrl(txtKolVo)
        If razmer > 0 And KolVo > 0 Then
            With RaschitatNormy
                If cboTip.ListIndex = 0 Then
                    .tZatirka = ZapolGravEmal_Znak_45(razmer, KolVo, Slozhnost, KolVoTonov)
                    .tKontrol = VisualnyyKontrol_Zatirka(razmer, KolVo, , 2)
                Else
                    .tZatirka = ZapolGravEmal_Riska_45(razmer, KolVoTonov, Slozhnost, KolVo)
                    .tKontrol = VisualKontrolZatirkiLiniy(razmer, 1)
                End If
            End With
        End If
    End If
    
End Function


Private Sub ZapisatItogi(KolVoTonov As Integer, tKontrol As Double)

    txtTsht_Zatirka.text = "-"
    txtTsht_Kontrol.text = "-"

    Dim KolVoSimvolov As Double, i As Integer
    Dim tZatirka As Double
    For i = 1 To KolVoStrok
        tZatirka = tZatirka + DblFromCtrl(Me.Controls("txtVremya_" & i))
    Next
    
    Dim Razryad As Integer
    Select Case KolVoTonov
        Case Is <= 2:   Razryad = 2
        Case Is <= 5:   Razryad = 3
        Case Else:      Razryad = 4
    End Select
    txtR_Zatirka.text = Razryad
    
    If tZatirka > 0 Then txtTsht_Zatirka.text = CStr(OkruglenieTsht(tZatirka))
    If tKontrol > 0 Then txtTsht_Kontrol.text = CStr(OkruglenieTsht(tKontrol))
    
End Sub
