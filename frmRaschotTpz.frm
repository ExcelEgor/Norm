VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRaschotTpz 
   Caption         =   "Ðàñ÷¸ò Òïç"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580.001
   OleObjectBlob   =   "frmRaschotTpz.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRaschotTpz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TablitsaTpzFrez As Range

Private Sub cboDlinaStola_Change(): Call RaschitatTpz: End Sub

Private Sub cboOborudovaniye_Frez_Change(): Call RaschitatTpz: End Sub

Private Sub cboTsekh_Frez_Change(): Call ZapolnitUchastok_Frez: End Sub

Private Sub cboUchastok_Change(): Call RaschitatTpz: End Sub

Private Sub cboUchastok_Frez_Change(): Call ZapolnitOborudovaniye_Frez: End Sub

Private Sub chkKulachki_Click(): Call RaschitatTpz: End Sub

Private Sub chkLunet_Click(): Call RaschitatTpz: End Sub

Private Sub chkPlanshayba_Click(): Call RaschitatTpz: End Sub

Private Sub MultiPage1_Change(): Call RaschitatTpz: End Sub


Private Sub opt1_0_Click(): Call RaschitatTpz: End Sub

Private Sub opt1_2_Click(): Call RaschitatTpz: End Sub

Private Sub opt1_3_Click(): Call RaschitatTpz: End Sub

Private Sub opt1_5_Click(): Call RaschitatTpz: End Sub


Private Sub txtKolVoInstrumentov_Change(): Call RaschitatTpz: End Sub

Private Sub txtKolVoOsnastki_Frez_Change(): Call RaschitatTpz: End Sub

Private Sub txtKolVoOsnastki_GRS_Change(): Call RaschitatTpz: End Sub

Private Sub txtTsht_Change(): Call RaschitatTpz: End Sub


Private Sub lblFrezernyy_Click(): PerekluchitVkladku 1, lblFrezernyy: End Sub

Private Sub lblGRS_Click(): PerekluchitVkladku 2, lblGRS: End Sub

Private Sub lblTokarnyy_Click(): PerekluchitVkladku 0, lblTokarnyy: End Sub


Private Sub lblTokarnyy_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): MouseDown lblTokarnyy: End Sub

Private Sub lblTokarnyy_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):  MouseUp lblTokarnyy: End Sub

Private Sub lblFrezernyy_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): MouseDown lblFrezernyy: End Sub

Private Sub lblFrezernyy_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): MouseUp lblFrezernyy: End Sub

Private Sub lblGRS_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): MouseDown lblGRS: End Sub

Private Sub lblGRS_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): MouseUp lblGRS: End Sub


Private Sub spnKolVoOsnastki_Frez_SpinDown(): SpinDown txtKolVoOsnastki_Frez: End Sub

Private Sub spnKolVoOsnastki_Frez_SpinUp(): SpinUp txtKolVoOsnastki_Frez: End Sub

Private Sub spnKolVoOsnastki_GRS_SpinDown(): SpinDown txtKolVoOsnastki_GRS: End Sub

Private Sub spnKolVoOsnastki_GRS_SpinUp(): SpinUp txtKolVoOsnastki_GRS: End Sub


Private Sub UserForm_Initialize()
    
    Call lblTokarnyy_Click

    With cboUchastok
        .ColumnCount = 2
        .List = wsTpz.Range("A2:B6").Value
        .ColumnWidths = "; 0pt"
        .ListIndex = 0
    End With
    
    
    Set TablitsaTpzFrez = wsTpz.Range("D2:G12")
    Dim i As Integer
    
    For i = 1 To TablitsaTpzFrez.Rows.Count
        With cboTsekh_Frez
            If i = 1 Then
                .AddItem TablitsaTpzFrez(i, 1)
            Else
                If TablitsaTpzFrez(i, 1) <> TablitsaTpzFrez(i - 1, 1) Then
                    .AddItem TablitsaTpzFrez(i, 1)
                End If
            End If
            .ListIndex = 0
        End With
    Next
    
    MultiPage1.Style = fmTabStyleNone
   
End Sub

Public Sub ZagruzitParametryVFormu(KolVoInstrumentov As Integer, Tsht As String)

    With frmRaschotTpz
        .txtKolVoInstrumentov.text = KolVoInstrumentov
        .txtKolVoInstrumentov.Tag = KolVoInstrumentov
        .txtTsht.text = Tsht
        .txtTsht.Tag = Tsht
        .Show
    End With
    
End Sub

Private Sub RaschitatTpz()

    txtTpz.text = Empty

    Dim KolVoInstrumentov As Integer
    KolVoInstrumentov = DblFromCtrl(txtKolVoInstrumentov)
    
    Dim kOtarabotki As Double
    Select Case True
        Case opt1_0: kOtarabotki = 1
        Case opt1_3: kOtarabotki = 1.3
        Case opt1_5: kOtarabotki = 1.5
    End Select
    
    Dim Tpz As Integer
    Select Case MultiPage1.Value
        Case 0
            With cboUchastok
                If .ListIndex = -1 Then Exit Sub
                Dim DiametrNadStaninoy As Double
                DiametrNadStaninoy = CDbl(.List(.ListIndex, 1))
                Tpz = TpzCHPU_Tokarnyy(DiametrNadStaninoy, KolVoInstrumentov, DblFromCtrl(txtTsht), kOtarabotki, chkPlanshayba, chkLunet, chkKulachki)
            End With
            
        Case 1
            With cboOborudovaniye_Frez
                If .ListIndex = -1 Then Exit Sub
                Dim DlinaStola As Double
                DlinaStola = CDbl(.List(.ListIndex, 1))
                Tpz = TpzCHPU_Frezernyy(DlinaStola, KolVoInstrumentov, DblFromCtrl(txtTsht), kOtarabotki, DblFromCtrl(txtKolVoOsnastki_Frez))
            End With
            
        Case 2
            Tpz = TpzCHPU_GorizantalnoRastochnoy(KolVoInstrumentov, DblFromCtrl(txtTsht), kOtarabotki, DblFromCtrl(txtKolVoOsnastki_GRS))
    End Select
    
    txtTpz.text = Tpz

End Sub

Private Sub PerekluchitVkladku(mltValue As Integer, lblPage As MSForms.Label)
    
    MultiPage1.Value = mltValue
    lblPage.BackColor = &H80000018
    
    Dim ArrLabels
    ArrLabels = Array(lblTokarnyy, lblFrezernyy, lblGRS)
    Dim i As Integer
    For i = LBound(ArrLabels) To UBound(ArrLabels)
        If ArrLabels(i).Name <> lblPage.Name Then ArrLabels(i).BackColor = vbWhite
    Next
    
End Sub

Private Sub MouseDown(lbl As MSForms.Label)
    lbl.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub MouseUp(lbl As MSForms.Label)
    lbl.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub SpinDown(txtKolVoOsnastki As MSForms.TextBox)
    If CInt(txtKolVoOsnastki) > 0 Then txtKolVoOsnastki.text = CInt(txtKolVoOsnastki) - 1
End Sub

Private Sub SpinUp(txtKolVoOsnastki As MSForms.TextBox)
    txtKolVoOsnastki.text = CInt(txtKolVoOsnastki) + 1
End Sub

Private Sub ZapolnitUchastok_Frez()

    Dim i As Integer
    With cboUchastok_Frez
        .Clear
        .Enabled = True
        For i = 1 To TablitsaTpzFrez.Rows.Count
            If CStr(TablitsaTpzFrez(i, 1)) = cboTsekh_Frez Then
                .AddItem TablitsaTpzFrez(i, 2)
                
            End If
        Next
        
        For i = .ListCount - 1 To 1 Step -1
            If .List(i, 0) = .List(i - 1, 0) Then .RemoveItem (i)
        Next
        
        .ListIndex = 0
        If .ListCount = 1 Then .Enabled = False
        
    End With
    
End Sub

Private Sub ZapolnitOborudovaniye_Frez()

    Dim i As Integer
    With cboOborudovaniye_Frez
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "; 0pt"
        .Enabled = True

        For i = 1 To TablitsaTpzFrez.Rows.Count
            If CStr(TablitsaTpzFrez(i, 1)) = cboTsekh_Frez And CStr(TablitsaTpzFrez(i, 2)) = cboUchastok_Frez Then
                .AddItem
                .List(.ListCount - 1, 0) = TablitsaTpzFrez(i, 3)
                .List(.ListCount - 1, 1) = TablitsaTpzFrez(i, 4)
            End If
        Next
        
        If cboTsekh_Frez.ListIndex <> -1 And cboUchastok_Frez.ListIndex <> -1 Then
            .ListIndex = 0
            If .ListCount = 1 Then .Enabled = False
        End If
        
    End With

End Sub
