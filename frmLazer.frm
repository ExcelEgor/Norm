VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLazer 
   Caption         =   "Лазерная гравировка"
   ClientHeight    =   11775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20535
   OleObjectBlob   =   "frmLazer.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmLazer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

' Константы цвета для элементов интерфейса
Private Const TSVET_KNOPKI_AKTIVNYY As Long = &H8000000D
Private Const TSVET_KNOPKI_PASSIIVNYY As Long = &H404040

Private tbLazer As ListObject

Private Type DannyeGravirovki
    Nadpis As String
    RazmerElementa As Double
    KolVo As Double
    Tip As TipyLazernoyGravirovki
End Type

'----------------------------------------
Private Sub cmdNovyyRaschet_Click()
    OtkrytFormyVTomzheMeste Me, Me.Top, Me.Left, Me.Name
End Sub

Private Sub cmdDobavit_Click(): DobavitVTablitsu: End Sub

Private Sub cmdDelete_Click()
    UdaleniyeStrokIzLisboksa_i_Tablitsy lstLazer, tbLazer
    ZapolnitRezultaty
End Sub

Private Sub chkPodPokrasku_Click(): ZapolnitRezultaty: End Sub

'----------------------------------------
Private Sub SbrosTsvetaKnopki()
    If cmdNovyyRaschet.ForeColor <> TSVET_KNOPKI_PASSIIVNYY Then
        cmdNovyyRaschet.ForeColor = TSVET_KNOPKI_PASSIIVNYY
    End If
End Sub

Private Sub cmdNovyyRaschet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): cmdNovyyRaschet.ForeColor = TSVET_KNOPKI_AKTIVNYY: End Sub

Private Sub fraRaschot_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SbrosTsvetaKnopki: End Sub

Private Sub fraRaschot_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SbrosTsvetaKnopki: End Sub

Private Sub fraVremya_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SbrosTsvetaKnopki: End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SbrosTsvetaKnopki: End Sub

'----------------------------------------
Private Sub optAlyuminiy_Click(): IzmenitMaterial ALUMINIYEVYYE_SPLAVY: End Sub

Private Sub optStalNerzh_Click(): IzmenitMaterial STAL_NERZHAVEYUSHCHAYA: End Sub

Private Sub optStalUglerod_Click(): IzmenitMaterial STAL_UGLERODISTAYA: End Sub

Private Sub optTitan_Click(): IzmenitMaterial TITANOVYYE_SPLAVY: End Sub

Private Sub IzmenitMaterial(ByVal VybrannyyMaterial As EnumMaterialy)
    wsLazer.Cells(1, 2).Value = CInt(VybrannyyMaterial)
    GlavnyyRaschot
    ZapolnitRezultaty
End Sub

'----------------------------------------

Private Sub UserForm_Initialize()

    VyravnitElementy

    Set tbLazer = wsLazer.ListObjects("tbLazer")
    
    If Not tbLazer.DataBodyRange Is Nothing Then
        tbLazer.DataBodyRange.Delete
    End If

    Dim SpisokShriftov As Variant
    SpisokShriftov = Split("2 3 4 5 6 8 10 12")
    
    With cboShrift
        .List = SpisokShriftov
        .ListIndex = 1
    End With
    
    With cboShrift_Simvoly
        .List = SpisokShriftov
        .ListIndex = 1
    End With
    
    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    DobavitVKlass_KontrolVvodaChisel Me
    
    optAlyuminiy.Value = True
    
    OchistitRezultaty
    
End Sub

Private Sub VyravnitElementy()
    
    VyravnitPoLevomuKrayuNazvaniyaVkladok mltLaser
        
    With fraMaterial
        .Left = 9
        .Top = 21
        .Width = 327
    End With

    With chkPodPokrasku
        .Left = 6
        .Top = 6
    End With

    lblTipMarkirovki.Left = 6
    RazmestitNizhe lblTipMarkirovki, chkPodPokrasku

    With mltLaser
        .Top = lblTipMarkirovki.Top + lblTipMarkirovki.Height + 3
        .Height = .TabFixedHeight * .Pages.Count + 2 * .TabFixedHeight
        .Left = 3
        .Width = fraMaterial.Width - .Left - 6
    End With
    
    With fraGravirovka
        .Left = fraMaterial.Left
        .Width = fraMaterial.Width
        .Height = mltLaser.Top + mltLaser.Height + 12
    End With
    RazmestitNizhe fraGravirovka, fraMaterial

    With fraVremya
        .Left = fraMaterial.Left
        .Top = fraGravirovka.Top + fraGravirovka.Height + 3
        .Width = fraMaterial.Width
    End With
    
    With fraRaschot
        .Top = fraMaterial.Top
        .Left = fraMaterial.Left + fraMaterial.Width + 3
        .Width = 357
        .Height = fraVremya.Top + fraVremya.Height - .Top
    End With
    
    With fraNormy
        .Top = fraRaschot.Height - .Height - 12
    End With
    
    With lstLazer
        .Height = fraNormy.Top - .Top - 3
    End With
    
    With cmdNovyyRaschet
        .Left = fraRaschot.Left + fraRaschot.Width - .Width
        .Top = 3
    End With

    Me.Width = fraRaschot.Left + fraRaschot.Width + PLUS_SHIRINA
    Me.Height = fraVremya.Top + fraVremya.Height + PLUS_VYSOTA
    
End Sub

Public Sub GlavnyyRaschot()

    wsLazer.Cells(2, 2).Value = chkPodPokrasku.Value
    txt_tPerehod.text = Empty
    
    Dim TekushchieDannye As DannyeGravirovki
    TekushchieDannye = SobratDannyeIzFormy()
    
    If TekushchieDannye.Tip = 0 Then Exit Sub

    Dim Normy As Variant
    With TekushchieDannye
        Normy = RaschotLazernoyGravirovki(OpredelitMaterial, chkPodPokrasku.Value, .Tip, .Nadpis, .RazmerElementa, .KolVo)
    End With
    
    If IsArray(Normy) Then txt_tPerehod.text = CStr(Round(CDbl(Normy(1)), 2))
    
End Sub

Private Function SobratDannyeIzFormy() As DannyeGravirovki

    Dim Rezultat As DannyeGravirovki
    
    Select Case mltLaser.SelectedItem.Name
        Case "pNadpis"
            Rezultat.Nadpis = txtNadpis.text
            Rezultat.RazmerElementa = DblFromCtrl(cboShrift)
            Rezultat.KolVo = Len(Rezultat.Nadpis)
            Rezultat.Tip = TipyLazernoyGravirovki.Nadpis
            
        Case "pLogotip"
            Rezultat.RazmerElementa = DblFromCtrl(txtShirinaLogotipa)
            Rezultat.KolVo = 1
            Rezultat.Tip = LogotipAvrora
            
        Case "pSpetsSimvoly"
            Rezultat.RazmerElementa = DblFromCtrl(txtVisotaSimvola)
            Rezultat.KolVo = DblFromCtrl(txtKolvoSimvolov)
            If optZemlya.Value Then
                Rezultat.Tip = Zemlya
            Else
                Rezultat.Tip = Napryazheniye
            End If
            
        Case "pSimvoly"
            Rezultat.RazmerElementa = DblFromCtrl(cboShrift_Simvoly)
            Rezultat.KolVo = DblFromCtrl(txtKolVoSimvolov_PoKolVu)
            If Rezultat.RazmerElementa > 0 And Rezultat.KolVo > 0 Then
                Rezultat.Tip = KolVoSimvolov
            End If
            
        Case "pLinii"
            Rezultat.RazmerElementa = DblFromCtrl(txtDlinaLinii)
            Rezultat.KolVo = DblFromCtrl(txtKolVoLiniy)
            Rezultat.Tip = Liniya
    End Select
    
    SobratDannyeIzFormy = Rezultat
    
End Function

Private Function OpredelitMaterial() As EnumMaterialy
    If optAlyuminiy.Value Then
        OpredelitMaterial = ALUMINIYEVYYE_SPLAVY
    ElseIf optStalNerzh.Value Then
        OpredelitMaterial = STAL_NERZHAVEYUSHCHAYA
    ElseIf optStalUglerod.Value Then
        OpredelitMaterial = STAL_UGLERODISTAYA
    ElseIf optTitan.Value Then
        OpredelitMaterial = TITANOVYYE_SPLAVY
    End If
End Function

Private Sub DobavitVTablitsu()

    If tbLazer Is Nothing Then Exit Sub

    Dim DannyeDlyaZapisi As DannyeGravirovki
    DannyeDlyaZapisi = SobratDannyeIzFormy()
    
    If DannyeDlyaZapisi.Tip <> 0 And DannyeDlyaZapisi.RazmerElementa > 0 Then
        With tbLazer
            .ListRows.Add
            .DataBodyRange(.ListRows.Count, 1).Value = CInt(DannyeDlyaZapisi.Tip)
            .DataBodyRange(.ListRows.Count, 3).Value = DannyeDlyaZapisi.Nadpis
            .DataBodyRange(.ListRows.Count, 4).Value = DannyeDlyaZapisi.RazmerElementa
            .DataBodyRange(.ListRows.Count, 5).Value = DannyeDlyaZapisi.KolVo
        End With
    End If

    ZapolnitRezultaty
    
End Sub

Private Sub OchistitRezultaty()
    Dim ctrl As Control
    For Each ctrl In fraNormy.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.text = "-"
        End If
    Next
End Sub

Private Sub ZapolnitRezultaty()

    lstLazer.Clear
    
    OchistitRezultaty
    
    ZapolnitTablitsu
        
    ZapolnitNormy

End Sub

Private Sub ZapolnitTablitsu()

    Dim i As Integer
    With lstLazer
        For i = 1 To tbLazer.ListRows.Count
            .AddItem
            .List(.ListCount - 1, 0) = tbLazer.DataBodyRange(i, 2).Value
            .List(.ListCount - 1, 1) = tbLazer.DataBodyRange(i, 4).Value
            .List(.ListCount - 1, 2) = tbLazer.DataBodyRange(i, 5).Value
            .List(.ListCount - 1, 3) = CStr(Round(tbLazer.DataBodyRange(i, 6).Value, 2))
            .List(.ListCount - 1, 4) = CStr(Round(tbLazer.DataBodyRange(i, 7).Value, 2))
            .List(.ListCount - 1, 5) = CStr(Round(tbLazer.DataBodyRange(i, 8).Value, 2))
        Next
    End With
    
End Sub

Private Sub ZapolnitNormy()

    If Not tbLazer.ListRows.Count > 0 Then Exit Sub

    txtR_Lazer.text = "3"
    txtR_Zatirka.text = "2"
    txtR_Kontrol.text = "3"
    
    txtTpz_Lazer.text = "10"
    txtTpz_Zatirka.text = "0"
    txtTpz_Kontrol.text = "0"

    txtTsht_Lazer.text = CStr(3 + RaschotSummyPoStolbtsu(6))
    txtTsht_Zatirka.text = CStr(RaschotSummyPoStolbtsu(7))
    txtTsht_Kontrol.text = CStr(RaschotSummyPoStolbtsu(8))
    
End Sub

Private Function RaschotSummyPoStolbtsu(ByVal NomerStolbtsa As Integer) As Double
    RaschotSummyPoStolbtsu = OkruglenieTsht(WorksheetFunction.Sum(tbLazer.ListColumns(NomerStolbtsa).DataBodyRange))
End Function
