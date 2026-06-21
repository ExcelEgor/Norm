VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetallografika 
   Caption         =   "Металлографика"
   ClientHeight    =   12060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21930
   OleObjectBlob   =   "frmMetallografika.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmMetallografika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Handler As clsGlavnyyRaschot
Private ControlHandlers As Collection

Private Const Pritupleniye As Boolean = True

Private Type GeometryaIzdediya
    DlinaDetali As Double
    ShirinaDetali As Double
    DlinaZagotovki As Double
    ShirinaZagotovki As Double
    TolshchinaZagotovki As Double
End Type

Private Sub chkDatron_Click()
    DostupnostDatron
End Sub

Private Sub cmdDelete_Click()
    UdalitStroku
End Sub

Private Sub cmdDobavit_Click()
    DobavitVTablitsu
End Sub

Private Sub cmdNovyyRaschet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovyyRaschet.ForeColor = &H8000000D
End Sub

Private Sub lblZagolovok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovyyRaschet.ForeColor = &H80000006
End Sub

Private Sub fraRaschot_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovyyRaschet.ForeColor = &H80000006
End Sub

Private Sub fraZagtovka_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovyyRaschet.ForeColor = &H80000006
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovyyRaschet.ForeColor = &H80000006
End Sub

Private Sub cmdNovyyRaschet_Click()
    OchistitFormu
End Sub

Private Sub UserForm_Initialize()

    Set ControlHandlers = New Collection
    DobavitVKlass_GlavnyyRaschot Me, ControlHandlers, Handler
    
    Call DobavitVKlass_KontrolVvodaChisel(Me)
    
    VyravnitElementy
    
    chkDatron_Click
    
End Sub

Private Sub VyravnitElementy()

    lblZagolovok.Top = 0
    lblZagolovok.Left = -1
    
    VyravnitElementy_Datron
    
    VyravnitPoLevomuKrayuNazvaniyaVkladok mltDatron
    
    VyravnitElementy_DatronZagotovkaDetal
    
    VyravnitElementy_RashotZagolovokForma
    
End Sub

Private Sub VyravnitElementy_Datron()

    Const SHIRINA_BLOKA_ELEMENTOV As Integer = 381
    
    fraDatron.Left = 3 * SMESHCHENIYE
    
    chkDatron.Left = 2 * SMESHCHENIYE
    chkDatron.Top = 3 * SMESHCHENIYE
    
    RazmestitNizhe fraElementy, chkDatron
    fraElementy.Left = chkDatron.Left
    fraElementy.Width = SHIRINA_BLOKA_ELEMENTOV
    
    RazmestitPravee lblTabl, fraElementy
    lblTabl.Top = fraElementy.Top
    
    lblElement.Left = SMESHCHENIYE
    lblElement.Top = SMESHCHENIYE
    RazmestitNizhe mltDatron, lblElement, 3
    
    VyravnitElementy_Mlt
    
    RazmestitPravee lstDatron, fraElementy

    fraElementy.Height = cmdDobavit.Top + cmdDobavit.Height + 3 * SMESHCHENIYE
    fraDatron.Height = fraElementy.Top + fraElementy.Height + 7 * SMESHCHENIYE

    lstDatron.Top = lblTabl.Top + lblTabl.Height
    lstDatron.Height = fraElementy.Height - lblTabl.Height
    
    RazmestitPravee cmdDelete, lstDatron, -SMESHCHENIYE
    cmdDelete.Top = lstDatron.Top
    
End Sub

Private Sub VyravnitElementy_Mlt()
    
    With mltDatron
        .Left = SMESHCHENIYE
        .Width = fraElementy.Width - (2 * SMESHCHENIYE)
        .Height = (.TabFixedHeight * .Pages.Count) + (2 * .TabFixedHeight)
    End With
    
    With cmdDobavit
        .Left = mltDatron.Left - 1
        .Top = mltDatron.Top + mltDatron.Height + (2 * SMESHCHENIYE)
        .Width = mltDatron.TabFixedWidth + 2
    End With
    
    Dim ctrl As Control, i As Long
    For i = 0 To mltDatron.Pages.Count - 1
        For Each ctrl In mltDatron.Pages(i).Controls
            If TypeOf ctrl Is MSForms.Frame Then
                With ctrl
                    .Left = SMESHCHENIYE
                    .Top = 0
                    .Width = mltDatron.Width - mltDatron.TabFixedWidth - .Left - SMESHCHENIYE
                    .Height = mltDatron.Height - SMESHCHENIYE
                End With
            End If
        Next
    Next
    
End Sub

Private Sub VyravnitElementy_DatronZagotovkaDetal()
    
    RazmestitNizhe fraZagtovka, lblZagolovok, 2 * SMESHCHENIYE
    
    fraZagtovka.Left = fraDatron.Left

    Dim obshchayaShirina As Single
    obshchayaShirina = cmdDelete.Left + cmdDelete.Width + (3 * SMESHCHENIYE)
    
    fraDatron.Width = obshchayaShirina
    fraZagtovka.Width = (fraDatron.Width / 2) - 1.5
    fraDetal.Width = fraZagtovka.Width
    fraDetal.Left = fraDatron.Left + fraDatron.Width - fraDetal.Width
    fraDetal.Top = fraZagtovka.Top
    RazmestitNizhe fraDatron, fraDetal
    
End Sub

Private Sub VyravnitElementy_RashotZagolovokForma()

    RazmestitPravee fraRaschot, fraDatron
    fraRaschot.Top = fraZagtovka.Top
    fraRaschot.Height = fraDatron.Top + fraDatron.Height - fraRaschot.Top
    lstMetallografika.Height = fraRaschot.Height - fraRaschot.Top - 3 * SMESHCHENIYE
    
    Me.Width = fraRaschot.Left + fraRaschot.Width + PLUS_SHIRINA
    Me.Height = fraDatron.Top + fraDatron.Height + PLUS_VYSOTA
    
    lblZagolovok.Width = Me.Width
    
    cmdNovyyRaschet.Top = 3 * SMESHCHENIYE
    cmdNovyyRaschet.Left = fraRaschot.Left + fraRaschot.Width - cmdNovyyRaschet.Width
    
End Sub

Public Sub GlavnyyRaschot()

    lstMetallografika.Clear

    Dim Geom As GeometryaIzdediya
    
    With Geom
        .DlinaDetali = DblFromCtrl(txtL_Detal)
        .ShirinaDetali = DblFromCtrl(txtB_Detal)
        .DlinaZagotovki = DblFromCtrl(txtL_Zagotovka)
        .ShirinaZagotovki = DblFromCtrl(txtB_Zagotovka)
        .TolshchinaZagotovki = DblFromCtrl(txtS_Zagotovka)
        
        SbrositTsveta
        
        Dim Oshibki As Collection
        Dim Raschot As New clsMetallografika_Frm
    
        Dim ZapolnenoVerno As Boolean
        ZapolnenoVerno = Raschot.ProverkaParametrov(.DlinaDetali, .ShirinaDetali, .DlinaZagotovki, .ShirinaZagotovki, .TolshchinaZagotovki, Oshibki)
        
        If Not ZapolnenoVerno Then
            Dim ctrlName As Variant
            For Each ctrlName In Oshibki
                Me.Controls(ctrlName).BackColor = vbRed
            Next
            Exit Sub
        End If
    
        Dim Marshrut As Variant
        Marshrut = Raschot.VychislitMarshrut( _
            Geom.DlinaDetali, Geom.ShirinaDetali, Geom.DlinaZagotovki, Geom.ShirinaZagotovki, Geom.TolshchinaZagotovki, _
            chkDatron.Value, _
            DblFromCtrl(txtSum_Datron), _
            DblFromCtrl(txtSum_Sles), _
            DblFromCtrl(txtSum_Kontrol), _
            DblFromCtrl(txtSum_Vyrez), _
            KolVoPoverhnosteiProhodov(txtKolVoDetaley), _
            Pritupleniye)
        
    End With
    
    If IsArray(Marshrut) Then
        FormatriovatTsht Marshrut
        lstMetallografika.List = Marshrut
    End If
    
    Set Raschot = Nothing
    
End Sub

Private Sub SbrositTsveta()

    Dim ctrl As Control
    For Each ctrl In fraDetal.Controls
        If TypeOf ctrl Is MSForms.TextBox Then ctrl.BackColor = vbWhite
    Next
    For Each ctrl In fraZagtovka.Controls
        If TypeOf ctrl Is MSForms.TextBox Then ctrl.BackColor = vbWhite
    Next
    
End Sub

Private Sub DostupnostDatron()

    Dim ctrl As Control, i As Long
    For i = 0 To mltDatron.Pages.Count - 1
        For Each ctrl In mltDatron.Pages(i).Controls
            ctrl.Enabled = chkDatron
        Next
    Next
    
    For Each ctrl In fraDatron.Controls
        ctrl.Enabled = chkDatron
    Next
    
    chkDatron.Enabled = True
End Sub

Private Sub FormatriovatTsht(ByRef Marshrut As Variant)
    Dim i As Long
    
    For i = LBound(Marshrut, 1) To UBound(Marshrut, 1)
        Marshrut(i, 6) = CStr(Marshrut(i, 6))
    Next i
End Sub

Private Sub DobavitVTablitsu()
    
    Dim Perekhod As String
    Dim Normy As Variant
    
    Normy = RaschitatDatron(Perekhod)
    
    If IsArray(Normy) Then
        If Normy(1) > 0 Then
            With lstDatron
                .AddItem
                .List(.ListCount - 1, 0) = Perekhod
                .List(.ListCount - 1, 1) = Normy(1)
                .List(.ListCount - 1, 2) = Normy(2)
                .List(.ListCount - 1, 3) = Normy(3)
                .List(.ListCount - 1, 4) = Normy(4)
            End With
            
            RasschitatSummu
            
            GlavnyyRaschot
            
        End If
    End If
    
End Sub

Private Function RaschitatDatron(ByRef Perekhod As String) As Variant
    
    Dim Normy As Variant
    
    Select Case mltDatron.SelectedItem.Name
        Case "pOkna"
            Normy = Datron_tMashPerimetr( _
                DblFromCtrl(txtDlina_Okno), DblFromCtrl(txtShirina_Okno), chkUgol_Okno.Value, Pritupleniye, KolVoPoverhnosteiProhodov(txtKolVo_Okno))
            Perekhod = "Окно " & txtDlina_Okno & "x" & txtShirina_Okno & " (х" & KolVoPoverhnosteiProhodov(txtKolVo_Okno) & ")"
            
        Case "pOtverstiya"
            Normy = Datron_tOtverstiye(DblFromCtrl(txtD_otv), chkZamok_Otv.Value, Pritupleniye, KolVoPoverhnosteiProhodov(txtKolVo_Otv))
            Perekhod = "Отверстие Ф" & txtD_otv & " (х" & KolVoPoverhnosteiProhodov(txtKolVo_Otv) & ")"
            
        Case "pKvadrat"
            Normy = Datron_tMashPerimetr( _
                DblFromCtrl(txtRazmer_Kvadrat), DblFromCtrl(txtRazmer_Kvadrat), True, Pritupleniye, KolVoPoverhnosteiProhodov(txtKolVo_Kvadrat))
            Perekhod = "Кнопка " & txtRazmer_Kvadrat & " (х" & KolVoPoverhnosteiProhodov(txtKolVo_Kvadrat) & ")"
            
        Case "pDorabotkaUglov"
            Normy = Datron_tMashUgol(DblFromCtrl(txtKolVo_Dorabotka))
            Perekhod = "Доработка угла  (х" & KolVoPoverhnosteiProhodov(txtKolVo_Dorabotka) & ")"
    End Select
    
    RaschitatDatron = Normy
    
End Function

Private Sub UdalitStroku()

    If lstDatron.ListIndex = -1 Then Exit Sub
        
    Dim i As Long
    
    With lstDatron
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                .RemoveItem (i)
            End If
        Next
    End With
    
    RasschitatSummu
    
    GlavnyyRaschot
    
End Sub

Private Sub RasschitatSummu()

    txtSum_Datron = "-"
    txtSum_Kontrol = "-"
    txtSum_Sles = "-"
    txtSum_Vyrez = "-"
    
    If chkDatron.Value = False Then Exit Sub
    
    Dim tDatron As Double, tSles As Double, tKontrol As Double, tVyrez As Double
    
    Dim i As Long
    With lstDatron
        For i = 0 To .ListCount - 1
            tDatron = tDatron + Val(.List(i, 1))
            tSles = tSles + Val(.List(i, 2))
            tKontrol = tKontrol + Val(.List(i, 3))
            tVyrez = tVyrez + Val(.List(i, 4))
        Next
    End With

    If tDatron > 0 Then
        txtSum_Datron = CStr(tDatron)
        txtSum_Kontrol = CStr(tKontrol)
        txtSum_Sles = CStr(tSles)
        txtSum_Vyrez = CStr(tVyrez)
    End If
    
    
End Sub

Private Sub OchistitFormu()

    lstDatron.Clear
    lstMetallografika.Clear
    
    RasschitatSummu

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ' Не трогаем поля сумм, так как RasschitatSummu уже записала туда "-"
            If Left(ctrl.Name, 6) <> "txtSum" Then
                ctrl.text = ""
                ctrl.BackColor = vbWhite
            End If
        ElseIf TypeOf ctrl Is MSForms.CheckBox Then
            ctrl.Value = False
        End If
    Next ctrl
    
    txtL_Zagotovka.SetFocus
    
End Sub

