VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTokFrez_Tablitsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pListBox As MSForms.ListBox
Private ptxtTsht_TokFrez As MSForms.TextBox
Private ptxtTsht_Slesar As MSForms.TextBox
Private ptxtTsht_Kontrol As MSForms.TextBox
Private pChkOkruglennyyeNomvy As MSForms.CheckBox

Private pTbRaschot As ListObject

Public Sub Init(ListBox As MSForms.ListBox, _
    txtTsht_TokFrez As MSForms.TextBox, txtTsht_Slesar As MSForms.TextBox, txtTsht_Kontrol As MSForms.TextBox, ChkOkruglennyyeNomvy As MSForms.CheckBox)

    Set pListBox = ListBox
    Set ptxtTsht_TokFrez = txtTsht_TokFrez
    Set ptxtTsht_Slesar = txtTsht_Slesar
    Set ptxtTsht_Kontrol = txtTsht_Kontrol
    Set pChkOkruglennyyeNomvy = ChkOkruglennyyeNomvy

End Sub

Public Sub UdalitStroki(tbRaschot As ListObject)

    Dim i As Integer, StrokaUdaleniya As Integer
    
    With pListBox
        If .ListIndex = -1 Then Exit Sub
        
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                StrokaUdaleniya = i + 1
                tbRaschot.ListRows(StrokaUdaleniya).Delete
            End If
        Next
        
    End With

    Call ZagruzkaVListBox(tbRaschot)

End Sub

Public Sub PeremestitStrokuVverkh(tbRaschot As ListObject)

    If pListBox.ListIndex = -1 Then Exit Sub

    Dim i As Integer
    Dim Stroka As Integer, StrokaPoslePeremeshcheniya As Integer

    For i = pListBox.ListCount - 1 To 0 Step -1
        If pListBox.Selected(i) Then
            Stroka = i + 1
            StrokaPoslePeremeshcheniya = i - 1
            Exit For
        End If
    Next

    If Stroka > 1 Then

        tbRaschot.ListRows(Stroka).Range.Cut
        tbRaschot.ListRows(Stroka - 1).Range.Insert Shift:=xlDown

        Call ZagruzkaVListBox(tbRaschot)

        pListBox.Selected(StrokaPoslePeremeshcheniya) = True
    End If

End Sub

Public Sub PeremestitStrokuVniz(tbRaschot As ListObject)

    If pListBox.ListIndex = -1 Then Exit Sub

    Dim i As Integer
    Dim Stroka As Integer, StrokaPoslePeremeshcheniya As Integer

    For i = pListBox.ListCount - 2 To 0 Step -1
        If pListBox.Selected(i) Then
            Stroka = i + 1
            StrokaPoslePeremeshcheniya = i + 1
            Exit For
        End If
    Next

    If Stroka > 0 And Stroka - 1 < pListBox.ListCount Then

        tbRaschot.ListRows(Stroka).Range.Cut

        Dim DiapazonPeremeshcheniya As Range
        If Stroka + 2 > tbRaschot.ListRows.Count Then
            Set DiapazonPeremeshcheniya = tbRaschot.ListRows(tbRaschot.ListRows.Count).Range.Offset(1)
        Else
            Set DiapazonPeremeshcheniya = tbRaschot.ListRows(Stroka + 2).Range
        End If
        DiapazonPeremeshcheniya.Insert Shift:=xlDown

        Call ZagruzkaVListBox(tbRaschot)

        pListBox.Selected(StrokaPoslePeremeshcheniya) = True

    End If

End Sub

Public Sub ZagruzkaVListBox(tbRaschot As ListObject)

    Set pTbRaschot = tbRaschot

    Dim StrokaListBox As Integer, StrokaTablitsa As Integer
    With pListBox
        .Clear
        For StrokaTablitsa = 1 To tbRaschot.ListRows.Count
            .AddItem
            StrokaListBox = .ListCount - 1
            .List(StrokaListBox, 0) = StrokaListBox + 1
            .List(StrokaListBox, 1) = tbRaschot.DataBodyRange(StrokaTablitsa, NomerStrolbtsa("Perekhod"))
            .List(StrokaListBox, 2) = VzyatVremyaIzTablitsty(StrokaTablitsa, "tTokFrez")
            .List(StrokaListBox, 3) = VzyatVremyaIzTablitsty(StrokaTablitsa, "tSles")
            .List(StrokaListBox, 4) = VzyatVremyaIzTablitsty(StrokaTablitsa, "tKontrol")
        Next
    End With
    
    Call ZapisatTsht

End Sub

Public Sub ZapisatTsht()

    ptxtTsht_TokFrez.text = 0
    ptxtTsht_Slesar.text = 0
    ptxtTsht_Kontrol.text = 0
    
    Dim Normy
    Normy = RaschotThst
    
    Dim Tsht_TokFrez As Double, Tsht_Slesar As Double, Tsht_Kontrol As Double
    Tsht_TokFrez = Normy(1)
    Tsht_Slesar = Normy(2)
    Tsht_Kontrol = Normy(3)

    If pListBox.ListCount > 0 Then

        If pChkOkruglennyyeNomvy Then
            ptxtTsht_TokFrez.text = CStr(OkruglenieTsht(Tsht_TokFrez))
            ptxtTsht_Slesar.text = CStr(OkruglenieTsht(Tsht_Slesar))
            ptxtTsht_Kontrol.text = CStr(OkruglenieTsht(Tsht_Kontrol))
        Else
            ptxtTsht_TokFrez.text = FormatVremeni(Tsht_TokFrez, 1)
            ptxtTsht_Slesar.text = FormatVremeni(Tsht_Slesar, 1)
            ptxtTsht_Kontrol.text = FormatVremeni(Tsht_Kontrol, 1)
        End If

    End If


End Sub

Private Function RaschotThst() As Variant
    
    Dim Normy(1 To 3) As Double
    
    If pListBox.ListCount > 0 Then

        Dim i As Integer
        For i = 0 To pListBox.ListCount - 1
            Normy(1) = Normy(1) + CDbl(pListBox.List(i, 2))
            Normy(2) = Normy(2) + CDbl(pListBox.List(i, 3))
            Normy(3) = Normy(3) + CDbl(pListBox.List(i, 4))
        Next

    End If
    
    RaschotThst = Normy

End Function

Private Function VzyatVremyaIzTablitsty(Stroka As Integer, NaimenovaniyeStolbtsa As String) As String
    VzyatVremyaIzTablitsty = FormatVremeni(pTbRaschot.DataBodyRange(Stroka, pTbRaschot.ListColumns(NaimenovaniyeStolbtsa).Index))
End Function

Private Function FormatVremeni(Vremya As Double, Optional ChisloZnakov As Integer = 2) As String
    FormatVremeni = CStr(Round(Vremya, ChisloZnakov))
End Function

Public Function ZapolnitParametryIzStrokiTablitsy(Stroka As Integer, pTbRaschot As ListObject) As clsTokFrez

    Dim Parametry As New clsTokFrez
    
    Dim rng As Range
    Set rng = pTbRaschot.ListRows(Stroka).Range
    
    On Error Resume Next
    With Parametry
        .id = rng(NomerStrolbtsa("id"))
        .OpisaniyePerekhoda = rng(NomerStrolbtsa("Perekhod"))
        .L = rng(NomerStrolbtsa("L"))
        .Lrezby = rng(NomerStrolbtsa("Lrezby"))
        .ShirinaPoverhnosti = rng(NomerStrolbtsa("ShirinaPoverhnosti"))
        .GlubinaPaza = rng(NomerStrolbtsa("GlubinaPaza"))
        .DlinaObnizki = rng(NomerStrolbtsa("DlinaObnizki"))
        .DBObnizki = rng(NomerStrolbtsa("DBObnizki"))
        .GlubinaObnizki = rng(NomerStrolbtsa("GlubinaObnizki"))
        .Pripusk = rng(NomerStrolbtsa("Pripusk"))
        .KalibrRezbSlesar = rng(NomerStrolbtsa("KalibrRezbSlesar"))
        .IT = rng(NomerStrolbtsa("IT"))
        .Ra = rng(NomerStrolbtsa("Ra"))
        .KolVoProkhodov = rng(NomerStrolbtsa("KolVoProkhodov"))
        .KolVoPoverkhnostey = rng(NomerStrolbtsa("KolVoPoverkhnostey"))
        .D1 = rng(NomerStrolbtsa("D1"))
        .Udar = rng(NomerStrolbtsa("Udar"))
        .Korka = rng(NomerStrolbtsa("Korka"))
        .GlubinaRezaniya = rng(NomerStrolbtsa("GlubinaRezaniya"))
        .ShagRezby = rng(NomerStrolbtsa("ShagRezby"))
        .d2 = rng(NomerStrolbtsa("d2"))
        .Otkloneniye = rng(NomerStrolbtsa("Otkloneniye"))
        .KolVoDetVPrutke = rng(NomerStrolbtsa("KolVoDetVPrutke"))
        .VUpor = rng(NomerStrolbtsa("Vupor"))
        .PodrezkaTortsa = rng(NomerStrolbtsa("PodrezkaTortsa"))
        .TocheniyeDiametra = rng(NomerStrolbtsa("TocheniyeDiametra"))
        .Vconst = rng(NomerStrolbtsa("Vconst"))
        .TipPoverkhnosti = rng(NomerStrolbtsa("TipPoverkhnosti"))
        .TipTsentrovaniya = rng(NomerStrolbtsa("TipTsentrovaniya"))
        .Massa = rng(NomerStrolbtsa("Massa"))
        .Sposob = rng(NomerStrolbtsa("Sposob"))
        .KharkterVyverki = rng(NomerStrolbtsa("KharkterVyverki"))
        .TochnostVyverki = rng(NomerStrolbtsa("TochnostVyverki"))
        .ZachistkaBokovykhStoron = rng(NomerStrolbtsa("ZachistkaBokovykhStoron"))
        .TipStanka = rng(NomerStrolbtsa("TipStanka"))
        .FrezPosleSverl = rng(NomerStrolbtsa("FrezPosleSverl"))
        .UchestFasku = rng(NomerStrolbtsa("UchestFasku"))
        .UchestSlesar = rng(NomerStrolbtsa("tSles*"))
        .UchestKontrol = rng(NomerStrolbtsa("tKontrol*"))
    End With
    On Error GoTo 0

    Set ZapolnitParametryIzStrokiTablitsy = Parametry

End Function

Public Function RaschotKolvaInstrumentov(tbRaschot As ListObject) As Integer

    If tbRaschot.ListColumns(NomerStrolbtsa("id")).DataBodyRange Is Nothing Then Exit Function

    Dim Perekhody As New Collection
    Set Perekhody = SozdatKollektsiyuPerekhodov(tbRaschot.ListColumns(NomerStrolbtsa("id")).DataBodyRange, tbRaschot.ListColumns(NomerStrolbtsa("D1")).DataBodyRange)

    Dim FrezPosleSverl As Boolean
    FrezPosleSverl = WorksheetFunction.CountIf(tbRaschot.ListColumns(NomerStrolbtsa("FrezPosleSverl")).DataBodyRange, True)

    Dim DobavitFrezuDluaOtverstiy As Boolean
    DobavitFrezuDluaOtverstiy = False
    
    Dim item
    Dim KolVoInstrumentov As Integer
    For Each item In Perekhody
        
        Select Case item
            Case "pUstRastFrez", "pUstTok_Samotsentr", "pUstTok_ChetyrekhTsilindr", "pUst"
            
            Case "pFrezOkon", "pObrabotkaRezbovihOtv", "pFrezPloskKonc", "pFrezUstup", "pObrabOtvSNulya"
                KolVoInstrumentov = KolVoInstrumentov + 2
                
            Case Else
                KolVoInstrumentov = KolVoInstrumentov + 1
                
        End Select
        
        If item = "pFrezOtverstii" And FrezPosleSverl = True Then
            DobavitFrezuDluaOtverstiy = True
        End If
        
    Next
    
    If DobavitFrezuDluaOtverstiy = True Then
        KolVoInstrumentov = KolVoInstrumentov + 1
    End If
    
    RaschotKolvaInstrumentov = KolVoInstrumentov
    
End Function

Private Function SozdatKollektsiyuPerekhodov(IdPerekhodov As Range, Diametry As Range) As Collection
    
    Dim Perekhody As New Collection
    Dim Key As String

    Dim i As Integer
    For i = 1 To IdPerekhodov.Rows.Count

        Select Case IdPerekhodov(i)
            Case "pPoperechnoye", "pProdolnoye", "pRastachivaniye"
                Key = IdPerekhodov(i)
            Case Else
                Key = IdPerekhodov(i) & Diametry(i)
        End Select
        
        On Error Resume Next
        Perekhody.Add IdPerekhodov(i), Key
        On Error GoTo 0
        
    Next
    
    Set SozdatKollektsiyuPerekhodov = Perekhody
    
End Function

Public Function RaschotUstanovki(tbRaschot As ListObject) As Double

    Dim i As Integer
    For i = 1 To tbRaschot.ListRows.Count
        If tbRaschot.ListColumns(NomerStrolbtsa("Perekhod")).DataBodyRange(i) Like "*становить*" Then
            RaschotUstanovki = RaschotUstanovki + tbRaschot.DataBodyRange(i, (NomerStrolbtsa("tTokFrez")))
        End If
    Next
 
End Function

Public Sub ProtyanutFormuly(tbRaschot As ListObject)

    With tbRaschot
        .DataBodyRange.Orientation = 0
        .ListColumns(.ListColumns("tTokFrez").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tTokFrez").Index).Formula
        .ListColumns(.ListColumns("tSles").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tSles").Index).Formula
        .ListColumns(.ListColumns("tKontrol").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tKontrol").Index).Formula
    End With

End Sub

Public Sub ZapolnitTablitsuParametrami(tbRaschot As ListObject, StrokaParametrov As Range, Perekhod As String, Parametry As clsTokFrez)

    Set pTbRaschot = tbRaschot

    Application.Calculation = xlCalculationManual
    
    On Error Resume Next

    With Parametry
        StrokaParametrov(NomerStrolbtsa("id")) = .id
        StrokaParametrov(NomerStrolbtsa("Perekhod")) = Perekhod
        StrokaParametrov(NomerStrolbtsa("L")) = .L
        StrokaParametrov(NomerStrolbtsa("Lrezby")) = .Lrezby
        StrokaParametrov(NomerStrolbtsa("ShirinaPoverhnosti")) = .ShirinaPoverhnosti
        StrokaParametrov(NomerStrolbtsa("GlubinaPaza")) = .GlubinaPaza
        StrokaParametrov(NomerStrolbtsa("DlinaObnizki")) = .DlinaObnizki
        StrokaParametrov(NomerStrolbtsa("DBObnizki")) = .DBObnizki
        StrokaParametrov(NomerStrolbtsa("GlubinaObnizki")) = .GlubinaObnizki
        StrokaParametrov(NomerStrolbtsa("Pripusk")) = .Pripusk
        StrokaParametrov(NomerStrolbtsa("KalibrRezbSlesar")) = .KalibrRezbSlesar
        StrokaParametrov(NomerStrolbtsa("IT")) = .IT
        StrokaParametrov(NomerStrolbtsa("Ra")) = .Ra
        StrokaParametrov(NomerStrolbtsa("KolVoProkhodov")) = .KolVoProkhodov
        StrokaParametrov(NomerStrolbtsa("KolVoPoverkhnostey")) = .KolVoPoverkhnostey
        StrokaParametrov(NomerStrolbtsa("D1")) = .D1
        StrokaParametrov(NomerStrolbtsa("Udar")) = .Udar
        StrokaParametrov(NomerStrolbtsa("Korka")) = .Korka
        StrokaParametrov(NomerStrolbtsa("GlubinaRezaniya")) = .GlubinaRezaniya
        StrokaParametrov(NomerStrolbtsa("ShagRezby")) = .ShagRezby
        StrokaParametrov(NomerStrolbtsa("d2")) = .d2
        StrokaParametrov(NomerStrolbtsa("Otkloneniye")) = .Otkloneniye
        StrokaParametrov(NomerStrolbtsa("KolVoDetVPrutke")) = .KolVoDetVPrutke
        StrokaParametrov(NomerStrolbtsa("Vupor")) = .VUpor
        StrokaParametrov(NomerStrolbtsa("PodrezkaTortsa")) = .PodrezkaTortsa
        StrokaParametrov(NomerStrolbtsa("TocheniyeDiametra")) = .TocheniyeDiametra
        StrokaParametrov(NomerStrolbtsa("Vconst")) = .Vconst
        StrokaParametrov(NomerStrolbtsa("TipPoverkhnosti")) = .TipPoverkhnosti
        StrokaParametrov(NomerStrolbtsa("TipTsentrovaniya")) = .TipTsentrovaniya
        StrokaParametrov(NomerStrolbtsa("Massa")) = .Massa
        StrokaParametrov(NomerStrolbtsa("Sposob")) = .Sposob
        StrokaParametrov(NomerStrolbtsa("KharkterVyverki")) = .KharkterVyverki
        StrokaParametrov(NomerStrolbtsa("TochnostVyverki")) = .TochnostVyverki
        StrokaParametrov(NomerStrolbtsa("ZachistkaBokovykhStoron")) = .ZachistkaBokovykhStoron
        StrokaParametrov(NomerStrolbtsa("TipStanka")) = .TipStanka
        StrokaParametrov(NomerStrolbtsa("FrezPosleSverl")) = .FrezPosleSverl
        StrokaParametrov(NomerStrolbtsa("UchestFasku")) = .UchestFasku
        StrokaParametrov(NomerStrolbtsa("tSles*")) = .UchestSlesar
        StrokaParametrov(NomerStrolbtsa("tKontrol*")) = .UchestKontrol
    End With
    
    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic

End Sub

Private Function NomerStrolbtsa(NaimenovaniyeStolbtsa As String) As Integer
    NomerStrolbtsa = pTbRaschot.ListColumns(NaimenovaniyeStolbtsa).Index
End Function
