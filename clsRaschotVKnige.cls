VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRaschotVKnige"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pKniga As clsKniga
Private Tablitsa As clsTokFrez_Tablitsa
Private pWbRaschot As Workbook

Private pWsRaschot As Worksheet
Private pListBox As MSForms.ListBox
Private pTbRaschot As ListObject

Private ptxtTsht_TokFrez As MSForms.TextBox
Private ptxtTsht_Slesar As MSForms.TextBox
Private ptxtTsht_Kontrol As MSForms.TextBox

Private pChkOkruglennyyeNomvy As MSForms.CheckBox
Private pCboMaterial As MSForms.ComboBox
Private pCboTipProizv As MSForms.ComboBox
Private pChkPoRezhimam As MSForms.CheckBox
Private pChkTonkosten As MSForms.CheckBox
Private pChkCHPU As MSForms.CheckBox
Private pTabOperatsii As MSForms.TabStrip

Public Property Set ListBox(ListBox As MSForms.ListBox): Set pListBox = ListBox: End Property

Public Property Set txtTsht_TokFrez(txtTsht_TokFrez As MSForms.TextBox): Set ptxtTsht_TokFrez = txtTsht_TokFrez: End Property

Public Property Set txtTsht_Slesar(txtTsht_Slesar As MSForms.TextBox): Set ptxtTsht_Slesar = txtTsht_Slesar: End Property

Public Property Set txtTsht_Kontrol(txtTsht_Kontrol As MSForms.TextBox): Set ptxtTsht_Kontrol = txtTsht_Kontrol: End Property

Public Property Set ChkOkruglennyyeNomvy(ChkOkruglennyyeNomvy As MSForms.CheckBox): Set pChkOkruglennyyeNomvy = ChkOkruglennyyeNomvy: End Property

Public Property Set cboMaterial(cboMaterial As MSForms.ComboBox): Set pCboMaterial = cboMaterial: End Property

Public Property Set cboTipProizv(cboTipProizv As MSForms.ComboBox): Set pCboTipProizv = cboTipProizv: End Property

Public Property Set chkPoRezhimam(chkPoRezhimam As MSForms.CheckBox): Set pChkPoRezhimam = chkPoRezhimam: End Property

Public Property Set chkCHPU(chkCHPU As MSForms.CheckBox): Set pChkCHPU = chkCHPU: End Property

Public Property Set chkTonkosten(chkTonkosten As MSForms.CheckBox): Set pChkTonkosten = chkTonkosten: End Property

Public Property Set tabOperatsii(tabOperatsii As MSForms.TabStrip): Set pTabOperatsii = tabOperatsii: End Property

Public Property Get wbRaschot() As Workbook: Set wbRaschot = pWbRaschot: End Property

Public Property Get wsRaschot() As Worksheet: Set wsRaschot = pWsRaschot: End Property

Public Property Get tbRaschot() As ListObject: Set tbRaschot = pTbRaschot: End Property

Public Sub SozdatNovyyRaschot(Kniga As clsKniga)
    
    Set pKniga = Kniga
    Call pKniga.NovyyRaschot
    Set pWbRaschot = pKniga.wbRaschot

    Application.ScreenUpdating = False

    Set pWsRaschot = pWbRaschot.Worksheets(1)
    pWsRaschot.Name = "005"

    Set pTbRaschot = pWsRaschot.ListObjects(1)
    
    Set Tablitsa = New clsTokFrez_Tablitsa
    Call Tablitsa.Init(pListBox, ptxtTsht_TokFrez, ptxtTsht_Slesar, ptxtTsht_Kontrol, pChkOkruglennyyeNomvy)

    Application.ScreenUpdating = True

    Call ZapisatParametryOperatsiiVListRaschota


End Sub

Public Sub OtkryRaschot(tbShablon As ListObject, Optional Raschot As Workbook)
    
    Dim i As Integer
    For i = pTabOperatsii.tabs.Count - 1 To 1 Step -1
        pTabOperatsii.tabs.Remove (i)
    Next
    
    If Not Raschot Is Nothing Then
        Set pWbRaschot = Raschot
    Else
        pKniga.OtkrytRaschot
        Set pWbRaschot = pKniga.wbRaschot
    End If
    
    For i = 1 To pWbRaschot.Worksheets.Count
        If i > 1 Then pTabOperatsii.tabs.Add
        pTabOperatsii.tabs(i - 1).Caption = pWbRaschot.Worksheets(i).Name
    Next
    
    Call ProveritStolbtsy(tbShablon)
    
    Call ZagruzkaVListBox
    
End Sub

Private Sub ProveritStolbtsy(tbShablon As ListObject)
    
    Dim ws As Worksheet
    Dim tbRaschot As ListObject
    
    For Each ws In pWbRaschot.Worksheets
        If ws.ListObjects.Count = 1 Then
            Set tbRaschot = ws.ListObjects(1)
            Call UdalitLishniyeStolbsty(tbShablon, tbRaschot)
            Call DobavitNedostayushchiyeStolbsty(tbShablon, tbRaschot)
        End If
    Next

End Sub

Private Sub UdalitLishniyeStolbsty(tbShablon As ListObject, tbRaschot As ListObject)
    
    Dim i As Integer
    For i = tbRaschot.ListColumns.Count To 1 Step -1
        If WorksheetFunction.CountIf(tbShablon.HeaderRowRange, tbRaschot.ListColumns(i).Name) = 0 Then
            tbRaschot.ListColumns(i).Delete
        End If
    Next

End Sub

Private Sub DobavitNedostayushchiyeStolbsty(tbShablon As ListObject, tbRaschot As ListObject)
    
    Dim i As Integer
    For i = 1 To tbShablon.ListColumns.Count
        If WorksheetFunction.CountIf(tbRaschot.HeaderRowRange, tbShablon.ListColumns(i).Name) = 0 Then
            tbRaschot.ListColumns.Add (i)
            tbRaschot.ListColumns(i).Name = tbShablon.ListColumns(i).Name
        End If
    Next

End Sub

Public Sub DobavitOperatsiyu()

    pKniga.KopirovatList pTabOperatsii.Value + 1

    Set pWsRaschot = pWbRaschot.Worksheets(pWbRaschot.Worksheets.Count)
    Set pTbRaschot = pWsRaschot.ListObjects(1)
    
    If pTbRaschot.ListRows.Count > 0 Then
        pTbRaschot.DataBodyRange.Rows.EntireRow.Delete
    End If

    Call DobavitTabVkladkuTab

    Application.ScreenUpdating = True

End Sub

Private Sub DobavitTabVkladkuTab()

    Dim NomerOperatsii As String

    With pTabOperatsii
        .tabs.Add
        .Value = .tabs.Count - 1
    End With

    pListBox.Clear

    NomerOperatsii = OpredelitNomerNovoyOperatsii

    pTabOperatsii.SelectedItem.Caption = NomerOperatsii
    pWsRaschot.Name = NomerOperatsii

End Sub

Private Function OpredelitNomerNovoyOperatsii() As String

    Dim i As Integer, Nomer As Double, NovyyNomer As Integer

    With pTabOperatsii
        For i = 0 To .tabs.Count - 1
            If IsNumeric(.tabs(i).Caption) Then
                Nomer = CDbl(.tabs(i).Caption)
                If Nomer > NovyyNomer Then
                    NovyyNomer = Nomer
                End If
            End If
        Next
    End With

    OpredelitNomerNovoyOperatsii = Format(5 + NovyyNomer, "000")

End Function

Public Sub UdalitOperatsiyu()

    If Not pTabOperatsii.tabs.Count > 1 Then Exit Sub
    
    pKniga.UdalitList (pTabOperatsii.Value + 1)
    
    pTabOperatsii.tabs.Remove (pTabOperatsii.Value)
    
    Call ZagruzkaVListBox

End Sub

Public Sub KopirovatStroki()

    If pListBox.ListIndex = -1 Then Exit Sub

    Dim i As Integer
    Dim Stroka As Integer

    For i = 0 To pListBox.ListCount - 1
        If pListBox.Selected(i) Then
            Stroka = i + 1
            With pTbRaschot
                .ListRows.Add
                .ListRows(Stroka).Range.Copy
                .ListRows(.ListRows.Count).Range.PasteSpecial (xlPasteAll)
            End With
        End If
    Next

    Call ZagruzkaVListBox

End Sub

Public Sub UdalitStroki()
    Call Tablitsa.UdalitStroki(pTbRaschot)
End Sub

Public Sub PeremestitStrokuVverkh()
    Call Tablitsa.PeremestitStrokuVverkh(pTbRaschot)
End Sub

Public Sub PeremestitStrokuVniz()
    Call Tablitsa.PeremestitStrokuVniz(pTbRaschot)
End Sub

Public Sub ZapisatParametryOperatsiiVListRaschota()

    If pWsRaschot Is Nothing Then Exit Sub

    Application.Calculation = xlCalculationManual

    With pWsRaschot
        .Range("F2").Value = MaterialFromCboMaterial(pCboMaterial)
        .Range("F3").Value = pCboTipProizv.ListIndex + 1
        .Range("F4").Value = pChkPoRezhimam.Value
        .Range("F5").Value = pChkTonkosten.Value
        .Range("F6").Value = pChkCHPU.Value
    End With

    Application.Calculation = xlCalculationAutomatic

    Call ZagruzkaVListBox

End Sub

Public Sub KopirovatOperatsiyu()

    pKniga.KopirovatList pTabOperatsii.Value + 1
    Set pWsRaschot = pWbRaschot.Worksheets(pWbRaschot.Worksheets.Count)

    Set pTbRaschot = pWsRaschot.ListObjects(1)

    With pTabOperatsii
        .tabs.Add
        .Value = .tabs.Count - 1
        .SelectedItem.Caption = Format(5 * .tabs.Count, "000")
    End With

End Sub

Sub PereklyucheniyeOperatsiy()

    Set pWsRaschot = pWbRaschot.Worksheets(pTabOperatsii.Value + 1)
    Set pTbRaschot = pWsRaschot.ListObjects(1)

    pCboTipProizv.ListIndex = pWsRaschot.Range("F3").Value - 1
    Select Case pWsRaschot.Range("F4").Value
        Case 0 = pTabOperatsii.Value = 1
        Case 1 = pTabOperatsii.Value = 2
        Case 2 = pTabOperatsii.Value = 1
    End Select

    pChkTonkosten.Value = pWsRaschot.Range("F5").Value
    pCboMaterial.ListIndex = pWsRaschot.Range("F2").Value - 1

    Call ZagruzkaVListBox

End Sub


Public Sub ZapisatTsht()
    Tablitsa.ZapisatTsht
End Sub

Public Sub DobavitVTablitsu(Perekhod As String, Parametry As clsTokFrez, Tsht_TokFrez As Double, RezhimRedaktirovaniya As Boolean)

    If Tsht_TokFrez <= 0 Then Exit Sub

    Set pWsRaschot = pWbRaschot.Worksheets(pTabOperatsii.Value + 1)
    Set pTbRaschot = pWsRaschot.ListObjects(1)

    Dim NomerStroki As Long
    NomerStroki = OpredelitNomerStroki(RezhimRedaktirovaniya)

    Call Tablitsa.ProtyanutFormuly(pTbRaschot)

    Dim DiapazonParametrov As Range
    Set DiapazonParametrov = pTbRaschot.ListRows(NomerStroki).Range

    Call Tablitsa.ZapolnitTablitsuParametrami(pTbRaschot, DiapazonParametrov, Perekhod, Parametry)

    Call ZagruzkaVListBox

End Sub

Private Function OpredelitNomerStroki(RezhimRedaktirovaniya As Boolean)

    Dim NomerStroki As Integer

    If RezhimRedaktirovaniya Then
        NomerStroki = pListBox.ListIndex + 1
    Else
        With pTbRaschot
            .ListRows.Add
            NomerStroki = .ListRows.Count
        End With
    End If

    OpredelitNomerStroki = NomerStroki

End Function

Public Function ZapolnitParametryIzStrokiTablitsy(Stroka As Integer) As clsTokFrez

    Set ZapolnitParametryIzStrokiTablitsy = Tablitsa.ZapolnitParametryIzStrokiTablitsy(Stroka, pTbRaschot)
    
End Function

Public Sub RaschotTpz()

    Call frmRaschotTpz.ZagruzitParametryVFormu(Tablitsa.RaschotKolvaInstrumentov(pTbRaschot), ptxtTsht_TokFrez.text)

End Sub

Public Sub ZagruzkaVListBox()

    Set pWsRaschot = pWbRaschot.Worksheets(frmTokFrez.tabOperatsii.Value + 1)

    Application.DisplayAlerts = False
    pWsRaschot.Cells.Replace What:=pWbRaschot.Name & "!", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
    Application.DisplayAlerts = True

    Set pTbRaschot = pWsRaschot.ListObjects(1)

    Call Tablitsa.ZagruzkaVListBox(pTbRaschot)

End Sub
