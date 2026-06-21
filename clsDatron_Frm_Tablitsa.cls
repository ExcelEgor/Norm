VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatron_Frm_Tablitsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private wbRaschot As Workbook
Private wsRaschot As Worksheet
Private tbRaschot As ListObject

Private lstRaschot As MSForms.ListBox
Private txtTsht_Datron As MSForms.TextBox
Private txtTsht_Slesar As MSForms.TextBox
Private txtTsht_Kontrol As MSForms.TextBox
Private cboMaterial As MSForms.ComboBox
Private txtPerekhod As MSForms.TextBox
Private txt_tDatron As MSForms.TextBox
Private tab_tDatron As MSForms.TabStrip

Dim Redaktirovanie As New clsDatron_Frm_Redaktirovaniye

Private Sub Class_Initialize()

    With frmDatron
        Set lstRaschot = .lstRaschot
        Set txtTsht_Datron = .txtTsht_Datron
        Set txtTsht_Slesar = .txtTsht_Slesar
        Set txtTsht_Kontrol = .txtTsht_Kontrol
        Set cboMaterial = .cboMaterial
        Set txtPerekhod = .txtPerekhod
        Set txt_tDatron = .txt_tDatron
        Set tab_tDatron = .tabOperatsii
    End With
    
    Call NovyyRaschot
  
End Sub

Public Sub NovyyRaschot()

    lstRaschot.Clear
    
    If Not wbRaschot Is Nothing Then wbRaschot.Close (False)
    wsTablitsa_Datron.Copy
    Set wbRaschot = ActiveWorkbook
    Set wsRaschot = wbRaschot.ActiveSheet
    Set tbRaschot = wsRaschot.ListObjects(1)
    wbRaschot.Windows(1).Visible = False
    
    wsRaschot.Name = tab_tDatron(0).Caption
    
    Call ZapistatMaterial
    
    
End Sub

Public Sub NovayaOperatsiya()

    lstRaschot.Clear
    
    Dim NomerOperatsii As Integer
    With tab_tDatron
        .tabs.Add
        .Value = tab_tDatron.tabs.Count - 1
        NomerOperatsii = .tabs.Count * 5
        .SelectedItem.Caption = Format(NomerOperatsii, "000")
    End With
    
    DobavitList NomerOperatsii
    
    PerekluchitOperatsiyu
    
    ZapistatMaterial
    
End Sub

Private Sub DobavitList(NomerOperatsii As Integer)

    wbRaschot.Windows(1).Visible = True
    wsTablitsa_Datron.Copy after:=wbRaschot.Worksheets(wbRaschot.Worksheets.Count)
    wbRaschot.Windows(1).Visible = False
    
    wbRaschot.Worksheets(wbRaschot.Worksheets.Count).Name = Format(NomerOperatsii, "000")
    
End Sub

Public Sub PerekluchitOperatsiyu()

    Dim ws As Worksheet
    For Each ws In wbRaschot.Worksheets
        If ws.Name = tab_tDatron.SelectedItem.Caption Then
        
            Set wsRaschot = ws
            Set tbRaschot = wsRaschot.ListObjects(1)
            ZagruzkaIzTablitsy_v_ListBox
            
            Exit For
        End If
    Next
    
End Sub

Public Sub ZapistatMaterial()
    
    With cboMaterial
        If .ListIndex <> -1 Then
            wsRaschot.Range("B1") = CInt(.List(.ListIndex, 1))
            Call ZagruzkaIzTablitsy_v_ListBox
        End If
    End With

End Sub

Public Sub DobavitZapis(ParametryDatron As clsDatron, RezhimRedaktirovaniya As Boolean)

    If DblFromCtrl(txt_tDatron) <= 0 Then Exit Sub
    
    Dim Parametry As Range
    If RezhimRedaktirovaniya Then
        Set Parametry = tbRaschot.ListRows(lstRaschot.ListIndex + 1).Range
    Else
        tbRaschot.ListRows.Add
        Set Parametry = tbRaschot.ListRows(tbRaschot.ListRows.Count).Range
    End If
    
    With ParametryDatron
        Parametry(tbRaschot.ListColumns("id").Index) = .id
        Parametry(tbRaschot.ListColumns("OpisaniyePerekhoda").Index) = frmDatron.txtPerekhod
        Parametry(tbRaschot.ListColumns("Dlina").Index) = .Dlina
        Parametry(tbRaschot.ListColumns("DlinaKrivoy").Index) = .DlinaKrivoy
        Parametry(tbRaschot.ListColumns("Shirina").Index) = .Shirina
        Parametry(tbRaschot.ListColumns("Diametr").Index) = .Diametr
        Parametry(tbRaschot.ListColumns("dObnizheniya").Index) = .Dobnizheniya
        Parametry(tbRaschot.ListColumns("lObnizheniya").Index) = .Lobnizheniya
        Parametry(tbRaschot.ListColumns("Glubina").Index) = .Glubina
        Parametry(tbRaschot.ListColumns("GlubinaRezby").Index) = .GlubinaRezby
        Parametry(tbRaschot.ListColumns("Pripusk").Index) = .Pripusk
        Parametry(tbRaschot.ListColumns("Glukhoye").Index) = .Glukhoye
        Parametry(tbRaschot.ListColumns("VysotaFaski").Index) = .VysotaFaski
        Parametry(tbRaschot.ListColumns("BezZachistki").Index) = .BezZachistki
        Parametry(tbRaschot.ListColumns("IT").Index) = .IT
        Parametry(tbRaschot.ListColumns("Ra").Index) = .Ra
        Parametry(tbRaschot.ListColumns("KolVoPoverkhnostey").Index) = .KolVoPoverkhnostey
        Parametry(tbRaschot.ListColumns("Massa").Index) = .Massa
        Parametry(tbRaschot.ListColumns("SposobUstanovki").Index) = .SposobUstanovki
        Parametry(tbRaschot.ListColumns("tSlesar*").Index) = .UchestSlesar
        Parametry(tbRaschot.ListColumns("tKontrol*").Index) = .UchestKontrol
    End With
       
    Call UdalitSsylkuNaKnigu
    
    Call ProtyanutFormuly
    
    Call ZagruzkaIzTablitsy_v_ListBox
    
    frmDatron.RezhimRedaktirovaniya = False
    lstRaschot.Enabled = True
     
End Sub

Private Sub UdalitSsylkuNaKnigu()

    Application.DisplayAlerts = False
    wsRaschot.Cells.Replace What:=wbRaschot.Name & "!", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
    Application.DisplayAlerts = True
    
End Sub

Private Sub ProtyanutFormuly()
    With tbRaschot
        .ListColumns(.ListColumns("tDatron").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tDatron").Index).Formula
        .ListColumns(.ListColumns("tSlesar").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tSlesar").Index).Formula
        .ListColumns(.ListColumns("tKontrol").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("tKontrol").Index).Formula
    End With
End Sub

Private Sub ZagruzkaIzTablitsy_v_ListBox()

    lstRaschot.Clear
    txtTsht_Datron.text = Empty
    txtTsht_Slesar.text = Empty
    txtTsht_Kontrol.text = Empty
    
    If Not tbRaschot.ListRows.Count > 0 Then Exit Sub

    Dim i As Integer
    With lstRaschot
        For i = 1 To tbRaschot.ListRows.Count
            .AddItem
            .List(.ListCount - 1, 0) = i
            .List(.ListCount - 1, 1) = tbRaschot.DataBodyRange(i, PoluchitNomerStolbtsa(tbRaschot, "OpisaniyePerekhoda"))
            .List(.ListCount - 1, 2) = CStr(tbRaschot.DataBodyRange(i, PoluchitNomerStolbtsa(tbRaschot, "tDatron")))
            .List(.ListCount - 1, 3) = CStr(tbRaschot.DataBodyRange(i, PoluchitNomerStolbtsa(tbRaschot, "tSlesar")))
            .List(.ListCount - 1, 4) = CStr(tbRaschot.DataBodyRange(i, PoluchitNomerStolbtsa(tbRaschot, "tKontrol")))
        Next
    End With
    
    txtTsht_Datron.text = RaschotSummy("tDatron")
    txtTsht_Slesar.text = RaschotSummy("tSlesar")
    txtTsht_Kontrol.text = RaschotSummy("tKontrol")

End Sub

Public Sub UdalitStroku()

    Call UdaleniyeStrokIzLisboksa_i_Tablitsy(lstRaschot, tbRaschot)
    
    Call ZagruzkaIzTablitsy_v_ListBox
    
End Sub

Private Function RaschotSummy(NaimenovaniyeStolbtsa As String) As String

    RaschotSummy = OkruglenieTsht(CStr(WorksheetFunction.Sum(tbRaschot.ListColumns(PoluchitNomerStolbtsa(tbRaschot, NaimenovaniyeStolbtsa)).DataBodyRange)))
    
End Function

Public Sub ZakrytRaschot()
    On Error Resume Next
    wbRaschot.Close (False)
    On Error GoTo 0
End Sub

Public Sub RedaktirovatPerekhod()

    Call Redaktirovanie.RedaktirovatPerekhod(wsRaschot, tbRaschot)
    
End Sub

