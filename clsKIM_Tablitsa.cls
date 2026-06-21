VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKIM_Tablitsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tbKIM As ListObject

Private txtTnsh As MSForms.TextBox
Private lstTablitsa As MSForms.ListBox
Attribute lstTablitsa.VB_VarHelpID = -1
Private txtTsht As MSForms.TextBox
Attribute txtTsht.VB_VarHelpID = -1
Private txtTpz As MSForms.TextBox

Private WithEvents optSloznost_Hera_1 As MSForms.OptionButton
Attribute optSloznost_Hera_1.VB_VarHelpID = -1
Private WithEvents optSloznost_Hera_2 As MSForms.OptionButton
Attribute optSloznost_Hera_2.VB_VarHelpID = -1
Private WithEvents optSloznost_Hera_3 As MSForms.OptionButton
Attribute optSloznost_Hera_3.VB_VarHelpID = -1
Private WithEvents txtKalibrovka_15 As MSForms.TextBox
Attribute txtKalibrovka_15.VB_VarHelpID = -1
Private WithEvents txtKalibrovka_40 As MSForms.TextBox
Attribute txtKalibrovka_40.VB_VarHelpID = -1
Private WithEvents txtKalibrovka_99 As MSForms.TextBox
Attribute txtKalibrovka_99.VB_VarHelpID = -1

Private Sub optSloznost_Hera_1_Click(): ZapisatParametryOperatsii: End Sub

Private Sub optSloznost_Hera_2_Click(): ZapisatParametryOperatsii: End Sub

Private Sub optSloznost_Hera_3_Click(): ZapisatParametryOperatsii: End Sub

Private Sub txtKalibrovka_15_Change(): ZapisatParametryOperatsii: End Sub

Private Sub txtKalibrovka_40_Change(): ZapisatParametryOperatsii: End Sub

Private Sub txtKalibrovka_99_Change(): ZapisatParametryOperatsii: End Sub

Private Sub Class_Initialize()
    
    Set tbKIM = wsKIM.ListObjects(1)
    
    With frmKIM
        Set txtTnsh = .txtVremya_KIM
        Set lstTablitsa = .lstHera
        Set txtTsht = .txtTsht
        Set txtTpz = .txtTpz
        Set optSloznost_Hera_1 = .optSloznost_1
        Set optSloznost_Hera_2 = .optSloznost_2
        Set optSloznost_Hera_3 = .optSloznost_3
        Set txtKalibrovka_15 = .txtKalibrovka_1
        Set txtKalibrovka_40 = .txtKalibrovka_2
        Set txtKalibrovka_99 = .txtKalibrovka_3
    End With

End Sub

Public Sub NovyyRaschot()

    Call OchistitFormu
    
    Dim i As Integer
    For i = tbKIM.ListRows.Count To 1 Step -1
        tbKIM.ListRows(i).Delete
    Next
    
    If tbKIM.ListRows.Count > 0 Then
        For i = 1 To tbKIM.ListColumns.Count - 2
            tbKIM.ListColumns(i).DataBodyRange.ClearContents
        Next
    End If
    
    Call ZapisatParametryOperatsii
    
End Sub

Public Sub DobavitNovuyuStroku(ParameryKIM As clsKIM)

    If DblFromCtrl(txtTnsh) <= 0 Then Exit Sub
    
    tbKIM.ListRows.Add

    Dim Parametry As Range
    Set Parametry = tbKIM.ListRows(tbKIM.ListRows.Count).Range
    
    With ParameryKIM
    
        Parametry(tbKIM.ListColumns("id").Index) = .id
        Parametry(tbKIM.ListColumns("Переход").Index) = .OpisaniyePerekhoda
        Parametry(tbKIM.ListColumns("Сложность").Index) = .Slozhost
        Parametry(tbKIM.ListColumns("D1").Index) = .D1
        Parametry(tbKIM.ListColumns("d2").Index) = .d2
        Parametry(tbKIM.ListColumns("Ширина").Index) = .Shirina
        Parametry(tbKIM.ListColumns("Длина").Index) = .Dlina
        Parametry(tbKIM.ListColumns("Кол-во поверхностей").Index) = .KolVoPoverkhnostey
        Parametry(tbKIM.ListColumns("Допуск размера").Index) = .DopuskRazmera
        Parametry(tbKIM.ListColumns("Допуск формы").Index) = .DopuskFormy
        Parametry(tbKIM.ListColumns("Допуск расположения").Index) = .DopuskRaspolozheniya
        Parametry(tbKIM.ListColumns("От оси").Index) = .OtOsi
        Parametry(tbKIM.ListColumns("Тригерный режим").Index) = .TrigernyyRezhim
        Parametry(tbKIM.ListColumns("Радиальное измерение").Index) = .RadialnoyeIzmereniye
        Parametry(tbKIM.ListColumns("Линейное измерение").Index) = .LieneynoyeIzmereniye
        Parametry(tbKIM.ListColumns("Кол-во сечений").Index) = .KolVoSecheniy
        Parametry(tbKIM.ListColumns("Масса").Index) = .Massa
        
    End With

    Call ZagruzitIzTablitsyVListBoks
 
End Sub

Public Sub ZapisatParametryOperatsii()

    With wsKIM
    
        If frmKIM.optSloznost_1 Then
            .Range("GruppaSlozhnosti") = 1
        ElseIf frmKIM.optSloznost_2 Then
            .Range("GruppaSlozhnosti") = 2
        Else
            .Range("GruppaSlozhnosti") = 3
        End If
        
        .Range("KolVoKalibrovok_15") = DblFromCtrl(txtKalibrovka_15)
        .Range("KolVoKalibrovok_40") = DblFromCtrl(txtKalibrovka_40)
        .Range("KolVoKalibrovok_99") = DblFromCtrl(txtKalibrovka_99)
        
    End With
    
    Call ZagruzitIzTablitsyVListBoks
    
End Sub

Public Sub UdalitStroki()

    Dim i As Integer, StrokaUdaleniya As Integer
    
    With lstTablitsa
        If .ListIndex = -1 Then Exit Sub
        
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                StrokaUdaleniya = i + 1
                tbKIM.ListRows(StrokaUdaleniya).Delete
            End If
        Next
        
    End With

    Call ZagruzitIzTablitsyVListBoks

End Sub

Private Sub ZagruzitIzTablitsyVListBoks()

    Call OchistitFormu
    
    If tbKIM.ListRows.Count = 0 Then Exit Sub
    
    With lstTablitsa
        
        Dim i As Long
        For i = 1 To tbKIM.ListRows.Count
            .AddItem
            .List(.ListCount - 1, 0) = tbKIM.DataBodyRange(i, tbKIM.ListColumns("Переход").Index)
            .List(.ListCount - 1, 1) = CStr(tbKIM.DataBodyRange(i, tbKIM.ListColumns("tнш").Index))
        Next
        
    End With
    
    With frmKIM
        .txtTpz = WorksheetFunction.RoundUp(wsKIM.Range("Tpz"), 0)
        .txtTsht = WorksheetFunction.RoundUp(WorksheetFunction.Sum(tbKIM.ListColumns("tнш").DataBodyRange), 0)
    End With
    
End Sub

Private Sub OchistitFormu()
    lstTablitsa.Clear
    txtTsht.text = Empty
    txtTpz.text = Empty
End Sub
