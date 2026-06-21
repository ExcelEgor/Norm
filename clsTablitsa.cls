VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTablitsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public lstRaschot As MSForms.ListBox
Public txtTsht_Mekhanika As MSForms.TextBox
Public txtTsht_Slesar As MSForms.TextBox
Public txtTsht_Kontrol As MSForms.TextBox
Public ChkOkruglennyyeNomvy As MSForms.CheckBox
Public Stolbets_Perekhod As Integer, Stolbets_tMekhanika As Integer, Stolbets_tSlesar As Integer, Stolbets_tKontrol As Integer

Public Sub PeremestitStrokuVverkh(tbRaschot As ListObject)

    If lstRaschot.ListIndex = -1 Then Exit Sub

    Dim i As Integer
    Dim Stroka As Integer, StrokaPoslePeremeshcheniya As Integer

    For i = lstRaschot.ListCount - 1 To 0 Step -1
        If lstRaschot.Selected(i) Then
            Stroka = i + 1
            StrokaPoslePeremeshcheniya = i - 1
            Exit For
        End If
    Next

    If Stroka > 1 Then

        tbRaschot.ListRows(Stroka).Range.Cut
        tbRaschot.ListRows(Stroka - 1).Range.Insert Shift:=xlDown

        Call ZagruzkaVListBox(tbRaschot)

        lstRaschot.Selected(StrokaPoslePeremeshcheniya) = True
    End If

End Sub

Public Sub PeremestitStrokuVniz(tbRaschot As ListObject)

    If lstRaschot.ListIndex = -1 Then Exit Sub

    Dim i As Integer
    Dim Stroka As Integer, StrokaPoslePeremeshcheniya As Integer

    For i = lstRaschot.ListCount - 2 To 0 Step -1
        If lstRaschot.Selected(i) Then
            Stroka = i + 1
            StrokaPoslePeremeshcheniya = i + 1
            Exit For
        End If
    Next

    If Stroka > 0 And Stroka - 1 < lstRaschot.ListCount Then

        tbRaschot.ListRows(Stroka).Range.Cut

        Dim DiapazonPeremeshcheniya As Range
        If Stroka + 2 > tbRaschot.ListRows.Count Then
            Set DiapazonPeremeshcheniya = tbRaschot.ListRows(tbRaschot.ListRows.Count).Range.Offset(1)
        Else
            Set DiapazonPeremeshcheniya = tbRaschot.ListRows(Stroka + 2).Range
        End If
        DiapazonPeremeshcheniya.Insert Shift:=xlDown

        Call ZagruzkaVListBox(tbRaschot)

        lstRaschot.Selected(StrokaPoslePeremeshcheniya) = True

    End If

End Sub

Public Sub UdalitStroki(tbRaschot As ListObject)

    Dim i As Integer, StrokaUdaleniya As Integer
    
    With lstRaschot
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

Public Sub ZagruzkaVListBox(tbRaschot As ListObject)

    Dim StrokaListBox As Integer, StrokaTablitsa As Integer
    With lstRaschot
        .Clear
        For StrokaTablitsa = 1 To tbRaschot.ListRows.Count
            .AddItem
            StrokaListBox = .ListCount - 1
            .List(StrokaListBox, 0) = StrokaListBox + 1
            .List(StrokaListBox, 1) = tbRaschot.DataBodyRange(StrokaTablitsa, Stolbets_Perekhod)
            .List(StrokaListBox, 2) = VzyatVremyaIzTablitsty(tbRaschot, StrokaTablitsa, Stolbets_tMekhanika)
            .List(StrokaListBox, 3) = VzyatVremyaIzTablitsty(tbRaschot, StrokaTablitsa, Stolbets_tSlesar)
            .List(StrokaListBox, 4) = VzyatVremyaIzTablitsty(tbRaschot, StrokaTablitsa, Stolbets_tKontrol)
        Next
    End With
    
    Call ZapisatTsht
    
End Sub

Public Sub ZapisatTsht()

    txtTsht_Mekhanika.text = 0
    txtTsht_Slesar.text = 0
    txtTsht_Kontrol.text = 0
    
    Dim Normy
    Normy = RaschotThst
    
    Dim Tsht_Mekhanika As Double, Tsht_Slesar As Double, Tsht_Kontrol As Double
    Tsht_Mekhanika = Normy(1)
    Tsht_Slesar = Normy(2)
    Tsht_Kontrol = Normy(3)
    
    Dim Okrugl As Boolean
    If Not ChkOkruglennyyeNomvy Is Nothing Then
        Okrugl = ChkOkruglennyyeNomvy.Value
    End If

    If lstRaschot.ListCount > 0 Then

        If Okrugl Then
            txtTsht_Mekhanika.text = CStr(OkruglenieTsht(Tsht_Mekhanika))
            txtTsht_Slesar.text = CStr(OkruglenieTsht(Tsht_Slesar))
            txtTsht_Kontrol.text = CStr(OkruglenieTsht(Tsht_Kontrol))
        Else
            txtTsht_Mekhanika.text = FormatVremeni(Tsht_Mekhanika, 1)
            txtTsht_Slesar.text = FormatVremeni(Tsht_Slesar, 1)
            txtTsht_Kontrol.text = FormatVremeni(Tsht_Kontrol, 1)
        End If

    End If


End Sub

Private Function RaschotThst() As Variant
    
    Dim Normy(1 To 3) As Double
    
    If lstRaschot.ListCount > 0 Then

        Dim i As Integer
        For i = 0 To lstRaschot.ListCount - 1
            Normy(1) = Normy(1) + CDbl(lstRaschot.List(i, 2))
            Normy(2) = Normy(2) + CDbl(lstRaschot.List(i, 3))
            Normy(3) = Normy(3) + CDbl(lstRaschot.List(i, 4))
        Next

    End If
    
    RaschotThst = Normy

End Function

Private Function VzyatVremyaIzTablitsty(tbRaschot As ListObject, StrokaTablitsa As Integer, NomerStolbtsa As Integer)
    VzyatVremyaIzTablitsty = FormatVremeni(tbRaschot.DataBodyRange(StrokaTablitsa, NomerStolbtsa))
End Function

Private Function FormatVremeni(Vremya As Double, Optional ChisloZnakov As Integer = 2) As String
    FormatVremeni = CStr(Round(Vremya, ChisloZnakov))
End Function

