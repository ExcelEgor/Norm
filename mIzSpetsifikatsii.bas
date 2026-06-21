Attribute VB_Name = "mIzSpetsifikatsii"
Option Explicit
Option Private Module

Sub PreobrazovatSpetsifikatsiyu(wb_spetsifikatsiya)
    
    Dim cell As Range
    
    Dim ws_spetsifikatsiya As Worksheet
    Set ws_spetsifikatsiya = wb_spetsifikatsiya.ActiveSheet
    
    Dim poslednyaya_stroka As Long
    
    With ws_spetsifikatsiya
        poslednyaya_stroka = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range("$F$2:$F$" & poslednyaya_stroka).FormulaR1C1 = "=IF(R[-1]C=""-"",""-"",IF(RC[-3]=""Примечания"",""-"",IF(AND(RC[-3]<>"""",RC[-4]="""",RC[-2]="""",RC[-5]=""""),RC[-3],IF(R[-1]C<>"""",R[-1]C))))"
        .Range("$F$2:$F$" & poslednyaya_stroka) = .Range("$F$2:$F$" & poslednyaya_stroka).Value
        .Range("A2").AutoFilter
        .Range("$A$2:$F$" & poslednyaya_stroka).AutoFilter field:=4, Criteria1:="="
        .Cells.EntireRow.Delete

        poslednyaya_stroka = .Cells(ws_spetsifikatsiya.Rows.Count, 4).End(xlUp).Row
        
        For Each cell In .Range("$D$2:$D$" & poslednyaya_stroka).Cells
            cell = cell.Value * 1
        Next
        
    End With
    
End Sub

Private Sub PerenestiIzSpetsifikatsiiVTablitsu(tb_komplektovaniye As ListObject, ws_spetsifikatsiya As Worksheet)
    
    Dim poslednyaya_stroka As Long
    With ws_spetsifikatsiya
        poslednyaya_stroka = .Cells(.Rows.Count, 2).End(xlUp).Row
    End With
    
    With tb_komplektovaniye
        If .ListRows.Count = 0 Then .ListRows.Add
        .DataBodyRange(1, .ListColumns("Поз.").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(1).Resize(poslednyaya_stroka).Value
        .DataBodyRange(1, .ListColumns("Обозначение").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(2).Resize(poslednyaya_stroka).Value
        .DataBodyRange(1, .ListColumns("Наименование").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(3).Resize(poslednyaya_stroka).Value
        .DataBodyRange(1, .ListColumns("Кол-во").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(4).Resize(poslednyaya_stroka).Value
        .DataBodyRange(1, .ListColumns("Примечание").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(5).Resize(poslednyaya_stroka).Value
        .DataBodyRange(1, .ListColumns("Тип").Index).Resize(poslednyaya_stroka) = ws_spetsifikatsiya.Columns(6).Resize(poslednyaya_stroka).Value
        .DataBodyRange.WrapText = False
    End With
    
End Sub

Sub RaschotBlokovStoyekModuley(wb_spetsifikatsiya As Workbook, cmd As Control)

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Dim ws_spetsifikatsiya As Worksheet
    Set ws_spetsifikatsiya = wb_spetsifikatsiya.Worksheets(1)
    
    Dim wb_blok_stoyka_modul As Workbook
    Set wb_blok_stoyka_modul = Workbooks.Open(PUT_K_SHABLONAM & "Электромонтаж блоков, стоек.xltm")
    
    Dim ImyaLista As String, tb_name As String
    Select Case cmd.Name
        Case frmSpetsifikatsiya.cmdStoyka.Name
            ImyaLista = "Стойка"
            tb_name = "tbMontazhStoyka"
        Case frmSpetsifikatsiya.cmdModul.Name
            ImyaLista = "Модуль"
            tb_name = "tbMontazhModul"
        Case frmSpetsifikatsiya.cmdBlok.Name
            ImyaLista = "Блок"
            tb_name = "tbMontazhBlok"
    End Select
    
    Dim ws_blok_stoyka_modul As Worksheet
    Set ws_blok_stoyka_modul = wb_blok_stoyka_modul.Worksheets(ImyaLista)
    
    Dim tb_komplektovaniye As ListObject
    Set tb_komplektovaniye = ws_blok_stoyka_modul.ListObjects(tb_name)
    
    wb_blok_stoyka_modul.Activate
    ws_blok_stoyka_modul.Select
    
    Call PreobrazovatSpetsifikatsiyu(wb_spetsifikatsiya)
    Call PerenestiIzSpetsifikatsiiVTablitsu(tb_komplektovaniye, ws_spetsifikatsiya)

    Call ZapisatMassu(ws_blok_stoyka_modul)
    
    Call ZapisatKolVoEtazhey(ws_blok_stoyka_modul)
    
    ws_blok_stoyka_modul.Cells(2, 2).Value = wb_spetsifikatsiya.Name
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
    wb_spetsifikatsiya.Close (False)
    Unload frmSpetsifikatsiya
    
End Sub

Private Sub ZapisatMassu(ws_blok_stoyka_modul As Worksheet)

    Dim MassaSborki As String
    Do While Not IsNumeric(MassaSborki)
        MassaSborki = Application.InputBox("Введите массу сборки, кг", "Масса сборки")
        If MassaSborki = "False" Then
            MassaSborki = 0
        Else
            MassaSborki = Replace(MassaSborki, ".", ",")
        End If
    Loop
    ws_blok_stoyka_modul.Range("Масса").Value = CDbl(MassaSborki)
    
    
End Sub

Private Sub ZapisatKolVoEtazhey(ws_blok_stoyka_modul As Worksheet)
    
    Dim KolVoEtazhey As String
    If ws_blok_stoyka_modul.Name = "Стойка" Then
        Do While Not IsNumeric(KolVoEtazhey)
            KolVoEtazhey = Application.InputBox("Введите кол-во этажей", "Кол-во этажей")
            If KolVoEtazhey = "False" Then
                KolVoEtazhey = 0
            Else
                KolVoEtazhey = Replace(KolVoEtazhey, ".", ",")
            End If
        Loop
        ws_blok_stoyka_modul.Range("КолВоЭтажей").Value = CDbl(KolVoEtazhey)
    End If
    
End Sub

Sub RaschotKomplektovaniyaSborki(wb_spetsifikatsiya)

    Dim ws_spetsifikatsiya As Worksheet
    Set ws_spetsifikatsiya = wb_spetsifikatsiya.ActiveSheet

    Dim wb_sborka As Workbook

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call PreobrazovatSpetsifikatsiyu(wb_spetsifikatsiya)
    
    wsKomplektovaniye1.Copy
    
    Set wb_sborka = ActiveWorkbook
    
    Dim ws_sborka As Worksheet
    Set ws_sborka = wb_sborka.ActiveSheet
    
    Dim tb_sborka As ListObject
    Set tb_sborka = ws_sborka.ListObjects("tbKomplektovaniye132")
    
    Call PerenestiIzSpetsifikatsiiVTablitsu(tb_sborka, ws_spetsifikatsiya)
    tb_sborka.ListColumns("Сборка").DataBodyRange.Formula = tb_sborka.ListColumns("Сборка").DataBodyRange.Formula
    
    'Call VvodMassyDlinyShnura(tb_sborka, ws_sborka)

    With wb_sborka.Worksheets(1)
        .Cells.Replace What:=ActiveWorkbook.Name & "!", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    End With

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
    Unload frmSpetsifikatsiya
    
    wb_spetsifikatsiya.Close (False)
    
    wb_sborka.Worksheets(1).Cells(1, 1).Select
    
End Sub

Private Sub VvodMassyDlinyShnura(tb_sborka As ListObject, ws_sborka As Worksheet)

    Dim MassaSborki As String
    Dim DlinaShnura As Variant
    Dim DlinaShnuraVvedenaPravilno As Boolean
    Dim rNaimenovaniye As Range
    
    Set rNaimenovaniye = tb_sborka.ListColumns("Наименование").DataBodyRange
    
    With ws_sborka
        
        Do While Not IsNumeric(MassaSborki)
            MassaSborki = Application.InputBox("Введите массу сборки, кг", "Масса сборки")
            If MassaSborki = "False" Then
                MassaSborki = 0
            Else
                MassaSborki = Replace(MassaSborki, ".", ",")
            End If
        Loop
        .Range("MassaSborki").Value = CDbl(MassaSborki)
        
        If WorksheetFunction.CountIf(rNaimenovaniye, "Шнур") > 0 Or WorksheetFunction.CountIf(rNaimenovaniye, "Прокладка уплотнительная") Then

            Do While DlinaShnuraVvedenaPravilno = False
                DlinaShnura = Application.InputBox("Введите длину уплотнительного шнура", "Уплотнительный шнур")
                If DlinaShnura = "False" Then
                    DlinaShnura = 0
                Else
                    DlinaShnura = Replace(DlinaShnura, ".", ",")
                    If IsNumeric(DlinaShnura) Then
                        If CDbl(DlinaShnura) > 0 Then
                            DlinaShnuraVvedenaPravilno = True
                        End If
                    End If
                End If
            Loop
            
            .Range("DlinaShnura").Value = CDbl(DlinaShnura)
        Else
            .Range("DlinaShnura").Value = 0
        End If
        
        DobavitChekBoksy tb_sborka, ws_sborka
        
    End With
    
End Sub

Private Sub DobavitChekBoksy(tb_sborka As ListObject, ws_sborka As Worksheet)

    Dim i As Integer
    Dim L As Double, t As Double, W As Double, H As Double
    Dim cell As Range
    
    Dim chk
    
    With tb_sborka
        For i = 1 To .ListRows.Count
            If .DataBodyRange(i, .ListColumns("Наименование").Index) Like "Винт*" Then
                Set cell = .DataBodyRange(i, .ListColumns("Заподлицо").Index)
                With cell
                    L = .Left
                    t = .Top - 1
                    W = .Width
                    H = .Height
                    .NumberFormat = ";;;"
                End With
                
                Set chk = ws_sborka.CheckBoxes.Add(L + W / 2 - H / 2, t, H, H)
                With chk
                    .Value = xlOff
                    .LinkedCell = cell.Address
                    .Characters.text = ""
                End With
        
            End If
        Next
    End With
        
End Sub

Public Sub ZapuskFormyIzSpetsifikatsii(wb As Workbook, Optional show_message As Boolean = False)

    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)

    If ws.Cells(2, 1) = "Позиция" And ws.Cells(2, 2) = "Обозначение" And ws.Cells(2, 3) = "Наименование" And _
        ws.Cells(2, 4) = "Количество" And ws.Cells(2, 5) = "Примечание" Then
        frmSpetsifikatsiya.Show
    Else
        If show_message Then
            MsgBox "Данная книга не является документом со спецификацией", vbExclamation, "Ошибка"
        End If
    End If
    
    Call InitializeAppEvents
    

End Sub
