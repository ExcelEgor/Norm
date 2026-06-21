Attribute VB_Name = "mSborkaIzBuferaObmena"
Option Explicit

Sub RaschotSborkiIzBufera()

    On Error GoTo ErrorHandler

    With Application
        .ScreenUpdating = False
        .CutCopyMode = False
        .Calculation = xlCalculationManual
    End With
    
    wsKomplektovaniye.Copy
    Dim wbRaschot As Workbook
    Set wbRaschot = ActiveWorkbook
    
    Dim wsRaschot As Worksheet
    Set wsRaschot = wbRaschot.ActiveSheet
    
    Dim tbKomplektovaniye As ListObject
    Set tbKomplektovaniye = wsRaschot.ListObjects("tbKomplektovaniye")
    
    Call UdalitSsylkuNaKnigu(wsRaschot)
    
    Call VstavitIzBuferaPerenestiVRaschot(wsRaschot)
    
    Call ProtyanutFormuly(tbKomplektovaniye)
    
    'Call ZapolnitMassuDlinuShnura(wsRaschot)
    
    Call DobavitChekBoksy(tbKomplektovaniye)
    
    Call RazvernutRezultatNaVesEkran(wbRaschot)
    
ErrorHandler:
    With Application
        .CutCopyMode = False
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
End Sub

Private Sub UdalitSsylkuNaKnigu(wsRaschot)

    wsRaschot.Cells.Replace What:=ActiveWorkbook.Name & "!", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
            
End Sub

Private Sub VstavitIzBuferaPerenestiVRaschot(wsRaschot As Worksheet)

    Dim wbBufer As Workbook
    Set wbBufer = Workbooks.Add
    
    Dim wsBufer As Worksheet
    Set wsBufer = wbBufer.ActiveSheet
    
    wsBufer.Cells(1, 5).Select
    wsBufer.PasteSpecial Format:="Текст"
    
    With wsBufer
        .Cells(1, 1).Select
        .Cells(1, 1) = 1
        .PasteSpecial Format:="Текст", link:=False, DisplayAsIcon:=False
        .Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
            :="¦", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
            1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
            , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
            Array(19, 1)), TrailingMinusNumbers:=True
    End With
    
    Call PerenestiVRaschot(Selection, wsRaschot)
    
    wbBufer.Close (False)
    
End Sub

Private Sub PerenestiVRaschot(DiapazonDannykh As Range, wsRaschot As Worksheet)

    DiapazonDannykh.Columns(3).Copy
    wsRaschot.Cells(4, 1).PasteSpecial Paste:=xlPasteValues
    
    DiapazonDannykh.Columns(8).Copy
    wsRaschot.Cells(4, 2).PasteSpecial Paste:=xlPasteValues
    
    DiapazonDannykh.Columns(18).Copy
    wsRaschot.Cells(4, 3).PasteSpecial Paste:=xlPasteValues
    
End Sub
 
Private Sub ProtyanutFormuly(tbKomplektovaniye As ListObject)
    With tbKomplektovaniye
        .ListColumns(.ListColumns("Сборка").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("Сборка").Index).Formula
        .ListColumns(.ListColumns("Есть в словаре").Index).DataBodyRange.Formula = .DataBodyRange(1, .ListColumns("Есть в словаре").Index).Formula
    End With
End Sub
 
Private Sub ZapolnitMassuDlinuShnura(wsRaschot As Worksheet)

    Call ZapolnitMassu(wsRaschot)
    Call ZapolnitDlinuShnura(wsRaschot)

End Sub

Private Sub ZapolnitMassu(wsRaschot)

    Dim MassaSborki As String
    Do While Not IsNumeric(MassaSborki)
    
        MassaSborki = Application.InputBox("Введите массу сборки, кг", "Масса сборки")
        
        If MassaSborki = "False" Then
            MassaSborki = 0
        Else
            MassaSborki = Replace(MassaSborki, ".", ",")
        End If
        
    Loop
    
    wsRaschot.Range("MassaSborki").Value = CDbl(MassaSborki)

End Sub

Private Sub ZapolnitDlinuShnura(wsRaschot)

    With wsRaschot
    
        .Range("DlinaShnura").Value = 0
        
        Dim rNaimenovaniye As Range
        Set rNaimenovaniye = .ListObjects(1).ListColumns(1).DataBodyRange

        If WorksheetFunction.CountIf(rNaimenovaniye, "Шнур") + WorksheetFunction.CountIf(rNaimenovaniye, "Прокладка уплотнительная") > 0 Then
                
            Dim DlinaShnura As Variant
            Dim DlinaVvedenaPravilno As Boolean
            Do While DlinaVvedenaPravilno = False
                
                DlinaShnura = Application.InputBox("Введите длину уплотнительного шнура", "Уплотнительный шнур")
                
                If DlinaShnura = "False" Then
                    DlinaShnura = 0
                Else
                    DlinaVvedenaPravilno = PolozhitelnyyeChisla(Replace(DlinaShnura, ".", ","))
                End If
                
            Loop
            
            .Range("DlinaShnura").Value = CDbl(DlinaShnura)

        End If
        
    End With
    
End Sub

Private Sub DobavitChekBoksy(tbKomplektovaniye As ListObject)

    Dim i As Integer
    Dim L As Double, t As Double, W As Double, H As Double
    Dim cell As Range
    
    For i = 1 To tbKomplektovaniye.ListRows.Count
        
        With tbKomplektovaniye
        
            If .DataBodyRange(i, .ListColumns("Наименование").Index) Like "Винт*" Then
                
                Set cell = .DataBodyRange(i, .ListColumns("Заподлицо").Index)
                
                With cell
                    L = .Left
                    t = .Top
                    W = .Width
                    H = .Height
                    .NumberFormat = ";;;"
                End With
        
                ActiveSheet.CheckBoxes.Add(L + W / 2 - H / 2, t - H / 8, H, H).Select
                
                With Selection
                    .Value = xlOff
                    .LinkedCell = cell.Address
                    .Characters.text = ""
                End With
        
            End If
            
        End With
        
    Next
    
    ActiveSheet.Cells(1, 1).Select
    
End Sub

Private Sub RazvernutRezultatNaVesEkran(wbRaschot As Workbook)

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = wbRaschot.Name Then
            wb.Application.WindowState = xlMaximized
        Else
            wb.Application.WindowState = xlMinimized
        End If
    Next
    
End Sub

