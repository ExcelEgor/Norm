Attribute VB_Name = "mUpakovkaIzBufera"
Option Explicit

Private Type NormaVremeni
    Tpz As Double
    Tsht As Double
End Type

Sub KomplektovaniyePoTsekhamBufera()
    
'    Application.ScreenUpdating = False

    Dim wbUpakovka As Workbook
    Set wbUpakovka = Workbooks.Open(PUT_K_SHABLONAM & "Комплект ЗИП, упаковка.xltm")
    
    Dim tbUpakovka As ListObject
    Set tbUpakovka = wbUpakovka.Worksheets("Упаковка").ListObjects("tbUpakovka")

    Dim newWB As Workbook
    Dim newWS As Worksheet
    Dim lastRow As Long
    Set newWB = Workbooks.Add
    Set newWS = newWB.Worksheets(1)

    newWS.Range("A1").Select
    ActiveSheet.PasteSpecial Format:="Текст в кодировке Unicode"
    
    newWS.Copy
    
    Dim wbDlyaTP As Workbook, wsDlyaTP As Worksheet
    Set wbDlyaTP = ActiveWorkbook
    Set wsDlyaTP = wbDlyaTP.Worksheets(1)
    
    newWB.Activate
    newWS.Select
    
    With newWS
    
        .Rows("1:1").Select
        Selection.Insert Shift:=xlDown
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Columns("C:D").Insert Shift:=xlToLeft

        
        .Range("$B$2:$B$" & lastRow).FormulaR1C1Local = "=ЕСЛИ(RC[-1]=""А"";RC[8];R[-1]C)"
        .Range("$C$2:$C$" & lastRow).FormulaR1C1Local = "=ЕСЛИ(RC[-2]=""А"";RC[2];R[-1]C)"
        .Range("$D$2:$D$" & lastRow).FormulaR1C1Local = "=ЕСЛИ(RC[-3]=""А"";RC[4];R[-1]C)"
        
        Dim Tablitsa As Range
        Set Tablitsa = .Range("$A$2:$U$" & lastRow)

    End With
    
    PerenestiKomplektovaniye Tablitsa, 33208, wbUpakovka
    PerenestiKomplektovaniye Tablitsa, 33430, wbUpakovka
    
    UpakovkaIzBufera tbUpakovka, wbDlyaTP, wsDlyaTP
    
    newWB.Close (False)

    wbDlyaTP.Worksheets(1).Copy after:=wbUpakovka.Worksheets(wbUpakovka.Worksheets.Count)
    
    wbDlyaTP.Close (False)

    Application.ScreenUpdating = True
    
End Sub

Private Sub PerenestiKomplektovaniye(Tablitsa As Range, Tsekh As Long, wbUpakovka As Workbook)

    Dim tbKomplekt As ListObject, tbKomplekt33430 As ListObject

    If Tsekh = 33208 Then
        Set tbKomplekt = wbUpakovka.Worksheets("Упаковка").ListObjects("tbKomplektovaniye_1")
    ElseIf Tsekh = 33430 Then
        Set tbKomplekt = wbUpakovka.Worksheets("Упаковка").ListObjects("tbKomplektovaniye_2")
    Else
        Exit Sub
    End If

    With Tablitsa
    
        .AutoFilter 1, "К", xlOr, "М"
        .AutoFilter 2, "Комплектование"
        .AutoFilter 3, CStr(Tsekh)
        
        If WorksheetFunction.Subtotal(103, .Columns(1)) = 1 Then Exit Sub
        tbKomplekt.ListRows.Add
        
        .Columns(5).Copy
        tbKomplekt.DataBodyRange(1, 1).PasteSpecial (xlPasteValues)
        
        .Columns(10).Copy
        tbKomplekt.DataBodyRange(1, 2).PasteSpecial (xlPasteValues)
        
        .Columns(20).Copy
        tbKomplekt.DataBodyRange(1, 3).PasteSpecial (xlPasteValues)
        
        .Columns(21).Copy
        tbKomplekt.DataBodyRange(1, 4).PasteSpecial (xlPasteValues)
        
        tbKomplekt.ListRows(1).Delete
        
        Dim i As Integer
        For i = 1 To tbKomplekt.ListRows.Count
            If tbKomplekt.DataBodyRange(i, 4) > 0 Then tbKomplekt.DataBodyRange(i, 4) = 1 * tbKomplekt.DataBodyRange(i, 4)
        Next
        
    End With
        
End Sub

Private Sub UpakovkaIzBufera(tbUpakovka As ListObject, wbDlyaTP As Workbook, wsDlyaTP As Worksheet)

    Application.ScreenUpdating = False
    
    PreobrazovatVstablennyyeDannyye wbDlyaTP, wsDlyaTP
    
    SozdatZagolovki wsDlyaTP
    
    PerenestiZnacheniya_Upakovka wsDlyaTP, tbUpakovka
    
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub PreobrazovatVstablennyyeDannyye(wbDlyaTP As Workbook, wsDlyaTP As Worksheet)
    
    Dim PoslednyayaStroka As Long
    
    wbDlyaTP.Activate
    With wsDlyaTP
        .Rows("1:1").Select
        Selection.Insert Shift:=xlDown
        PoslednyayaStroka = wsDlyaTP.Cells(wsDlyaTP.Rows.Count, 1).End(xlUp).Row
    
        .Range("$A$1:$S$" & PoslednyayaStroka).AutoFilter field:=1, Criteria1:="<>А"
        .Cells.EntireRow.Delete
        PoslednyayaStroka = .Cells(.Rows.Count, 1).End(xlUp).Row

        .Columns("A:E").Delete Shift:=xlToLeft
        .Columns("B:B").Delete Shift:=xlToLeft
        .Columns("C:J").Delete Shift:=xlToLeft
    
        .Range("D1:D" & PoslednyayaStroka).FormulaR1C1 = "=LEFT(SUBSTITUTE(RC[-1],"" "",""""),16)"
        .Range("E1:E" & PoslednyayaStroka).FormulaR1C1 = "=IFERROR(RIGHT(RC[-2],3)*1,"""")"
        .Range("A1:E" & PoslednyayaStroka) = .Range("A1:E" & PoslednyayaStroka).Value
        .Columns("C:C").Delete Shift:=xlToLeft
    End With
    
End Sub

Private Sub SozdatZagolovki(wsDlyaTP As Worksheet)
    With wsDlyaTP
        .Rows("1:1").Select
        Selection.Insert Shift:=xlDown
        .Cells(1, 1) = "№ оп."
        .Cells(1, 2) = "Операция"
        .Cells(1, 3) = "ТТП"
        .Cells(1, 4) = "№ оп. по ТТП"
        .Cells(1, 5) = "Тпз"
        .Cells(1, 6) = "Тшт"
        .Columns("A:F").EntireColumn.AutoFit
        .Cells(1, 1).Select
    End With
End Sub

Private Sub PerenestiZnacheniya_Upakovka(wsDlyaTP As Worksheet, tbUpakovka As ListObject)
    
    Dim PoslednyayaStroka As Long
    PoslednyayaStroka = wsDlyaTP.Cells(wsDlyaTP.Rows.Count, 1).End(xlUp).Row
    
    Dim Norma As NormaVremeni
    
    Dim i As Integer
    For i = 2 To PoslednyayaStroka
        Norma = NaytiNormyVremeni_Upakovka(tbUpakovka, wsDlyaTP.Cells(i, 3), wsDlyaTP.Cells(i, 4))
        wsDlyaTP.Cells(i, 5) = Norma.Tpz
        wsDlyaTP.Cells(i, 6) = Norma.Tsht
    Next
    
End Sub

Private Function NaytiNormyVremeni_Upakovka(tbUpakovka As ListObject, TTP As String, NomerOperatsii As Integer) As NormaVremeni
    
    Dim NomerOperatsiiDlyaPoiska As Variant
    NomerOperatsiiDlyaPoiska = NomerOperatsii
    If TTP = "ДАИЕ.01208.00002" Then
        Select Case NomerOperatsii
            Case 105, 175
                NomerOperatsiiDlyaPoiska = "105, 175"
            Case 110, 180
                NomerOperatsiiDlyaPoiska = "110, 180"
            Case Else
        End Select
    End If
    
    Dim i As Integer
    With tbUpakovka
        For i = 1 To .ListRows.Count
            If .DataBodyRange(i, 2) = TTP And .DataBodyRange(i, 3) = NomerOperatsiiDlyaPoiska Then
                NaytiNormyVremeni_Upakovka.Tpz = .DataBodyRange(i, 5)
                NaytiNormyVremeni_Upakovka.Tsht = .DataBodyRange(i, 6)
                Exit For
            End If
        Next
    End With
    
End Function
