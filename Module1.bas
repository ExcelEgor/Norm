Attribute VB_Name = "Module1"
Option Explicit

Sub DeleteUnwantedFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim fileToKeep As String
    
    ' Укажите путь к папке (обязательно с обратным слэшем на конце)
    folderPath = "E:\Новая папка\"
    
    ' Проверка существования папки
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Указанная папка не существует.", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' Получаем первый файл в папке
    fileName = Dir(folderPath & "*.*")
    
    Do While fileName <> ""
        ' Приводим имя файла к нижнему регистру для исключения ошибок регистра
        Select Case fileName
            
            ' Список файлов, которые НУЖНО ОСТАВИТЬ:
            Case _
                "clsFigura_Krug.cls", _
                "clsFigura_List.cls", _
                "clsFigura_PryamougolnayaTruba.cls", _
                "clsFigura_Shestigrnnik.cls", _
                "clsFigura_Shveller_P.cls", _
                "clsFigura_Shveller_U.cls", _
                "clsFigura_Truba.cls", _
                "clsFigura_Ugolnik.cls", _
                "clsLentOtrez_Forma.cls", _
                "clsLentOtrez_Raschot.cls", _
                "clsLentOtrez_TokPOp.cls", _
                "frmLentochnoOtreznieStanki.frm", _
                "ITechnologicalShape.cls", _
                "mLentochnoOtreznaya.bas"
                
            Case Else
                ' Удаляем все остальные файлы
                On Error Resume Next
                Kill folderPath & fileName
                On Error GoTo 0
        End Select
        
        ' Переходим к следующему файлу
        fileName = Dir()
    Loop
    
    MsgBox "Очистка папки завершена.", vbInformation, "Готово"
End Sub

Sub GetFileList()
    Dim folderPath As String
    Dim fileName As String
    Dim i As Long
    
    ' Укажите путь к папке (обязательно с обратным слэшем на конце)
    folderPath = "E:\Новая папка\"
    
    ' Проверка существования папки
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Указанная папка не существует.", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' Подготовка листа для вывода данных
    Application.ScreenUpdating = False
    ActiveSheet.Cells.Clear
    ActiveSheet.Range("A1").Value = "Имя файла"
    ActiveSheet.Range("A1").Font.Bold = True
    
    i = 2
    ' Получаем первый файл в папке
    fileName = Dir(folderPath & "*.*")
    
    ' Цикл по всем файлам
    Do While fileName <> ""
        ActiveSheet.Cells(i, 1).Value = fileName
        i = i + 1
        ' Переход к следующему файлу
        fileName = Dir()
    Loop
    
    Application.ScreenUpdating = True
    
    ' Вывод результата
    If i > 2 Then
        MsgBox "Список файлов успешно получен. Найдено файлов: " & (i - 2), vbInformation, "Готово"
    Else
        MsgBox "Папка пуста.", vbExclamation, "Внимание"
    End If
End Sub

Function KolVoUnikalnykhZnacheniyVStolbtse(Stolbets As Range) As Long

    Dim col As New Collection, i As Long
    On Error Resume Next
        For i = 1 To Stolbets.Rows.Count
            If Stolbets(i).text <> "" Then col.Add Stolbets(i).text, Stolbets(i).text
        Next
    On Error GoTo 0
    
    KolVoUnikalnykhZnacheniyVStolbtse = col.Count
End Function

Sub ZakrytSkrytyyeKnigi()
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Windows(1).Visible = False Then wb.Close (False)
    Next
End Sub

Sub SokhranitProchiyeRaboty()
    
    If Not ActiveSheet.Range("A2") = "ID плана" Then Exit Sub
    
    Dim God As String
    God = Split(Range("H1"))(1) & Split(Range("H1"))(2)
    
    Dim Mesyatsa As Variant
    Mesyatsa = Array("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
    
    Dim Mesyats As String
    Mesyats = Split(Split(Range("H1"), ":")(1))(0)
    
    Dim i As Integer
    For i = 0 To 11
        If Mesyatsa(i) = Mesyats Then
            Mesyats = Format(i + 1, "00")
            Exit For
        End If
    Next
    
    Dim Tsekh As String
    Tsekh = Split(Range("G1"), ":")(1)
    
    Dim ImyaFayla As String
    ImyaFayla = God & " " & Mesyats & " " & Tsekh
    ImyaFayla = "D:\Проекты\Расчёт коээфициента к норме времени\Прочие работы\" & ImyaFayla & ".xlsx"
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=ImyaFayla, FileFormat:=xlOpenXMLWorkbook
    If Workbooks.Count > 1 Then ActiveWorkbook.Close (False)
    Application.DisplayAlerts = True
    
End Sub

Sub CloseTestXLAM2()
    Workbooks("test — копия.xlam").Close (False)
End Sub


Sub ZapisatTekushcheyeVremya()
Attribute ZapisatTekushcheyeVremya.VB_ProcData.VB_Invoke_Func = "N\n14"
    ActiveCell.Value = Format(Time, "h:mm")
End Sub

Sub DobavitVSlovarSborki()
    frmSlovarSborki.Show
End Sub

Private Sub frmGravirovkaShow(Dlina As Double, Shirina As Double)
    frmGravirovka.Show
End Sub

Sub ConverAllValuesTotext()
    Dim cell As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.Value = CStr(WorksheetFunction.Round(cell.Value, 3))
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function ZamenitLatinitsuNaKirillitsu(ByVal text As String) As String
    
    Dim LatinskiyeSimvoly As String
    Dim KirillicheskiyeSimvoly As String
    Dim i As Integer
    Dim Simvol As String
    Dim result As String
    
    LatinskiyeSimvoly = "AEKMHOPCTYXaekmhopctyx"
    KirillicheskiyeSimvoly = "АЕКМНОРСТУХаекмнорстух"
    
    result = text
    For i = 1 To Len(LatinskiyeSimvoly)
        Simvol = Mid(LatinskiyeSimvoly, i, 1)
        result = Replace(text, result, Mid(KirillicheskiyeSimvoly, i, 1))
    Next
    
    ZamenitLatinitsuNaKirillitsu = result
    
End Function

Sub ImportTableFromPDFToExcel()
    
    Dim i As Integer
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim KolVoStrok As Integer
    
    Dim newWS As Worksheet
    
    Dim pdfPath As String
    pdfPath = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Паспорт ЭМ\Перечень.pdf"
    
    Dim wordTable

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    Set wordDoc = wordApp.Documents.Open(pdfPath, ReadOnly:=True)
    
    If wordDoc.Tables.Count > 0 Then
    
        Workbooks.Add
        Set newWS = ActiveWorkbook.Worksheets(1)
        Set wordTable = wordDoc.Tables(1)
        wordTable.Range.Copy
        
        newWS.Paste
        
        wordDoc.Close SaveChanges:=False
        
        With newWS
            .Cells.UnMerge
            .Range("B:C,E:Q").Delete Shift:=xlToLeft
            .Range("A:B").EntireColumn.AutoFit
            
            KolVoStrok = Selection.Rows.Count
            For i = KolVoStrok To 1 Step -1
                If Not .Cells(i, 2) Like "Вилка*" And Not .Cells(i, 2) Like "Розетка*" Then
                    .Rows(i).Delete
                End If
            Next
            
        .Columns("A:A").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
            
        End With

        
    End If
    
End Sub


Sub Testd()
    
    Dim TimeBeg As Double
    Dim TimeEnd As Double
    
    TimeBeg = Now
    Dim tNsh As Double

    Dim i As Long
    For i = 1 To 100000
        ActiveSheet.Calculate
    Next

    TimeEnd = Now
    
    MsgBox Format(TimeEnd - TimeBeg, "h:mm:ss")
    
End Sub

Sub SozdatSoderchaniye()
    
    Dim i  As Integer
    Dim wb As Workbook
    Dim wsName As String
    Dim wsSoderzhaniye As Worksheet
    
    Set wb = ActiveWorkbook
    Set wsSoderzhaniye = wb.Worksheets(1)
    
    For i = 2 To wb.Worksheets.Count
        wsName = wb.Worksheets(i).Name
        wsSoderzhaniye.Hyperlinks.Add _
        Anchor:=wsSoderzhaniye.Cells(i - 1, 1), _
        Address:="", _
        SubAddress:="'" & wsName & "'!A1", TextToDisplay:=wsName
    Next

End Sub

Sub AddCheckboxInSelection()

    Dim L As Double, t As Double, W As Double, H As Double
    Dim cell As Range
    
    For Each cell In Selection

        With cell
            L = .Left
            t = .Top
            W = .Width
            H = .Height
            .NumberFormat = ";;;"
        End With

        ActiveSheet.CheckBoxes.Add(L + W / 2 - H / 2, t, H, H).Select
        
        With Selection
            .Value = xlOff
            .LinkedCell = cell.Address
            .Characters.text = ""
        End With

    Next
    
End Sub

Sub SaveImageFromUserFrom()
    Dim img As StdPicture
    Dim FilePath As String
    
    Set img = frmLazer.imgLogotip.Picture
    
    FilePath = "D:\img.emf"
    
    SavePicture img, FilePath
End Sub

Sub TpIzBufera()

    Application.ScreenUpdating = False
    
    Dim newWB As Workbook
    Dim newWS As Worksheet
    Dim lastRow As Long
    Set newWB = Workbooks.Add
    Set newWS = newWB.Worksheets(1)

    
    newWS.Range("A1").Select
    ActiveSheet.PasteSpecial Format:="Текст в кодировке Unicode"
    
    With newWS
    
        .Rows("1:1").Select
        Selection.Insert Shift:=xlDown
    
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("$A$1:$S$" & lastRow).AutoFilter field:=1, Criteria1:="<>А"
        .Cells.EntireRow.Delete
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        .Columns("A:E").Delete Shift:=xlToLeft
        .Columns("B:B").Delete Shift:=xlToLeft
        .Columns("C:L").Delete Shift:=xlToLeft
    
        .Range("D1:D" & lastRow).FormulaR1C1 = "=LEFT(SUBSTITUTE(RC[-1],"" "",""""),16)"
        .Range("E1:E" & lastRow).FormulaR1C1 = "=IFERROR(RIGHT(RC[-2],3)*1,"""")"
        .Range("A1:E" & lastRow) = .Range("A1:E" & lastRow).Value
        .Columns("C:C").Delete Shift:=xlToLeft
        
        'Заголовки
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

    Application.ScreenUpdating = True
    
End Sub

Sub SnyatZashchitySListov()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Unprotect
    Next
    On Error GoTo 0
End Sub
