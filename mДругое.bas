Attribute VB_Name = "mДругое"
Option Explicit

Public Sub OktryRaschotBazovogoKarkasa()
    
    Dim wbRaschot As Workbook
    Set wbRaschot = Workbooks.Open(PoluchitPutKShablonam & "АМИЕ.301241.096.xltx", , , , , , , , , False)
    wbRaschot.Windows(1).Visible = False
    
    With frmTokFrez
        .tabForMlt.Value = 1
        .Show
        Call .Raschot.OtkryRaschot(wsTablitsa_TokarnoFrezer.ListObjects(1), wbRaschot)
    End With

    
End Sub

Sub СкопироватьФЗП()
    ActiveSheet.Range("B75:G75,J75").Copy
End Sub

Sub ZakrytVseRaschotySborkiKomplektovaniya()

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.ActiveSheet.Cells(1, 1) = "sborka" Then
            wb.Close (False)
        End If
    Next
    
End Sub


Function CalcPrice(Trudoemskost As Double, Optional Material As Double = 0, Optional PKI As Double = 0, Optional MZK As Double = 0, Optional UchetPribili As Boolean = True)
    
    Dim StoimostNormoChasa As Double
    StoimostNormoChasa = 467.16
    
    Dim PKI_MZK
    PKI_MZK = PKI + MZK
    
    Dim VozvrOth As Double
    VozvrOth = Material * 0.05
    
    Dim ItogMatRashod
    ItogMatRashod = Material - VozvrOth + PKI_MZK
    
    Dim OsnZarPlata As Double
    OsnZarPlata = Trudoemskost * StoimostNormoChasa
    
    Dim DopZarPlata As Double
    DopZarPlata = OsnZarPlata * 0.087
    
    Dim RashNaOplatRabotnikov
    RashNaOplatRabotnikov = OsnZarPlata + DopZarPlata
    
    Dim SocObespech As Double
    SocObespech = RashNaOplatRabotnikov * 0.304
    
    Dim SpeCTehnOsnast As Double
    SpeCTehnOsnast = OsnZarPlata * 0.03
    
    Dim ObsheProizvRashod As Double
    ObsheProizvRashod = RashNaOplatRabotnikov * 1.651
    
    Dim ObsheHozRashod As Double
    ObsheHozRashod = RashNaOplatRabotnikov * 0.3
    
    Dim PolnSebest As Double
    PolnSebest = ItogMatRashod + OsnZarPlata + DopZarPlata + SocObespech + SpeCTehnOsnast + ObsheProizvRashod + ObsheHozRashod
    
    Dim Pribil As Double
    If UchetPribili Then
        Pribil = ItogMatRashod * 0.01 + (PolnSebest - ItogMatRashod) * 0.25
    Else
        Pribil = 0
    End If
    
    CalcPrice = PolnSebest + Pribil
    
End Function


Function OboznachTochka(Oboznach As String) As String
    Oboznach = Left(Oboznach, 4) & "." & Mid(Oboznach, 5)
    Oboznach = Left(Oboznach, 11) & "." & Mid(Oboznach, 12)
    If Len(Oboznach) > 15 Then
        OboznachTochka = Left(Oboznach, 15) & "-" & Mid(Oboznach, 16)
    Else
        OboznachTochka = Oboznach
    End If
End Function


Sub ШиринаСтолбцов()
    Debug.Print Selection.Columns.Count
    Dim selcount
    selcount = Selection.Columns.Count
    Dim maxcol As Double
    Dim i As Integer, j As Integer
    For i = 2 To selcount
        For j = i To 1 Step -1
            If Selection.Columns(j).ColumnWidth > Selection.Columns(i).ColumnWidth Then maxcol = Selection.Columns(j).ColumnWidth
        Next
    Next
    For i = 1 To selcount
        Selection.Columns(i).ColumnWidth = maxcol
    Next
End Sub


Function Scepka(rText As Range, Razdelitel As String) As String
    Dim c As Range
    For Each c In rText
        Scepka = Scepka & Razdelitel & " " & Left(c.Value, 4)
    Next
    Scepka = Right(Scepka, Len(Scepka) - 2)
End Function

Function RaschetKolichestvaSimvolov(c As Range) As Double
        
    If c.text <> Empty Then
        Dim str As String
        str = Replace(c.text, " ", "")
        If str Like "=*" Then
            RaschetKolichestvaSimvolov = Application.Evaluate(Replace(str, "=", ""))
        Else
            RaschetKolichestvaSimvolov = Len(str)
        End If
    End If
    
End Function

Function SortArray(myArray As Variant) As Variant
    Dim L As Long, u As Long, i1 As Long, i2 As Long, Im As Long, tmp As Variant
    'Определение наименьшего индекса массива
    L = LBound(myArray)
    'Определение наибольшего индекса массива
    u = UBound(myArray)
    For i1 = L To u
        Im = i1
        For i2 = i1 To u
            'Поиск наименьшего элемента массива, начиная с элемента myArray(i1)
            If myArray(i2) < myArray(Im) Then Im = i2
        Next
        'Если наименьший элемент не является текущим (im <> i1),
        'тогда наименьший элемент и текущий меняются местами
        If Im <> i1 Then
            tmp = myArray(i1)
            myArray(i1) = myArray(Im)
            myArray(Im) = tmp
        End If
    Next
    SortArray = myArray
End Function


