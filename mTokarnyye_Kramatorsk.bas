Attribute VB_Name = "mTokarnyye_Kramatorsk"
Option Explicit
'Нормативы времени и режимов резания для нормирования работ токарных станках с наибольшим диаметром изделия, устанавливаемого над станиной, до 1000 мм
'Единичное и мелкосерийное производство
'Часть 1
'Режимы резания, нормы времени вспомогательного, на обслуживание рабочего места и подготовительно-заключительного
'Краматорск 184

Const Koeffitsient As Double = 1.1
Public Const DMAX_KRAMATORSK As Integer = 1000

Function PolirovaniyeNaTokranomStankeShkurkoy(d As Double, L As Double, Ra As Double, TipPoverkhnosti As Integer) As Double
    'Полирование поверхностей наждачным полотном. Сталь углеродистая и легированная, чугн и медные сплавы. Карта 83
    
    If d > DMAX_KRAMATORSK Then Exit Function
    
    Dim a As Double, X As Double
    Select Case d
        Case Is <= 30:  a = 0.133792304884813:  X = 0.586371919212006
        Case Is <= 40:  a = 0.143406325276126:  X = 0.599229056251756
        Case Is <= 50:  a = 0.164986238358309:  X = 0.593052229136354
        Case Is <= 63:  a = 0.187021385460527:  X = 0.590547191928651
        Case Is <= 80:  a = 0.204745609234213:  X = 0.599354226632251
        Case Is <= 100: a = 0.224052866966898:  X = 0.599938084575839
        Case Is <= 125: a = 0.24417630953762:   X = 0.600950349015101
        Case Is <= 160: a = 0.267202509018292:  X = 0.606818481437288
        Case Is <= 200: a = 0.297427642921033:  X = 0.605505712977424
        Case Is <= 250: a = 0.323513047524501:  X = 0.61141349485436
        Case Is <= 315: a = 0.351474258820695:  X = 0.61660107247504
        Case Is <= 400: a = 0.394922377027482:  X = 0.614406017742761
        Case Is <= 500: a = 0.430305144103749:  X = 0.61964358366594
        Case Is <= 630: a = 0.473942619973737:  X = 0.621902875038521
        Case Is <= 800: a = 0.520068135803267:  X = 0.625377653938578
        Case Else:      a = 0.578557638726126:  X = 0.626583862060593
    End Select
    
    Dim kRa As Double
    Select Case Ra
        Case Is > 0.63: kRa = 1
        Case Is > 0.32: kRa = 1.2
        Case Else:      kRa = 1.4
    End Select

    Dim kTipPov As Double
    Select Case TipPoverkhnosti
        Case 1: kTipPov = 1     '1 - Наружная
        Case 2: kTipPov = 1.2   '2 - Внутренняя
        Case 3: kTipPov = 0.6   '3 - Торцовая
    End Select

    PolirovaniyeNaTokranomStankeShkurkoy = Koeffitsient * kRa * kTipPov * (a * L ^ X)

End Function

