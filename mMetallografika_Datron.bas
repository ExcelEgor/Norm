Attribute VB_Name = "mMetallografika_Datron"
'Option Explicit
'
'Public Function Metallografika_Marshrut_Datron(DlinaZagotovki As Double, ShirinaZagotovki As Double, TolshchinaZagotovki As Double, _
'    DlinaDetali As Double, ShirinaDetali As Double, _
'    KolVoDetaley As Double, _
'    tDatron As Double, tSles As Double, tKontrol As Double, tVyrez As Double) As Variant
''
''№ оп.   Наименование    ЕН  Тпз Тшт
''5   Комплектование  1   10  0
''10  Контроль внешнего вида  9   0   0,3
''15  Вырубка 9   10  6
''20  Нанесение рисунка   9   10  15
''25  Наполнение покрытия 9   5   10
''30  Контроль внешнего вида  9   0   0,4
''35  Упаковывание    9   5   3
''40  Гравировально-фрезерная 1   20  3,5
''45  Слесарная   1   5   3,5
''50  Контрольная 1   0   0,9
''55  Наполнение покрытия 9   5   5
''60  Сушка температурная 9   5   4
''65  Контроль внешнего вида  1   0   0,2
''70  Упаковывание    1   5   0,8
''75  Слесарно-сборочная  1   5   4
''80  Контроль внешнего вида  1   0   0,2
''85  Хранение    1   0   0
'
'    Const KOLVO_OPERATSIY = 17
'
'    Dim ArrParametry(1 To KOLVO_OPERATSIY)  As ParametryOperatsii
'    ArrParametry(1) = op005_Komplektovaniye
'
'
'    Dim i As Integer
'    Dim ArrTablitsa(1 To KOLVO_OPERATSIY, 1 To 6) As Variant
'    For i = 1 To KOLVO_OPERATSIY
'        ArrTablitsa(i, 1) = ArrParametry(i).Nomer
'        ArrTablitsa(i, 2) = ArrParametry(i).Naimenovaniye
'        ArrTablitsa(i, 3) = ArrParametry(i).Razryad
'        ArrTablitsa(i, 4) = ArrParametry(i).EN
'        ArrTablitsa(i, 5) = ArrParametry(i).Tpz
'        ArrTablitsa(i, 6) = CStr(OkruglenieTsht(ArrParametry(i).Tsht))
'    Next
'
'    NormyVremeni = ArrTablitsa
'
'
'End Function
'
'Private Function op005_Komplektovaniye() As ParametryOperatsii
'
'    With op005_Komplektovaniye
'        .Nomer = 5
'        .Naimenovaniye = "Комплектование"
'        .Razryad = 2
'        .EN = 1
'        .Tpz = 10
'        .Tsht = 0
'    End With
'
'End Function
'
'Private Function op010_Kontrol() As ParametryOperatsii
'
'    With op010_Kontrol
'        .Nomer = 10
'        .Naimenovaniye = "Контроль внешнего вида"
'        .Razryad = 3
'        .EN = 1
'        .Tpz = 10
'        .Tsht = 0
'    End With
'
'End Function
