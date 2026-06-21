VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetallografika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Технологические константы и коэффициенты ---
Private Const KOEF_STASA As Double = 2.4
Private Const LIMIT_TOLSHCHINY_PEREMYCHKI As Double = 0.8
Private Const SHAG_PEREMYCHKI As Double = 50
Private Const RAZRYAD_KONTROLYA As Long = 3

' --- Внутреннее состояние объекта ---
Private pDlinaZagotovki As Double
Private pShirinaZagotovki As Double
Private pTolshchinaZagotovki As Double
Private pDlinaDetali As Double
Private pShirinaDetali As Double
Private pKolVoDetaley As Long
Private ptDatron As Double
Private ptSles As Double
Private ptKontrol As Double
Private ptVyrez As Double
Private kTipProizv As Double

Public Sub Init( _
    ByVal DlinaZagotovki As Double, _
    ByVal ShirinaZagotovki As Double, _
    ByVal TolshchinaZagotovki As Double, _
    ByVal DlinaDetali As Double, _
    ByVal ShirinaDetali As Double, _
    ByVal KolVoDetaley As Long, _
    ByVal tDatron As Double, _
    ByVal tSles As Double, _
    ByVal tKontrol As Double, _
    ByVal tVyrez As Double, _
    Optional ByVal TipProizv As eTipyProizvodstva = Edinichnyy)
    
    pDlinaZagotovki = DlinaZagotovki
    pShirinaZagotovki = ShirinaZagotovki
    pTolshchinaZagotovki = TolshchinaZagotovki
    pDlinaDetali = DlinaDetali
    pShirinaDetali = ShirinaDetali
    ptDatron = tDatron
    ptSles = tSles
    ptKontrol = tKontrol
    ptVyrez = tVyrez
    
    If KolVoDetaley = 0 Then
        pKolVoDetaley = 1
    Else
        pKolVoDetaley = KolVoDetaley
    End If
    
    kTipProizv = KoefTipProizv(TipProizv)
End Sub

Public Function Marshrut() As Variant

    If ptDatron > 0 Then
        Marshrut = Marshrut_Datron
        Exit Function
    End If

    If Not PolozhitelnyyeChisla(pDlinaZagotovki, pShirinaZagotovki, pTolshchinaZagotovki, pDlinaDetali, pShirinaDetali, pKolVoDetaley) Then Exit Function

    Const KOLVO_OPERATSIY As Long = 16
    Dim ArrParametry(1 To KOLVO_OPERATSIY) As ParametryOperatsii
    
    ' Заполнение инвариантной (повторяющейся) части маршрута
    ZapolnitBazovyyMarshrut ArrParametry
    
    ' Специфические операции базового техпроцесса
    ArrParametry(7) = op035
    ArrParametry(8) = op040
    ArrParametry(9) = op045
    ArrParametry(10) = op050
    ArrParametry(11) = op055
    ArrParametry(12) = KontrolVneshnegoVidaDetali(60)
    ArrParametry(13) = op065
    ArrParametry(14) = PrikleivaniyeLipkoyLentyNaPlankuSNadpisyu(False)
    ArrParametry(15) = KontrolVneshnegoVidaDetali(75)
    ArrParametry(16) = Khraneniye(80)
    
    Marshrut = SformirovatMassiv(KOLVO_OPERATSIY, ArrParametry)
End Function

Private Function Marshrut_Datron() As Variant
    If Not PolozhitelnyyeChisla(pDlinaZagotovki, pShirinaZagotovki, pTolshchinaZagotovki, pDlinaDetali, pShirinaDetali, pKolVoDetaley, _
                               ptDatron, ptKontrol, ptSles, ptVyrez) Then Exit Function

    Const KOLVO_OPERATSIY As Long = 17
    Dim ArrParametry(1 To KOLVO_OPERATSIY) As ParametryOperatsii
    
    ' Заполнение инвариантной части маршрута
    ZapolnitBazovyyMarshrut ArrParametry
    
    ' Специфические операции техпроцесса Datron
    ArrParametry(7) = op035_Datron
    ArrParametry(8) = op040_Datron
    ArrParametry(9) = op045_Datron
    ArrParametry(10) = op050_Datron
    ArrParametry(11) = op050
    ArrParametry(12) = op055
    ArrParametry(13) = KontrolVneshnegoVidaDetali(65)
    ArrParametry(14) = op070_Datron
    ArrParametry(15) = PrikleivaniyeLipkoyLentyNaPlankuSNadpisyu(True)
    ArrParametry(16) = KontrolVneshnegoVidaDetali(80)
    ArrParametry(17) = Khraneniye(85)
    
    Marshrut_Datron = SformirovatMassiv(KOLVO_OPERATSIY, ArrParametry)
End Function

Private Sub ZapolnitBazovyyMarshrut(ByRef Arr() As ParametryOperatsii)
    Arr(1) = op005
    Arr(2) = op010
    Arr(3) = op015
    Arr(4) = op020
    Arr(5) = op025
    Arr(6) = op030
End Sub

Private Function SformirovatMassiv(ByVal KolVoOperatsiy As Long, ByRef ArrParametry() As ParametryOperatsii) As Variant
    Dim i As Long
    Dim ArrTablitsa() As Variant
    ReDim ArrTablitsa(1 To KolVoOperatsiy, 1 To 6)
    
    For i = 1 To KolVoOperatsiy
        ArrTablitsa(i, 1) = ArrParametry(i).Nomer
        ArrTablitsa(i, 2) = ArrParametry(i).Naimenovaniye
        ArrTablitsa(i, 3) = ArrParametry(i).Razryad
        ArrTablitsa(i, 4) = ArrParametry(i).EN
        ArrTablitsa(i, 5) = ArrParametry(i).Tpz
        ArrTablitsa(i, 6) = OkruglenieTsht(ArrParametry(i).Tsht)
    Next i
    
    SformirovatMassiv = ArrTablitsa
End Function

Private Function op005() As ParametryOperatsii
    With op005
        .Nomer = 5
        .Naimenovaniye = "Комплектование"
        .Razryad = 2
        .EN = 1
        .Tpz = 10
        .Tsht = 0
    End With
End Function

Private Function op010() As ParametryOperatsii
    With op010
        .Nomer = 10
        .Naimenovaniye = "Контроль внешнего вида"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 0
        .Tsht = kTipProizv * VisualnyyKontrol(2 * (pDlinaZagotovki + pShirinaZagotovki), pTolshchinaZagotovki)
    End With
End Function

Private Function op015() As ParametryOperatsii
    With op015
        .Nomer = 15
        .Naimenovaniye = "Вырубка"
        .Razryad = 2
        .EN = pKolVoDetaley
        .Tpz = 10
        .Tsht = 2
    End With
End Function

Private Function op020() As ParametryOperatsii
    With op020
        .Nomer = 20
        .Naimenovaniye = "Нанесение рисунка"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 10
        .Tsht = 10
    End With
End Function

Private Function op025() As ParametryOperatsii
    With op025
        .Nomer = 25
        .Naimenovaniye = "Наполнение покрытия"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 5
        .Tsht = 10
    End With
End Function

Private Function op030() As ParametryOperatsii
    With op030
        .Nomer = 30
        .Naimenovaniye = "Контроль внешнего вида"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 0
        .Tsht = kTipProizv * IzmerShtangenCircul_CPL(pDlinaZagotovki, pShirinaZagotovki)
    End With
End Function

Private Function op035() As ParametryOperatsii
    With op035
        .Nomer = 35
        .Naimenovaniye = "Вырубка"
        .Razryad = 2
        .EN = 1
        .Tpz = 10
        .Tsht = KOEF_STASA * OtrezkaNaGilotinNozhnitsZaDvaReza(ALUMINIYEVYYE_SPLAVY, pDlinaDetali, pShirinaDetali, pTolshchinaZagotovki)
    End With
End Function

Private Function op035_Datron() As ParametryOperatsii
    With op035_Datron
        .Nomer = 35
        .Naimenovaniye = "Упаковывание"
        .Razryad = 2
        .EN = pKolVoDetaley
        .Tpz = 5
        .Tsht = 3
    End With
End Function

Private Function op040_Datron() As ParametryOperatsii
    With op040_Datron
        .Nomer = 40
        .Naimenovaniye = "Гравировально-фрезерная"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = 1
        .Tpz = 20
        .Tsht = ptDatron
    End With
End Function

Private Function op040() As ParametryOperatsii
    With op040
        .Nomer = 40
        .Naimenovaniye = "Слесарная"
        .Razryad = 2
        .EN = 1
        .Tpz = 5
        .Tsht = kTipProizv * SnyatiyeZausentsevList(pDlinaDetali, pShirinaDetali, pTolshchinaZagotovki, ALUMINIYEVYYE_SPLAVY)
    End With
End Function

Private Function op045() As ParametryOperatsii
    With op045
        .Nomer = 45
        .Naimenovaniye = "Контрольная"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = 1
        .Tpz = 0
        .Tsht = kTipProizv * IzmerShtangenCircul_CPL(pDlinaDetali, pShirinaDetali)
    End With
End Function

Private Function op045_Datron() As ParametryOperatsii
    With op045_Datron
        .Nomer = 45
        .Naimenovaniye = "Слесарная"
        .Razryad = 2
        .EN = 1
        .Tpz = 5
        
        Dim tPeremychki As Double
        If pTolshchinaZagotovki <= LIMIT_TOLSHCHINY_PEREMYCHKI Then
            tPeremychki = 0.55 * (2 * (pDlinaDetali + pShirinaDetali) / SHAG_PEREMYCHKI)
        End If
        
        Dim tUpakovka As Double
        tUpakovka = UpakovyvaniyePramougolnika(pDlinaDetali, pShirinaDetali, pTolshchinaZagotovki)
        
        .Tsht = kTipProizv * (ptSles + tUpakovka) + tPeremychki
    End With
End Function

Private Function op050() As ParametryOperatsii
    With op050
        .Nomer = 50
        .Naimenovaniye = "Наполнение покрытия"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 5
        .Tsht = 5
    End With
End Function

Private Function op050_Datron() As ParametryOperatsii
    With op050_Datron
        .Nomer = 50
        .Naimenovaniye = "Контрольная"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = 1
        .Tpz = 0
        .Tsht = kTipProizv * (ptKontrol + IzmerShtangenCircul_CPL(pDlinaDetali, pShirinaDetali))
    End With
End Function

Private Function op055() As ParametryOperatsii
    With op055
        .Nomer = 55
        .Naimenovaniye = "Сушка температурная"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = pKolVoDetaley
        .Tpz = 5
        .Tsht = 4
    End With
End Function

Private Function KontrolVneshnegoVidaDetali(ByVal NomerOperatsii As Integer) As ParametryOperatsii
    With KontrolVneshnegoVidaDetali
        .Nomer = NomerOperatsii
        .Naimenovaniye = "Контроль внешнего вида"
        .Razryad = RAZRYAD_KONTROLYA
        .EN = 1
        .Tpz = 0
        .Tsht = kTipProizv * VisualnyyKontrol(2 * (pDlinaDetali + pShirinaDetali), pTolshchinaZagotovki, 3)
    End With
End Function

Private Function op065() As ParametryOperatsii
    With op065
        .Nomer = 65
        .Naimenovaniye = "Упаковывание"
        .Razryad = 2
        .EN = pKolVoDetaley
        .Tpz = 5
        .Tsht = kTipProizv * pKolVoDetaley * (0.467 / (pKolVoDetaley ^ 0.5))
    End With
End Function

Private Function PrikleivaniyeLipkoyLentyNaPlankuSNadpisyu(ByVal IsDatron As Boolean) As ParametryOperatsii
    With PrikleivaniyeLipkoyLentyNaPlankuSNadpisyu
        If IsDatron Then .Nomer = 75 Else .Nomer = 70
        
        .Naimenovaniye = "Слесарно-сборочная"
        .Razryad = 2
        .EN = 1
        .Tpz = 5
        
        Dim tObezzhirivaniye As Double
        tObezzhirivaniye = 2 * Galvanika("Обезжиривание", (pDlinaDetali * pShirinaDetali) / 1000000)
        
        Dim tVyrezki As Double
        tVyrezki = VyrezaniyeProkladokVruchnuyu(1, 1, 2 * (pDlinaDetali + pShirinaDetali), 1, 1, 1)
    
        .Tsht = kTipProizv * (tObezzhirivaniye + tVyrezki + ptVyrez)
    End With
End Function

Private Function op070_Datron() As ParametryOperatsii
    With op070_Datron
        .Nomer = 70
        .Naimenovaniye = "Упаковывание"
        .Razryad = 2
        .EN = 1
        .Tpz = 5
        .Tsht = kTipProizv * UpakovyvaniyePramougolnika(pDlinaDetali, pShirinaDetali, pTolshchinaZagotovki)
    End With
End Function

Private Function Khraneniye(ByVal NomerOperatsii As Integer) As ParametryOperatsii
    With Khraneniye
        .Nomer = NomerOperatsii
        .Naimenovaniye = "Хранение"
        .Razryad = 2
        .EN = 1
        .Tpz = 0
        .Tsht = 0
    End With
End Function

