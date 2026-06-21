VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTermograf300_400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DLINA_ZAGOTOVKI As Integer = 400
Private Const SHIRINA_ZAGOTOVKI As Integer = 300
Private Const TOLSHCHINA_ZAGOTOVKI As Integer = 1.5
Private MATERIAL_ZAGOTOVKI As EnumMaterialy

Private pMassaZagotovki As Double
Private pKolVoDet As Integer
Private ptMin As Double
Private ptSek As Double
Private pArrDiametry() As Double
Private pArrKolVo() As Double
Private pArrZenkovaniye() As Boolean
Private pDlinaPlanki As Double
Private pShirnaPlanki As Double

Public Property Let KolVoDet(KolVoDetaleyVZagotovke As Integer): pKolVoDet = KolVoDetaleyVZagotovke: End Property

Public Property Let tMin(tMin_LazerPress As Double): ptMin = tMin_LazerPress: End Property

Public Property Let tSek(tSek_LazerPress As Double): ptSek = tSek_LazerPress: End Property

Public Property Let ArrDiametry(ArrDiametry As Variant): pArrDiametry = ArrDiametry: End Property

Public Property Let ArrKolVo(ArrKolVo As Variant): pArrKolVo = ArrKolVo: End Property

Public Property Let ArrZenkovaniye(ArrZenkovaniye As Variant): pArrZenkovaniye = ArrZenkovaniye: End Property

Public Property Let DlinaPlanki(DlinaPlanki As Double): pDlinaPlanki = DlinaPlanki: End Property

Public Property Let ShirnaPlanki(ShirnaPlanki As Double): pShirnaPlanki = ShirnaPlanki: End Property

Public Function NormyVremeni() As Variant

    If pKolVoDet <= 0 Or pDlinaPlanki <= 0 Or pShirnaPlanki <= 0 Then Exit Function
    If ptMin <= 0 And ptSek <= 0 Then Exit Function
    
    MATERIAL_ZAGOTOVKI = ALUMINIYEVYYE_SPLAVY
    pMassaZagotovki = MassaLista(DLINA_ZAGOTOVKI, SHIRINA_ZAGOTOVKI, TOLSHCHINA_ZAGOTOVKI, MATERIAL_ZAGOTOVKI)

    Const KOLVO_OPERATSIY As Integer = 16
    
    Dim ArrParametry(1 To KOLVO_OPERATSIY)  As ParametryOperatsii
    ArrParametry(1) = op005_PodgotovkaDlyuaProbivochnoVyreznoy
    ArrParametry(2) = op010_Kontrolnaya
    ArrParametry(3) = op015_ProbivochnoVyreznaya
    ArrParametry(4) = op020_Slesarnaya
    ArrParametry(5) = op025_Kontrolnaya
    ArrParametry(6) = op030_Galvanika
    ArrParametry(7) = op035_Kontrol
    ArrParametry(8) = op040_NaneseniyeRisunka
    ArrParametry(9) = op045_Pressovaniye
    ArrParametry(10) = op050_Kontrol
    ArrParametry(11) = op055_NapolneniyePokrytiya
    ArrParametry(12) = op060_Slesarnaya
    ArrParametry(13) = op065_Kontrol
    ArrParametry(14) = op070_NapolneniyePokrytiya
    ArrParametry(15) = op075_Kontrol
    ArrParametry(16) = op080_Khraneniye
    
    Dim i As Integer
    Dim ArrTablitsa(1 To KOLVO_OPERATSIY, 1 To 6) As Variant
    For i = 1 To KOLVO_OPERATSIY
        ArrTablitsa(i, 1) = ArrParametry(i).Nomer
        ArrTablitsa(i, 2) = ArrParametry(i).Naimenovaniye
        ArrTablitsa(i, 3) = ArrParametry(i).Razryad
        ArrTablitsa(i, 4) = ArrParametry(i).EN
        ArrTablitsa(i, 5) = ArrParametry(i).Tpz
        ArrTablitsa(i, 6) = CStr(OkruglenieTsht(ArrParametry(i).Tsht))
    Next
    
    NormyVremeni = ArrTablitsa
    
End Function

Private Function op005_PodgotovkaDlyuaProbivochnoVyreznoy() As ParametryOperatsii
    With op005_PodgotovkaDlyuaProbivochnoVyreznoy
        .Nomer = 5
        .Naimenovaniye = "Подгот. для пробивочно-вырезного пресса"
        .Razryad = 2
        .EN = 1
        .Tpz = 10
        .Tsht = 0
    End With
End Function

Private Function op010_Kontrolnaya() As ParametryOperatsii
    With op010_Kontrolnaya
        .Nomer = 10
        .Naimenovaniye = "Контрольная"
        .Razryad = 3
        .EN = 1
        .Tpz = 2
        .Tsht = 0
    End With
End Function

Private Function op015_ProbivochnoVyreznaya() As ParametryOperatsii
    
    Dim Tsht As Double
    Tsht = TshtLaserPress(2500, 1200, 1.5, 300, 400, ptMin, ptSek, MATERIAL_ZAGOTOVKI)
    
    With op015_ProbivochnoVyreznaya
        .Nomer = 15
        .Naimenovaniye = "Пробивочно -вырезная"
        .Razryad = 3
        .EN = 1
        .Tpz = 25
        .Tsht = Tsht
    End With
End Function

Private Function op020_Slesarnaya() As ParametryOperatsii
    
    Dim Tsht As Double
    Dim EstZenkovaniye As Boolean
    Dim tZausGabarity As Double, tOtv As Double, tPravka As Double, tObertyvaniye As Double
    
    tZausGabarity = SnyatiyeZausentsevList(DLINA_ZAGOTOVKI, SHIRINA_ZAGOTOVKI, TOLSHCHINA_ZAGOTOVKI, MATERIAL_ZAGOTOVKI)
    tOtv = RaschotOtverstiy(EstZenkovaniye)
    tPravka = PravkaPloskikhDetaley(TOLSHCHINA_ZAGOTOVKI, DLINA_ZAGOTOVKI, SHIRINA_ZAGOTOVKI, MATERIAL_ZAGOTOVKI)
    tObertyvaniye = ObertyvaniyeBumagoy_PloskieDetalie(SHIRINA_ZAGOTOVKI, DLINA_ZAGOTOVKI)
    
    Tsht = 1.3 * (tZausGabarity + tOtv + tPravka + tObertyvaniye)

    With op020_Slesarnaya
        .Nomer = 20
        .Naimenovaniye = "Слесарная"
        .Razryad = 2
        .EN = pKolVoDet
        .Tpz = IIf(EstZenkovaniye, 10, 5)
        .Tsht = Tsht
    End With
    
End Function

Private Function RaschotOtverstiy(ByRef EstZenkovaniye As Boolean) As Double
    
    Dim i As Integer
    Dim tZausOtv As Double, tUst As Double, tZenk As Double
    Dim d As Double, KolVo As Double
    
    For i = LBound(pArrDiametry) To UBound(pArrDiametry)
    
        d = CDbl(pArrDiametry(i))
        KolVo = CDbl(pArrKolVo(i))
        EstZenkovaniye = CBool(pArrZenkovaniye(i))

        tZausOtv = tZausOtv + ZachistkaZausentsev_IzTablitsy(MATERIAL_ZAGOTOVKI, 0, d, 0, 0, 0, KolVo, 2)
        
        If EstZenkovaniye Then
            tZenk = tZenk + ObrabotkaOtverstiySleasar(ZenkovaniyeFasok, MATERIAL_ZAGOTOVKI, d, KolVo, 0, 0, False)(1)
            If tUst = 0 Then tUst = Ustanov_NaStoleKreplBoltPlank(pMassaZagotovki)
        End If
        
    Next
    
    RaschotOtverstiy = tUst + pKolVoDet * (tZausOtv + tZenk)
    
End Function

Private Function op025_Kontrolnaya() As ParametryOperatsii
    
    Dim i As Integer
    Dim tKontrolOtv As Double
    Dim d As Double, KolVo As Double
    For i = LBound(pArrDiametry) To UBound(pArrDiametry)
        d = CDbl(pArrDiametry(i))
        KolVo = CDbl(pArrKolVo(i))
        tKontrolOtv = tKontrolOtv + KolVo * KontrolOtverstiy(d, TOLSHCHINA_ZAGOTOVKI, 11, False)
    Next
    
    Dim Tsht As Double
    Tsht = 1.3 * KoefPartiynosti_Kontrol(pKolVoDet) * (pKolVoDet * tKontrolOtv + IzmerShtangenCircul_CPL(DLINA_ZAGOTOVKI, SHIRINA_ZAGOTOVKI))
    
    With op025_Kontrolnaya
        .Nomer = 25
        .Naimenovaniye = "Котрольная"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 0
        .Tsht = Tsht
    End With
    
End Function

Private Function op030_Galvanika() As ParametryOperatsii
    With op030_Galvanika
        .Nomer = 30
        .Naimenovaniye = "Окисление анодное"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 5
        .Tsht = 25
    End With
End Function

Private Function op035_Kontrol() As ParametryOperatsii
    With op035_Kontrol
        .Nomer = 35
        .Naimenovaniye = "Контроль внешнего вида"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 0
        .Tsht = 1.3 * VisualnyyKontrol(SHIRINA_ZAGOTOVKI, DLINA_ZAGOTOVKI)
    End With
End Function

Private Function op040_NaneseniyeRisunka() As ParametryOperatsii
    With op040_NaneseniyeRisunka
        .Nomer = 40
        .Naimenovaniye = "Нанесение рисунка"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 20
        .Tsht = 5
    End With
End Function

Private Function op045_Pressovaniye() As ParametryOperatsii
    With op045_Pressovaniye
        .Nomer = 45
        .Naimenovaniye = "Прессование"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 5
        .Tsht = 2 + pKolVoDet
    End With
End Function

Private Function op050_Kontrol() As ParametryOperatsii
    With op050_Kontrol
        .Nomer = 50
        .Naimenovaniye = "Контрольная"
        .Razryad = 3
        .EN = pKolVoDet
        .Tpz = 0
        .Tsht = pKolVoDet * 1.3 * VisualnyyKontrol(pShirnaPlanki, pDlinaPlanki)
    End With
End Function

Private Function op055_NapolneniyePokrytiya() As ParametryOperatsii
    With op055_NapolneniyePokrytiya
        .Nomer = 55
        .Naimenovaniye = "Наполнение покрытия"
        .Razryad = 3
        .EN = IIf(pKolVoDet > 9, pKolVoDet, 1)
        .Tpz = 5
        .Tsht = IIf(pKolVoDet > 9, 20, 2)
    End With
End Function

Private Function op060_Slesarnaya() As ParametryOperatsii
    
    Dim Tsht As Double
    Dim tPeremychki As Double, tZausGabarity As Double, tPravka As Double, tObertyvaniye As Double
    
    tPeremychki = 4 * T_PEREMYCHKA
    tZausGabarity = SnyatiyeZausentsevList(pDlinaPlanki, pShirnaPlanki, TOLSHCHINA_ZAGOTOVKI, MATERIAL_ZAGOTOVKI)
    tPravka = PravkaPloskikhDetaley(TOLSHCHINA_ZAGOTOVKI, pDlinaPlanki, pShirnaPlanki, MATERIAL_ZAGOTOVKI)
    tObertyvaniye = ObertyvaniyeBumagoy_PloskieDetalie(pShirnaPlanki, pDlinaPlanki)
    
    Tsht = tPeremychki + 1.3 * (tZausGabarity + tPravka + tObertyvaniye)

    With op060_Slesarnaya
        .Nomer = 60:  .Naimenovaniye = "Слесарная": .Razryad = 2: .EN = 1: .Tpz = 5: .Tsht = Tsht
    End With
    
End Function

Private Function op065_Kontrol() As ParametryOperatsii
    With op065_Kontrol
        .Nomer = 65
        .Naimenovaniye = "Контрольная"
        .Razryad = 3
        .EN = 1
        .Tpz = 0
        .Tsht = 1.3 * KoefPartiynosti_Kontrol(pKolVoDet) * IzmerShtangenCircul_CPL(pDlinaPlanki, pShirnaPlanki)
    End With
End Function

Private Function op070_NapolneniyePokrytiya() As ParametryOperatsii
    With op070_NapolneniyePokrytiya
        .Nomer = 70
        .Naimenovaniye = "Наполнение покрытия"
        .Razryad = 3
        .EN = IIf(pKolVoDet > 9, pKolVoDet, 1)
        .Tpz = 5
        .Tsht = IIf(pKolVoDet > 9, 30, 3)
    End With
End Function

Private Function op075_Kontrol() As ParametryOperatsii
    With op075_Kontrol
        .Nomer = 70
        .Naimenovaniye = "Контрольная"
        .Razryad = 3
        .EN = 1
        .Tpz = 0
        .Tsht = 1.3 * KoefPartiynosti_Kontrol(pKolVoDet) * VisualnyyKontrol(pShirnaPlanki, pDlinaPlanki)
    End With
End Function

Private Function op080_Khraneniye() As ParametryOperatsii
    With op080_Khraneniye
        .Nomer = 80
        .Naimenovaniye = "Хранение"
        .Razryad = 2
        .EN = 1
        .Tpz = 0
        .Tsht = 0
    End With
End Function



