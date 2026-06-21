Attribute VB_Name = "mKoordinat"
Option Explicit
'РАБОТЫ, ВЫПОЛНЯЕМЫЕ НА КООРДИНАТНО-РАСТОЧНЫХ СТАНКАХ
'УКРУПНЕННЫЕ ОДНОСТРОЧНЫЕ НОРМАТИВЫ ВРЕМЕНИ. Мелкосерийное производство. 1985 г.

'Material - Материал:
'1 - ЛС59-1
'2 - АЛ2
'3 - Сталь 35
'4 - 25Х13Н2
'5 - ВТ1-0
'6 - 12Х18Н9Т
'7 - ВНМ3
    
Public Enum EnumSposobSovmeshcheniyaOsi
    PoNoniusnoyIliMasshtabnoyLineyke = 1
    PoShtikhmasuIliIzmeritelnymPlitkam = 2
    BezSovmeshcheniyaOsey = 3
End Enum

Type Koordinat
    Material As Integer
    d As Double
    L As Double
    SposobSovmeshcheniyaOsi As EnumSposobSovmeshcheniyaOsi
    PredvaritelnayaTsentrovka As Boolean
    Glukhoye As Boolean
    NaklonnayaPloskost As Boolean
    TonkostennyaDetal As Boolean
    NaUdar As Boolean
    IT As Integer
    Ra As Double
    BieniyeaSoosnostKruglost As Double
    ZakalonnayaStal As Boolean
    KolVoDopProhodov As Double
End Type

Type ParameryUstanovki_Koordinat
    Massa As Double
    Sposob As Integer
    HarakterViverki As Integer
    Nezestk As Boolean
    TochnostViverki As Integer
    KolVoDopBoltov As Integer
    KolVoDopDomkratov As Integer
    KolVoUstDetaley As Integer
End Type

Public Const DMAX_SVERLENIYE_KOORDINAT As Integer = 30
Public Const DMAX_RASSVERLIVANIYE_KOORDINAT As Integer = 40
Public Const LMAX_RASSVERLIVANIYE_KOORDINAT As Integer = 150
Public Const DMAX_RASTACHIVANIYE_KOORDINAT As Integer = 250

Private Function MaterialKorrektnyy(Material As EnumMaterialy) As Boolean
    
    Select Case Material
    
        Case ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY, STAL_UGLERODISTAYA, STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY
            MaterialKorrektnyy = True
        
        Case Else
            MaterialKorrektnyy = False
            
    End Select

End Function

Function PoprKoef_Koordinat_TipProizv(TipProizv As Integer) As Double
    Select Case TipProizv
        Case 1: PoprKoef_Koordinat_TipProizv = 1.3  'единичное
        Case 2: PoprKoef_Koordinat_TipProizv = 1    'мелкосерийное
        Case 3: PoprKoef_Koordinat_TipProizv = 0.8  'среднесерийное
        Case 4: PoprKoef_Koordinat_TipProizv = 0.6  'массовое и крупносер.
    End Select
End Function

Function PoprKoef_Koordinat_TipStanka(TipStanka As Integer) As Double
    '1 одностоечный
    '2 двухстоечный
    PoprKoef_Koordinat_TipStanka = IIf(TipStanka = 2, 1.2, 1)
End Function

Function PoprKoef_Koordinat_Vozrast(Vozrast As Integer) As Double
    Select Case Vozrast
        Case 1: PoprKoef_Koordinat_Vozrast = 1      'до 10 лет
        Case 2: PoprKoef_Koordinat_Vozrast = 1.1    'свыше 10 до 20 лет
        Case 3: PoprKoef_Koordinat_Vozrast = 1.15   'свыше 20 лет
    End Select
End Function

Function PoprKoef_Koordinat_MatrialInstrument(MatrialInstrument As Integer) As Double
    Select Case MatrialInstrument
        Case 1: PoprKoef_Koordinat_MatrialInstrument = 1    'Т15К6, ВК8, Р18
        Case 2: PoprKoef_Koordinat_MatrialInstrument = 1.1  'Р6М5
        Case 3: PoprKoef_Koordinat_MatrialInstrument = 1.2  '9ХС, У10, У12
    End Select
End Function

Function MaxMassaUstanovki_Koordinat(Sposob As Integer, HarakterViverki As Integer) As Double
    Dim tb As Range
    Set tb = ThisWorkbook.Worksheets("Координат_2").ListObjects("tbUstID_Koordinat").DataBodyRange
    Dim i As Integer
    For i = 1 To tb.Rows.Count
        If tb(i, 1) = Sposob And tb(i, 2) = HarakterViverki Then
            MaxMassaUstanovki_Koordinat = tb(i, 3)
            Exit For
        End If
    Next
End Function

Function Ustanovka_Koordinat(Massa As Double, Sposob As Integer, HarakterViverki As Integer, Optional Nezestk As Boolean = False, _
    Optional TochnostViverki As Integer = 2, Optional KolVoDopBoltov As Integer = 0, Optional KolVoDopDomkratov As Integer = 0, Optional KolVoUstDetaley As Integer = 1) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_ustanovka")
    
    Dim i As Long
    Dim X As Double, Y As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 2) = Sposob And Normativ(i, 4) = HarakterViverki Then
            X = IIf(Massa <= 20, Normativ(i, 5), Normativ(i, 7))
            Y = IIf(Massa <= 20, Normativ(i, 6), Normativ(i, 8))
            Exit For
        End If
    Next
        
    Dim Vremya As Double
    Vremya = X * Massa ^ Y
        
    Select Case TochnostViverki
        Case 1: Vremya = Vremya + 2
        Case 2: Vremya = Vremya
        Case 3: Vremya = Vremya - 0.5
        Case 4: Vremya = Vremya - 2.5
    End Select
        
    Dim K As Double
    Select Case KolVoUstDetaley
        Case Is <= 1:   K = 1
        Case 2:         K = 1.2
        Case 3:         K = 1.4
        Case 4:         K = 1.6
        Case 5:         K = 1.8
        Case 6:         K = 2.1
        Case 7:         K = 2.4
        Case 8:         K = 2.6
        Case Is <= 10:  K = 3
        Case Is <= 15:  K = 3.5
        Case Else:      K = 4
    End Select
        
    Vremya = Vremya * IIf(Nezestk, 1.2, 1) * K
    Vremya = Vremya + (KolVoDopBoltov * 0.5) + (KolVoDopDomkratov * 4)
        
    Ustanovka_Koordinat = Vremya

End Function

Function MaksMassaUstanovki_Koordinat(SposobUstanovki As Integer, KharakterVyverki As Integer) As Integer

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_ustanovka")
    
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 2) = SposobUstanovki And Normativ(i, 4) = KharakterVyverki Then
            MaksMassaUstanovki_Koordinat = Normativ(i, 9)
            Exit For
        End If
    Next
    
End Function


Function Lmax_Rastachivaniye_Koordinat(d As Double) As Double
    
    If d <= DMAX_RASTACHIVANIYE_KOORDINAT Then
        Select Case d
            Case Is <= 3:   Lmax_Rastachivaniye_Koordinat = 20
            Case Is <= 5:   Lmax_Rastachivaniye_Koordinat = 55
            Case Is <= 7:   Lmax_Rastachivaniye_Koordinat = 75
            Case Is <= 10:  Lmax_Rastachivaniye_Koordinat = 100
            Case Is <= 20:  Lmax_Rastachivaniye_Koordinat = 100
            Case Else:      Lmax_Rastachivaniye_Koordinat = 200
        End Select
    End If
    
End Function
Function Lmax_Vytocheki(Material As EnumMaterialy, d As Double) As Integer

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_vytochki")
    
    Dim i As Long
    Dim Lmax As Integer
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
             Normativ(i, 2) >= d Then
        
            If Normativ(i, 3) > Lmax Then
                Lmax = Normativ(i, 3)
                If PoslednyayaStrokaVMassive(Normativ, i, 2) = True Then Exit For
            End If
            
        End If
    Next
   
    Lmax_Vytocheki = Lmax
    
End Function

Function Dmax_Vytocheki(Material As EnumMaterialy) As Integer
    Dmax_Vytocheki = IIf(Material = 7, 200, 250)
End Function

Function RastachivaniyeVytochek(Material As EnumMaterialy, ByVal d As Double, L As Double, _
    Optional IT As Integer = 11, Optional SovmeshcheniyeOsi As Boolean = True, Optional vNaklonnoyPloskosti As Boolean = False, _
    Optional TonkostenDetal As Boolean = False, Optional Udar As Boolean = False, Optional BieniyeaSoosnostKruglost As Double = 0, Optional ZakalonnayaStal As Boolean = False)
    
    If d > Dmax_Vytocheki(Material) Or PolozhitelnyyeChisla(d, L) = False Or IT < 6 Then Exit Function
    If L > Lmax_Vytocheki(Material, d) Then Exit Function
    
    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_vytochki")

    Dim i As Long
    Dim tKoordinat As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 3) >= L And _
            Normativ(i, 4) = SovmeshcheniyeOsi And _
            Normativ(i, 7) <= IT Then
            tKoordinat = Normativ(i, 8) * RaschotKoeffitsientaPriObrabotkeOtverstiy(False, vNaklonnoyPloskosti, TonkostenDetal, Udar, BieniyeaSoosnostKruglost, ZakalonnayaStal)
            Exit For
        End If
    Next
    
    Dim tSles As Double
    tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, 1, False)

    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(d, L, IT)

    Dim Normy(1 To 3) As Double
    Normy(1) = tKoordinat
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    RastachivaniyeVytochek = Normy
    
End Function

Function Rastachivaniye_Koordinat(Material As EnumMaterialy, ByVal d As Double, L As Double, Sposob As EnumSposobSovmeshcheniyaOsi, IT As Integer, Ra As Double, _
    Optional Glukhoye As Boolean = False, Optional vNaklonnoyPloskosti As Boolean = False, _
    Optional TonkostenDetal As Boolean = False, Optional Udar As Boolean = False, _
    Optional BieniyeaSoosnostKruglost As Double = 0, Optional ZakalonnayaStal As Boolean = False, _
    Optional KolVoDopProhodov As Double = 0) As Variant
    
    If d > DMAX_RASTACHIVANIYE_KOORDINAT Or IT < 6 Or PolozhitelnyyeChisla(d, L) = False Or MaterialKorrektnyy(Material) = False Then Exit Function
    If L > Lmax_Rastachivaniye_Koordinat(d) Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_rastachivaniye")
    
    Dim tOsnovnoe As Double
    tOsnovnoe = Rastachivaniye_Koordinat_tNsh(Normativ, Material, d, L, IT, Ra, Sposob)

    Dim tDopProhod As Double
    tDopProhod = KolVoDopProhodov * Rastachivaniye_Koordinat_tNsh(Normativ, Material, d, L, 11, 5, BezSovmeshcheniyaOsey)

    Dim Koefficient As Double
    Koefficient = RaschotKoeffitsientaPriObrabotkeOtverstiy(Glukhoye, vNaklonnoyPloskosti, TonkostenDetal, Udar, BieniyeaSoosnostKruglost, ZakalonnayaStal)
    
    Dim tKoordinat As Double
    tKoordinat = Koefficient * (tOsnovnoe + tDopProhod)
    
    Dim tSles As Double
    tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, IIf(Glukhoye, 1, 2), False)

    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(d, L, IT)

    Dim Normy(1 To 3) As Double
    Normy(1) = tKoordinat
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    Rastachivaniye_Koordinat = Normy
    
End Function

Private Function Rastachivaniye_Koordinat_tNsh(Normativ, Material As EnumMaterialy, d As Double, L As Double, IT As Integer, Ra As Double, Sposob As EnumSposobSovmeshcheniyaOsi) As Double
    
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And Normativ(i, 3) >= L And _
            Normativ(i, 4) <= IT And Normativ(i, 5) <= Ra And Normativ(i, 7) = Sposob Then
            Rastachivaniye_Koordinat_tNsh = Normativ(i, 8)
            Exit For
        End If
    Next

End Function

Private Function RaschotKoeffitsientaPriObrabotkeOtverstiy(Optional Glukhoye As Boolean = False, Optional vNaklonnoyPloskosti As Boolean = False, _
    Optional TonkostenDetal As Boolean = False, Optional Udar As Boolean = False, _
    Optional BieniyeaSoosnostKruglost As Double = 0, Optional ZakalonnayaStal As Boolean = False) As Double
    
    Dim K1 As Double:   K1 = IIf(Glukhoye, 1.1, 1)
    Dim K2 As Double:   K2 = IIf(vNaklonnoyPloskosti, 1.1, 1)
    Dim K3 As Double:   K3 = IIf(TonkostenDetal, 1.3, 1)
    Dim K4 As Double:   K4 = IIf(Udar, 1.15, 1)
    Dim K5 As Double:   K5 = IIf(ZakalonnayaStal, 1.15, 1)
    
    Dim K6 As Double
    If BieniyeaSoosnostKruglost = 0 Then
        K6 = 1
    Else
        
        Select Case BieniyeaSoosnostKruglost
            Case Is <= 0.002:   K6 = 1.35
            Case Is <= 0.005:   K6 = 1.25
            Case Else:          K6 = 1.2
        End Select
    End If
    
    RaschotKoeffitsientaPriObrabotkeOtverstiy = WorksheetFunction.Product(K1, K2, K3, K4, K5, K6)

End Function

Private Function NomerStolbcaPoMaterialu(Material As Integer) As Integer
    Dim KolVoMaterialov As Integer
    KolVoMaterialov = 7
    NomerStolbcaPoMaterialu = (KolVoMaterialov - Material) + 1
End Function

Function Sverlenie_Koordinat(ByVal Material As EnumMaterialy, d As Double, L As Double, Sposob As EnumSposobSovmeshcheniyaOsi, _
    Optional PredvaritelnayaTsentrovka As Boolean = False, Optional Glukhoye As Boolean = False, _
    Optional vNaklonnoyPloskosti As Boolean = False, Optional TonkostenDetal As Boolean = False, Optional Udar As Boolean = False) As Variant
    
    If PolozhitelnyyeChisla(d, L) = False Or MaterialKorrektnyy(Material) = False Then Exit Function
    If d > DMAX_SVERLENIYE_KOORDINAT Then Exit Function
    If L > Lmax_Sverlenie_Koordinat(d) Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_sverleniye")

    Dim tKoordinat As Double
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And Normativ(i, 3) >= L And _
            Normativ(i, 4) = Sposob And Normativ(i, 5) = PredvaritelnayaTsentrovka Then
            tKoordinat = Normativ(i, 6) * RaschotKoeffitsientaPriObrabotkeOtverstiy(Glukhoye, vNaklonnoyPloskosti, TonkostenDetal, Udar)
            Exit For
        End If
    Next
    
    Dim tSles As Double
    tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, IIf(Glukhoye, 1, 2), False)

    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(d, L)

    Dim Normy(1 To 3) As Double
    Normy(1) = tKoordinat
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    Sverlenie_Koordinat = Normy
    
End Function

Function Lmax_Sverlenie_Koordinat(d As Double) As Integer

    If d <= DMAX_SVERLENIYE_KOORDINAT Then
        Select Case d
            Case Is <= 4:   Lmax_Sverlenie_Koordinat = 40
            Case Is <= 6:   Lmax_Sverlenie_Koordinat = 60
            Case Is <= 8:   Lmax_Sverlenie_Koordinat = 80
            Case Is <= 10:  Lmax_Sverlenie_Koordinat = 100
            Case Is <= 12:  Lmax_Sverlenie_Koordinat = 125
            Case Else:      Lmax_Sverlenie_Koordinat = 150
        End Select
    End If

End Function

Function Rassverlivaniye_Koordinat(Material As EnumMaterialy, d As Double, L As Double, _
    Sposob As EnumSposobSovmeshcheniyaOsi, _
    Optional Glukhoye As Boolean = False, _
    Optional vNaklonnoyPloskosti As Boolean = False, Optional TonkostenDetal As Boolean = False, Optional Udar As Boolean = False) As Variant
    
    If d > DMAX_RASSVERLIVANIYE_KOORDINAT Or _
        L > LMAX_RASSVERLIVANIYE_KOORDINAT Or _
        PolozhitelnyyeChisla(d, L) = False Or _
        MaterialKorrektnyy(Material) = False Then Exit Function

    Dim Normativ
    Normativ = ZagruzitNormativ("koordinat_rassverlivaniye")
    
    Dim tKoordinat As Double
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = IIf(Material = MEDNYYE_SPLAVY, ALUMINIYEVYYE_SPLAVY, Material) And _
            Normativ(i, 2) >= d And _
            Normativ(i, 5) >= L And _
            Normativ(i, 6) = Sposob Then
            tKoordinat = Normativ(i, 7) * RaschotKoeffitsientaPriObrabotkeOtverstiy(Glukhoye, vNaklonnoyPloskosti, TonkostenDetal, Udar)
            Exit For
        End If
    Next
    
    Dim tSles As Double
    tSles = ZachistkaZausencevSOtverstiiVruchnuyu(CInt(Material), d, 2, IIf(Glukhoye, 1, 2), False)

    Dim tKontrol As Double
    tKontrol = KontrolOtverstiy(d, L)

    Dim Normy(1 To 3) As Double
    Normy(1) = tKoordinat
    Normy(2) = tSles
    Normy(3) = tKontrol
    
    Rassverlivaniye_Koordinat = Normy
    
End Function


Function NametkaOtverstiy_Koordinat(Sposob As EnumSposobSovmeshcheniyaOsi, Optional NaklonnayaPloskost As Boolean = False) As Double
    'Сверление отверстий. Карта 5

    'SovmeshcheniyeOsi - Способ совмещения оси:
    '1 - По нониусной или масштабной линейке
    '2 - По штихмасу или измерительными плитками и с помощью контрольного валика
    '3 - Без совмещения осей
    
    Dim tNsh As Double
    Select Case Sposob
        Case 1: tNsh = 1
        Case 2: tNsh = 1.8
        Case 3: tNsh = 0.3
    End Select
    
    Dim K As Double
    K = IIf(NaklonnayaPloskost, 1.1, 1)
    
    NametkaOtverstiy_Koordinat = tNsh * K
    
End Function
