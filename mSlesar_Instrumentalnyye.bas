Attribute VB_Name = "mSlesar_Instrumentalnyye"
'ОБЩЕМАШИНОСТРОИТЕЛЬНЫЕ НОРМАТИВЫ ВРЕМЕНИ НА СЛЕСАРНО-ИНСТРУМЕНТАЛЬНЫЕ РАБОТЫ, ВЫПОЛНЯЕМЫЕ НА СТАНКАХИ ВРУЧНУЮ. МОСКВА 1990

Option Explicit

Private Const MAKS_MASSA_UST As Integer = 20

Public Enum InstrumentyZachistki
    NAPILNIK_SHABER = 1
    PNEVM_MASHINA = 2
    SVERLO = 3
    NADFIL = 4
    ABRAZIVNYY_BRUSOK = 5
End Enum

Function RezkaPravkaProvoloki(Diametr As Double, Dlina As Double, Material As EnumMaterialy) As Double
    'Карта 2. Резка ручными рычажными ножницами + Карта 7. Правка деталей
    
    If Diametr <= 0 Or Dlina <= 0 Then Exit Function
    
    Dim kMaterialRezka As Double, kMaterialPravka As Double
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        kMaterialRezka = 0.7
        kMaterialPravka = kMaterialRezka
    
    ElseIf Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Then
        kMaterialRezka = 1
        kMaterialPravka = kMaterialRezka
    
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterialRezka = 1.2
        kMaterialPravka = 1.1
    
    Else
        Exit Function
    End If
    
    Dim kDlina As Double
    Select Case Dlina
        Case Is <= 200
            kDlina = 1
        Case Is <= 500
            kDlina = 1.1
        Case Is <= 1000
            kDlina = 1.2
        Case Else
            kDlina = 1.3
    End Select
      
    Dim tOtrez As Double
    tOtrez = kMaterialRezka * kDlina * 0.07 * Diametr ^ 0.5
    
    Dim kTochnost As Double
    kTochnost = 1.4

    Dim tPravka As Double
    tPravka = kMaterialPravka * kTochnost * 0.335 * Diametr ^ 0.5 * (Dlina / 1000) ^ 0.84
    
    RezkaPravkaProvoloki = tOtrez + tPravka
    
End Function

Function ZachistkaShvovPnevmatShlifMashinoy(Material As EnumMaterialy, DlinaZachistki As Double) As Double
    'Карта 26.Зачистка сварных швов пневматической шлифовальной машиной

    Dim kMaterial As Double
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.7
    ElseIf Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1
    Else
        Exit Function
    End If
    
    Dim kTavrovyy As Double
    kTavrovyy = 1.3
    
    ZachistkaShvovPnevmatShlifMashinoy = kMaterial * kTavrovyy * 0.028 * DlinaZachistki ^ 0.77
    
End Function

Function ProgonkaRezby(Material As EnumMaterialy, dRezby As Double, Lrezby As Double) As Double
    'Карта 33. IV. Прогонка резьбы метчиком

    If Not dRezby * Lrezby > 0 Then Exit Function
    
    Dim kMaterial As Double
    
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.7
    ElseIf Material = STAL_LEGIROVANNAYA Or Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.2
    ElseIf Material = CHUGUN Then
        kMaterial = 0.8
    Else
        Exit Function
    End If
    
    ProgonkaRezby = kMaterial * 0.315 * (Lrezby ^ 0.62 / dRezby ^ 0.5)
    
End Function

Function PravkaShestigrannika(ByVal S As Double, ByVal Dlina As Double, Material As EnumMaterialy, Optional Tochnost As Double) As Double

    PravkaShestigrannika = PravkaPoPloshchadi(Material, S, PloshchadShestigrannika(S), Dlina, Tochnost)
    
End Function

Function PravkaPloskikhDetaley(ByVal Tolshchina As Double, ByVal Dlina As Double, ByVal Shirina As Double, Material As EnumMaterialy, Optional Tochnost As Double) As Double
    'Карта 7. Правка деталей

    If Tolshchina <= 0 Or Dlina <= 0 Or Shirina <= 0 Then Exit Function
    
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina, Tolshchina)
    
    Tolshchina = Gabarity.Tolshchina
    Dlina = Gabarity.Dlina
    Shirina = Gabarity.Shirina

    PravkaPloskikhDetaley = PravkaPoPloshchadi(Material, Tolshchina, Dlina * Shirina, Dlina, Tochnost)

End Function

Function PravkaPoPloshchadi(Material As EnumMaterialy, Tolshchina As Double, Ploschad As Double, Dlina As Double, Optional Tochnost As Double) As Double
    'Карта 7. Правка деталей
    
    If Tolshchina <= 0 Or Ploschad <= 0 Then Exit Function
    
    Dim Ploschad_sm2 As Double
    Ploschad_sm2 = Ploschad / 100
    
    Dim kMaterial As Double
    Select Case Material
        Case STAL_LEGIROVANNAYA, STAL_UGLERODISTAYA
            kMaterial = 1
            
        Case STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY
            kMaterial = 1.1
            
        Case ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY
            kMaterial = 0.7
            
        Case Else
            Exit Function
    End Select
    
    Dim kTochnost As Double 'Коэффициент на точность правки при стреле прогиба на 1 метр
    Select Case Tochnost
        Case Is <= 0.3: kTochnost = 1.4
        Case Is <= 0.5: kTochnost = 1.2
        Case Is <= 1: kTochnost = 1
        Case Else: kTochnost = 0.85
    End Select

    Dim kDopTochnost As Double  'Требуется плоскостность и прямолинейность 0,2, а в нормативе 0,3 на метр
    kDopTochnost = 0.3 / (0.2 / Dlina * 1000)
    If kDopTochnost < 1 Then kDopTochnost = 1
    
    PravkaPoPloshchadi = kTochnost * kDopTochnost * kMaterial * 0.055 * (Ploschad_sm2 ^ 0.55) * (Tolshchina ^ 0.26)
    
End Function

Function ZenkovaniyeFasokNaNastolnoSverlilnomStanke(DiametrOtverstiya As Double, Material As EnumMaterialy, ByVal KolVoOtverstiy As Double) As Double
    'Карта 32. Зенкование фасок на настольно-сверлильном станке
        
    If DiametrOtverstiya <= 0 Or KolVoOtverstiy < 0 Then Exit Function
    If KolVoOtverstiy = 0 Then KolVoOtverstiy = 1

    Dim kMaterial As Double
    kMaterial = OpredelitKoeffitsientNaMaterialPriZenkovaniiFasok(Material)

    Dim kKolvo As Double
    kKolvo = OpredelitKoeffitsientNaKolVoPriZenkovaniiFasok(KolVoOtverstiy)
    
    ZenkovaniyeFasokNaNastolnoSverlilnomStanke = kMaterial * kKolvo * KolVoOtverstiy * 0.0284 * (DiametrOtverstiya ^ 0.73)
    
End Function

Function ZenkovaniyeFasokPnevmaticheskoySverlilnoyMashinoy(Material As EnumMaterialy, DiametrOtverstiya As Double, _
    ByVal KolVoOtverstiy As Double, Optional ByVal Glubina As Double = 0) As Double
    'Карта 32. Зенкование фасок пневматической сверлильной машиной
    
    If DiametrOtverstiya <= 0 Or KolVoOtverstiy < 0 Then Exit Function
    If KolVoOtverstiy = 0 Then KolVoOtverstiy = 1

    Dim kMaterial As Double
    kMaterial = OpredelitKoeffitsientNaMaterialPriZenkovaniiFasok(Material)

    Dim kKolvo As Double
    kKolvo = OpredelitKoeffitsientNaKolVoPriZenkovaniiFasok(KolVoOtverstiy)
    
    If Glubina = 0 Then
        Select Case DiametrOtverstiya
            Case Is <= 5
                Glubina = 1.5
            Case Is <= 8
                Glubina = 2
            Case Else
                Glubina = 3
        End Select
    End If
    
    ZenkovaniyeFasokPnevmaticheskoySverlilnoyMashinoy = kMaterial * kKolvo * KolVoOtverstiy * 0.052 * (Glubina ^ 0.33) * (DiametrOtverstiya ^ 0.43)
    
End Function

Private Function OpredelitKoeffitsientNaMaterialPriZenkovaniiFasok(Material As EnumMaterialy) As Double

    Dim kMaterial As Double
    If Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.2
    ElseIf Material = CHUGUN Then
        kMaterial = 0.8
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.7
    ElseIf Material = TEKSTOLIT Or Material = GETINAKS Then
        kMaterial = 0.5
    Else
        Exit Function
    End If
    
    OpredelitKoeffitsientNaMaterialPriZenkovaniiFasok = kMaterial
    
End Function

Private Function OpredelitKoeffitsientNaKolVoPriZenkovaniiFasok(KolVoOtverstiy As Double) As Double

    Dim kKolvo As Double
    Select Case KolVoOtverstiy
        Case Is <= 3:   kKolvo = 1
        Case Is <= 10:  kKolvo = 0.9
        Case Else:      kKolvo = 0.85
    End Select

    OpredelitKoeffitsientNaKolVoPriZenkovaniiFasok = kKolvo
    
End Function

Function SnyatiyeZausentsevPoKonturu(DlinaZachistki As Double, Material As EnumMaterialy, Krivolineynyy As Boolean, _
    Instrument As InstrumentyZachistki, VnutrenniyKontur As Boolean) As Double
    'Карта 27. Снятие заусенцев и притупление кромок

    If DlinaZachistki <= 0 Then Exit Function

    Dim kMaterial As Double
    Select Case Material
        Case STAL_LEGIROVANNAYA, STAL_UGLERODISTAYA
            kMaterial = 1
            
        Case STAL_NERZHAVEYUSHCHAYA, TITANOVYYE_SPLAVY
            kMaterial = 1.2
            
        Case CHUGUN
            kMaterial = 0.8
            
        Case ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY, REZINA, TEKSTOLIT, STEKLOTEKSTOLIT, GETINAKS, POLIAMID, ORGSTEKLO
            kMaterial = 0.7
            
        Case Else
            Exit Function
    End Select
    
    Dim a As Double
    If Instrument = PNEVM_MASHINA Then
        a = IIf(Krivolineynyy, 0.021, 0.0186)
    ElseIf Instrument = NAPILNIK_SHABER Then
        a = IIf(Krivolineynyy, 0.03, 0.027)
    ElseIf Instrument = NADFIL Then
        a = IIf(Krivolineynyy, 0.043, 0.038)
    ElseIf Instrument = ABRAZIVNYY_BRUSOK Then
        a = IIf(Krivolineynyy, 0.06, 0.05)
    Else
        Exit Function
    End If
    
    Dim kVnutr As Double
    If VnutrenniyKontur Then kVnutr = IIf(Krivolineynyy, 1.5, 1.3) Else kVnutr = 1

    SnyatiyeZausentsevPoKonturu = kVnutr * kMaterial * a * DlinaZachistki ^ 0.59
    
End Function

