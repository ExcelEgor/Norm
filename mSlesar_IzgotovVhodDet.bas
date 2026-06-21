Attribute VB_Name = "mSlesar_IzgotovVhodDet"
Option Explicit
'ИЗГОТОВЛЕНИЕ ВХОДЯЩИХ ДЕТАЛЕЙ И СБОРКА КОРПУСОВ. СЛЕСАРНО-КАРКАСНЫЕ РАБОТЫ В ПРИБОРОСТРОЕНИИ
'Нормативы времени. Мелкосерийное производство. 1990 г.
Public Const MAX_DLINA_PESKOSTRUY As Integer = 3000     'Увеличил с 2000 до 3000, чтобы считалось время после отрезки на гильотинных ножницах
Public Const MAX_SHIRINA_PESKOSTRUY As Integer = 1500   'Увеличил с 1000 до 1500, чтобы считалось время после отрезки на гильотинных ножницах
Public Const MAX_VYSOTA_PESKOSTRUY As Integer = 1500
Public Const MAX_KOLVO_DETALEY_PESKOSTRUY As Integer = 105
Public Const MAKS_TOLSHCHINA_PUKLEVKA As Integer = 2

Const MIN_DLINA_PESKOSTRUY As Integer = 100
Const MIN_SHIRINA_PESKOSTRUY As Integer = 50

Const K_ZAPILOVKA As Double = 0.8

Private Const NADFIL As Integer = 1
Private Const NAPILNIK As Integer = 2
Private Const SHABER As Integer = 3
Private Const SHKURKA As Integer = 3

Public Const MAX_SHIRINA_PRAVKI_POSLE_SVARKI As Integer = 900
Public Const MAX_DLINA_PRAVKI_POSLE_SVARKI As Integer = 2000
Public Const MAX_TOLSHCHINA_PRAVKI_POSLE_SVARKI As Integer = 6

Public Const MAKS_MASSA_UST_SLESAR_TISKI As Integer = 20

Function KalibrovaniyeRezbyPlashkoy(Material As EnumMaterialy, Diametr As Double, Dlina As Double) As Double
    'Карта 51. Нарезание и калибрование резьбы.
    'Калибрование плашкой, инструмент - вороток
   
    Dim kMaterial As Double
    kMaterial = KoefMaterial_SlesarRezbonarez(Material)
    
    If kMaterial = 0 Then Exit Function
    
    Dim tKalibr As Double
    If Diametr <= 3 Then
        tKalibr = (0.2079 * Dlina ^ 0.32) / Diametr ^ 0.29
    Else
        tKalibr = 0.074 * Dlina ^ 0.27 * Diametr ^ 0.31
    End If
    
    Dim kPlashka As Double
    kPlashka = 0.75

    KalibrovaniyeRezbyPlashkoy = kPlashka * tKalibr
    
End Function

Function ZachistkaPodSvarkuNapilnikom(Material As EnumMaterialy, Dlina As Double) As Double
    'Карта 69. Зачистка под сварку или пайку
    
    Dim kMaterial As Double
    kMaterial = kMaterialZachistka(Material)
    
    If kMaterial > 0 Then
    
        ZachistkaPodSvarkuNapilnikom = kMaterial * 0.1614 * Dlina ^ 0.38
    
    End If
    
End Function

Function ZachistkaShvovNapilnikom(Material As EnumMaterialy, Dlina As Double, _
    Optional Vnutrenniy As Boolean = False, Optional KatetBolshe10 As Boolean = False) As Double
    'Карта 70. Зачистка сварных швов после сварки или пайки
    
    Dim kMaterial As Double
    kMaterial = kMaterialZachistka(Material)
    
    If kMaterial > 0 Then
    
        Dim kVnutrenniy As Double
        kVnutrenniy = IIf(Vnutrenniy, 1.2, 1)
    
        Dim kKatet As Double
        kKatet = IIf(KatetBolshe10, 1.2, 1)

        ZachistkaShvovNapilnikom = kKatet * kVnutrenniy * 0.2248 * Dlina ^ 0.61
    
    End If
    
End Function

Function ZachistkaShvovShlifovalnoyMashinkoy(Material As EnumMaterialy, Dlina As Double, _
    Optional Vnutrenniy As Boolean = False, Optional KatetBolshe10 As Boolean = False, Optional KrivolineynyyShov As Boolean = True) As Double
    'Карта 70. Зачистка сварных швов после сварки или пайки
    
    Dim kMaterial As Double
    kMaterial = kMaterialZachistka(Material)
    
    If kMaterial > 0 Then
    
        Dim kVnutrenniy As Double
        kVnutrenniy = IIf(Vnutrenniy, 1.2, 1)
    
        Dim kKatet As Double
        kKatet = IIf(KatetBolshe10, 1.2, 1)
        
        Dim kKrivolineynyy As Double
        kKrivolineynyy = IIf(KrivolineynyyShov, 1.25, 1)

        ZachistkaShvovShlifovalnoyMashinkoy = kMaterial * kKatet * kVnutrenniy * kKrivolineynyy * 0.2169 * Dlina ^ 0.54
    
    End If
    
End Function

Private Function kMaterialZachistka(Material As EnumMaterialy)

    Dim kMaterial As Double
    If Material = ALUMINIYEVYYE_SPLAVY Then
        kMaterial = 0.6
    ElseIf Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.8
    ElseIf Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1.1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = 1.6
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.8
    Else
        Exit Function
    End If
    
    kMaterialZachistka = kMaterial
    
End Function

Function ZenkovaniePodPotaynuyuGolovku(Material As Integer, d As Double, L As Double, Optional ZakalonnayaStal As Boolean = False) As Double
    'Карта 50. Позиция 4-5
    
    If Not (d > 0 And L > 0) Then Exit Function
    
    Dim D_Max As Integer
    D_Max = 14
    
    Dim L_Max As Integer
    L_Max = 3
    
    If d > D_Max Or L > L_Max Then Exit Function
    
    Dim D_MIN As Double
    D_MIN = 1.5
    If d < D_MIN Then d = D_MIN
    
    Dim L_Min As Double
    L_Min = 1.5
    If L < L_Min Then L = L_Min
        
    '1 - Алюминиевые сплавы
    '2 - Медные сплавы
    '3 - Сталь Ст3, Ст10 и т.д.
    '4 - Сталь У8А, 12ХН3А, 40Х и т.д.
    '5 - Сталь Э12, Э31 и т.д.
    '6 - Сталь 14Х17Н2, 25Х13Н2 и т.д.
    '7 - Сплав 52КФТМ и т.д.
    
    Dim K1 As Double
    Select Case Material
        Case 1: K1 = 0.7
        Case 2: K1 = 0.8
        Case 3: K1 = 1
        Case 4: K1 = 1.1
        Case 5: K1 = 1.4
        Case 6: K1 = 1.5
        Case 7: K1 = 1.6
    End Select
    
    Dim K2 As Integer
    K2 = IIf(ZakalonnayaStal, 1.2, 1)
    
    ZenkovaniePodPotaynuyuGolovku = (0.0249 * d ^ 0.39 * L ^ 0.35) * K1 * K2
    
End Function

Function UstanovkaSlesar_vSlesarnyyeTiski_UstanovkaSyom(Massa As Double) As Double

    UstanovkaSlesar_vSlesarnyyeTiski_UstanovkaSyom = UstanovkaSlesar_vSlesarnyyeTiski(Massa)(1)
    
End Function

Function UstanovkaSlesar_vSlesarnyyeTiski_Povorot(Massa As Double) As Double

    UstanovkaSlesar_vSlesarnyyeTiski_Povorot = UstanovkaSlesar_vSlesarnyyeTiski(Massa)(2)
    
End Function

Private Function UstanovkaSlesar_vSlesarnyyeTiski(Massa As Double) As Variant

    If Massa > MAKS_MASSA_UST_SLESAR_TISKI Then Exit Function
    
    Dim tUst As Double, tPovorot As Double
    Select Case Massa
        Case Is <= 1
            tUst = 0.25
            tPovorot = 0.15
        Case Is <= 3
            tUst = 0.34
            tPovorot = 0.2
        Case Is <= 5
            tUst = 0.39
            tPovorot = 0.25
        Case Is <= 8
            tUst = 0.43
            tPovorot = 0.3
        Case Is <= 12
            tUst = 0.49
            tPovorot = 0.35
        Case Else
            tUst = 0.55
            tPovorot = 0.4
    End Select
    
    Dim Normy(1 To 3) As Variant
    Normy(1) = tUst
    Normy(2) = tPovorot
    
    UstanovkaSlesar_vSlesarnyyeTiski = Normy
    
End Function

Function UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(Massa As Double) As Double
    'КАРТА 2. УСТАНОВКА, КРЕПЛЕНИЕ И СНЯТИЕ ДЕТАЛИ
    'На столе без крепления. Установить и снять

    Dim tUst As Double
    Select Case Massa
        Case Is <= 1:   tUst = 0.15
        Case Is <= 3:   tUst = 0.2
        Case Is <= 5:   tUst = 0.25
        Case Is <= 8:   tUst = 0.3
        Case Is <= 12:  tUst = 0.35
        Case Is <= 20:  tUst = 0.4
        Case Is <= 30:  tUst = 2.5
        Case Is <= 50:  tUst = 2.7
        Case Else:      tUst = 3.3
    End Select
    
    If Massa > 20 Then tUst = tUst * 2
    
    UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom = tUst
    
End Function

Function UstanovkaSlesar_NaStoleBezKrepleniya_Povorot(Massa As Double) As Double
    
    Dim tUst As Double
    Select Case Massa
        Case Is <= 1:   tUst = 0.1
        Case Is <= 3:   tUst = 0.15
        Case Is <= 5:   tUst = 0.2
        Case Is <= 8:   tUst = 0.25
        Case Is <= 12:  tUst = 0.3
        Case Is <= 20:  tUst = 0.35
        Case Is <= 30:  tUst = 2.1
        Case Is <= 50:  tUst = 2.3
        Case Else:      tUst = 2.8
    End Select
    
    If Massa > 20 Then tUst = tUst * 2
    UstanovkaSlesar_NaStoleBezKrepleniya_Povorot = tUst
    
End Function

Private Function VremyaNaUstanovkuPerevoroty(Massa As Double, KolVoPerevorotov As Double) As Double
    If Massa > 0 Then
        Dim tUst As Double
        tUst = UstanovkaSlesar_NaStoleBezKrepleniya_UstanovkaSyom(Massa)
        
        Dim tPerevorot As Double
        tPerevorot = UstanovkaSlesar_NaStoleBezKrepleniya_Povorot(Massa) * KolVoPerevorotov
        
        VremyaNaUstanovkuPerevoroty = tUst + tPerevorot
    End If
End Function
Function Koeffitsient_KolVoDetaleyVPartii_IzgotVhodDetaley(KolVoDetaleyVPartii As Double) As Double
    Dim K1 As Double
    Select Case KolVoDetaleyVPartii
        Case Is <= 2:   K1 = 1.3
        Case Is <= 5:   K1 = 1.2
        Case Is <= 10:  K1 = 1
        Case Is <= 15:  K1 = 0.95
        Case Is <= 24:  K1 = 0.9
        Case Is <= 39:  K1 = 0.85
        Case Else:      K1 = 0.8
    End Select
    Koeffitsient_KolVoDetaleyVPartii_IzgotVhodDetaley = K1
    
End Function

Function RaschotKoefUsloviya_IzgotVhodDetaley(UsloviyaVypolneniyaRaboty As Integer) As Double

    Dim K2 As Double
    Select Case UsloviyaVypolneniyaRaboty
        Case 2:     K2 = 1.3
        Case 3:     K2 = 1.5
        Case Else:  K2 = 1
    End Select
    RaschotKoefUsloviya_IzgotVhodDetaley = K2
    
End Function

Function GlavnyyKoefftsient_IzgotVhodDetaley(Optional KolVoDetaleyVPartii As Double = 10, Optional UsloviyaVypolneniyaRaboty As Integer = 1, Optional MassaBolshe20 As Boolean = False) As Double

    Dim K1 As Double
    K1 = Koeffitsient_KolVoDetaleyVPartii_IzgotVhodDetaley(KolVoDetaleyVPartii)

    Dim K2 As Double
    K2 = RaschotKoefUsloviya_IzgotVhodDetaley(UsloviyaVypolneniyaRaboty)
    
    Dim K3 As Double
    K3 = IIf(MassaBolshe20, 1.2, 1)
    
    GlavnyyKoefftsient_IzgotVhodDetaley = K1 * K2 * K3
    
End Function

Function PeskostruynayaOchistka(ByVal DiametrIliShirina As Double, ByVal Dlina As Double, _
    Optional TipZagotovki As Integer = 1, Optional GruppaSlozhnosti As Integer = 1, _
    Optional Massa As Double = 1, Optional Material As EnumMaterialy = STAL_UGLERODISTAYA, Optional Usloviya As Integer = 2) As Double
    'КАРТА 66. ОЧИСТКА ДРОБЕСТРУЙНАЯ ЗАГОТОВОК И ДЕТАЛЕЙ

    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(DiametrIliShirina, Dlina)
    DiametrIliShirina = Gabarity.Shirina
    Dlina = WorksheetFunction.Max(Gabarity.Dlina, MIN_DLINA_PESKOSTRUY)
    
    Dim kGruppaMaterialov As Double
    Select Case Material
        Case EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.TITANOVYYE_SPLAVY
            kGruppaMaterialov = 1
        Case EnumMaterialy.MEDNYYE_SPLAVY, EnumMaterialy.ALUMINIYEVYYE_SPLAVY
            kGruppaMaterialov = 0.9
        Case EnumMaterialy.GETINAKS
            kGruppaMaterialov = 0.8
        Case EnumMaterialy.TEKSTOLIT, EnumMaterialy.STEKLOTEKSTOLIT, EnumMaterialy.POLIAMID
            kGruppaMaterialov = 0.7
        Case EnumMaterialy.ORGSTEKLO
            kGruppaMaterialov = 0.5
        Case Else
            Err.Raise vbObjectError + 1, "PeskostruynayaOchistka", "Недопустимый материал"
    End Select
    
    Dim X As Double, Y As Double, z As Double
    Select Case TipZagotovki
        Case 1 'Угольники, швеллеры, оси, планки
            X = 0.0294: Y = 0.15:   z = 0.49
        Case 2  'Втулки, колеса, обоимы, рамы, кольца, фланцы
            X = 0.004:  Y = 0.36:   z = 0.77
        Case 3  'Платы, крышки, сварные кронштейны
            X = 0.0042: Y = 0.6:    z = 0.61
        Case Else
            Err.Raise vbObjectError + 1, "PeskostruynayaOchistka", "Недопустимый тип заготовки"
    End Select
    
    Dim kGruppaSlozhnosti As Double
    Select Case GruppaSlozhnosti
        Case 1:     kGruppaSlozhnosti = 1
        Case 2:     kGruppaSlozhnosti = IIf(TipZagotovki > 1, 1.2, 1)
        Case Else:  Err.Raise vbObjectError + 1, "PeskostruynayaOchistka", "Недопустимая группа сложности"
    End Select
    
    Dim kUsloviya As Double
    kUsloviya = RaschotKoefUsloviya_IzgotVhodDetaley(Usloviya)
    
    Dim tUstanovkiPerevorotov As Double
    tUstanovkiPerevorotov = VremyaNaUstanovkuPerevoroty(Massa, 1)

    PeskostruynayaOchistka = kGruppaMaterialov * kGruppaSlozhnosti * kUsloviya * (X * DiametrIliShirina ^ Y * Dlina ^ z) + tUstanovkiPerevorotov
    Exit Function
    
End Function

Function PeskostruyKarkasov(ByVal Dlina As Double, ByVal Shirina As Double, ByVal Vysota As Double, _
    Optional ByVal KolVoDetaleyVKarkase As Double = 15, Optional Massa As Double = 1) As Double
    'КАРТА 67. ОЧИСТКА ДРОБЕСТРУЙНАЯ СВАРНЫХ КАРКАСОВ И КОРПУСОВ

    Dim Arr_DlinaShirinaVysota
    Arr_DlinaShirinaVysota = Array(Dlina, Shirina, Vysota)
    
    Dlina = WorksheetFunction.Large(Arr_DlinaShirinaVysota, 1)
    Vysota = WorksheetFunction.Large(Arr_DlinaShirinaVysota, 2)
    Shirina = WorksheetFunction.Large(Arr_DlinaShirinaVysota, 3)
    
    Dlina = MinimalnoeZnachenie(Dlina, 600)
    Shirina = MinimalnoeZnachenie(Shirina, 300)
    Vysota = MinimalnoeZnachenie(Vysota, 400)
    KolVoDetaleyVKarkase = MinimalnoeZnachenie(KolVoDetaleyVKarkase, 15)
    Massa = MinimalnoeZnachenie(Massa, 1)
    
    If Dlina > MAX_DLINA_PESKOSTRUY Or _
        Shirina > MAX_SHIRINA_PESKOSTRUY Or _
        Vysota > MAX_VYSOTA_PESKOSTRUY Or _
        KolVoDetaleyVKarkase > MAX_KOLVO_DETALEY_PESKOSTRUY Then
        Exit Function
    End If

    Dim tPeskostruy As Double
    tPeskostruy = 0.0549 * (Dlina + Shirina + Vysota) ^ 0.45 * KolVoDetaleyVKarkase ^ 0.73
    
    Dim tUstanovkiPerevorotov As Double
    tUstanovkiPerevorotov = VremyaNaUstanovkuPerevoroty(Massa, 5)
    
    PeskostruyKarkasov = WorksheetFunction.RoundUp(tPeskostruy + tUstanovkiPerevorotov, 1)
    
End Function

Function Polirovanie(Dlina As Double, Shirina As Double, Material As EnumMaterialy, Optional Sloznost As Integer = 3, Optional Ra As Double = 2.5) As Double
    'Карта 44

    Dim K1 As Double
    Select Case Material
        Case PLASTMASSA
            K1 = 1
        Case ORGSTEKLO
            K1 = 1.2
        Case ALUMINIYEVYYE_SPLAVY
            K1 = 1.3
        Case MEDNYYE_SPLAVY
            K1 = 1.5
        Case STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, STAL_UGLERODISTAYA
            K1 = 1.7
        Case Else
            Exit Function
    End Select
    
    Dim PloschadVSantimetrakh As Double
    PloschadVSantimetrakh = Dlina / 10 * Shirina / 10
    
    Dim X As Double, Y As Double
    Select Case Sloznost
        Case 1
            X = 0.072
            Y = 0.57
        Case 2
            X = 0.16
            Y = 0.49
        Case Else
            X = 0.2
            Y = 0.49
    End Select
    Polirovanie = X * PloschadVSantimetrakh ^ Y
    
    Dim K2 As Double
    Select Case Ra
        Case Is <= 0.32:    K2 = 2.3
        Case Is <= 1.25:    K2 = 1.55
        Case Else:          K2 = 1
    End Select
    
    Polirovanie = Polirovanie * K1 * K2

End Function

Function ZachistkaZausencev_PoKonturu_Napilnikom(Material As EnumMaterialy, ByVal DlinaZachistki As Double, _
    Optional ObrabotkaPoH5H7 As Boolean = False, Optional MassaBolshe20 As Boolean = False) As Double
    
    ZachistkaZausencev_PoKonturu_Napilnikom = ZachistkaZausencev(Material, DlinaZachistki, NAPILNIK, ObrabotkaPoH5H7, MassaBolshe20)
    
End Function

Function ZachistkaZausencev(Material As EnumMaterialy, ByVal DlinaZachistki As Double, Optional Instrumet As Integer = NAPILNIK, _
    Optional ObrabotkaPoH5H7 As Boolean = False, Optional MassaBolshe20 As Boolean = False) As Double
    'Зачистка заусенцев после резки, штамповки, механической обработки. Карта 11
    
    'Instrumet - Инструмент:
    '1-Надфиль
    '2-Напильник
    '3-Шабер
    '4-Шлифовальная шкурка
    
    If Not DlinaZachistki > 0 Then Exit Function
    
    DlinaZachistki = MinimalnoeZnachenie(DlinaZachistki, IIf(Instrumet = 1, 12, 30))
       
    Dim a As Double, b As Double
    If Instrumet = NADFIL Then
        a = 0.0434
        b = 0.52
    Else
        a = 0.053
        b = 0.48
    End If
    
    Dim K1 As Double
    If Instrumet = SHABER Then
        K1 = 0.9
    ElseIf Instrumet = SHKURKA Then
        K1 = 0.6
    Else    'NAPILNIK
        K1 = 1
    End If
    
    Dim K2 As Double
    K2 = IIf(ObrabotkaPoH5H7, 1.2, 1)
    
    Dim K3 As Double
    K3 = KoefMaterial_SlesarRezbonarez(Material)
    
    Dim K4 As Double
    K4 = IIf(MassaBolshe20, 1.2, 1)
    
    ZachistkaZausencev = K1 * K2 * K3 * K4 * (a * DlinaZachistki ^ b)
    
End Function

Function SnyatiyeZausentsevList(Dlina As Double, Shirina As Double, Tolshchina As Double, Material As EnumMaterialy) As Double

    Dim List As New clsFigura_List
    With List
        .Init Material, Shirina, Tolshchina, Dlina

        SnyatiyeZausentsevList = .tZaus
    End With

End Function

Function ZachistkaZausencevSOtverstiiVruchnuyu(Material As EnumMaterialy, ByVal DiametrOtverstiya As Double, _
    Optional Instrumet As Integer = 2, Optional KolVoStoron As Integer = 1, Optional MassaBolshe20 As Boolean = False) As Double
    'Зачистка заусенцев с отверстий. Карта 12

    If Not DiametrOtverstiya > 0 Then Exit Function
    
    Dim DiametrOtverstiya_Min As Integer: DiametrOtverstiya_Min = 5
    If DiametrOtverstiya < DiametrOtverstiya_Min Then DiametrOtverstiya = DiametrOtverstiya_Min
    
    If KolVoStoron < 1 Then KolVoStoron = 1
    
    Dim a As Double, b As Double
    Select Case Instrumet
        Case 1: a = 0.019:  b = 0.71    'Сверло
        Case 2: a = 0.062:  b = 0.3     'Шабер
        Case 3: a = 0.0417: b = 0.7     'Шлифовальная шкурка
        Case Else: Exit Function
    End Select
    
    Dim kMaterial As Double:    kMaterial = KoefMaterial_SlesarRezbonarez(Material)
    Dim kMassa As Double:       kMassa = IIf(MassaBolshe20, 1.2, 1)
    
    ZachistkaZausencevSOtverstiiVruchnuyu = kMaterial * kMassa * KolVoStoron * (a * DiametrOtverstiya ^ b)
    
End Function

Function PuklevkaOtverstiiVruchnuyu(Material As EnumMaterialy, Tolschina As Double, KolvoOtverstii As Double, Optional Sposob As Integer = 2) As Double
    'Пуклевка отверстий. Карта 57
    
    Dim kMaterial As Double
    If Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
        kMaterial = 0.8
    ElseIf Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.9
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1.1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = 1.7
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.8
    Else
        Exit Function
    End If

    If Tolschina > MAKS_TOLSHCHINA_PUKLEVKA Then Exit Function

    PuklevkaOtverstiiVruchnuyu = kMaterial * ((0.598 * Tolschina ^ 0.16) + ((0.336 * Tolschina ^ 0.28) * (KolvoOtverstii - 1)))

End Function

Function ZachistkaZausencevPosleFrezPazovShlicev(Material As Integer, L As Double, Optional Instrumet As Integer = 2, _
    Optional ObrabotkaPoH5H7 As Boolean = False, Optional TipPaza As Integer = 1) As Double
Attribute ZachistkaZausencevPosleFrezPazovShlicev.VB_Description = "Зачистка заусенцев после фрезерования пазов, выемок, уступов, шлицев.  Карта 11."
Attribute ZachistkaZausencevPosleFrezPazovShlicev.VB_ProcData.VB_Invoke_Func = " \n19"
    'Зачистка заусенцев после фрезерования пазов, выемок, уступов, шлицев.  Карта 11
    
    'Материал - Материал
    '1 - Алюминиевые сплавы
    '2 - Медные сплавы
    '3 - Ст 35
    '4 - 40Х
    '5 - 20Х13
    '6 - 12Х18Н9Т
    '7 - ВТ1-0
    
    'Instrumet - Инструмент:
    '1-Надфиль
    '2-Напильник
    '3-Шабер
    '4-Шлифовальная шкурка
    
    'TipPaza - Тип паза, выемки, уступа, шлица:
    '1 - открытый
    '2 - полузакрытый
    '3 - закрытый
    
    L = MinimalnoeZnachenie(L, IIf(Instrumet = 1, 5, 12))
       
    Dim a As Double, b As Double
    If Instrumet = 1 Then
        a = 0.106
        b = 0.52
    Else
        a = 0.039
        b = 0.6
    End If
    
    Dim K1 As Double
    If Instrumet = 3 Then
        K1 = 0.9
    ElseIf Instrumet = 4 Then
        K1 = 0.6
    Else
        K1 = 1
    End If
    
    Dim K2 As Double
    K2 = IIf(ObrabotkaPoH5H7, 1.2, 1)
    
    Dim K3 As Double
    K3 = KoefMaterial_SlesarRezbonarez(CInt(Material))
    
    Dim K4 As Double
    K4 = Choose(TipPaza, 1, 1.4, 1.9)
    
    ZachistkaZausencevPosleFrezPazovShlicev = (a * L ^ b) * K1 * K2 * K3 * K4
    
End Function

Private Function KoefMaterial_SlesarRezbonarez(Material As EnumMaterialy) As Double
    
    Dim K1 As Double
    
    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, POLIAMID, TEKSTOLIT, STEKLOTEKSTOLIT, ORGSTEKLO
            K1 = 0.7
        Case MEDNYYE_SPLAVY
            K1 = 0.8
        Case STAL_UGLERODISTAYA
            K1 = 1
        Case STAL_LEGIROVANNAYA
            K1 = 1.1
        Case STAL_NERZHAVEYUSHCHAYA
            K1 = 1.7
        Case TITANOVYYE_SPLAVY
            K1 = 1.8
    End Select
    
    KoefMaterial_SlesarRezbonarez = K1

End Function

Private Function KoefMarkaStali(Stal As Integer) As Double
    Dim K1 As Double
    Select Case Stal
        Case EnumMaterialy.STAL_UGLERODISTAYA:        K1 = 1
        Case EnumMaterialy.STAL_LEGIROVANNAYA:        K1 = 1.1
        Case EnumMaterialy.STAL_NERZHAVEYUSHCHAYA:    K1 = 1.7
        Case EnumMaterialy.TITANOVYYE_SPLAVY:         K1 = 1.8
        Case Else:                          Err.Raise vbObjectError + 1, KoefMarkaStali, "Недопустимый материал"
    End Select
    KoefMarkaStali = K1
End Function

Function VypilivaniyeVnutrennegoRadiusa(ByVal Tolshchina As Double, ByVal Radius As Double, Material As EnumMaterialy) As Double
Attribute VypilivaniyeVnutrennegoRadiusa.VB_Description = "Выпиливание внутреннего радиуса под 90° (после фрезерования). Карта 12."
Attribute VypilivaniyeVnutrennegoRadiusa.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim kMaterial As Double
    If Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Then
        kMaterial = 0.7
    ElseIf Material = MEDNYYE_SPLAVY Then
        kMaterial = 0.8
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1.1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = 1.7
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.8
    Else
        Exit Function
    End If

'    Tolshchina = MinimalnoeZnachenie(Tolshchina, 2)
    If Radius = 0 Then Radius = 5
    
    Dim tVypilivaniye As Double
    tVypilivaniye = 0.12 * Radius ^ 0.64 * Tolshchina ^ 0.53

    VypilivaniyeVnutrennegoRadiusa = kMaterial * tVypilivaniye
    
End Function

Function ProshivkaVruchnuyu_m56(Material As Integer, S As Double, TipKontura As Integer, KolKonturov As Double) As Double
Attribute ProshivkaVruchnuyu_m56.VB_Description = "Прошивка окон, пазов, отверстий вручную. Карта 56."
Attribute ProshivkaVruchnuyu_m56.VB_ProcData.VB_Invoke_Func = " \n19"
    
    If Material < 1 Or Material > 3 Then Exit Function
    If Material = 1 And S > 2 Then S = 2
    If Material > 1 And S > 1 Then S = 1
    If TipKontura < 1 Or TipKontura > 3 Then Exit Function
    
    Dim rTime As Range
    Set rTime = wsIzgotVhodDet_56_1.Range("E8:M9").Offset(((TipKontura * 2) - 2))
    
    Dim NumCol As Integer
    Select Case S
        Case Is <= 0.3: NumCol = 1
        Case Is <= 0.4: NumCol = 2
        Case Is <= 0.5: NumCol = 3
        Case Is <= 0.7: NumCol = 4
        Case Is <= 1:   NumCol = 5
        Case Is <= 1.5: NumCol = 6
        Case Is <= 2:   NumCol = 7
    End Select
    NumCol = NumCol + ((Material * 2) - 2)
    
    Dim tFirst As Double, tNext As Double
    tFirst = rTime(1, NumCol)
    tNext = rTime(2, NumCol)
    
    ProshivkaVruchnuyu_m56 = tFirst + (tNext * (KolKonturov - 1))
    
End Function

Function VyrezaniyeProkladokVruchnuyu(Material As Integer, ByVal Tolshchina As Double, ByVal SummarnayaDlinaReza As Double, _
    Optional Sposob As Integer = 1, Optional KrivolineynyyKontur As Boolean = False, Optional NaruzVnutr As Integer = 1) As Double
    'Карта 52. Вырезание прокладок вручную
    'Содержание работы
    'Взять заготовку, шаблон. Установить шаблон на заготовку. Взять инструмент (нож, ножницы), вырезать прокладку. Отложить инструмент, шаблон, прокладку.

    'Sposob:        1-По шаблону; 2-По разметке
    'PriamKriv:     1-Прямолинейный; 2-Криволинейный
    'NaruzVnutr:    1-Наружный; 2-Внутренний
    'Material:      1-Картон , бумага, дерматин, полиэтилен, фторопласт; 2-прессшпан , паранит, асбест; 3-кожа , резина, брезент; 4-губчатая резина, фетр, войлок
       
    If SummarnayaDlinaReza <= 0 Then Exit Function
    
    If Sposob = 0 Then Sposob = 1
    If NaruzVnutr = 0 Then NaruzVnutr = 1
      
    If 1 > Sposob Or Sposob > 2 _
        Or 1 > NaruzVnutr Or NaruzVnutr > 2 _
        Or 1 > Material Or Material > 4 _
        Or Tolshchina > 20 Then
        Exit Function
    End If
    
    If Tolshchina < 1 Then Tolshchina = 1
    If SummarnayaDlinaReza < 120 Then SummarnayaDlinaReza = 120
    
    Dim X As Double, Y As Double, z As Double, W As Double
    
    If Sposob = 1 Then
        X = 1.78264314692331E-03
        Y = 0.705604478216146
        z = 2.57036337061542E-02
        W = 0.496880881241297
    Else
        X = 1.62153756434872E-03
        Y = 0.679991051348182
        z = 2.19390119161153E-02
        W = 0.514836170031468
    End If
    
    Dim K1 As Double:   K1 = IIf(NaruzVnutr = 2, 1.3, 1)
    Dim K2 As Double:   K2 = IIf(KrivolineynyyKontur, 1.2, 1)
    
    Dim K3 As Double
    Select Case Material
        Case 1: K3 = 1
        Case 2: K3 = 1.2
        Case 3: K3 = 2
        Case 4: K3 = 2.2
    End Select
    
    Dim Vremya As Double
    Vremya = K1 * K2 * K3 * (X * SummarnayaDlinaReza ^ Y * Tolshchina + z * SummarnayaDlinaReza ^ W)
    
    Dim ChisloZnakov As Integer
    ChisloZnakov = IIf(Vremya > 1, 1, 2)

    VyrezaniyeProkladokVruchnuyu = WorksheetFunction.Round(Vremya, ChisloZnakov)

End Function

Function Obduvka_m65(FormaDetali As Integer, AD As Double, L As Double, Optional Sloznost As Integer = 2) As Double
    
    'FormaDetali - Форма детали
    '1 - Плоская
    '2 - Цилиндрическая
        
    'AD - ширина или диаметр поверхности, мм
    'L - длина поверхности, мм

    'Sloznost - Поверхность детали
    '1 - Простой конфигурации
    '2 - Сложной конфигурации
    
    Dim CalcAD As Double, CalcL As Double
    CalcAD = WorksheetFunction.Min(AD, L)
    CalcL = WorksheetFunction.Max(AD, L)
    
    Dim a As Double, b As Double, c As Double
    
    If FormaDetali = 1 Then
        a = 3.78237892934291E-03
        b = 0.377845203479438
        c = 0.375566429367064
    Else
        a = 5.89818996366569E-03
        b = 0.375225539991688
        c = 0.375301473461325
    End If
    
    Obduvka_m65 = ((a * CalcAD ^ b) * CalcL ^ c) * IIf(Sloznost = 2, 1.3, 1)


End Function
    
Function Promivka_m64(FormaDetali As Integer, TipPromivki As Integer, AD As Double, L As Double, Sloznost As Integer) As Double
Attribute Promivka_m64.VB_Description = "Промывка деталей в ванне. Карта 64."
Attribute Promivka_m64.VB_ProcData.VB_Invoke_Func = " \n19"
    
    'FormaDetali - Форма детали
    '1 - Плоская
    '2 - Цилиндрическая
    
    'TipPromivki - Тип промывки
    '1 - От пыли и стружки
    '2 - От масла
    '3 - От консистентной смазки
        
    'AD - ширина или диаметр промываемой поверхности, мм
    'L - длина промываемой поверхности, мм

    'Sloznost - Поверхность детали
    '1 - Простой конфигурации
    '2 - Сложной конфигурации

    If 1 > FormaDetali Or FormaDetali > 2 Or 1 > TipPromivki Or TipPromivki > 3 Or 1 > Sloznost Or Sloznost > 2 Then Exit Function
    
    Dim CalcAD As Double, CalcL As Double
    CalcAD = WorksheetFunction.Min(AD, L)
    CalcL = WorksheetFunction.Max(AD, L)
    
    Static tbPromivka As ListObject
    Set tbPromivka = ThisWorkbook.Worksheets("ИзготВход_64").ListObjects("tbPromivka")
    
    Dim id As String
    id = Sloznost & FormaDetali & TipPromivki
    
    Dim idRow As Integer
    
    With tbPromivka
        
        On Error Resume Next
        idRow = WorksheetFunction.Match(id, .ListColumns(4).DataBodyRange, 0)
        
        If Err.Number = 0 Then
            With .ListRows(WorksheetFunction.Match(id, .ListColumns(4).DataBodyRange, 0)).Range
                Dim a As Double, b As Double, c As Double
                a = .Columns(5)
                b = .Columns(6)
                c = .Columns(7)
            End With
            
            Promivka_m64 = ((a * CalcAD ^ b) * CalcL ^ c)
        End If
            
    End With

    
End Function

Function GibkaDetaleyVruchnuyu_Range(Material As EnumMaterialy, Tolshchina As Double, rDlinaLiniiGiba As Range, rVisotaBorta As Range, rKolvoUglov As Range, _
    Optional rUgol As Range, Optional rKrivolineynyyKonturGibki As Range)
     
    Dim i As Integer
    Dim t As Double
    Dim KolVo As Double
    Dim Ugol As Double
    Dim KrivolineynyyKonturGibki As Boolean

    For i = 1 To rDlinaLiniiGiba.Rows.Count
        If rDlinaLiniiGiba(i) > 0 Then
        
            If rUgol Is Nothing Then
                Ugol = 0
            Else
                Ugol = rUgol(i)
            End If
            
            If rKrivolineynyyKonturGibki Is Nothing Then
                KrivolineynyyKonturGibki = False
            Else
                KrivolineynyyKonturGibki = rKrivolineynyyKonturGibki(i) = "Да"
            End If
            
            t = GibkaDetaleyVruchnuyu(Material, Tolshchina, rDlinaLiniiGiba(i), rVisotaBorta(i), Ugol, KrivolineynyyKonturGibki)
            KolVo = IIf(rKolvoUglov(i) = 0, 1, rKolvoUglov(i))
            
            If i = 1 Then
                t = (t * (KolVo - 1) * 1.1) + t
            Else
                t = t * KolVo * 1.1
            End If
            
            GibkaDetaleyVruchnuyu_Range = GibkaDetaleyVruchnuyu_Range + t
            
        End If
    Next

End Function

Function GibkaDetaleyVruchnuyu(Material As EnumMaterialy, ByVal Tolshchina As Double, ByVal DlinaLiniiGiba As Double, ByVal VysotaBorta As Variant, _
    Optional UgolGibki As Double = 90, Optional KrivolineynyyKonturGibki As Boolean = False) As Double
    
    Dim X As Double, Y As Double, z As Double, W As Double
    
    Select Case VysotaBorta
        Case Is <= 5
            X = 0.002779:   Y = 0.65856
            z = 0.533597:   W = 0.313305
        Case Is <= 10
            X = 0.002692:   Y = 0.71928
            z = 0.703991:   W = 0.18517
        Case Is <= 20
            X = 0.002657:   Y = 0.780782
            z = 0.81077:    W = 0.103569
        Case Is <= 50
            X = 0.003024:   Y = 0.766114
            z = 0.776497:   W = 0.132098
        Case Is <= 150
            X = 0.003357:   Y = 0.749823
            z = 0.873237:   W = 0.148808
        Case Else
            X = 0.003732:   Y = 0.747432
            z = 0.907511:   W = 0.173613
    End Select
    
    Dim tGibka As Double
    tGibka = WorksheetFunction.Round((X * Tolshchina ^ Y) * DlinaLiniiGiba + (z * Tolshchina ^ W), 1)
    
    Dim PopravochnyyKoeffitsient As Double
    PopravochnyyKoeffitsient = RaschotKoeffitsienta_GibkaDetaleyVruchnuyu(Material, Tolshchina, UgolGibki, KrivolineynyyKonturGibki)
    
    GibkaDetaleyVruchnuyu = PopravochnyyKoeffitsient * tGibka

End Function

Private Function RaschotKoeffitsienta_GibkaDetaleyVruchnuyu(Material As EnumMaterialy, Tolshchina As Double, UgolGibki As Double, KrivolineynyyKonturGibki As Boolean)

    Dim kMaterial As Double
    
    Dim kStal As Double
    kStal = 1.2
    
    If Material = ALUMINIYEVYYE_SPLAVY Then
        kMaterial = 0.8
    ElseIf Material = MEDNYYE_SPLAVY Then
        kMaterial = 1
    ElseIf Material = STAL_UGLERODISTAYA Then
        kMaterial = kStal
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = kStal * 1.1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = kStal * 1.7
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = kStal * 1.8
    Else
        Exit Function
    End If
        
    Dim kUgol As Double
    kUgol = IIf(UgolGibki > 90, 1.4, 1)
    
    Dim kKrivolineynyy As Double
    kKrivolineynyy = IIf(KrivolineynyyKonturGibki, 2, 1)
    
    Dim kOPytnyy  'Поправочный коэффициент, полученный опытно-статистическим путем
    If Tolshchina >= 8 Then
        kOPytnyy = 1.3
    Else
        kOPytnyy = 1
    End If
    
    RaschotKoeffitsienta_GibkaDetaleyVruchnuyu = kMaterial * kUgol * kKrivolineynyy * kOPytnyy
        
End Function


Function ZapilovkaTortsovPosleGibki(tGibka As Double) As Double
    'Функция на основе карты 25 - гибка деталей вручную

    ZapilovkaTortsovPosleGibki = tGibka / K_ZAPILOVKA - tGibka
    
End Function


Function PravkaPosleTochechnoySvarki(ByVal Shirina As Double, ByVal Dlina As Double, ByVal Vysota As Double, Material As EnumMaterialy, _
    Optional StrelaProgiba As Double = 2, Optional KolVoStoron As Integer = 2, Optional BoleePyatiDetaleyVSborke As Boolean = False) As Double
    
    'Material:
    '1-Алюминий
    '2-Ст35
    '3-40Х
    '4-20Х13
    '5-12Х18Н9Т
    '6-ВТ1-0
    
    If Material < 1 Or Material > 6 Or Vysota > 500 > 6 Or Shirina > 900 Or Dlina > 2000 Then Exit Function
    
    Dim Gabarity As Variant
    Gabarity = DlinaShirina(Dlina, Shirina, Vysota)
    Dlina = Gabarity(1)
    Shirina = Gabarity(2)
    Vysota = Gabarity(3)
    
    Shirina = MinimalnoeZnachenie(Shirina, 120)
    Dlina = MinimalnoeZnachenie(Dlina, 120)
    Vysota = MinimalnoeZnachenie(Vysota, 20)

    Dim a As Double:    a = 0.123
    Dim X As Double:    X = 0.2
    Dim Y As Double:    Y = 0.17
    Dim z As Double:    z = 0.21
    
    Dim K As Double
    K = Koeffitsient_PravkaPosleSvarki(Material, StrelaProgiba, KolVoStoron, BoleePyatiDetaleyVSborke)
    
    PravkaPosleTochechnoySvarki = (a * Shirina ^ X * Dlina ^ Y * Vysota ^ z) * K
    
    
End Function

Private Function Koeffitsient_PravkaPosleSvarki(Material As EnumMaterialy, Optional StrelaProgiba As Double = 2, Optional KolVoStoron As Integer = 2, _
    Optional BoleePyatiDetaleyVSborke As Boolean = False) As Double

    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
    
    If Material = ALUMINIYEVYYE_SPLAVY Or _
            Material = MEDNYYE_SPLAVY Then
        K1 = 1
    ElseIf Material = STAL_UGLERODISTAYA Or _
            Material = STAL_LEGIROVANNAYA Or _
            Material = STAL_NERZHAVEYUSHCHAYA Or _
            Material = TITANOVYYE_SPLAVY Then
        K1 = 1.3
    Else
        Exit Function
    End If

    Select Case StrelaProgiba
        Case Is <= 0.6: K2 = 1.5
        Case Is <= 1.3: K2 = 1.3
        Case Is <= 2:   K2 = 1
        Case Is <= 3:   K2 = 0.8
        Case Else:      K2 = 0.7
    End Select

    Select Case KolVoStoron
        Case 1, 2:      K3 = 1
        Case 3:         K3 = 1.4
        Case 4:         K3 = 1.5
    End Select

    K4 = IIf(BoleePyatiDetaleyVSborke, 1.5, 1)
    
    Koeffitsient_PravkaPosleSvarki = K1 * K2 * K3 * K4
    
End Function

Function PeskostruynayaOchistkaSvarnukhIzdeleiy(Dlina As Double, Shirina As Double, Vysota As Double, KolVoDetaleyVSborke As Double, Optional OchistkaVruchnuyu As Boolean = True) As Double
    'Карта 67. Очистка дробеструйная сварных каркасов и корпусов.
    
    Dim Gabarity As GabarityZagotovki
    Gabarity = RaschotGabaritov(Dlina, Shirina, Vysota)
    
    Dim tOchistka As Double
    With Gabarity
        tOchistka = 0.0549 * (.Dlina + .Shirina + .Tolshchina) ^ 0.45 * KolVoDetaleyVSborke ^ 0.73
    End With
    
    Dim kVruchnuyu As Double
    kVruchnuyu = IIf(OchistkaVruchnuyu, 1.2, 1)
    
    PeskostruynayaOchistkaSvarnukhIzdeleiy = kVruchnuyu * tOchistka
    
End Function

Function PravkaZagIzListMatVRuch_Range(a As Range, L As Range, iSurf As Range, t As Double, Material As Integer, Optional StrealProgiba As Integer = 1, Optional PosleGibki As Boolean = False) As Double
    If a.Rows.Count = L.Rows.Count And a.Rows.Count = iSurf.Rows.Count Then
        Dim i As Long
        For i = 1 To a.Rows.Count
            PravkaZagIzListMatVRuch_Range = _
                PravkaZagIzListMatVRuch_Range + (PravkaZagIzListMatVRuch(a(i), L(1), t, Material, StrealProgiba, PosleGibki) * IIf(iSurf(i) = 0, 1, iSurf(i)))
        Next
    End If
End Function
Function PravkaZagIzListMatVRuch(a As Double, L As Double, t As Double, Material As Integer, Optional StrealProgiba As Integer = 1, Optional PosleGibki As Boolean = False) As Double
Attribute PravkaZagIzListMatVRuch.VB_Description = "Вырезание прокладок вручную. Карта 52."
Attribute PravkaZagIzListMatVRuch.VB_ProcData.VB_Invoke_Func = " \n19"
    
    'A - Ширина заготовки, мм
    'L - Длина заготовки, мм
    't - Толщина материала, мм
    
    'Material - Материал:
    '1 - Алюминиевые сплавы
    '2 - Медные сплавы
    '3 - Ст 35
    '4 - 40Х
    '5 - 20Х13
    '6 - 12Х18Н9Т
    '7 - ВТ1-0
    
    'StrealProgiba - Стрела прогиба на 1м, мм:
    '1 - 0,3 - 0,6 мм
    '2 - 0,6 - 1,0 мм
    '3 - 1,0 - 2,0 мм
    '4 - 2,0 - 3,0 мм
    '5 - свыше 3,0 мм
    
    'PosleGibki - Правка деталей после гибки
    
    If a > 0 And L > 0 And t > 0 And Material >= 1 And Material <= 7 Then
    
        Dim X As Double, Y As Double, z As Double, W As Double
        
        Dim MinA As Integer: MinA = 25
        Dim minL As Integer: minL = 120
        Dim CalcA As Double, CalcL As Double
        CalcA = WorksheetFunction.Min(a, L): If CalcA < MinA Then CalcA = MinA
        CalcL = WorksheetFunction.Max(a, L): If CalcL < minL Then CalcL = minL
        
        If t <= 2 Then
        
            Select Case Material
                Case 1
                    X = 5.82432275922029E-04
                    Y = 0.740070046665948
                    z = 0.647518865728401
                    W = 0.45459395648134
                Case 2
                    X = 6.08245250859789E-04
                    Y = 0.746476198065395
                    z = 0.665373012893885
                    W = 0.411977696164669
                Case Else
                    X = 6.0591198480024E-04
                    Y = 0.773305678701967
                    z = 0.678351242583209
                    W = 0.424315299703244
            End Select
            PravkaZagIzListMatVRuch = (X * CalcA ^ Y * CalcL ^ z) / t ^ W
            
        Else
        
            Select Case Material
                Case 1
                    X = 2.92168253028256E-04
                    Y = 0.708007645652566
                    z = 0.682790989971501
                    W = 0.584962657462639
                Case 2
                    X = 3.6044078569642E-04
                    Y = 0.740176096982452
                    z = 0.656022951481244
                    W = 0.584962500721156
                Case Else
                    X = 3.82722492181706E-04
                    Y = 0.773305680619784
                    z = 0.653297650211473
                    W = 0.584962500721156
            End Select
            PravkaZagIzListMatVRuch = (X * CalcA ^ Y * CalcL ^ z) * t ^ W
        
        End If
        
        Dim K1 As Double
        K1 = KoefMaterial_SlesarRezbonarez(CInt(Material))
        
        Dim K2 As Double
        Select Case StrealProgiba
            Case 1: K2 = 1.5
            Case 2: K2 = 1.3
            Case 3: K2 = 1
            Case 4: K2 = 0.8
            Case 5: K2 = 0.7
        End Select
        
        PravkaZagIzListMatVRuch = PravkaZagIzListMatVRuch * K1 * K2 * IIf(PosleGibki, 1.3, 1)
    
    End If

End Function

Function ZachistkaShlifovalnoyShkurkoy(Material As Integer, ByVal DlinaPoverkhnosti As Double, ByVal ShirinaPoverkhnosti As Double, _
    Optional TipPoverhnosti As Integer = EnumTipPoverhnosti.Ploskaya, _
    Optional TipShtrikha As Integer = EnumTipShtrikha.Pryamoy, _
    Optional Ra As Double = 1.25, Optional GlubokieRiski As Boolean = True) As Double
    'Зачистка поверхности детали шлифовальной шкуркой. Карта 13
    
    Dim ZachistkaPoverkhnosti As New clsZachistkaPoverkhnosti
    
    With ZachistkaPoverkhnosti
        .Dlina = DlinaPoverkhnosti
        .GlubokieRiski = GlubokieRiski
        .Material = Material
        .Ra = Ra
        .Shirina = ShirinaPoverkhnosti
        .SposobZachistki = Shkurkoy
        .TipPoverhnosti = TipPoverhnosti
        .TipShtrikha = TipShtrikha
        
        ZachistkaShlifovalnoyShkurkoy = .tZachistka
    End With
    
End Function

Function ZachistkaShlifovalnoyMashinkoy(Material As Integer, ByVal DlinaPoverkhnosti As Double, ByVal ShirinaPoverkhnosti As Double, _
    Optional TipPoverhnosti As Integer = EnumTipPoverhnosti.Ploskaya, _
    Optional Ra As Double = 1.25, Optional GlubokieRiski As Boolean = True) As Double
    'Зачистка поверхности детали шлифовальной машинкой. Карта 13
    
    Dim ZachistkaPoverkhnosti As New clsZachistkaPoverkhnosti
    
    With ZachistkaPoverkhnosti
        .Dlina = DlinaPoverkhnosti
        .GlubokieRiski = GlubokieRiski
        .Material = Material
        .Ra = Ra
        .Shirina = ShirinaPoverkhnosti
        .SposobZachistki = ShifovalnoyMashinkoy
        .TipPoverhnosti = TipPoverhnosti
        .TipShtrikha = Putannyy
        
        ZachistkaShlifovalnoyMashinkoy = .tZachistka

    End With
    
End Function



