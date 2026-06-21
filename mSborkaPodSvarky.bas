Attribute VB_Name = "mSborkaPodSvarky"
Option Explicit

Function SborkaMetallokonstruktsiyIzListovogoMetalla(MassaIzdeliya As Double, KolVoSobiraemykhDetaley As Double) As Double
    'НОРМАТИВЫ ВРЕМЕНИ НА СБОРКУ МЕТАЛЛОКОНСТРУКЦИЙ ПОД СВАРКУ. КРАМАТОРСК 1988
    'Карта 16. Сборка металлоконструкций из листового металла.
    'Содерражине работы:
    '1. Подать детали узлы к месту сборки; 2. Установить базовые детали, узлы на место сборки; 3. Разметить, установить детали или узлы, выдержав размеры по чертежу;
    '4. Подогнать сопряжение деталей и прихватить электросваркой, зачистить после прихватки; 5. Сдать металлоконструкцию ОТК.
    
    Dim X As Double, Y As Double, z As Double
    
    Select Case MassaIzdeliya
        Case Is <= 100
            X = 0.036
            Y = 0.25
            z = 0.48
        Case Is <= 1000
            X = 0.037
            Y = 0.72
            z = 0.51
        Case Is <= 15000
            X = 0.0117
            Y = 0.56
            z = 0.52
        Case Else
            X = 0.0115
            Y = 0.59
            z = 0.51
    End Select
    
    Dim kTochnost As Double
    kTochnost = 1.2
    
    Dim kMekhObrabotka As Double
    kMekhObrabotka = 1.4

    Dim kPolozheniyaUdobstva As Double
    kPolozheniyaUdobstva = 1.4
    
    SborkaMetallokonstruktsiyIzListovogoMetalla = 60 * kPolozheniyaUdobstva * kTochnost * kMekhObrabotka * (X * MassaIzdeliya ^ Y * KolVoSobiraemykhDetaley ^ z)

End Function

Function PravkaSvarnykhUzlovVruchnuyuNaPlite(Material As EnumMaterialy, Dlina As Double, Shirina As Double, Tolshchina As Double) As Double
    'ОБЩЕМАШИНОСТРОИТЕЛЬНЫЕ НОРМАТИВЫ ВРЕМЕНИ НА СЛЕСАРНО-СБОРОЧНЫЕ РАБОТЫ ПРИ СБОРКЕ МЕТАЛЛОКОНСТРУКЦИЙ ПОД СВАРКУ
    '(УТВ. ГОСКОМТРУДОМ СССР)
    'Карта 38. Правка сварных узлов вручную на плите
   
    Dim kMaterial As Double
    If Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.2
    ElseIf Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Or STAL_UGLERODISTAYA Then
        kMaterial = 1
    Else
        Exit Function
    End If
        
    If ParametyBolsheNulya(Array(Dlina, Shirina, Tolshchina)) = False Then Exit Function

    Dim kPolnayaPravka As Double
    kPolnayaPravka = 1.6
    
    Dim kVysota As Double
    kVysota = 1.2
    
    Dim kNeudobnyyeUsloviya As Double
    kNeudobnyyeUsloviya = 1.1
    
    
    Dim tPravka As Double
    tPravka = 0.0575 * Dlina ^ 0.37 * Shirina ^ 0.23 * Tolshchina ^ 0.21
    
    PravkaSvarnykhUzlovVruchnuyuNaPlite = WorksheetFunction.Product(kMaterial, kPolnayaPravka, kVysota, kNeudobnyyeUsloviya, tPravka)
    
End Function
