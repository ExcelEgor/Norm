Attribute VB_Name = "mZagotovitelnyye"
Option Explicit
Public Const MAX_DLINA_CIRCULYARKA As Integer = 1100
Public Const MAX_SHIRINA_GILOTIONA As Integer = 1500
Public Const MAX_DLINA_GILOTIONA As Integer = 3000
Public Const MAX_TOLSHCHINA_GILOTIONA As Integer = 16

'ЗАГОТОВИТЕЛЬНЫЕ РАБОТЫ. УКРУПНЁННЫЕ ОДНОСТРОЧНЕ НОРМАТИВЫ ВРЕМЕНИ. МЕЛКОСЕРИЙНОЕ И ЕДИНИЧНОЕ ПРОИЗВОДСТВО. 1988 Г.

Function OtrezkaNaTsirkulyarnykhPilakh_Marshrut(Material As EnumMaterialy, Dlina As Double, Shirina As Double, Tolshchina As Double, KolVoDetVzagotovke As Double) As Variant

    Dim List As New clsFigura_List
    With List
        .Init Material, Shirina, Tolshchina, Dlina

        OtrezkaNaTsirkulyarnykhPilakh_Marshrut = .OtrezkaNaTsirkulyarnykhPilakh_Marshrut(KolVoDetVzagotovke)
    End With

End Function

Function OtrezkaTsirkulyarnoyPiloy_List(Material As EnumMaterialy, ByVal DlinaReza As Double, ByVal Tolshchina As Double) As Double
    'КАРТА 4. ОТРЕЗКА ЗАГОТОВОК ИЗ ЛИСТОВОГО И ПРОФИЛЬНОГО МАТЕРИАЛА НА ЦИРКУЛЯРНЫХ ПИЛАХ (ЛИСТ)
    
    If PolozhitelnyyeChisla(DlinaReza, Tolshchina) = False Or DlinaReza > MAX_DLINA_CIRCULYARKA Then Exit Function
    If Tolshchina > OtrezkaTsirkulyarnoyPiloy_List_MaksTolshchina(Material) Then Exit Function

    Dim a As Double, b As Double, c As Double
    Select Case Material
        Case STEKLOTEKSTOLIT, ORGSTEKLO
            a = 0.002:  b = 0.824:  c = 0.728
            
        Case ALUMINIYEVYYE_SPLAVY, GETINAKS
            a = 0.0016: b = 0.846:  c = 0.73
            
        Case MEDNYYE_SPLAVY, TEKSTOLIT
            a = 0.0014: b = 0.821:  c = 0.727
            
        Case Else
            Exit Function
    End Select

    
    Const MinTolshchina As Integer = 3
    If Tolshchina < MinTolshchina Then Tolshchina = MinTolshchina

    Const MinDlinaReza = 20
    If DlinaReza < MinDlinaReza Then DlinaReza = MinDlinaReza
    
    OtrezkaTsirkulyarnoyPiloy_List = a * DlinaReza ^ b * Tolshchina ^ c
    
End Function
Function OtrezkaTsirkulyarnoyPiloy_List_MaksTolshchina(Material As EnumMaterialy) As Integer
    'КАРТА 4. ОТРЕЗКА ЗАГОТОВОК ИЗ ЛИСТОВОГО И ПРОФИЛЬНОГО МАТЕРИАЛА НА ЦИРКУЛЯРНЫХ ПИЛАХ (ЛИСТ).
    'Вспомогательная функция для определения максимальной толщины отрезки.
    
    Dim MaksTolshchina As Integer
    Select Case Material
        Case STEKLOTEKSTOLIT, ORGSTEKLO
            MaksTolshchina = 15
            
        Case ALUMINIYEVYYE_SPLAVY, GETINAKS
            MaksTolshchina = 45
            
        Case MEDNYYE_SPLAVY, TEKSTOLIT
            MaksTolshchina = 60
            
        Case Else
            Err.Raise vbObjectError + 1, "OtrezkaTsirkulyarnoyPiloy_List_MaksTolshchina", "Недопустимый материал"
            Exit Function
    End Select
    
    OtrezkaTsirkulyarnoyPiloy_List_MaksTolshchina = MaksTolshchina
    
End Function

Function OtrezkaNaGilotinNozhnitsZaDvaReza(Material As EnumMaterialy, ByVal Dlina As Double, ByVal Shirina As Double, ByVal Tolshcina As Double) As Double
    'КАРТА 9. ОТРЕЗКА ЗАГОТОВОК НА ГИЛЬОТИННЫХ НОЖНИЦАХ ПО РАЗМЕТКЕ ЗА ДВА РЕЗА
    
    If Dlina <= 0 Or Shirina <= 0 Or Tolshcina <= 0 Then
        Err.Raise vbObjectError + 1, "OtrezkaNaGilotinNozhnitsZaDvaReza", "Длина, ширина и толщина должны быть положительными числами >0"
    End If
    
    Dim a As Double, X As Double, Y As Double, z As Double
    Select Case Material
        Case EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA, EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.MEDNYYE_SPLAVY
            a = 0.004:  X = 0.409: Y = 0.342: z = 0.87
        Case EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.REZINA
            a = 0.0036: X = 0.409: Y = 0.339: z = 0.821
        Case EnumMaterialy.KARTON, EnumMaterialy.GETINAKS, EnumMaterialy.STEKLOTEKSTOLIT
            a = 0.0036: X = 0.411: Y = 0.344: z = 0.703
        Case Else
            Err.Raise vbObjectError + 1, "OtrezkaNaGilotinNozhnitsZaDvaReza", "Недопустимый материал"
            Exit Function
    End Select

    Dlina = WorksheetFunction.Max(Array(Dlina, Shirina))
    Shirina = WorksheetFunction.Min(Array(Dlina, Shirina))
    
    If Dlina > MAX_DLINA_GILOTIONA Then
        Err.Raise vbObjectError + 1, "OtrezkaNaGilotinNozhnitsZaDvaReza", "Длина заготовки не должна превышать " & MAX_DLINA_GILOTIONA
    ElseIf Shirina > MAX_SHIRINA_GILOTIONA Then
        Err.Raise vbObjectError + 1, "OtrezkaNaGilotinNozhnitsZaDvaReza", "Ширина заготовки не должна превышать " & MAX_SHIRINA_GILOTIONA
    ElseIf Tolshcina > MAX_TOLSHCHINA_GILOTIONA Then
        Err.Raise vbObjectError + 1, "OtrezkaNaGilotinNozhnitsZaDvaReza", "Толщина заготовки не должна превышать " & MAX_TOLSHCHINA_GILOTIONA
    End If

    Dim Shirina_Min As Integer
    Shirina_Min = 60
    If Shirina < Shirina_Min Then Shirina = Shirina_Min
    
    Dim Dlina_Min As Integer
    Dlina_Min = 150
    If Dlina < Dlina_Min Then Dlina = Dlina_Min
    
    Dim Tolshcina_Min As Integer
    Tolshcina_Min = 2
    If Tolshcina < Tolshcina_Min Then Tolshcina = Tolshcina_Min
    
    Dim tOtrezka As Double
    tOtrezka = a * Dlina ^ X * Shirina ^ Y * Tolshcina ^ z
    
    Dim tRazmetka As Double
    Const kVypolneniyeRabotySlesarem As Double = 1.15
    tRazmetka = kVypolneniyeRabotySlesarem * RazmetkaKonturaBezShablona_Lineykoy(Material, MassaLista(Dlina, Shirina, Tolshcina, Material), 2 * (Dlina + Shirina))
    
    OtrezkaNaGilotinNozhnitsZaDvaReza = tRazmetka + tOtrezka

End Function

