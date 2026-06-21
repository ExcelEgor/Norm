VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMillingContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Frezerovaniye

Option Explicit

' ==========================================================
' ПУБЛИЧНЫЕ СВОЙСТВА (Простые данные без сложной валидации)
' ==========================================================
' --- Оборудование и материал ---
Public TipStanka As EnumTipStankaRastFrez
Public Material As EnumMaterialy
Public CHPU As Boolean

' --- Геометрия обработки ---
Public DlinaPoverkhnosti As Double
Public ShirinaPoverkhnosti As Double

' --- Специфические условия ---
Public NaUdar As Boolean
Public Korka As Boolean
Public TonkostennayaKonfiguraciya As Boolean

' --- Параметры управления расчетом ---
Public PoRezhimam As Boolean
Public AvtoRezhimy As Boolean

' --- Данные, формируемые в процессе расчета ---
Public Dfrezy As Double
Public Rezhimy As clsMillingRegimes

' ==========================================================
' СКРЫТЫЕ ПОЛЯ (Для свойств с валидацией)
' ==========================================================
Private pPripusk As Double
Private pIT As Integer
Private pRa As Double

' ==========================================================
' ИНИЦИАЛИЗАЦИЯ КЛАССА
' ==========================================================
Private Sub Class_Initialize()
    ' Объект режимов создается автоматически при создании контекста
    Set Rezhimy = New clsMillingRegimes
    
    ' Задаем базовые значения, чтобы класс изначально был в безопасном состоянии
    pPripusk = MIN_PRIPUSK_RASTFREZ
    pIT = 11
End Sub

' ==========================================================
' СВОЙСТВА С АВТОКОРРЕКЦИЕЙ (Истинная инкапсуляция)
' ==========================================================
Public Property Let Pripusk(ByVal Value As Double)
    ' Если припуск 0 или отрицательный, ставим минимальный
    pPripusk = IIf(Value <= 0, MIN_PRIPUSK_RASTFREZ, Value)
End Property

Public Property Get Pripusk() As Double
    Pripusk = pPripusk
End Property

Public Property Let IT(ByVal Value As Integer)
    ' Если квалитет не задан, ставим 11 по умолчанию
    pIT = IIf(Value <= 0, 11, Value)
    
    ' Если Ra уже был задан ранее, пересчитываем его под новый квалитет
    If pRa > 0 Then pRa = MinRa_RastFrez(pRa, pIT)
End Property

Public Property Get IT() As Integer
    IT = pIT
End Property

Public Property Let Ra(ByVal Value As Double)
    ' Корректируем Ra с учетом текущего квалитета
    pRa = MinRa_RastFrez(Value, pIT)
End Property

Public Property Get Ra() As Double
    Ra = pRa
End Property

' ==========================================================
' ВАЛИДАЦИЯ КОНТЕКСТА (Self-validating object)
' ==========================================================
Public Property Get IsValid() As Boolean
    
    ' 1. Базовая геометрия детали не может быть нулевой или отрицательной
    If DlinaPoverkhnosti <= 0 Or ShirinaPoverkhnosti <= 0 Or pPripusk <= 0 Then
        IsValid = False
        Exit Property
    End If
    
    ' 2. Проверка диаметра фрезы
    ' Если режимы задаются вручную, диаметр уже должен быть определен
    If PoRezhimam And Not AvtoRezhimy Then
        If Dfrezy <= 0 Then
            IsValid = False
            Exit Property
        End If
    End If
    
    ' Дополнительные проверки можно добавлять сюда
    ' If MaterialKorrektnyy_RastFrez(Material) = False Then ...
    
    ' Если все проверки пройдены
    IsValid = True
    
End Property
