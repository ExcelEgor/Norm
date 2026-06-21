VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPritirkaPloskikhPov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const RA_MIN As Double = 0.01
Private Const OTLONENIYE_MIN As Double = 0.001

Private pMaterial As EnumMaterialy
Private pPloshchadSM2 As Double
Private pRa As Double
Private pPripusk As Double
Private pOtkloneniye As Double
Private pParallelnostPerpendikulyarnost As Boolean
Private pPreryvistayaIliAzotirovaniye As Boolean
Private pPritirkaVUporIliPaza As Boolean
Private pPritirkaNaStanke As Boolean
Private pTolshchina As Double

Private pPripuskMax As Double

Public Property Let Material(Value As EnumMaterialy):                   pMaterial = Value: End Property

Public Property Let PloshchadSM2(Value As Double):                      pPloshchadSM2 = Value: End Property

Public Property Let Otkloneniye(Value As Double):                       pOtkloneniye = Value: End Property

Public Property Let ParallelnostPerpendikulyarnost(Value As Boolean):   pParallelnostPerpendikulyarnost = Value: End Property

Public Property Let PreryvistayaIliAzotirovaniye(Value As Boolean):     pPreryvistayaIliAzotirovaniye = Value: End Property

Public Property Let PritirkaVUporIliPaza(Value As Boolean):             pPritirkaVUporIliPaza = Value: End Property

Public Property Let PritirkaNaStanke(Value As Boolean):                 pPritirkaNaStanke = Value: End Property

Public Property Let Tolshchina(Value As Double):                        pTolshchina = Value: End Property

Public Property Let Pripusk(Value As Double):                           pPripusk = Value: End Property

Public Property Let Ra(Value As Double)
    pRa = Value
    pPripuskMax = IIf(pRa < 0.025, 0.02, 0.1)
End Property

Function Vremya() As Double

    If pPloshchadSM2 <= 0 Then Exit Function

    If pRa < RA_MIN Then
        Err.Raise vbObjectError + 1, "PritirkaPloskikhPoverkhnostey", "Значение шерозоватости не может быть меньше " & RA_MIN
    End If
    
    If pPripusk > pPripuskMax Then
        Err.Raise vbObjectError + 1, "PritirkaPloskikhPoverkhnostey", "Превышено значение припуска"
    End If
    
    If pPloshchadSM2 > PloshchadMax() Then
        Err.Raise vbObjectError + 1, "PritirkaPloskikhPoverkhnostey", "Превышено значение площади"
    End If
    
    If pOtkloneniye < OTLONENIYE_MIN Then
        Err.Raise vbObjectError + 1, "PritirkaPloskikhPoverkhnostey", "Значение отклонения не может быть меньше " & OTLONENIYE_MIN
    End If
            
    Dim Normativ
    Normativ = ZagruzitNormativ("slesar_pritir_plosk")
    
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) >= pPloshchadSM2 _
            And Normativ(i, 2) <= pRa _
            And Normativ(i, 3) <= pOtkloneniye _
            And Normativ(i, 4) >= pPripusk Then
            Vremya = Normativ(i, 5) * Koeffitsient()
            Exit For
        End If
    Next
    

End Function

Private Function Koeffitsient() As Double
    
    Dim K(1 To 6) As Double
    
    K(1) = IIf(pParallelnostPerpendikulyarnost, 1.3, 1)
    K(2) = IIf(pPreryvistayaIliAzotirovaniye, 1.2, 1)
    K(3) = IIf(pPritirkaVUporIliPaza, 1.5, 1)
    K(4) = IIf(pPritirkaNaStanke, 0.5, 1)
    
    If pMaterial = STAL_UGLERODISTAYA Or pMaterial = STAL_LEGIROVANNAYA Then
        K(5) = 1
    ElseIf pMaterial = MEDNYYE_SPLAVY Then
        K(5) = 0.8
    ElseIf pMaterial = ALUMINIYEVYYE_SPLAVY Then
        K(5) = 1
    ElseIf pMaterial = STAL_NERZHAVEYUSHCHAYA Then
        K(5) = 1.9
    Else
        Err.Raise vbObjectError + 1, "PritirDovodPloskPoverh", "Некорректный материал"
    End If

    Select Case pTolshchina
        Case Is <= 1:   K(6) = 1.4
        Case Is <= 3:   K(6) = 1.2
        Case Is <= 5:   K(6) = 1.1
        Case Else:      K(6) = 1
    End Select
        
    
    Koeffitsient = WorksheetFunction.Product(K)
    
End Function

Function PloshchadMax() As Double
    
    Dim Smax As Double
    
    Select Case pRa
        Case Is < 0.025
            Smax = IIf(pPripusk <= 0.01, 300, 100)
        Case Is < 0.1
            Select Case pPripusk
                Case Is <= 0.02
                    Smax = 500
                Case Is <= 0.04
                    Smax = 30
                Case Else
                    Smax = 100
            End Select
        Case Else
            Smax = 500
    End Select
    
    PloshchadMax = Smax
    
End Function

Function DopustimyyeZnacheniyaPripuska() As Variant

    Dim ArrPripusk
    ArrPripusk = Array(0.01, 0.02, 0.04, 0.1)
    
    Dim result() As Variant
    ReDim result(-1 To -1)

    
    Dim i As Integer
    For i = LBound(ArrPripusk) To UBound(ArrPripusk)
        If Not ArrPripusk(i) > pPripuskMax Then
            If UBound(result) = -1 Then
                ReDim result(0 To 0)
            Else
                ReDim Preserve result(0 To UBound(result) + 1)
            End If
            result(UBound(result)) = ArrPripusk(i)

        End If
    Next
    
    DopustimyyeZnacheniyaPripuska = result
        
End Function

