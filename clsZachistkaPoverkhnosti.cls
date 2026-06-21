VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZachistkaPoverkhnosti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Зачистка поверхности детали шлифовальной шкуркой. Карта 13

Private pMaterial As EnumMaterialy
Private pDlina As Double
Private pShirina As Double
Private pTipPoverhnosti As EnumTipPoverhnosti
Private pTipShtrikha As EnumTipShtrikha
Private pRa As Double
Private pGlubokieRiski As Boolean
Private pSposobZachistki As SposobZachistki

Public Enum EnumTipShtrikha
    Pryamoy = 1
    Putannyy = 2
End Enum

Public Enum EnumTipPoverhnosti
    Ploskaya = 1
    Tsilindricheskaya = 2
End Enum

Public Enum SposobZachistki
    Shkurkoy = 1
    ShifovalnoyMashinkoy = 2
End Enum

Public Property Let Material(Value As EnumMaterialy): pMaterial = Value: End Property

Public Property Let Dlina(Value As Double): pDlina = Value: End Property
Public Property Get Dlina() As Double: Dlina = pDlina: End Property

Public Property Let Shirina(Value As Double): pShirina = Value: End Property
Public Property Get Shirina() As Double: Shirina = pShirina: End Property

Public Property Let TipPoverhnosti(Value As EnumTipPoverhnosti): pTipPoverhnosti = Value: End Property

Public Property Let TipShtrikha(Value As EnumTipShtrikha): pTipShtrikha = Value: End Property

Public Property Let Ra(Value As Double): pRa = Value: End Property

Public Property Let GlubokieRiski(Value As Boolean): pGlubokieRiski = Value: End Property

Public Property Let SposobZachistki(Value As SposobZachistki): pSposobZachistki = Value: End Property


Public Function tZachistka() As Double

    If pDlina <= 0 Or pShirina <= 0 Then
        Err.Raise vbObjectError + 1, "ZachistkaShlifovalnoyShkurkoy", "Длина и ширина должны быть положительными числами больше 0"
    End If
    
    
    Dim Koeffitsient As Double
    Koeffitsient = WorksheetFunction.Product(kMaterial, kTipShtrikha, kRa, kGlubina, kPoverhnost, kSposob)
    
    tZachistka = Koeffitsient * 0.002 * pDlina ^ 0.58 * pShirina ^ 0.63
        
End Function
    
Private Property Get kMaterial() As Double

    Select Case pMaterial
    
        Case EnumMaterialy.MEDNYYE_SPLAVY
            kMaterial = 1
            
        Case EnumMaterialy.ALUMINIYEVYYE_SPLAVY, EnumMaterialy.TEKSTOLIT, EnumMaterialy.STEKLOTEKSTOLIT, EnumMaterialy.ORGSTEKLO, EnumMaterialy.GETINAKS
            kMaterial = 1.2
        
        Case EnumMaterialy.STAL_UGLERODISTAYA, EnumMaterialy.STAL_LEGIROVANNAYA, EnumMaterialy.STAL_NERZHAVEYUSHCHAYA
            kMaterial = 1.4
        
        Case Else
            Err.Raise vbObjectError + 1, "ZachistkaShlifovalnoyShkurkoy", "Недопустимый материал"
    End Select
End Property

Private Property Get kTipShtrikha() As Double

    If pSposobZachistki = ShifovalnoyMashinkoy Then pTipShtrikha = Putannyy
    
    If pTipShtrikha = Pryamoy Then
        kTipShtrikha = 1
    ElseIf pTipShtrikha = Putannyy Then
        kTipShtrikha = 0.8
    Else
        Err.Raise vbObjectError + 1, "ZachistkaShlifovalnoyShkurkoy", "Недопустимый тип штриха"
    End If
    
End Property

Private Property Get kPoverhnost() As Double
    If pTipPoverhnosti = Ploskaya Then
        kPoverhnost = 1
    ElseIf pTipPoverhnosti = Tsilindricheskaya Then
        kPoverhnost = 1.2
    Else
        Err.Raise vbObjectError + 1, "ZachistkaShlifovalnoyShkurkoy", "Недопустимый тип поверхности"
    End If
End Property

Private Property Get kRa() As Double
    If pRa < 1.25 Then
        Err.Raise vbObjectError + 1, "ZachistkaShlifovalnoyShkurkoy", "Ra должно быть числом >=1,25"
    Else
        kRa = IIf(pRa > 2.5, 1, 1.5)
    End If
End Property

Private Property Get kGlubina() As Double
    kGlubina = IIf(pGlubokieRiski = True, 1.2, 1)
End Property

Private Property Get kSposob() As Double
    kSposob = IIf(pSposobZachistki = ShifovalnoyMashinkoy, 0.8, 1)
End Property


