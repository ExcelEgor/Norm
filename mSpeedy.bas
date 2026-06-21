Attribute VB_Name = "mSpeedy"
Option Explicit

Function VyrezkaSpeedy(Material As EnumMaterialy, Tolshchina As Double, DlinaVyrezki As Double) As Double
    
    If Tolshchina > RaschotMakcsimalnoyTolshchiny_Speedy(Material) Then Exit Function
    
    VyrezkaSpeedy = 1.1 * (DlinaVyrezki / RaschotMinutnoyPodachi_Speedy(Material, Tolshchina))
    
End Function

Function RaschotMinutnoyPodachi_Speedy(Material As EnumMaterialy, Tolshchina As Double) As Double

    If Tolshchina > RaschotMakcsimalnoyTolshchiny_Speedy(Material) Then Exit Function
    
    Dim SmmMin As Double
    Dim Ks As Double 'Поправочный коэффициент на подачу, полученный опытным путем
    
    Ks = 1
    
    If Material = PARONIT Then
        Select Case Tolshchina
            Case Is <= 1:   SmmMin = 206
            Case Is <= 2:   SmmMin = 76
            Case Else:      SmmMin = 38
        End Select

    ElseIf Material = KARTON Then
        Select Case Tolshchina
            Case Is <= 0.5: SmmMin = 635
            Case Is <= 1:   SmmMin = 588
            Case Else:      SmmMin = 588
        End Select

    ElseIf Material = FIBRA Then
        Select Case Tolshchina
            Case Is <= 0.6: SmmMin = 611
            Case Is <= 1:   SmmMin = 588
            Case Is <= 1.5: SmmMin = 567
            Case Else:      SmmMin = 397
        End Select
    
    ElseIf Material = REZINA Or Material = ORGSTEKLO Then
        Select Case Tolshchina
            Case Is <= 1:   SmmMin = 645
            Case Is <= 2:   SmmMin = 567
            Case Is <= 3
                SmmMin = 135
                Ks = 0.7
            Case Else:      Exit Function
        End Select
               
        If Material = ORGSTEKLO Then
            Ks = Ks * 1.5 'Для толщицы 2 коэф. 1.6 относительно резины. Для других толщин неизвестно
        End If

    Else
        Exit Function
    End If
    
    RaschotMinutnoyPodachi_Speedy = Ks * SmmMin
    
End Function

Function RaschotMakcsimalnoyTolshchiny_Speedy(Material As EnumMaterialy) As Double
    
    Dim hMax As Double
    
    If Material = PARONIT Then
        hMax = 3

    ElseIf Material = KARTON Then
        hMax = 1.5

    ElseIf Material = FIBRA Then
        hMax = 2.5
    
    ElseIf Material = REZINA Or Material = ORGSTEKLO Then
        hMax = 3

    Else
        Exit Function
    End If
    
    RaschotMakcsimalnoyTolshchiny_Speedy = hMax
    
End Function
