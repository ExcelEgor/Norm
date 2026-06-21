Attribute VB_Name = "mShlifovaniye_KoefPublic"
Option Explicit
Option Private Module

'Материалы
'1 Сталь 45
'2 25Х13Н2
'3 ШХ15
'4 40Х13
'5 АНК04
'6 ВНМ3


Function GlavnyyKoeffitsient_Shlifovaniye(KolVoDetaleyVPartii As Double, NezhestkayaSPID As Boolean, VozrastStanka As Double, ShlifovaniyeDvukhPoverhnostey As Boolean) As Double

    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
    
    Select Case KolVoDetaleyVPartii
        Case Is <= 3:       K1 = 1.3
        Case Is <= 20:      K1 = 1
        Case Is <= 100:     K1 = 0.8
        Case Else:          K1 = 0.6
    End Select
    
    K2 = IIf(NezhestkayaSPID, 1.3, 1)
    
    Select Case VozrastStanka
        Case Is <= 10:  K3 = 1
        Case Is <= 20:  K3 = 1.2
        Case Is > 20:   K3 = 1.4
    End Select
    
    K4 = IIf(ShlifovaniyeDvukhPoverhnostey, 1.2, 1)
    
    GlavnyyKoeffitsient_Shlifovaniye = K1 * K2 * K3 * K4

End Function

Private Function Koeffitsient_VosrastStanka(VozrastStanka As Double) As Double
    Select Case VozrastStanka
        Case Is <= 10:  Koeffitsient_VosrastStanka = 1
        Case Is <= 20:  Koeffitsient_VosrastStanka = 1.2
        Case Is > 20:   Koeffitsient_VosrastStanka = 1.4
    End Select
End Function

Private Function Koeffitsient_ShirinaKruga(ShirinaKruga As Double) As Double
    Select Case ShirinaKruga
        Case Is <= 50: Koeffitsient_ShirinaKruga = 1
        Case Is <= 63: Koeffitsient_ShirinaKruga = 0.85
        Case Is <= 80: Koeffitsient_ShirinaKruga = 0.7
    End Select
End Function

Private Function Koeffitsient_NaUdarIliVUpor(NaUdarIliVUpor As Boolean) As Double
    Koeffitsient_NaUdarIliVUpor = IIf(NaUdarIliVUpor, 1.15, 1)
End Function

Private Function Koeffitsient_NezhestkayaSPID(NezhestkayaSPID As Boolean) As Double
    Koeffitsient_NezhestkayaSPID = IIf(NezhestkayaSPID, 1.3, 1)
End Function

Private Function Koeffitsient_ShlifovaniyeDvukhPoverhnostey(ShlifovaniyeDvukhPoverhnostey As Boolean) As Double
    Koeffitsient_ShlifovaniyeDvukhPoverhnostey = IIf(ShlifovaniyeDvukhPoverhnostey, 1.2, 1)
End Function

