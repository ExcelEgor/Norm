Attribute VB_Name = "mDatron_Drugiye"
Option Explicit

Private Const K_OBS As Double = 1.1
Private Const T_VSPOM As Double = 0.3

Public Enum EnumUstanovkaDatron
    VakuumnyyStol = 1
    Tiski = 2
    Prisposobleniye = 3
    Patron4Os = 4
End Enum

Function Datron_RaschotTshIzTablitsy(Material As EnumMaterialy, TipProizvodstva As Integer, Zagolovki As Range, StorkaDannykh As Range) As Variant

    Dim Parametry As New clsDatron
    Set Parametry = Parametry.PrisvoitParametryIzTablitsy(Material, Zagolovki, StorkaDannykh)
    
    Dim Normy As NormyVremeni
    Normy = Parametry.RaschotIzParametrov
    
    Dim kTipProizvodstva As Double
    Select Case TipProizvodstva
        Case 2:     kTipProizvodstva = 1
        Case 3:     kTipProizvodstva = 0.7
        Case Else:  kTipProizvodstva = 1.3
    End Select
    
    Dim ArrNormy(1 To 3) As Double
    ArrNormy(1) = WorksheetFunction.RoundUp(Normy.tMekhanika, 2)
    ArrNormy(2) = WorksheetFunction.RoundUp(kTipProizvodstva * Normy.tSlesar, 2)
    ArrNormy(3) = WorksheetFunction.RoundUp(kTipProizvodstva * Normy.tKontrol, 2)
    
    Datron_RaschotTshIzTablitsy = ArrNormy
    
End Function

Public Function Datron_Ustanovka(ByVal Massa As Double, SposobUstanovki As EnumUstanovkaDatron) As Double
    
    Const MAKS_MASSA As Integer = 20
    
    If Massa <= 0 Or Massa > MAKS_MASSA Then Exit Function
    If Massa < 0.3 Then Massa = 0.3
    
    Dim tUst As Double
    
    If SposobUstanovki = VakuumnyyStol Then
        tUst = Ustanov_NaStoleSUporomBezKrepleniya(Massa, True)
    ElseIf SposobUstanovki = Tiski Then
        tUst = Ustanov_Vtiskah(Massa, 1, 2)
    ElseIf SposobUstanovki = Prisposobleniye Then
        tUst = Ustanov_NaStoleKreplBoltPlank(Massa, 2, 0.1)
    ElseIf SposobUstanovki = Patron4Os Then
        tUst = Ustanov_ChetirehKulachkovyi_FasonnayaForma(Massa)
    Else
        Exit Function
    End If
    
    Datron_Ustanovka = K_OBS * (1 + tUst)
    
End Function

Public Function MaterialKorrektnyy_Datron(Material As EnumMaterialy) As Boolean

    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY, TEKSTOLIT, POLIAMID, STEKLOTEKSTOLIT, ORGSTEKLO
            MaterialKorrektnyy_Datron = True
        Case Else
            MaterialKorrektnyy_Datron = False
    End Select
    
    
End Function

Function RaschotTsht_Datron(tMash As Double) As Double
    RaschotTsht_Datron = Round(K_OBS * (tMash + T_VSPOM), 2)
End Function

