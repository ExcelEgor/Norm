Attribute VB_Name = "mKim"
Option Explicit
Option Private Module

Public Const KIM_T_VSPOMOGATELNOYE As Double = 0.05
Public Const KIM_T_ANALIZA_RAZMERA As Integer = 1

Function KIM_TpzOsnovnoe(Sloznost As Integer) As Integer
    
    Select Case Sloznost
        Case 1: KIM_TpzOsnovnoe = 15
        Case 2: KIM_TpzOsnovnoe = 30
        Case 3: KIM_TpzOsnovnoe = 40
    End Select
    
End Function

Function KIM_KolVoSecheniy(L As Double, TipDopuska As Integer) As Double
    
    'TipDopuska
    '1-Размер
    '2-Диаметр, расположение
    '3-Форма
 
    Dim MinKolVoSecheniy As Integer
    MinKolVoSecheniy = 3

    Dim Shag As Double
    Select Case TipDopuska
        Case 1
            KIM_KolVoSecheniy = MinKolVoSecheniy
        Case 2, 3
            Select Case L
                Case Is <= 7:   Shag = 1
                Case Is <= 8:   Shag = 1.25
                Case Is <= 10:  Shag = 1.5
                Case Is <= 12:  Shag = 1.75
                Case Is <= 16:  Shag = 2
                Case Is <= 22:  Shag = 2.5
                Case Is <= 24:  Shag = 3
                Case Is <= 30:  Shag = 3.5
                Case Is <= 36:  Shag = 4
                Case Is <= 42:  Shag = 4.5
                Case Is <= 52:  Shag = 5
                Case Is <= 56:  Shag = 5.5
                Case Else:      Shag = 6
            End Select
            
            If TipDopuska = 3 Then Shag = Shag * 0.5
            If Shag < 1 Then Shag = 1
            
            KIM_KolVoSecheniy = WorksheetFunction.RoundUp(L / Shag, 0)
            If KIM_KolVoSecheniy < MinKolVoSecheniy Then KIM_KolVoSecheniy = MinKolVoSecheniy
            
    End Select
    
End Function

Function KIM_RaschotTshtIzParametrov(ParametryKIM As clsKIM) As Double

    Dim Vremya As Double
    With ParametryKIM
        Select Case .id
            Case "pHera_Ustanovka"
                Vremya = KIM_Hera_Ustanovka(.Massa)
            
            Case "pHera_Cilindr"
                Vremya = KIM_Hera_RaschotTsilindra(.D1, .Dlina, .DopuskRazmera, .DopuskFormy, .DopuskRaspolozheniya, .OtOsi)
                
            Case "pHera_Ploskost"
                Vremya = KIM_Hera_RaschotPloskosti(.Shirina, .Dlina, .TrigernyyRezhim, .DopuskFormy, .DopuskRaspolozheniya, .DopuskRazmera)
            
            Case "pHera_Krug"
                Vremya = KIM_Hera_RaschotKrugloyPloskosti(.D1, .d2, .TrigernyyRezhim, .DopuskFormy, .DopuskRaspolozheniya, .DopuskRazmera)
            
            Case "pKruglomer_Ustanovka"
                Vremya = KIM_Kruglomer_Ustanovka(.Massa)

            Case "pKruglomer_Vyravnivaniye"
                Vremya = KIM_Kruglomer_Vyravnivaniye(.Dlina)
            
            Case "pKruglomer_Izmereniye"
                Vremya = KIM_Kruglomer_Izmereniye(.Dlina, .RadialnoyeIzmereniye, .LieneynoyeIzmereniye)
            
            Case "pMikroskop_Ustanovka"
                Vremya = KIM_Mikroskop_Ustanovka(.Massa)
            
            Case "pMikroskop_Kontur"
                Vremya = KIM_Mikroskop_IzmereniyeKontura(.Dlina)
            
            Case "pMikroskop_Diametr"
                Vremya = KIM_Mikroskop_IzmereniyeKontura(.D1)
            
        End Select
        
        Dim tObsuzhOtdykhaLichnNadobn As Double
        tObsuzhOtdykhaLichnNadobn = 1.1
        
        Vremya = tObsuzhOtdykhaLichnNadobn * .KolVoPoverkhnostey * Vremya
    
    End With
    
    KIM_RaschotTshtIzParametrov = WorksheetFunction.RoundUp(Vremya, 2)

End Function




