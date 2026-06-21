Attribute VB_Name = "mVspomVremya_Ustanovka"
Option Explicit

Public Const MaksMassaUstanov_Tok As Integer = 3000

Public Enum KharakterVyverki_SamotsentrPatron
    BezVyverki = 1
    sVyverkoyPoDiametruItortsu = 2
End Enum

Function tObsOtd_GorVertFrezStanki(LTable As Double) As Double
    'Время на обслуживание рабочего места, отдых и личные надобности в % от оперативного
    
    'LTable - длина стола, мм
    
    If LTable > 2500 Then
        tObsOtd_GorVertFrezStanki = CVErr(xlErrNA)
    Else
        Select Case LTable
            Case Is <= 750:     tObsOtd_GorVertFrezStanki = 4
            Case Is <= 1500:    tObsOtd_GorVertFrezStanki = 4.5
            Case Is <= 2500:    tObsOtd_GorVertFrezStanki = 5
        End Select
        tObsOtd_GorVertFrezStanki = tObsOtd_GorVertFrezStanki + 4
        tObsOtd_GorVertFrezStanki = tObsOtd_GorVertFrezStanki / 100
    End If
        
End Function

Function Ustanov_Vtiskah(Massa As Double, Optional Kreplenie As Integer = 1, Optional HarakterViverki As Integer = 2) As Double
Attribute Ustanov_Vtiskah.VB_Description = "Установка. В тисках"
Attribute Ustanov_Vtiskah.VB_ProcData.VB_Invoke_Func = " \n18"
    
    'Massa - Масса заготовки, кг
    
    'Kreplenie - Способ крепления детали:
    '1 - Винтом
    '2 - Эксцентриком
    '3 - Пневмогидравлическим цилиндром
    
    'HarakterViverki - Характер выверки:
    '1 - Без выверки
    '2 - С выверкой в одной плоскости с точностью 0,5 мм
    
    Dim a As Double, X As Double
    
    If Massa > 20 Or Not Massa > 0 Or Kreplenie > 3 Or Kreplenie < 1 Or HarakterViverki > 2 Or HarakterViverki < 1 Then Exit Function
    
    Select Case Kreplenie
        Case 1
            If Massa <= 0.1 Then
                a = IIf(HarakterViverki = 1, 0.2, 0.4)
                X = 0.09
            Else
                a = IIf(HarakterViverki = 1, 0.4, 0.96)
                X = IIf(HarakterViverki = 1, 0.24, 0.2)
            End If
        Case 2
            If Massa <= 0.1 Then
                a = IIf(HarakterViverki = 1, 0.132, 0.264)
                X = 0.1
            Else
                a = IIf(HarakterViverki = 1, 0.31, 0.69)
                X = IIf(HarakterViverki = 1, 0.29, 0.28)
            End If
        Case 3
            If Massa <= 0.1 Then
                a = IIf(HarakterViverki = 1, 0.115, 0.25)
                X = 0.11
            Else
                a = IIf(HarakterViverki = 1, 0.25, 0.56)
                X = IIf(HarakterViverki = 1, 0.26, 0.27)
            End If
    End Select
    
    If Massa <= 0.1 Then
        Ustanov_Vtiskah = a / (Massa ^ X)
    Else
        Ustanov_Vtiskah = a * Massa ^ X
    End If
    
End Function

Function Ustanov_NaStoleKreplBoltPlank(Massa As Double, Optional HarakterViverki As Integer = 2, Optional TochnostViverki As Double = 0.1) As Double
Attribute Ustanov_NaStoleKreplBoltPlank.VB_Description = "Установка. На столе (плите) с креплением болтами и планками"
Attribute Ustanov_NaStoleKreplBoltPlank.VB_ProcData.VB_Invoke_Func = " \n18"
    
    'Massa - Масса заготовки, кг
    
    'HarakterViverki - Характер выверки:
    '1 - Без выверки
    '2 - С выверкой в одной плоскости
    '3 - С выверкой в двух плоскостях
    
    'TochnostViverki - Точность выверки, мм
    
    If TochnostViverki = 0 Then TochnostViverki = 0.1
    If HarakterViverki > 3 Or HarakterViverki < 1 Or TochnostViverki < 0.1 Then Exit Function
    
    Dim arr_a As Variant, arr_x As Variant
    Select Case Massa
        Case Is <= 20
            arr_a = Split("0,73 1,04 1,38 1,52 1,66 2 2,2")
            arr_x = Split("0,26 0,26 0,26 0,26 0,27 0,27 0,27")
        Case Is <= 3000
            arr_a = Split("2 2,87 3,42 3,76 4,2 5,6 6,18")
            arr_x = Split("0,22 0,22 0,22 0,22 0,22 0,2 0,2")
        Case Is <= 30000
            arr_a = Split("1,05 1,5 1,32 1,45 4,9 2,89 3,18")
            arr_x = Split("0,3 0,3 0,34 0,34 0,2 0,28 0,28")
        Case Else
            arr_a = Split("0,233 0,333 0,346 0,38 0,245 0,338 0,372")
            arr_x = Split("0,45 0,45 0,47 0,47 0,49 0,49 0,49")
    End Select
    
    Dim a As Double, X As Double
    Select Case HarakterViverki
        Case 1
            a = CDbl(arr_a(0))
            X = CDbl(arr_x(0))
        Case Else
            Select Case TochnostViverki
                Case Is > 0.5
                    a = CDbl(arr_a(IIf(HarakterViverki = 2, 1, 4)))
                    X = CDbl(arr_x(IIf(HarakterViverki = 2, 1, 4)))
                Case Is > 0.1
                    a = CDbl(arr_a(IIf(HarakterViverki = 2, 2, 5)))
                    X = CDbl(arr_x(IIf(HarakterViverki = 2, 2, 5)))
                Case Else
                    a = CDbl(arr_a(IIf(HarakterViverki = 2, 3, 6)))
                    X = CDbl(arr_x(IIf(HarakterViverki = 2, 3, 6)))
            End Select
    End Select
    
    Ustanov_NaStoleKreplBoltPlank = a * Massa ^ X
    
End Function

Function Ustanov_NaStoleSUporomBezKrepleniya(Massa As Double, Vyverka As Boolean) As Double
    'Приложение 8, поз. 1, 2
    
    Dim arr_a As Variant, X  As Double
    Select Case Massa
        Case Is <= 0.1
            arr_a = Split("0,048 0,096")
            X = 0.2
        Case Is <= 20
            arr_a = Split("0,17 0,34")
            X = 0.35
        Case Is <= 3000
            arr_a = Split("1,62 2,32527847836389")  'В нормативе скорее всего опечатка, потому что при значении 3,24 значения времени не сходятся с картой 45
            X = 0.19
        Case Else
            arr_a = Split("0,6 0,93092024439775")  'В нормативе скорее всего опечатка, потому что при значении 1,21 значения времени не сходятся с картой 45
            X = 0.31
    End Select

    Dim a As Double
    a = CDbl(arr_a(IIf(Vyverka, 1, 0)))
    
    If Massa <= 0.1 Then
        Ustanov_NaStoleSUporomBezKrepleniya = a / Massa ^ X
    Else
        Ustanov_NaStoleSUporomBezKrepleniya = a * Massa ^ X
    End If
    
End Function

Function Ustanov_NaStole_s_Domkratami_i_KreplBoltami_i_Plankami(Massa As Double, Optional HarakterViverki As Integer = 4, Optional TochnostViverki As Double = 0.01) As Double
    
    If HarakterViverki > 3 Or HarakterViverki < 0 Or (TochnostViverki < 0.01 And TochnostViverki <> 0) Then Exit Function
    
    Ustanov_NaStole_s_Domkratami_i_KreplBoltami_i_Plankami = Ustanovka_Frezernyye_i_Rastochnyye(Massa, 4, HarakterViverki, TochnostViverki)
    
End Function

Function MaksMassaUstanovki_RastFrez(Sposob As Integer)
    
    If Sposob = 0 Then Exit Function

    Dim rUstFrez_MinMaksMassa As Range
    Set rUstFrez_MinMaksMassa = wsUst_Frez.Range("UstFrez_Sposob")
    
    Dim NomerStroki As Integer
    NomerStroki = WorksheetFunction.Match(Sposob, rUstFrez_MinMaksMassa.Columns(4), 0)

    Dim mMax As Integer
    mMax = rUstFrez_MinMaksMassa(NomerStroki, 3)
    
    MaksMassaUstanovki_RastFrez = mMax
    
End Function

Function Ustanovka_Frezernyye_i_Rastochnyye(ByVal Massa As Double, Sposob As Integer, HarakterViverki As Integer, ByVal TochnostViverki As Double) As Double

    'HarakterViverki:
    '0 - Без выверки
    '1 - С выверкой в одной плоскости
    '2 - С выверкой в двух плоскостях
    '3 - С выверкой в трёх плоскостях
    
    If Massa <= 0 Or Sposob = 0 Then Exit Function
    
    Dim rUstFrez_MinMaksMassa As Range
    Set rUstFrez_MinMaksMassa = wsUst_Frez.Range("UstFrez_Sposob")
    
    Dim NomerStroki As Integer
    NomerStroki = WorksheetFunction.Match(Sposob, rUstFrez_MinMaksMassa.Columns(4), 0)
    
    Dim mMin As Integer
    mMin = rUstFrez_MinMaksMassa(NomerStroki, 2)
    If Massa < mMin Then Massa = mMin
    
    Dim mMax As Integer
    mMax = rUstFrez_MinMaksMassa(NomerStroki, 3)
    
    If Massa > mMax And mMax <> 0 Then Exit Function
    
    Select Case TochnostViverki
        Case 0:         TochnostViverki = 0
        Case Is > 0.5:  TochnostViverki = 1
        Case Is > 0.1:  TochnostViverki = 0.5
        Case Is > 0.05: TochnostViverki = 0.1
        Case Is > 0.01: TochnostViverki = 0.05
        Case Else:      TochnostViverki = 0.01
    End Select
    
    Dim id As String
    id = Sposob & HarakterViverki & TochnostViverki
    
    Dim rUst As Range
    Set rUst = wsUst_Frez.Range("UstFrez")
    
    On Error Resume Next
    NomerStroki = WorksheetFunction.Match(id, rUst.Columns(14), 0)
    
    If Err.Number = 0 Then
    On Error GoTo 0
    
        Dim a As Double, X As Double
        Select Case Massa
            Case Is <= 0.1
                a = rUst(NomerStroki, 4)
                X = rUst(NomerStroki, 5)
            Case Is <= 20
                a = rUst(NomerStroki, 6)
                X = rUst(NomerStroki, 7)
            Case Is <= 3000
                a = rUst(NomerStroki, 8)
                X = rUst(NomerStroki, 9)
            Case Is <= 30000
                a = rUst(NomerStroki, 10)
                X = rUst(NomerStroki, 11)
            Case Else
                a = rUst(NomerStroki, 12)
                X = rUst(NomerStroki, 13)
        End Select
        
        Dim Vremya As Double
        If Massa <= 0.1 Then
            Vremya = a / Massa ^ X
        Else
            Vremya = a * Massa ^ X
        End If
        
        Ustanovka_Frezernyye_i_Rastochnyye = Vremya
    
    End If
    
End Function

Private Function Ustanov_ChetirehKulachkovyi(Massa As Double, _
    Optional FormaZagotovki As Integer = 1, _
    Optional HarakterViverki As Integer = 3, _
    Optional TochnostViverki As Double = 0.01) As Double
    
    'Massa - Масса заготовки, кг
    
    'FormaZagotovki: 1 - детали цилиндрической формы, 2 - детали коробчатой и фасонной формы
    
    'HarakterViverki - Характер выверки:
    '1 - Без выверки
    '2 - С выверкой по диаметру
    '3 - С выверкой по диаметру и торцу
    
    'TochnostViverki - Точность выверки, мм

    If HarakterViverki > 3 Or HarakterViverki < 1 Or (TochnostViverki < 0.01 And TochnostViverki <> 0) Or Massa > MaksMassaUstanov_Tok Then Exit Function
    
    Dim arr_a As Variant, arr_x As Variant
    If FormaZagotovki = 1 Then
        Select Case Massa
            Case Is <= 20
                arr_a = Split("0,565 0,81 0,89 1,08 1,3 1,56 1,3 1,43 1,73 2,08 2,5")
                arr_x = Split("0,31 0,34 0,34 0,34 0,34 0,34 0,34 0,34 0,34 0,34 0,34")
            Case Else
                arr_a = Split("2,4 3,3 3,63 5 5,9 7,08 3 3,83 4,8 5,71 6,56")
                arr_x = Split("0,17 0,19 0,19 0,16 0,16 0,16 0,16 0,16 0,16 0,16 0,16")
        End Select
    Else
        Select Case Massa
            Case Is <= 20
                arr_a = Split("0,78 1,02 0,12 0,143 1,73 2,08 1,87 2,06 1,65 3,06 3,67")
                arr_x = Split("0,31 0,38 0,38 0,35 0,35 0,35 0,33 0,33 0,34 0,34 0,34")
            Case Else
                arr_a = Split("2,89 4,5 4,94 6,3 7,43 8,9 6,7 7,36 8,36 9,85 11,8")
                arr_x = Split("0,19 0,19 0,19 0,16 0,16 0,16 0,18 0,18 0,18 0,18 0,18")
        End Select
    End If

    Dim PoziciyaVMassive As Integer
    Select Case HarakterViverki
        Case 1
            PoziciyaVMassive = 0
        Case Else
            If TochnostViverki = 0 Then TochnostViverki = 0.01
            Select Case TochnostViverki
                Case Is > 0.5:  PoziciyaVMassive = 1
                Case Is > 0.1:  PoziciyaVMassive = 2
                Case Is > 0.05: PoziciyaVMassive = 3
                Case Is > 0.01: PoziciyaVMassive = 4
                Case Else:      PoziciyaVMassive = 5
            End Select
            PoziciyaVMassive = PoziciyaVMassive + IIf(HarakterViverki = 3, 5, 0)
    End Select
    
    Dim a As Double, X As Double
    a = CDbl(arr_a(PoziciyaVMassive))
    X = CDbl(arr_x(PoziciyaVMassive))
    
    Ustanov_ChetirehKulachkovyi = a * Massa ^ X
    
End Function

Function Ustanov_ChetirehKulachkovyi_TsilindrForma(Massa As Double, _
    Optional HarakterViverki As Integer = 3, _
    Optional TochnostViverki As Double = 0.01) As Double
    
    Ustanov_ChetirehKulachkovyi_TsilindrForma = Ustanov_ChetirehKulachkovyi(Massa, 1, HarakterViverki, TochnostViverki)
    
End Function

Function Ustanov_ChetirehKulachkovyi_FasonnayaForma(Massa As Double, _
    Optional HarakterViverki As Integer = 3, _
    Optional TochnostViverki As Double = 0.01) As Double
    
    Ustanov_ChetirehKulachkovyi_FasonnayaForma = Ustanov_ChetirehKulachkovyi(Massa, 2, HarakterViverki, TochnostViverki)
    
End Function

Function Ustanov_SamotsentrPatron(Massa As Double, _
    Optional HarakterViverki As KharakterVyverki_SamotsentrPatron = sVyverkoyPoDiametruItortsu, _
    Optional TochnostViverki As Double = 0.1) As Double
    
    'Massa - Масса заготовки, кг
    
    'TochnostViverki - Точность выверки, мм
    
    If HarakterViverki <> BezVyverki And HarakterViverki <> sVyverkoyPoDiametruItortsu Then Exit Function
     
    If Massa > MaksMassaUstanov_Tok Or (TochnostViverki < 0.1 And TochnostViverki <> 0) Then Exit Function

    Dim arr_a As Variant, arr_x As Variant
    Select Case Massa
        Case Is <= 0.1
            arr_a = Split("0,154 0,252 0,42")
            arr_x = Split("0,13 0,13 0,13")
        Case Is <= 20
            arr_a = Split("0,36 0,74 1,53")
            arr_x = Split("0,37 0,27 0,24")
        Case Else
            arr_a = Split("2,14 1,216 1,57")
            arr_x = Split("0,2 0,4 0,4")
    End Select

    Dim PoziciyaVMassive As Integer
    Select Case HarakterViverki
        Case 1
            PoziciyaVMassive = 0
        Case Else
            If TochnostViverki > 0.1 Then
                PoziciyaVMassive = 1
            Else
                PoziciyaVMassive = 2
            End If
    End Select
    
    Dim a As Double, X As Double
    a = CDbl(arr_a(PoziciyaVMassive))
    X = CDbl(arr_x(PoziciyaVMassive))
    
'    If Massa < 0.1 Then
'        Ustanov_SamotsentrPatron = a / Massa ^ x
'    Else
'        Ustanov_SamotsentrPatron = a * Massa ^ x
'    End If
    
    Ustanov_SamotsentrPatron = a * Massa ^ X

End Function

Function Ustanov_TokFrez(Sposob As Integer, Massa As Double, _
    Optional HarakterViverki As Integer = 2, _
    Optional TochnostViverki As Double = 0.1) As Double
    
    Dim tUst As Double
    Select Case Sposob
        Case 1  'В самоцентрирующем патроне
            tUst = Ustanov_SamotsentrPatron(Massa, CInt(HarakterViverki), TochnostViverki)
        Case 2  'В 4-х кулачковом патроне (детали цилиндрической формы)
            tUst = Ustanov_ChetirehKulachkovyi_TsilindrForma(Massa, HarakterViverki, TochnostViverki)
        Case 3  'В 4-х кулачковом патроне (детали фасонной и коробчатой формы)
            tUst = Ustanov_ChetirehKulachkovyi_FasonnayaForma(Massa, HarakterViverki, TochnostViverki)
    End Select
    
    Ustanov_TokFrez = tUst

End Function




