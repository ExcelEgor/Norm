Attribute VB_Name = "mLakokras"
Option Explicit

Private Const MAKS_RAZMER_SIMVOLA As Integer = 15

Function MarkirovkaKistyu(VisotaShrifta As Double, KolVoSimvolov As Double, Optional Sloznost As Integer = 2, _
    Optional KolvoCvetov As Integer = 1, Optional KolvoNadpisei As Double = 1, Optional Lakirovanie As Boolean = False)
      
    If KolvoNadpisei < 1 Or Sloznost < 1 Or Sloznost > 2 Or KolvoCvetov < 1 Or KolvoCvetov > 2 Or VisotaShrifta > 10 Then Exit Function
    
    VisotaShrifta = BlizhBolshRavn_Znachenie_Array(VisotaShrifta, Array(3, 5, 10))
    
    Dim Key As Long
    Key = CLng(Sloznost & KolvoCvetov & VisotaShrifta)

    Dim tbLakokras41 As ListObject
    Set tbLakokras41 = ThisWorkbook.Worksheets("Лакокрас_41").ListObjects("tbLakokras41")
    
    Dim NumRow As Integer
    Dim a As Double, b As Double
    With tbLakokras41
        NumRow = WorksheetFunction.Match(Key, .ListColumns(4).DataBodyRange, 0)
        a = .DataBodyRange(NumRow, 5)
        b = .DataBodyRange(NumRow, 6)
    End With
    
    MarkirovkaKistyu = (a * (KolVoSimvolov / KolvoNadpisei) + b) * KolvoNadpisei
    
    Dim tLakirovanie As Double
    If Lakirovanie Then
        tLakirovanie = LakirovanieNadpisi(VisotaShrifta, KolVoSimvolov, KolvoNadpisei)
        MarkirovkaKistyu = MarkirovkaKistyu + tLakirovanie
    End If
    
End Function

Function ZapolGravEmal_Znak_45(VistotaShrifta As Double, KolZnakov As Double, Optional Slozhnost As Integer = 2, Optional KolTonov As Integer = 1) As Double

    If KolTonov > 8 Or Slozhnost > 2 Or KolZnakov = 0 Or VistotaShrifta > MAKS_RAZMER_SIMVOLA Then Exit Function
    
    If KolTonov < 1 Then KolTonov = 1
    If Slozhnost < 1 Then Slozhnost = 2
    
    Dim rTime As Range
    Set rTime = wsLakokras_45.Range("E8:F12").Offset(0, IIf(Slozhnost = 2, 2, 0))
    
    If VistotaShrifta > 3 Then
        If VistotaShrifta > 7 Then
            Set rTime = rTime.Offset(10)
        Else
            Set rTime = rTime.Offset(5)
        End If
    End If
    
    Dim NumRow As Integer
    Select Case KolTonov
        Case Is <= 1:   NumRow = 1
        Case Is <= 2:   NumRow = 2
        Case Is <= 3:   NumRow = 3
        Case Is <= 5:   NumRow = 4
        Case Else:      NumRow = 5
    End Select
    
    Dim tFirst As Double, tNext As Double
    tFirst = rTime(NumRow, 1)
    tNext = rTime(NumRow, 2) * (KolZnakov - 1)
    ZapolGravEmal_Znak_45 = tFirst + tNext
    
End Function

Function ZapolGravEmal_Znak_Range(rVistotaShrifta As Range, rKolZnakov As Range) As Double
    Dim i As Long
    For i = 1 To rVistotaShrifta.Rows.Count
        If rVistotaShrifta(i) > 0 And rKolZnakov(i) > 0 Then
            ZapolGravEmal_Znak_Range = ZapolGravEmal_Znak_Range + ZapolGravEmal_Znak_45(rVistotaShrifta(i), rKolZnakov(i))
        End If
    Next
End Function

Function ZapolGravEmal_Riska_45(L As Double, KolTonov As Integer, Slozhnost As Integer, KolRisok As Double)

    Dim NumRow As Integer
    
    If KolTonov > 8 Or Slozhnost > 2 Or KolRisok = 0 Then Exit Function
    If KolTonov < 1 Then KolTonov = 1
    If Slozhnost < 1 Then Slozhnost = 2
    
    With wsLakokras_45
        Dim rTime As Range
        Set rTime = .Range("E8:F12").Offset(0, IIf(Slozhnost = 2, 2, 0))
    
        If L > 10 Then
            Select Case L
                Case Is <= 25: Set rTime = rTime.Offset(5)
                Case Is <= 50: Set rTime = rTime.Offset(10)
                Case Else: Set rTime = rTime.Offset(15)
            End Select
            If L > 100 Then
                Dim rTime100 As Range
                Set rTime100 = .Range("E28:F32").Offset(0, IIf(Slozhnost = 2, 2, 0))
            End If
        End If
    End With
    
    Select Case KolTonov
        Case Is <= 1:   NumRow = 1
        Case Is <= 2:   NumRow = 2
        Case Is <= 3:   NumRow = 3
        Case Is <= 5:   NumRow = 4
        Case Else:      NumRow = 5
    End Select
    
    Dim tFirst As Double, tNext As Double, t100First As Double, t100next As Double
    If L <= 100 Then
        tFirst = rTime(NumRow, 1)
        tNext = rTime(NumRow, 2) * (KolRisok - 1)
        ZapolGravEmal_Riska_45 = tFirst + tNext
    Else
        t100First = rTime100(NumRow, 1) * ((L - 100) / 100)
        tFirst = rTime(NumRow, 1) + t100First
        t100next = rTime100(NumRow, 2) * ((L - 100) / 100) * (KolRisok - 1)
        tNext = rTime(NumRow, 2) * (KolRisok - 1) + t100next
        ZapolGravEmal_Riska_45 = tFirst + tNext
    End If
    
End Function

Function ZapolneniyeGravirovki_Linii_Range(RangeDlinaLinii As Range, RangeKolichestvoLiniy As Range) As Double
    
    Dim DlinaLinii As Double, KolichestvoLiniy As Double
    With WorksheetFunction
        If .Sum(RangeDlinaLinii) > 0 And .Sum(RangeKolichestvoLiniy) > 0 Then
            KolichestvoLiniy = .Sum(RangeKolichestvoLiniy)
            DlinaLinii = .SumProduct(RangeDlinaLinii, RangeKolichestvoLiniy) / .Sum(RangeKolichestvoLiniy)
            ZapolneniyeGravirovki_Linii_Range = ZapolGravEmal_Riska_45(DlinaLinii, 1, 1, KolichestvoLiniy)
        End If
    End With


End Function


Function Gruntovaniye(S As Double, Optional Pokrit As String = "ВЛ-02", _
    Optional KolSloev As Integer = 1, Optional Sposob As Integer = 2, Optional Sloznost As Integer = 3, Optional TipPoverhnosti As Integer = 1) As Double
    
    'TipPoverhnosti: 1-Наружная; 2-Внутренняя
   
    Static tbGrunt As ListObject, tbMapGrunt As ListObject
    With wsLakokras_Grunt
        Set tbGrunt = .ListObjects("tbФормулы_Грунт")
        Set tbMapGrunt = .ListObjects("tbСлой_Грунт")
    End With
    
    Dim Map As Integer
    With tbMapGrunt
        Map = .DataBodyRange(WorksheetFunction.Match(Pokrit & KolSloev, .ListColumns(3).DataBodyRange, 0), 4)
    End With
    
    Dim i As Integer
    Dim a As Double, b As Double
    With tbGrunt
        i = WorksheetFunction.Match(Map & Sposob & Sloznost, .ListColumns(4).DataBodyRange, 0)
        a = .DataBodyRange(i, IIf(S <= 7.5, 5, 7))
        b = .DataBodyRange(i, IIf(S <= 7.5, 6, 8))
    End With
    
    Gruntovaniye = (a * S ^ b) * IIf(TipPoverhnosti = 2, 1.15, 1)
    
End Function

Function Okrashivaniye(S As Double, Optional Pokrit As String = "ПФ-218ГС", _
    Optional Sloznost As Integer = 3, Optional KolSloev As Integer = 3, Optional Sposob As Integer = 2, Optional TipPoverhnosti As Integer = 1) As Double
    
    Static tbOkrash As ListObject, tbMapOkrash As ListObject, tbMaksKolSloev As ListObject
    With wsLakokras_Okrash
        Set tbOkrash = .ListObjects("tbФормулы_Окраш")
        Set tbMapOkrash = .ListObjects("tbСлой_Окраш")
        Set tbMaksKolSloev = .ListObjects("tbMaksKolSloev")
    End With
    
    If KolSloev = 0 Then
        KolSloev = WorksheetFunction.VLookup(Pokrit, tbMaksKolSloev.DataBodyRange, 2, 0)
    End If
    
    Dim Map As Integer
    With tbMapOkrash
        Map = .DataBodyRange(WorksheetFunction.Match(Pokrit & KolSloev, .ListColumns(3).DataBodyRange, 0), 4)
    End With
    
    Dim a As Double, b As Double
    Dim i As Integer
    With tbOkrash
        i = WorksheetFunction.Match(Map & Sposob & Sloznost, .ListColumns(4).DataBodyRange, 0)
        a = .DataBodyRange(i, IIf(S <= 7.5, 5, 7))
        b = .DataBodyRange(i, IIf(S <= 7.5, 6, 8))
    End With
    
    Okrashivaniye = (a * S ^ b) * IIf(TipPoverhnosti = 2, 1.15, 1)
    
End Function
Function Lakirovanie(Sdm As Double, Optional Sposob As Integer = 2, Optional Sloznost As Integer = 3) As Double
    
    If Not Sdm > 0 Then Exit Function
    
    If Sdm < 0.05 Then Sdm = 0.05
    
    Dim Pokazateli(1 To 2, 1 To 3, 1 To 2, 1 To 2) As Double
    
    'Пневматическим распылителем
    '1 группа сложности
    Pokazateli(1, 1, 1, 1) = 0.696658725557781
    Pokazateli(1, 1, 1, 2) = 0.337870238870545
    Pokazateli(1, 1, 2, 1) = 0.387639942590545
    Pokazateli(1, 1, 2, 2) = 0.63453889428507
    '2 группа сложности
    Pokazateli(1, 2, 1, 1) = 0.835076180342227
    Pokazateli(1, 2, 1, 2) = 0.32019016957371
    Pokazateli(1, 2, 2, 1) = 0.430354207954542
    Pokazateli(1, 2, 2, 2) = 0.648180232951234
    '3 группа сложности
    Pokazateli(1, 3, 1, 1) = 0.971356636025651
    Pokazateli(1, 3, 1, 2) = 0.306147534704333
    Pokazateli(1, 3, 2, 1) = 0.471215251491972
    Pokazateli(1, 3, 2, 2) = 0.673127995807029

    'Кистью
    '1 группа сложности
    Pokazateli(2, 1, 1, 1) = 1.29999243686171
    Pokazateli(2, 1, 1, 2) = 0.325621470952861
    Pokazateli(2, 1, 2, 1) = 0.64791147814242
    Pokazateli(2, 1, 2, 2) = 0.690777039640291
    '2 группа сложности
    Pokazateli(2, 2, 1, 1) = 1.58099355387989
    Pokazateli(2, 2, 1, 2) = 0.317902174102287
    Pokazateli(2, 2, 2, 1) = 0.802851477145644
    Pokazateli(2, 2, 2, 2) = 0.694182915003908
    '3 группа сложности
    Pokazateli(2, 3, 1, 1) = 2.00117599273683
    Pokazateli(2, 3, 1, 2) = 0.310260481462106
    Pokazateli(2, 3, 2, 1) = 0.95284694265008
    Pokazateli(2, 3, 2, 2) = 0.709064344146066
    
    Dim a As Double, b As Double
    a = Pokazateli(Sposob, Sloznost, IIf(Sdm <= 7.5, 1, 2), 1)
    b = Pokazateli(Sposob, Sloznost, IIf(Sdm <= 7.5, 1, 2), 2)
    
    Lakirovanie = a * Sdm ^ b

End Function

Function IzolirovaniyeOtverstiy(Sposob As Integer, d As Double, Sloznost As Integer) As Double
    
    Dim a As Double, b As Double
    
    Select Case Sposob
        Case 1  'Винтом
            a = 0.218: b = 0.19
        Case 2 'Трубкой
            a = 0.091: b = 0.38
        Case 3 'Липкой лентой
            a = 0.147: b = 0.58
        Case 4 'Шаблоном
            a = 0.0757: b = 0.28
        Case 5 'Заглушкой или пробкой
            a = 0.0738: b = 0.28
        Case 6 'Бумажной пробкой
            a = 0.1418: b = 0.28
        Case 7 'Гипсовой замазкой
            a = 0.1168: b = 0.42
    End Select
    
    IzolirovaniyeOtverstiy = (a * d ^ b) * fIzolKoef(Sloznost)
    
End Function

Function fIzolCilindrPov(Sposob As Integer, S As Double, Sloznost As Integer) As Double

    Dim a As Double, b As Double
    
    Select Case Sposob
        Case 1 'Шаблоном
            a = 0.234: b = 0.37
        Case 2 'Липкой лентой
            a = 1.41: b = 0.37
        Case 3 'Бумага с креплением ниткой
            a = 1.128: b = 0.37
        Case 4 'Бумага с креплением липкой лентой
            a = 0.959: b = 0.37
        Case 5 'Лаком, смазкой
            a = 2.769: b = 0.37
    End Select
    
    fIzolCilindrPov = (a * S ^ b) * fIzolKoef(Sloznost)
    
End Function

Function fIzolPloskPov(Sposob As Integer, S As Double, Sloznost As Integer) As Double

    Dim a As Double, b As Double
    
    Select Case Sposob
        Case 1 'Шаблоном
            a = 0.2017: b = 0.37
        Case 2 'Липкой лентой
            a = 1.185: b = 0.37
        Case 3 'Бумага с креплением ниткой
            a = 0.9145: b = 0.38
        Case 4 'Бумага с креплением липкой лентой
            a = 0.935: b = 0.33
        Case 5 'Лаком, смазкой
            a = 2.33: b = 0.36
    End Select
    
    fIzolPloskPov = (a * S ^ b) * fIzolKoef(Sloznost)
    
End Function

Function fIzolKoef(Sloznost As Integer)
    Select Case Sloznost
        Case 1:     fIzolKoef = 1
        Case 2:     fIzolKoef = 1.3
        Case Else:  fIzolKoef = 1.5
    End Select
End Function

Private Function LakirovanieNadpisi(VisotaShrifta As Double, KolVoSimvolov As Double, Optional KolvoNadpisei As Double = 1) As Double
    
    If VisotaShrifta > 0 And KolvoNadpisei > 0 Then

        
        Dim PloshchadLakirovaniya As Double
        PloshchadLakirovaniya = PloshchadPramougolnika(VisotaShrifta, DlinaNadpisi(VisotaShrifta, KolVoSimvolov) / KolvoNadpisei) / 10000
        
        LakirovanieNadpisi = Lakirovanie(PloshchadLakirovaniya) * KolvoNadpisei

    Else
        LakirovanieNadpisi = CVErr(xlErrNA)
    End If


End Function
