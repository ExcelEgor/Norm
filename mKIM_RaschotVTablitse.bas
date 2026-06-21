Attribute VB_Name = "mKIM_RaschotVTablitse"
Option Explicit

Function RaschotKIMIzTablitsy(Zagolovki As Range, StorkaDannykh As Range) As Double

    If ZnacheniyaParamtrovKorrektny(StorkaDannykh, Zagolovki) = False Then Exit Function

    Dim KIM As clsKIM
    Set KIM = PrisvoitParametry(StorkaDannykh, Zagolovki)

    RaschotKIMIzTablitsy = KIM_RaschotTshtIzParametrov(KIM)

End Function

Function KIM_RaschotTpzIzTablitsy(Zagolovki As Range, Dannye As Range, GruppaSlozhosti As Integer, _
    KolVoKalibrovokSchupa_15 As Integer, KolVoKalibrovokSchupa_40 As Integer, KolVoKalibrovokSchupa_99 As Integer) As Double
    
    Dim i As Integer
    Dim KIM As clsKIM
    
    Dim Tup_Kruglomer As Double, Tup_Hera As Double
    Dim HeraEstVRaschote As Boolean, KruglomerEstVRaschote As Boolean, MikroskopEstVraschote As Boolean
    For i = 1 To Dannye.Rows.Count

        Dim StorkaDannykh As Range
        Set StorkaDannykh = Dannye.Rows(i)
        If ZnacheniyaParamtrovKorrektny(StorkaDannykh, Zagolovki) = False Then Exit Function
        Set KIM = PrisvoitParametry(Dannye.Rows(i), Zagolovki)

        With KIM
            If LCase(.id) Like "*kruglomer*" Then
                KruglomerEstVRaschote = True
                Tup_Kruglomer = Tup_Kruglomer + KIM_Kruglomer_VremyaNapisaniyaUPOdnogoPerekhoda(.Dlina, .Slozhost, .KolVoPoverkhnostey, .RadialnoyeIzmereniye, .LieneynoyeIzmereniye)
            
            ElseIf LCase(.id) Like "*mikroskop*" Then
                MikroskopEstVraschote = True
                
            ElseIf LCase(.id) Like "*hera*" Then
                HeraEstVRaschote = True
                Tup_Hera = KIM_Hera_VremyaNapisaniyaUPOdnogoPerekhoda(.Slozhost, .KolVoPoverkhnostey)
            
            End If
        End With

    Next
    
    Dim Tpz_Kruglomer As Double
    If KruglomerEstVRaschote Then
        Tpz_Kruglomer = KIM_Kruglomer_Tpz(GruppaSlozhosti, Tup_Kruglomer)
    End If
    
    Dim Tpz_Hera As Double
    If HeraEstVRaschote Then
        Tpz_Hera = KIM_Hera_Tpz(GruppaSlozhosti, Tup_Hera, KolVoKalibrovokSchupa_15, KolVoKalibrovokSchupa_40, KolVoKalibrovokSchupa_99)
    End If
    
    Dim Tpz_Mikroskop As Double
    If MikroskopEstVraschote Then
        Tpz_Mikroskop = KIM_Mikroskop_Tpz(GruppaSlozhosti)
    End If
    
    KIM_RaschotTpzIzTablitsy = WorksheetFunction.RoundUp(Tpz_Kruglomer + Tpz_Mikroskop + Tpz_Hera, 0)

End Function

Private Function ZnacheniyaParamtrovKorrektny(StorkaDannykh As Range, Zagolovki As Range) As Boolean

    Dim i As Integer
    ZnacheniyaParamtrovKorrektny = True
    For i = 1 To Zagolovki.Columns.Count
        Select Case Zagolovki(i)
            
            Case "tнш", "Переход"
            
            Case "id"
                If VarType(StorkaDannykh.Columns(i)) <> vbString Or VarType(StorkaDannykh.Columns(i)) = vbEmpty Then
                    ZnacheniyaParamtrovKorrektny = False
                    Exit Function
                End If
                
            Case "D1", "d2", "Длина", "Кол-во поверхностей", "Ширина", "Сложность", "Кол-во сечений", "Масса"
                If VarType(StorkaDannykh.Columns(i)) <> vbEmpty Then
                    If Not IsNumeric(StorkaDannykh.Columns(i)) Then
                        ZnacheniyaParamtrovKorrektny = False
                        Exit Function
                    End If
                End If
            
            Case "Допуск формы", "Допуск расположения", "Допуск размера", "Линейное измерение", "От оси", "Радиальное измерение", "Тригерный режим"
                If VarType(StorkaDannykh.Columns(i)) <> vbEmpty Then
                    If VarType(StorkaDannykh.Columns(i)) <> vbBoolean Then
                        ZnacheniyaParamtrovKorrektny = False
                        Exit Function
                    End If
                End If
            
            Case Else
                ZnacheniyaParamtrovKorrektny = False
                Exit Function
                
        End Select
    Next
    
End Function

Private Function PrisvoitParametry(StorkaDannykh As Range, Zagolovki As Range) As clsKIM

    Dim KIM As New clsKIM
    With KIM
        .D1 = PoluchitZnacheniye("D1", StorkaDannykh, Zagolovki)
        .d2 = PoluchitZnacheniye("d2", StorkaDannykh, Zagolovki)
        .Dlina = PoluchitZnacheniye("Длина", StorkaDannykh, Zagolovki)
        .DopuskFormy = PoluchitZnacheniye("Допуск формы", StorkaDannykh, Zagolovki)
        .DopuskRaspolozheniya = PoluchitZnacheniye("Допуск расположения", StorkaDannykh, Zagolovki)
        .DopuskRazmera = PoluchitZnacheniye("Допуск размера", StorkaDannykh, Zagolovki)
        .id = PoluchitZnacheniye("id", StorkaDannykh, Zagolovki)
        .KolVoPoverkhnostey = PoluchitZnacheniye("Кол-во поверхностей", StorkaDannykh, Zagolovki)
        .KolVoSecheniy = PoluchitZnacheniye("Кол-во сечений", StorkaDannykh, Zagolovki)
        .LieneynoyeIzmereniye = PoluchitZnacheniye("Линейное измерение", StorkaDannykh, Zagolovki)
        .OtOsi = PoluchitZnacheniye("От оси", StorkaDannykh, Zagolovki)
        .RadialnoyeIzmereniye = PoluchitZnacheniye("Радиальное измерение", StorkaDannykh, Zagolovki)
        .Shirina = PoluchitZnacheniye("Ширина", StorkaDannykh, Zagolovki)
        .Slozhost = PoluchitZnacheniye("Сложность", StorkaDannykh, Zagolovki)
        .TrigernyyRezhim = PoluchitZnacheniye("Тригерный режим", StorkaDannykh, Zagolovki)
        .Massa = PoluchitZnacheniye("Масса", StorkaDannykh, Zagolovki)
    End With
    
    Set PrisvoitParametry = KIM
    
End Function


