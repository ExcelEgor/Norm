Attribute VB_Name = "mPrivate"

'Данный модуль используется для хранения общих процедур и функций,
'которые должны быть доступны из других модулей, классов и форм,
'но при этом не должны быть видны при штатном использовании Excel

Option Explicit
Option Private Module
Public AppEvents As clsAppEvents

#If Win64 Then
    Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
#Else
    Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
#End If

Type GabarityZagotovki
    Dlina As Double
    Shirina As Double
    Tolshchina As Double
End Type
Sub InitializeAppEvents()
    Set AppEvents = New clsAppEvents
    Set AppEvents.App = Application
End Sub


Sub dobavit_checkbox_v_stolbets(Diapazon As Range, NomerStolbtsa As Integer)

    Application.ScreenUpdating = False

    Dim i As Integer
    Dim L As Double, t As Double, W As Double, H As Double
    Dim cell As Range
    
    For i = 1 To Diapazon.Rows.Count
        
        Set cell = Diapazon(i, NomerStolbtsa)
        
        With cell
            H = .Height
            L = .Left
            t = .Top - 1
            W = .Width
            .NumberFormat = ";;;"
        End With

        ActiveSheet.CheckBoxes.Add(L + W / 2 - H / 2, t, H, H).Select
        
        With Selection
            .Value = xlOff
            .LinkedCell = cell.Address
            .Characters.text = ""
        End With
        
    Next
    
    Diapazon(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub
Function OKRUGLVVERKH(Chislo As Double, Optional ByVal KolVoZnakovPosleZapyatoy As Integer = 0) As Double

    Dim Factor As Double
    Factor = 10 ^ KolVoZnakovPosleZapyatoy
    OKRUGLVVERKH = -Int(-Chislo * Factor) / Factor
    
End Function
    
Function OKRUGLVNIZ(Chislo As Double) As Long
    OKRUGLVNIZ = Int(Chislo)
End Function

Function OkruglVverkhSTochnostyu(ByVal Chislo As Double, ByVal Tochnost As Integer) As Double

    If Chislo <= 0 Or Tochnost <= 0 Then Exit Function


    OkruglVverkhSTochnostyu = -Int(-Chislo / Tochnost) * Tochnost


End Function

Function RaschotGabaritov(Dlina As Double, Shirina As Double, Optional Vysota As Double) As GabarityZagotovki

    RaschotGabaritov.Dlina = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 1)
    RaschotGabaritov.Shirina = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 2)
    RaschotGabaritov.Tolshchina = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 3)
    
End Function


Function DlinaShirina(Dlina As Double, Shirina As Double, Optional Vysota As Double)
    
    Dim MinMax(1 To 3)
    MinMax(1) = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 1)
    MinMax(2) = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 2)
    MinMax(3) = WorksheetFunction.Large(Array(Dlina, Shirina, Vysota), 3)
    
    DlinaShirina = MinMax

End Function

Function KoefficientLmaxL(Lmax As Double, L As Double) As Double
    If Lmax > 0 Then
        If L > Lmax Then
            KoefficientLmaxL = L / Lmax
        Else
            KoefficientLmaxL = 1
        End If
    End If
End Function
Function PerekhodSUchotomKolVaPoverkhnostey(TekstPerekhoda As String, KolVoPoverkhnostey As Double)
    If KolVoPoverkhnostey > 1 Then
        PerekhodSUchotomKolVaPoverkhnostey = TekstPerekhoda & " (x" & KolVoPoverkhnostey & ")"
    Else
        PerekhodSUchotomKolVaPoverkhnostey = TekstPerekhoda
    End If
End Function



Function Num_NOt_Stable() As Boolean
    
    
    Dim keystat(0 To 255) As Byte
    Dim state As String
    
    GetKeyboardState keystat(0)
    state = keystat(vbKeyNumlock)
    If (state = 0) Then
        Num_NOt_Stable = False
    Else
        Num_NOt_Stable = True
    End If
    
End Function

Function MinimalnoeZnachenie(Znachenie As Double, Min As Double)
    MinimalnoeZnachenie = IIf(Znachenie < Min, Min, Znachenie)
End Function

Function PlotnostMaterialaIzTokarnihNormativov(cbo As MSForms.ComboBox)
    PlotnostMaterialaIzTokarnihNormativov = CDbl(cbo.List(cbo.ListIndex, 2))
End Function
Function NaimenovanieOperaciiGalvaniki(cbo As MSForms.ComboBox)
    Select Case cbo.ListIndex
        Case 0:         NaimenovanieOperaciiGalvaniki = "Цинкование электрохимическое"
        Case 1 To 3:    NaimenovanieOperaciiGalvaniki = "Окисление анодное"
        Case 4:         NaimenovanieOperaciiGalvaniki = "Пассивирование"
        Case 5, 6:      NaimenovanieOperaciiGalvaniki = "Никелирование "
        Case 7:         NaimenovanieOperaciiGalvaniki = "Нанесение сплава олово-висмут "
    End Select
End Function


Function DlinaNadpisi(VisotaShrifta As Double, KolVoSimvolov As Double) As Double

    Dim a As Double, b As Double
    a = 0.709448504983389
    b = 4.87264673307841E-05
    
    DlinaNadpisi = (a * VisotaShrifta + b) * KolVoSimvolov
    
End Function




Function KnigaUzheOtkryta(ImyaKnigi As String) As Boolean

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = ImyaKnigi Then
            KnigaUzheOtkryta = True
            Exit For
        End If
    Next
    
End Function

Function ParametyBolsheNulya(Parametry) As Boolean
    
    ParametyBolsheNulya = True
    
    Dim i As Integer
    For i = 0 To UBound(Parametry)
        If Parametry(i) <= 0 Then
            ParametyBolsheNulya = False
            Exit Function
        End If
    Next
    
End Function

Function PolozhitelnyyeChisla(ParamArray Znacneniya()) As Boolean
    
    Dim i As Integer
    
    PolozhitelnyyeChisla = True
    For i = 0 To UBound(Znacneniya)
        If IsNumeric(Znacneniya(i)) Then
            If Not Znacneniya(i) > 0 Then
                PolozhitelnyyeChisla = False
                Exit For
            End If
        Else
            PolozhitelnyyeChisla = False
            Exit For
        End If
    Next
    
End Function

Function KolVoPoverhnosteiProhodov(ctrl As Control) As Double
    'Если количество обрабатываемых поверхностей не указано то сделать его равным 1
    
    If IsNumeric(ctrl) Then
        KolVoPoverhnosteiProhodov = CDbl(ctrl)
    Else
        KolVoPoverhnosteiProhodov = 1
    End If
    
End Function

Function FormatTime(t As Double, Optional UpDownMath As Integer = 1, Optional MinTime As Double = 0.1, Optional NumDigits As Integer = 1) As String
    If t < MinTime Then t = MinTime
    Select Case UpDownMath
        Case 1
            t = WorksheetFunction.RoundUp(t, NumDigits)
        Case 2
            t = WorksheetFunction.RoundDown(t, NumDigits)
        Case 3
            t = WorksheetFunction.Round(t, NumDigits)
    End Select
    If NumDigits = 0 Then
        FormatTime = Format(t, "#,##")
    Else
        FormatTime = Format(t, "#,##0." & String(NumDigits, "0"))
    End If
    
End Function



Function BlizhBolshRavn_Pozic_Array(Val As Double, Arr As Variant) As Double
    Dim findVal As Boolean
    Dim i As Integer
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = Replace(Arr(i), ".", ",")
        If IsNumeric(Arr(i)) Then
            If CDbl(Arr(i)) >= Val Then
                findVal = True
                Exit For
            End If
        End If
    Next
    If findVal = True Then
        BlizhBolshRavn_Pozic_Array = i + 1
    Else
        BlizhBolshRavn_Pozic_Array = CVErr(xlErrValue)
    End If
End Function

Function BlizhBolshRavn_Znachenie_Array(ByVal Val As Double, ByVal Arr As Variant) As Double
    
    Dim i As Integer
    
    For i = LBound(Arr) To UBound(Arr)
'        Arr(i) = Replace(Arr(i), ".", ",")
        If IsNumeric(Arr(i)) Then
            If Arr(i) >= Val Then
                BlizhBolshRavn_Znachenie_Array = Arr(i)
                Exit Function
            End If
        End If
    Next
    
End Function


Function NomerStolbca_BlizhBolshRavn(ByVal Znachenie As Double, ByVal Diapazon As Range, Optional MaxZnachenie As Double = 0) As Variant

    If Diapazon.Rows.Count > 1 Then Exit Function
    
    If MaxZnachenie <> 0 Then
        If Znachenie > MaxZnachenie Then Znachenie = MaxZnachenie
    End If
    
    Dim findVal As Boolean
    Dim i As Long
    For i = 1 To Diapazon.Columns.Count
        If IsNumeric(Diapazon.Columns(i)) Then
            If Diapazon.Columns(i) >= Znachenie Then
                findVal = True
                Exit For
            End If
        End If
    Next
    If findVal = True Then
        NomerStolbca_BlizhBolshRavn = i
    Else
        NomerStolbca_BlizhBolshRavn = CVErr(xlErrValue)
    End If
    
End Function
Function BlizhBolshRavnZnacheniye_V_Diapazone(ByVal Znachenie As Double, ByVal Diapazon As Range, Optional MaxZnachenie As Double = 0) As Double
    BlizhBolshRavnZnacheniye_V_Diapazone = Diapazon(NomerStroki_BlizhBolshRavn(Znachenie, Diapazon, MaxZnachenie))
End Function

Function NomerStroki_BlizhBolshRavn(ByVal Znachenie As Double, ByVal Diapazon As Range, Optional MaxZnachenie As Double = 0) As Long

    Dim ArrDiapazon()

    If Diapazon.Columns.Count > 1 Then Exit Function
    
    If MaxZnachenie <> 0 Then
        If Znachenie > MaxZnachenie Then Znachenie = MaxZnachenie
    End If
    
    Dim PoziciyaTochnogoSovpadeniya As Variant
    On Error Resume Next
    PoziciyaTochnogoSovpadeniya = WorksheetFunction.Match(Znachenie, Diapazon, 0)
    On Error GoTo 0
    
    Dim findVal As Boolean
    Dim i As Long
    If Not IsEmpty(PoziciyaTochnogoSovpadeniya) Then
        NomerStroki_BlizhBolshRavn = PoziciyaTochnogoSovpadeniya
    Else
        ArrDiapazon = Diapazon.Value
        For i = LBound(ArrDiapazon) To UBound(ArrDiapazon)
            If IsNumeric(ArrDiapazon(i, 1)) Then
                If ArrDiapazon(i, 1) >= Znachenie Then
                    findVal = True
                    Exit For
                End If
            End If
        Next
        If findVal = True Then
            NomerStroki_BlizhBolshRavn = i
        Else
            NomerStroki_BlizhBolshRavn = CVErr(xlErrValue)
        End If
    End If
    
End Function

Function LineynayaInterpolyatisiya(VremyaBolshe As Double, VremyaMenshe As Double, DlinaBolshe As Double, DlinaMenshe As Double, DlinaFakt As Double) As Double
    LineynayaInterpolyatisiya = VremyaMenshe + (DlinaFakt - VremyaMenshe) * ((VremyaBolshe - VremyaMenshe) / (DlinaBolshe - DlinaMenshe))
End Function

Function TsvetMateriala(Material As EnumMaterialy) As Double

    Dim MaterialColor As Double
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = POLIAMID Or Material = MEDNYYE_SPLAVY Then
        MaterialColor = &HDAEFE2   'Зеленый
    ElseIf Material = STAL_UGLERODISTAYA Or Material = STAL_LEGIROVANNAYA Then
        MaterialColor = &HF7EBDD   'Синий
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        MaterialColor = &H80000018    'Жёлтый
    ElseIf Material = TITANOVYYE_SPLAVY Then
        MaterialColor = &HF2F2F2   'Серый
    End If
    
    TsvetMateriala = MaterialColor

End Function

Function MaterialFromCboMaterial(cboMaterial As MSForms.ComboBox) As EnumMaterialy
    MaterialFromCboMaterial = CInt(cboMaterial.List(cboMaterial.ListIndex, 1))
End Function

Function PoluchitNomerStolbtsa(Tablitsa As ListObject, NaimenovaniyeStolbtsa As String) As Integer
    On Error Resume Next
    PoluchitNomerStolbtsa = Tablitsa.ListColumns(NaimenovaniyeStolbtsa).Index
    On Error GoTo 0
End Function
