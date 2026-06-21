VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShlifovaniye_Naruzhnoye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Diametr_NaruzhnoyeProdolnoye As String = "1 2 3 5 10 18 25 35 45 55 70 90 110 135 160 200 250 320 400"
Private Const Diametr_NaruzhnoyeRadialnoye As String = "0.5 1 2 3 5 10 18 25 35 45 55 70 90 110 135 160 200 250"

Function Shlifovanie_NaruzhnoyeProdolnoye(Material As Integer, ByVal Diametr As Double, Pripusk As Double, DlinaShlifovaniya As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, _
    Optional KolVoGalteley As Integer = 0, Optional ShirinaKruga As Double = 0) As Double
    
    If Diametr > D_MAKS_SHLIF_NARUZH_PROD Then Exit Function
    
    Dim tbShlifNaruzhProd As ListObject
    Set tbShlifNaruzhProd = wsShlif_NaruzhProd.ListObjects("tbShlifNaruzhProd")
    
    Diametr = BlizhBolshRavn_Znachenie_Array(Diametr, Split(Diametr_NaruzhnoyeProdolnoye))
    
    Dim MaksimalnyyPripusk As Double
    With tbShlifNaruzhProd
        MaksimalnyyPripusk = MaksZnacheniyePoNaboruUsloviy(.ListColumns(4).DataBodyRange, .ListColumns(1).DataBodyRange, Material, .ListColumns(3).DataBodyRange, Diametr)
    End With

    If Pripusk > MaksimalnyyPripusk Then Exit Function
    
    Dim i As Integer
    Dim Sm As Double, Stx As Double
    With tbShlifNaruzhProd
        i = WorksheetFunction.Match(Material & Diametr, .ListColumns(7).DataBodyRange, 0)
        For i = i To .ListRows.Count
            If .DataBodyRange(i, 4) >= Pripusk Then
                Stx = .DataBodyRange(i, 5)
                Sm = .DataBodyRange(i, 6)
                Exit For
            End If
        Next
    End With
    
    Dim tShlif As Double
    tShlif = ((DlinaShlifovaniya * Pripusk) / (Stx * Sm) + (0.4 * DlinaShlifovaniya ^ 0.1)) * 1.1
    
    Dim Shlifovaniye_Koeff As New clsShlifovaniye_Koeff, K As Double
    K = Shlifovaniye_Koeff.RaschotKoeff_NauzhProd(IT, Ra, KolVoGalteley, DopuskBieniyaSoosnostiKruglosti, KonicheskayaPoverhnost, ShirinaKruga)
    
    Shlifovanie_NaruzhnoyeProdolnoye = K * tShlif

End Function

Function Shlifovanie_NaruzhnoyeRadialnoye(Material As Integer, ByVal Diametr As Double, Pripusk As Double, DlinaShlifovaniya As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, Optional KolVoGalteley As Integer = 0) As Double
    
    If Diametr > D_MAKS_SHLIF_NARUZH_RADIAL Then Exit Function
    Diametr = BlizhBolshRavn_Znachenie_Array(Diametr, Split(Diametr_NaruzhnoyeRadialnoye))
    
    If DlinaShlifovaniya > LmaxRadialnoye(Material, Diametr) Or Pripusk > MaksPripusk_Radialnoye(Material, Diametr) Then Exit Function
    
    Dim tbShlifNaruzhRadial As ListObject
    Set tbShlifNaruzhRadial = wsShlif_NaruzhRadial.ListObjects(1)
    
    With tbShlifNaruzhRadial

        Dim i As Integer
        Dim st As Double
        i = WorksheetFunction.Match(Material & Diametr, .ListColumns(7).DataBodyRange, 0) - 1
        
        For i = i To .ListRows.Count
            If .DataBodyRange(i, 4) >= DlinaShlifovaniya And .DataBodyRange(i, 5) >= Pripusk Then
                st = .DataBodyRange(i, 6)
                Exit For
            End If
        Next
        
    End With
    
    Dim tShlif As Double
    If Diametr <= 110 Then
        tShlif = (Pripusk / st + 0.35) * 1.1
    Else
        tShlif = ((Pripusk / st) + (0.35 * DlinaShlifovaniya ^ 0.1)) * 1.1
    End If
    
    Dim Shlifovaniye_Koeff As New clsShlifovaniye_Koeff, K As Double
    K = Shlifovaniye_Koeff.RaschotKoeff_NaruzhRadial(IT, Ra, KolVoGalteley, DopuskBieniyaSoosnostiKruglosti)

    Shlifovanie_NaruzhnoyeRadialnoye = K * tShlif
    
End Function

Public Function LmaxRadialnoye(ByVal Material As Integer, ByVal Diametr As Double) As Double

    Diametr = BlizhBolshRavn_Znachenie_Array(Diametr, Split(Diametr_NaruzhnoyeRadialnoye))
 
    Dim tbShlifNaruzhRadial As ListObject
    Set tbShlifNaruzhRadial = wsShlif_NaruzhRadial.ListObjects(1)
    
    With tbShlifNaruzhRadial
        LmaxRadialnoye = MaksZnacheniyePoNaboruUsloviy(.ListColumns(4).DataBodyRange, _
            .ListColumns(1).DataBodyRange, Material, _
            .ListColumns(3).DataBodyRange, Diametr)
    End With
    
End Function

Function MaksPripusk_Radialnoye(ByVal Material As Integer, ByVal Diametr As Double) As Double

    Diametr = BlizhBolshRavn_Znachenie_Array(Diametr, Split(Diametr_NaruzhnoyeRadialnoye))
 
    Dim tbShlifNaruzhRadial As ListObject
    Set tbShlifNaruzhRadial = wsShlif_NaruzhRadial.ListObjects(1)
    
    With tbShlifNaruzhRadial
        MaksPripusk_Radialnoye = MaksZnacheniyePoNaboruUsloviy(.ListColumns(5).DataBodyRange, _
            .ListColumns(1).DataBodyRange, Material, _
            .ListColumns(3).DataBodyRange, Diametr)
    End With
    
End Function

