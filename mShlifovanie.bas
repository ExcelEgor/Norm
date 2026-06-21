Attribute VB_Name = "mShlifovanie"
Option Explicit

Public Const D_MAKS_SHLIF_NARUZH_PROD As Integer = 400
Public Const D_MAKS_SHLIF_NARUZH_RADIAL As Integer = 250
Public Const MAKS_SHIRINA_SHLIFOVANIYA As Integer = 300
Public Const D_MAKS_VNUTR_SHLIF As Integer = 150
Public Const IMYA_KNIGI_SHLIFOVANIYE As String = "shlifovaniye.xlsb"

Public Function Shlifovanie_Vnutrenneye(Material As EnumMaterialy, ByVal Diametr As Double, Pripusk As Double, Dlina As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, Optional ShirinaShlifKruga As Double = 40) As Double
    
    Dim ShlifVnutr As New clsShlifovaniye_Vnutrenneye
    
    Shlifovanie_Vnutrenneye = ShlifVnutr.RaschotTsht(Material, Diametr, Pripusk, Dlina, IT, Ra, KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, ShirinaShlifKruga)

End Function

Public Function Shlifovanie_Ploskoe(Material As EnumMaterialy, ByVal Shirina As Double, ByVal Dlina As Double, _
    ByVal IT As Integer, Ra As Double, Pripusk As Double, _
    Optional ShirinaShlifKruga As Double, Optional TolshchinaDetali As Double, Optional DopuskPloskostnostiParallelnosti As Double = 0) As Double
    
    Dim ShlifPlosk As New clsShlifovaniye_Ploskoye
    
    Shlifovanie_Ploskoe = ShlifPlosk.RaschotTsht(Material, Shirina, Dlina, IT, Ra, Pripusk, ShirinaShlifKruga, TolshchinaDetali, DopuskPloskostnostiParallelnosti)

End Function

Public Function Shlifovanie_NaruzhnoyeProdolnoye(Material As Integer, ByVal Diametr As Double, Pripusk As Double, DlinaShlifovaniya As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, _
    Optional KolVoGalteley As Integer = 0, Optional ShirinaKruga As Double = 0) As Double
    
    Dim ShlifNaruzh As New clsShlifovaniye_Naruzhnoye
    Shlifovanie_NaruzhnoyeProdolnoye = ShlifNaruzh.Shlifovanie_NaruzhnoyeProdolnoye(Material, Diametr, Pripusk, _
        DlinaShlifovaniya, IT, Ra, KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, KolVoGalteley, ShirinaKruga)

End Function

Public Function Shlifovanie_NaruzhnoyeRadialnoye(Material As Integer, ByVal Diametr As Double, Pripusk As Double, DlinaShlifovaniya As Double, IT As Integer, Ra As Double, _
    Optional KonicheskayaPoverhnost As Boolean = False, Optional DopuskBieniyaSoosnostiKruglosti As Double = 0, Optional KolVoGalteley As Integer = 0) As Double
    
    Dim ShlifRad As New clsShlifovaniye_Naruzhnoye
    Shlifovanie_NaruzhnoyeRadialnoye = ShlifRad.Shlifovanie_NaruzhnoyeRadialnoye(Material, Diametr, Pripusk, _
        DlinaShlifovaniya, IT, Ra, KonicheskayaPoverhnost, DopuskBieniyaSoosnostiKruglosti, KolVoGalteley)
    
End Function

Public Function MaterialKorrektnyy_Shlifovaniye(Material As EnumMaterialy) As Boolean
    Select Case Material
        Case STAL_LEGIROVANNAYA, STAL_NERZHAVEYUSHCHAYA, STAL_UGLERODISTAYA, TITANOVYYE_SPLAVY
            MaterialKorrektnyy_Shlifovaniye = True
        Case Else
            MaterialKorrektnyy_Shlifovaniye = False
    End Select
End Function
