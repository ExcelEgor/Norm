VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShlifovaniye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public OpisaniyePerekhoda As String
Public Material As EnumMaterialy
Public Diametr As Double
Public Dlina As Double
Public Shirina As Double
Public Pripusk As Double
Public Kvalitet As Integer
Public Ra As Double
Public ShirinaShlifKruga As Double
Public TolshchinaDetali As Double
Public DopuskPloskostParallel As Double
Public KonicheskayaPoverhnost As Boolean
Public DopuskBieniyaSoosKrugl As Double
Public KolVoGalteley As Integer
Public KolVoPoverkhnostey As Double
Public UchestSlesar As Boolean
Public UchestKontrol As Boolean

Public Function RaschotIzParametrov() As RezultatyTokFrez

    Dim Normy As Variant
    Dim OpisaniyePerekhoda As String
    
    If KolVoPoverkhnostey = 0 Then KolVoPoverkhnostey = 1
    
    Select Case LCase(id)
        Case "ploskoye"
            Normy = Shlifovanie_Ploskoe(Material, Shirina, Dlina, Kvalitet, Ra, Pripusk, ShirinaShlifKruga, TolshchinaDetali, DopuskPloskostParallel)
            OpisaniyePerekhoda = PerekhodSUchotomKolVaPoverkhnostey("Шлифовать плоскость " & Shirina & "x" & Dlina, KolVoPoverkhnostey)
    End Select
    
    With RaschotIzParametrov
        If IsArray(Normy) Then
            .Normy(1) = Normy(1) * KolVoPoverkhnostey
            .Normy(2) = IIf(UchestSlesar, Normy(2) * KolVoPoverkhnostey, 0)
            .Normy(3) = IIf(UchestKontrol, Normy(3) * KolVoPoverkhnostey, 0)
        Else
            .Normy(1) = Normy * KolVoPoverkhnostey
            .Normy(2) = 0
            .Normy(3) = 0
        End If
        .OpisaniyePerekhoda = OpisaniyePerekhoda
    End With
    
End Function
