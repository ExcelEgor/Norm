Attribute VB_Name = "mTipyDannykh"
Option Explicit

Public Type ParametryOperatsii
    Nomer As Double
    Naimenovaniye As String
    Razryad As Double
    EN As Double
    Tpz As Double
    Tsht As Double
End Type

Public Enum eTipyProizvodstva
    Edinichnyy = 1
    Melkoseriynyy = 2
    Krupnoseriynyy = 3
End Enum

Public Type NormyVremeni
    tMekhanika As Double
    tSlesar As Double
    tKontrol As Double
End Type

Public Function KoefTipProizv(TipProizv As eTipyProizvodstva) As Double
    Select Case TipProizv
        Case Edinichnyy
            KoefTipProizv = 1.3
        Case Melkoseriynyy
            KoefTipProizv = 1
        Case Krupnoseriynyy
            KoefTipProizv = 0.7
    End Select
End Function

