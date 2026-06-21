VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public OpisaniyePerekhoda As String
Public Dlina As Double
Public DlinaKrivoy As Double
Public Shirina As Double
Public Glubina As Double
Public GlubinaRezby As Double
Public Diametr As Double
Public Pripusk As Double
Public Glukhoye As Boolean
Public IT As Integer
Public Ra As Double
Public Dobnizheniya As Double
Public Lobnizheniya As Double
Public KolVoPoverkhnostey As Double
Public Radius As Double
Public Massa As Double
Public SposobUstanovki As EnumUstanovkaDatron
Public Material As EnumMaterialy
Public UchestSlesar As Boolean
Public UchestKontrol As Boolean
Public VysotaFaski As Double
Public BezZachistki As Boolean

Public Function RaschotIzParametrov() As NormyVremeni

    Dim Normy As Variant
    Select Case id
        Case "FrezerovaniyeObnizheniy"
            Normy = Datron_FrezerovaniyeObnizheniy(Material, Dlina, Shirina, Glubina, Radius, BezZachistki)
        Case "VyrezaniyePoKonturu"
            Normy = Datron_VyrezaniyePoKonturu(Material, Dlina, Glubina)
        Case "FrezerovaniyeTortsevoyFrezoy"
            Normy = Datron_FrezerovaniyeTortsevoyFrezoy(Material, Dlina, Shirina, Pripusk)
        Case "ObrabotkaRezbovykhOtverstiy"
            Normy = Datron_ObrabotkaRezbovykhOtverstiy(Material, Diametr, Glubina, GlubinaRezby, Glukhoye, Dobnizheniya, Lobnizheniya)
        Case "ObrabotkaOtverstiy"
            Normy = Datron_ObrabotkaOtverstiy(Material, Diametr, Glubina, Glukhoye, VysotaFaski, Dobnizheniya, Lobnizheniya)
        Case "StupenchatoyeFrezerovaniye"
            Normy = Datron_PostrochnoyeFrezerovaniye(Material, Dlina, DlinaKrivoy)
        Case "Ustanovka"
            Normy = Datron_Ustanovka(Massa, SposobUstanovki)
        Case "FrezPazov"
            Normy = Datron_FrezerovaniyePazov(Material, Shirina, Glubina, Dlina)
        Case "FrezUstup"
            Normy = Datron_FrezerovaniyeUstupov(Material, Dlina, Shirina, Glubina)
    End Select
    
    With RaschotIzParametrov
        If IsArray(Normy) Then
            .tMekhanika = KolVoPoverkhnostey * Normy(1)
            .tSlesar = IIf(UchestSlesar, Normy(2) * KolVoPoverkhnostey, 0)
            .tKontrol = IIf(UchestKontrol, Normy(3) * KolVoPoverkhnostey, 0)
        Else
            .tMekhanika = Normy * KolVoPoverkhnostey
            .tSlesar = 0
            .tKontrol = 0
        End If
    End With
    
End Function

Public Function PrisvoitParametryIzTablitsy(Material As EnumMaterialy, Zagolovki As Range, StrokaDannykh As Range) As clsDatron
    
    Dim Parametry As New clsDatron
    
    Dim i As Integer
    For i = 1 To Zagolovki.Columns.Count
        With Parametry
            .Diametr = PoluchitZnacheniye("Diametr", StrokaDannykh, Zagolovki)
            .Dlina = PoluchitZnacheniye("Dlina", StrokaDannykh, Zagolovki)
            .DlinaKrivoy = PoluchitZnacheniye("DlinaKrivoy", StrokaDannykh, Zagolovki)
            .Dobnizheniya = PoluchitZnacheniye("dObnizheniya", StrokaDannykh, Zagolovki)
            .Glubina = PoluchitZnacheniye("Glubina", StrokaDannykh, Zagolovki)
            .GlubinaRezby = PoluchitZnacheniye("GlubinaRezby", StrokaDannykh, Zagolovki)
            .Glukhoye = PoluchitZnacheniye("Glukhoye", StrokaDannykh, Zagolovki)
            .id = PoluchitZnacheniye("id", StrokaDannykh, Zagolovki)
            .VysotaFaski = PoluchitZnacheniye("VysotaFaski", StrokaDannykh, Zagolovki)
            .BezZachistki = PoluchitZnacheniye("BezZachistki", StrokaDannykh, Zagolovki)
            .IT = PoluchitZnacheniye("IT", StrokaDannykh, Zagolovki)
            
            .KolVoPoverkhnostey = PoluchitZnacheniye("KolVoPoverkhnostey", StrokaDannykh, Zagolovki)
            If .KolVoPoverkhnostey = 0 Then .KolVoPoverkhnostey = 1
            
            .Lobnizheniya = PoluchitZnacheniye("lObnizheniya", StrokaDannykh, Zagolovki)
            .Massa = PoluchitZnacheniye("Massa", StrokaDannykh, Zagolovki)
            .Material = Material
            .OpisaniyePerekhoda = PoluchitZnacheniye("OpisaniyePerekhoda", StrokaDannykh, Zagolovki)
            .Pripusk = PoluchitZnacheniye("Pripusk", StrokaDannykh, Zagolovki)
            .Ra = PoluchitZnacheniye("Ra", StrokaDannykh, Zagolovki)
            .Radius = PoluchitZnacheniye("Radius", StrokaDannykh, Zagolovki)
            .Shirina = PoluchitZnacheniye("Shirina", StrokaDannykh, Zagolovki)
            .SposobUstanovki = PoluchitZnacheniye("SposobUstanovki", StrokaDannykh, Zagolovki)
            .UchestKontrol = PoluchitZnacheniye("tKontrol*", StrokaDannykh, Zagolovki)
            .UchestSlesar = PoluchitZnacheniye("tSlesar*", StrokaDannykh, Zagolovki)
        End With
    Next
    
    Set PrisvoitParametryIzTablitsy = Parametry
    
End Function
