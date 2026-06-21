Attribute VB_Name = "mRaschotVTablitse"
Option Explicit
Option Private Module

Function PoluchitZnacheniye(Naimenovaniye As String, StrokaDannykh As Range, ZagolovkiDlyaPoiska As Range) As Variant
    Dim NomerStolbtsa As Integer
    NomerStolbtsa = OpredelitNomerStolbtsa(Naimenovaniye, ZagolovkiDlyaPoiska)
    
    If NomerStolbtsa = 0 Then
        Exit Function
    Else
        PoluchitZnacheniye = StrokaDannykh.Columns(NomerStolbtsa)
    End If
End Function

Function OpredelitNomerStolbtsa(Naimenovaniye As String, ZagolovkiDlyaPoiska As Range) As Integer
    On Error Resume Next
    OpredelitNomerStolbtsa = WorksheetFunction.Match(Naimenovaniye, ZagolovkiDlyaPoiska, 0)
    On Error GoTo 0
End Function
