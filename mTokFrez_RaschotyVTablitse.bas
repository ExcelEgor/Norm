Attribute VB_Name = "mTokFrez_RaschotyVTablitse"
Option Explicit

Function RaschotObrabotki_TokarnoFrezernyy(idPerehoda As String, NaimenovaniyaParametrov As Range, ZnacheniyaParametrov As Range, Material As Integer, _
    TipProizvodstva As Integer, Tonkostennaya As Boolean, PoRezhimam As Boolean, CHPU As Boolean) As Variant

    Dim kTipProizvTokFrez As Double
    kTipProizvTokFrez = PoprKoef_TipProizv_TokFrez(TipProizvodstva)
    
    Dim kTipProizvSles As Double
    kTipProizvSles = PoprKoef_TipProizv_Sles(TipProizvodstva)

    Dim Parametry As clsTokFrez
    Set Parametry = PrisvoitZnacheniyaParametram(idPerehoda, ZnacheniyaParametrov, Material, Tonkostennaya, PoRezhimam, NaimenovaniyaParametrov)
    
    Dim Normy
    Normy = Parametry.RaschotNormTokFrezIzParametrov.Normy
    
    Normy(1) = Normy(1) * kTipProizvTokFrez
    Normy(2) = Normy(2) * kTipProizvSles
    Normy(3) = Normy(3) * kTipProizvTokFrez
    
    RaschotObrabotki_TokarnoFrezernyy = Normy
    
End Function

Private Function PoprKoef_TipProizv_TokFrez(TipProizv As Integer) As Double
    Dim K1 As Double
    Select Case TipProizv
        Case 1: K1 = 1.3
        Case 2: K1 = 1
        Case 3: K1 = 0.7
    End Select
    PoprKoef_TipProizv_TokFrez = K1
End Function

Private Function PoprKoef_TipProizv_Sles(TipProizv As Integer) As Double
    Dim K1 As Double
    Select Case TipProizv
        Case 1: K1 = 1.3
        Case 2: K1 = 1
        Case 3: K1 = 0.8
    End Select
    PoprKoef_TipProizv_Sles = K1
End Function

Private Function PrisvoitZnacheniyaParametram(idPerehoda As String, ZnacheniyaParametrov As Range, Material As Integer, Tonkostennaya As Boolean, PoRezhimam As Boolean, NaimenovaniyaParametrov As Range) As clsTokFrez
    
    On Error Resume Next
    
    Dim col As New Collection
    Dim i As Integer
    For i = 1 To ZnacheniyaParametrov.Columns.Count
        col.Add ZnacheniyaParametrov(i), NaimenovaniyaParametrov(i)
    Next

    Dim Parametry As New clsTokFrez
    
    With Parametry
        
        .id = idPerehoda
        .D1 = col("D1")
        .d2 = col("d2")
        .DBObnizki = col("DBObnizki")
        .DlinaObnizki = col("DlinaObnizki")
        .GlubinaObnizki = col("GlubinaObnizki")
        .GlubinaRezaniya = col("GlubinaRezaniya")
        .IT = col("IT")
        .KolVoDetVPrutke = col("KolVoDetVPrutke")
        .KolVoPoverkhnostey = col("KolVoPoverkhnostey")
        .KolVoProkhodov = col("KolVoProkhodov")
        .Korka = col("Korka")
        .L = col("L")
        .Lrezby = col("Lrezby")
        .Otkloneniye = col("Otkloneniye")
        .PodrezkaTortsa = col("PodrezkaTortsa")
        .PoRezhimam = PoRezhimam
        .Pripusk = col("Pripusk")
        .Ra = col("Ra")
        .ShagRezby = col("ShagRezby")
        .ShirinaPoverhnosti = col("ShirinaPoverhnosti")
        .GlubinaPaza = col("GlubinaPaza")
        .TipPoverkhnosti = col("TipPoverkhnosti")
        .TipStanka = col("TipStanka")
        .TipTsentrovaniya = col("TipTsentrovaniya")
        .TocheniyeDiametra = col("TocheniyeDiametra")
        .Udar = col("Udar")
        .VUpor = col("Vupor")
        .Vconst = col("Vconst")
        .KalibrRezbSlesar = col("KalibrRezbSlesar")
        .Massa = col("Massa")
        .Material = Material
        .Sposob = col("Sposob")
        .KharkterVyverki = col("KharkterVyverki")
        .TochnostVyverki = col("TochnostVyverki")
        .ZachistkaBokovykhStoron = col("ZachistkaBokovykhStoron")
        .FrezPosleSverl = col("FrezPosleSverl")
        .Tonkostennaya = Tonkostennaya
        .UchestFasku = col("UchestFasku")
        .UchestKontrol = col("tKontrol*")
        .UchestSlesar = col("tSles*")
  

    End With
    
    Set PrisvoitZnacheniyaParametram = Parametry
    
    On Error GoTo 0
    
End Function





