Attribute VB_Name = "mKoordinat_RaschotVTablitse"
Option Explicit

Function RaschotObrabotki_Koordinat(idPerehoda As Integer, NaimenovaniyaParametrov As Range, ZnacheniyaParametrov As Range, Material As Integer, KolVo As Double, _
    TipProizvodstva As Integer, TipStanka As Integer, VozrastStanka As Integer, MaterialInstrumenta As Integer) As Variant
    
    Dim varNormy As Variant
     
    Dim kTipProizv As Double
    kTipProizv = PoprKoef_Koordinat_TipProizv(TipProizvodstva)
    
    If KolVo = 0 Then KolVo = 1
    
    If idPerehoda = 1 Then
        varNormy = RaschotUstanovki(NaimenovaniyaParametrov, ZnacheniyaParametrov)
        
    Else
        varNormy = RaschotMekhObrabotki(idPerehoda, NaimenovaniyaParametrov, ZnacheniyaParametrov, Material, TipStanka, VozrastStanka, MaterialInstrumenta)
        
    End If
    
    Dim Normy As Variant
    Normy = PreobrazovatNormyVMassiv(varNormy)
    
    Dim i As Integer
    For i = LBound(Normy) To UBound(Normy)
        Normy(i) = Normy(i) * KolVo * kTipProizv
    Next
    
    RaschotObrabotki_Koordinat = Normy
    
End Function
Private Function RaschotUstanovki(NaimenovaniyaParametrov As Range, ZnacheniyaParametrov As Range) As Double

    Dim ParametryUstanovki As ParameryUstanovki_Koordinat

    With ParametryUstanovki
        .HarakterViverki = ZnacheniyaParametrov(WorksheetFunction.Match("HarakterViverki", NaimenovaniyaParametrov, 0))
        .KolVoDopBoltov = ZnacheniyaParametrov(WorksheetFunction.Match("KolVoDopBoltov", NaimenovaniyaParametrov, 0))
        .KolVoDopDomkratov = ZnacheniyaParametrov(WorksheetFunction.Match("KolVoDopDomkratov", NaimenovaniyaParametrov, 0))
        .KolVoUstDetaley = ZnacheniyaParametrov(WorksheetFunction.Match("KolVoUstDetaley", NaimenovaniyaParametrov, 0))
        .Massa = ZnacheniyaParametrov(WorksheetFunction.Match("Massa", NaimenovaniyaParametrov, 0))
        .Nezestk = ZnacheniyaParametrov(WorksheetFunction.Match("Nezestk", NaimenovaniyaParametrov, 0))
        .Sposob = ZnacheniyaParametrov(WorksheetFunction.Match("Sposob", NaimenovaniyaParametrov, 0))
        .TochnostViverki = ZnacheniyaParametrov(WorksheetFunction.Match("TochnostViverki", NaimenovaniyaParametrov, 0))
    
        RaschotUstanovki = Ustanovka_Koordinat(.Massa, .Sposob, .HarakterViverki, .Nezestk, .TochnostViverki, .KolVoDopBoltov, .KolVoDopDomkratov, .KolVoUstDetaley)
    
    End With
        
End Function

Private Function RaschotMekhObrabotki(idPerehoda As Integer, NaimenovaniyaParametrov As Range, ZnacheniyaParametrov As Range, Material As Integer, _
    TipStanka As Integer, VozrastStanka As Integer, MaterialInstrumenta As Integer) As Variant

    

    Dim Parametry As Koordinat
    
        With Parametry
    
            .d = ZnacheniyaParametrov(WorksheetFunction.Match("D", NaimenovaniyaParametrov, 0))
            .L = ZnacheniyaParametrov(WorksheetFunction.Match("L", NaimenovaniyaParametrov, 0))
            .SposobSovmeshcheniyaOsi = ZnacheniyaParametrov(WorksheetFunction.Match("SovmeshcheniyeOsi", NaimenovaniyaParametrov, 0))
            .PredvaritelnayaTsentrovka = ZnacheniyaParametrov(WorksheetFunction.Match("PredvaritelnayaTsentrovka", NaimenovaniyaParametrov, 0))
            .Glukhoye = ZnacheniyaParametrov(WorksheetFunction.Match("Gluhoe", NaimenovaniyaParametrov, 0))
            .NaklonnayaPloskost = ZnacheniyaParametrov(WorksheetFunction.Match("NaklonnayaPloskost", NaimenovaniyaParametrov, 0))
            .TonkostennyaDetal = ZnacheniyaParametrov(WorksheetFunction.Match("TonkostennyaDetal", NaimenovaniyaParametrov, 0))
            .NaUdar = ZnacheniyaParametrov(WorksheetFunction.Match("NaUdar", NaimenovaniyaParametrov, 0))
            .IT = ZnacheniyaParametrov(WorksheetFunction.Match("IT", NaimenovaniyaParametrov, 0))
            .Ra = ZnacheniyaParametrov(WorksheetFunction.Match("Ra", NaimenovaniyaParametrov, 0))
            .BieniyeaSoosnostKruglost = ZnacheniyaParametrov(WorksheetFunction.Match("BieniyeaSoosnostKruglost", NaimenovaniyaParametrov, 0))
            .ZakalonnayaStal = ZnacheniyaParametrov(WorksheetFunction.Match("ZakalonnayaStal", NaimenovaniyaParametrov, 0))
            .KolVoDopProhodov = ZnacheniyaParametrov(WorksheetFunction.Match("KolVoDopProhodov", NaimenovaniyaParametrov, 0))
            
            Dim varNormy As Variant
            Select Case idPerehoda
                Case 2  'Наметить
                    varNormy = NametkaOtverstiy_Koordinat(.SposobSovmeshcheniyaOsi, .NaklonnayaPloskost)
                Case 3  'Сверлить
                    varNormy = Sverlenie_Koordinat(CInt(Material), .d, .L, .SposobSovmeshcheniyaOsi, .PredvaritelnayaTsentrovka, .Glukhoye, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar)
                Case 5  'Растачивание
                    varNormy = Rastachivaniye_Koordinat(CInt(Material), .d, .L, .SposobSovmeshcheniyaOsi, .IT, .Ra, .Glukhoye, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar, .BieniyeaSoosnostKruglost, .ZakalonnayaStal, .KolVoDopProhodov)
                Case 6  'Растачивание выточек
                    varNormy = RastachivaniyeVytochek(CInt(Material), .d, .L, .IT, .SposobSovmeshcheniyaOsi <> BezSovmeshcheniyaOsey, .NaklonnayaPloskost, .TonkostennyaDetal, .NaUdar, .BieniyeaSoosnostKruglost, .ZakalonnayaStal)
            End Select
            
            Dim kIntsrument As Double, kTipStanka As Double, kVozrast As Double
            kIntsrument = PoprKoef_Koordinat_MatrialInstrument(MaterialInstrumenta)
            kTipStanka = PoprKoef_Koordinat_TipStanka(TipStanka)
            kVozrast = PoprKoef_Koordinat_Vozrast(VozrastStanka)
            
            Dim Normy As Variant
            Normy = PreobrazovatNormyVMassiv(varNormy)
            
            Normy(1) = Normy(1) * kIntsrument * kTipStanka * kVozrast
            Normy(2) = Normy(2)
            Normy(3) = Normy(3)

    
        End With
        
        RaschotMekhObrabotki = Normy
        
End Function

Private Function PreobrazovatNormyVMassiv(varNormy As Variant) As Variant
    
    Dim Normy(1 To 3) As Double
    
    If IsArray(varNormy) Then
        Normy(1) = varNormy(1)
        Normy(2) = varNormy(2)
        Normy(3) = varNormy(3)
    Else
        Normy(1) = varNormy
        Normy(2) = 0
        Normy(3) = 0
    End If
    
    PreobrazovatNormyVMassiv = Normy
    
End Function
