Attribute VB_Name = "mTpzCHPU"
Option Explicit

Function TpzCHPU_Tokarnyy(NaibolshiyDiametrNadStaninoy As Double, KolVoInstrumentov As Integer, Tsht As Double, KoefPervoyDetali As Double, _
    Optional Planshayba As Boolean = False, Optional Lunet As Boolean = False, Optional RastochkaKulachkov As Boolean = False) As Integer
    
    If KolVoInstrumentov <= 0 Then Exit Function

    Dim tBaza As Double, tInstrument As Double, tPlanshayba As Double, tLunet As Double, tRastochkaKulachkov As Double
    Select Case NaibolshiyDiametrNadStaninoy
        Case Is <= 250
            tBaza = 28
            tInstrument = 5.85
            tPlanshayba = 8
            tLunet = 3
            tRastochkaKulachkov = 9
        Case Is <= 400
            tBaza = 32.7
            tInstrument = 7
            tPlanshayba = 10
            tLunet = 4
            tRastochkaKulachkov = 9.4
        Case Is <= 630
            tBaza = 35.95
            tInstrument = 8.55
            tPlanshayba = 12
            tLunet = 5
            tRastochkaKulachkov = 10.6
        Case Else
            Exit Function
    End Select
    
    Dim tDopolnitelnyyePriemy As Double
    tDopolnitelnyyePriemy = IIf(Planshayba, tPlanshayba, 0) + IIf(Lunet, tLunet, 0) + IIf(RastochkaKulachkov, tRastochkaKulachkov, 0)
    
    TpzCHPU_Tokarnyy = RaschotTpz(Tsht, KoefPervoyDetali, tBaza, KolVoInstrumentov, tInstrument, 1, tDopolnitelnyyePriemy)
    
End Function

Function TpzCHPU_GorizantalnoRastochnoy(KolVoInstrumentov As Integer, Tsht As Double, KoefPervoyDetali As Double, Optional KolVoDopolnitelnoyOsnastki As Integer = 0) As Integer

    If KolVoInstrumentov <= 0 Then Exit Function
    
    Const tBaza As Double = 27.4
    Const tInstrument As Double = 7.1
    Const tOsnastka As Double = 14
    
    TpzCHPU_GorizantalnoRastochnoy = RaschotTpz(Tsht, KoefPervoyDetali, tBaza, KolVoInstrumentov, tInstrument, KolVoDopolnitelnoyOsnastki, tOsnastka)
    
End Function

Function TpzCHPU_Frezernyy(DlinaStola As Double, KolVoInstrumentov As Integer, Tsht As Double, KoefPervoyDetali As Double, Optional KolVoDopolnitelnoyOsnastki As Integer = 0) As Integer

    If DlinaStola <= 0 Or KolVoInstrumentov <= 0 Then Exit Function
    
    Dim tBaza As Double, tInstrument As Double, tOsnastka As Double
    Select Case DlinaStola
        Case Is <= 630
            tBaza = 22.9
            tInstrument = 6.9
            tOsnastka = 6
        Case 1250
            tBaza = 24.9
            tInstrument = 7.9
            tOsnastka = 11

        Case 2500
            tBaza = 27.7
            tInstrument = 8.6
            tOsnastka = 14

        Case Else
            Exit Function
    End Select
    
    TpzCHPU_Frezernyy = RaschotTpz(Tsht, KoefPervoyDetali, tBaza, KolVoInstrumentov, tInstrument, KolVoDopolnitelnoyOsnastki, tOsnastka)
    
End Function

Private Function RaschotTpz(Tsht As Double, KoefPervoyDetali As Double, tBaza As Double, KolVoInstrumentov As Integer, tInstrument As Double, _
    KolVoDopolnitelnoyOsnastki As Integer, tOsnastka As Double)
    
    RaschotTpz = OkruglVverkhSTochnostyu((KoefPervoyDetali - 1) * Tsht + tBaza + KolVoInstrumentov * tInstrument + KolVoDopolnitelnoyOsnastki * tOsnastka, 5)
    
End Function

