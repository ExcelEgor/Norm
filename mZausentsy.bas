Attribute VB_Name = "mZausentsy"
Option Explicit

Function ZachistkaZausentsev_Otverstiye(Material As EnumMaterialy, DiametrOtverstiya As Double, _
    Optional Glukhoye As Boolean = False, Optional Instrument As InstrumentyZachistki = NAPILNIK_SHABER) As Double
    'Карта 27. II, III. Обработка кромок отверстий детали после сверления при входе сверла и выходе сверла

    Dim tVkhod As Double
    tVkhod = ZachistkaZausentsev_Otverstiye_PriVkhodeSverla(Material, DiametrOtverstiya, Instrument)

    Dim tVykhod As Double
    If Glukhoye = False Then
        tVykhod = ZachistkaZausentsev_Otverstiye_PriVykhodeSverla(Material, DiametrOtverstiya, Instrument)
    Else
        tVykhod = 0
    End If

    ZachistkaZausentsev_Otverstiye = tVkhod + tVykhod

End Function

Private Function ZachistkaZausentsev_Otverstiye_PriVkhodeSverla(Material As EnumMaterialy, DiametrOtverstiya As Double, Instrument As InstrumentyZachistki) As Double
    'Карта 27. II. Обработка кромок отверстий детали после сверления при входе сверла
    
    If DiametrOtverstiya > 0 Then Exit Function
    
    Dim kMaterial As Double
    kMaterial = RaschotKoefNaMatrial(Material)
    If kMaterial = 0 Then Exit Function
    
    Dim koefA As Double
    Select Case Instrument
        Case NAPILNIK_SHABER
            koefA = 0.15
        Case PNEVM_MASHINA
            koefA = 0.05
        Case SVERLO
            koefA = 0.09
        Case Else
            Exit Function
    End Select
    
    ZachistkaZausentsev_Otverstiye_PriVkhodeSverla = koefA * kMaterial * DiametrOtverstiya ^ 0.31
    
End Function

Private Function ZachistkaZausentsev_Otverstiye_PriVykhodeSverla(Material As EnumMaterialy, DiametrOtverstiya As Double, Instrument As InstrumentyZachistki) As Double
    'Карта 27. III. Обработка кромок отверстий детали после сверления при выходе сверла
    
    If Not DiametrOtverstiya > 0 Then Exit Function
    
    Dim kMaterial As Double
    kMaterial = RaschotKoefNaMatrial(Material)
    If kMaterial = 0 Then Exit Function
    
    Dim koefA As Double
    Select Case Instrument
        Case NAPILNIK_SHABER
            koefA = 0.18
        Case PNEVM_MASHINA
            koefA = 0.07
        Case SVERLO
            koefA = 0.126
        Case Else
            Exit Function
    End Select
    
    ZachistkaZausentsev_Otverstiye_PriVykhodeSverla = koefA * kMaterial * DiametrOtverstiya ^ 0.32
    
End Function

Private Function RaschotKoefNaMatrial(Material As EnumMaterialy) As Double
    
    Dim kMaterial As Double
    Select Case Material
        Case STAL_LEGIROVANNAYA, STAL_UGLERODISTAYA
            kMaterial = 1
        Case STAL_NERZHAVEYUSHCHAYA
            kMaterial = 1.2
        Case CHUGUN
            kMaterial = 0.8
        Case ALUMINIYEVYYE_SPLAVY, MEDNYYE_SPLAVY
            kMaterial = 0.7
        Case Else
            Exit Function
    End Select
    
    RaschotKoefNaMatrial = kMaterial
    
End Function

