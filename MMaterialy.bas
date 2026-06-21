Attribute VB_Name = "MMaterialy"
Option Explicit

Public Enum EnumMaterialy

    ALUMINIYEVYYE_SPLAVY = 1
    MEDNYYE_SPLAVY = 2
    STAL_UGLERODISTAYA = 3
    STAL_LEGIROVANNAYA = 4
    STAL_NERZHAVEYUSHCHAYA = 5
    TITANOVYYE_SPLAVY = 6
    REZINA = 7
    TEKSTOLIT = 8
    STEKLOTEKSTOLIT = 9
    ORGSTEKLO = 10
    GETINAKS = 11
    POLIAMID = 12
    KARTON = 13
    CHUGUN = 14
    PARONIT = 15
    FIBRA = 16
    PLASTMASSA = 17
    
End Enum

Public Function idMaterialPoNaimenovaniyu(Naimenovaniye As String) As Integer

    Dim i As Integer
    For i = 1 To TB_MATERIALY.ListRows.Count
        If Replace(LCase(TB_MATERIALY.DataBodyRange(i, 1)), " ", "") = Replace(LCase(Naimenovaniye), " ", "") Then
            idMaterialPoNaimenovaniyu = TB_MATERIALY.DataBodyRange(i, 2)
            Exit For
        End If
    Next
    
End Function

Public Function RaschotMassyPoMaterialu_i_Obyomu(Material As EnumMaterialy, Obyom As Double) As Double

    Dim Plotnost As Double
    Plotnost = PlotnostMateriala(CInt(Material))
    
    If Plotnost = 0 Then
        Err.Raise vbObjectError + 1, , "Некорректный материал"
    Else
        RaschotMassyPoMaterialu_i_Obyomu = (Plotnost * Obyom) / 10 ^ 9
    End If
    
End Function

Function PlotnostMateriala(Material As EnumMaterialy) As Double
    
    Dim ArrMateria()
    ArrMateria = wsMaterial.ListObjects("tbMainMaterial").DataBodyRange
    
    Dim i As Integer
    For i = LBound(ArrMateria) To UBound(ArrMateria)
        If ArrMateria(i, 2) = Material Then
            PlotnostMateriala = ArrMateria(i, 3)
            Exit For
        End If
    Next
    
End Function

Function NazvaniyeMateriala(Material As EnumMaterialy) As String
    
    Select Case Material
        Case ALUMINIYEVYYE_SPLAVY
            NazvaniyeMateriala = "Алюминиевые сплавы"
        
        Case MEDNYYE_SPLAVY
            NazvaniyeMateriala = "Медные сплавы"
        
        Case STAL_UGLERODISTAYA
            NazvaniyeMateriala = "Сталь углеродистая"
        
        Case STAL_LEGIROVANNAYA
            NazvaniyeMateriala = "Сталь легированная"
        
        Case STAL_NERZHAVEYUSHCHAYA
            NazvaniyeMateriala = "Сталь нержавеющая"
        
        Case TITANOVYYE_SPLAVY
            NazvaniyeMateriala = "Титановые сплавы"
        
        Case REZINA
            NazvaniyeMateriala = "Резина"
        
        Case TEKSTOLIT
            NazvaniyeMateriala = "Текстолит"
        
        Case STEKLOTEKSTOLIT
            NazvaniyeMateriala = "Стеклотекстолит"
        
        Case ORGSTEKLO
            NazvaniyeMateriala = "Оргстекло"
        
        Case GETINAKS
            NazvaniyeMateriala = "Гетинакс"
        
        Case POLIAMID
            NazvaniyeMateriala = "Полиамид"
        
        Case KARTON
            NazvaniyeMateriala = "Картон"
        
        Case CHUGUN
            NazvaniyeMateriala = "Чугун"

        Case PARONIT
            NazvaniyeMateriala = "Паронит"
            
        Case FIBRA
            NazvaniyeMateriala = "Фибра"
            
    End Select
    
End Function
