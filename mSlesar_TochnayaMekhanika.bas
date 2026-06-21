Attribute VB_Name = "mSlesar_TochnayaMekhanika"
'СЛЕСАРНО-СБОРОЧНЫЕ РАБОТЫ ПРИ ИЗГОТОВЛЕНИИ БЛОКОВ И ПРИБОРОВ ТОЧНОЙ МЕХАНИКИ
'Нормативы времени. Мелкосерийное производство. 1984 г.

Option Explicit

Public Const D_MAX_REZBA_METCHIKOM As Integer = 12
Public Const D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA As Integer = 10

Public Enum EnumSposobySverleniya
    NaStanke = 1
    RuchnoyDrelyu = 2
End Enum

Public Type ParamOtv_TochnayaMekhanika
    Material As Integer
    Diametr As Double
    GlubinaOtverstiya As Double
    GlubinaRezby As Double
    Glukhoye As Boolean
    IT As Integer
    Sposob As EnumSposobySverleniya
End Type
    
Public Enum TipObrabotkiOtverstiya
    Sverleniye = 1
    Razvertyvaniye = 2
    NarezaniyeRezby = 3
    SverleniyeRezba = 4
    SverleniyeRazvertyvaniye = 5
    KalibrovaniyeRezby = 6
    Puklevka = 7
    ZenkovaniyeFasok = 8
End Enum

 Function SverleniyeSlesar(Material As EnumMaterialy, Diametr As Double, GlubinaOtverstiya As Double, _
    Optional Sposob As EnumSposobySverleniya = RuchnoyDrelyu, Optional Glukhoye As Boolean = False) As Double
    'Карта 13. Обработка отверстий
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_sverleniye")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)

    If Diametr > D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA Or GruppaMateriala = 0 Then Exit Function
    
    If GlubinaOtverstiya > LmaxSverleniyeSlesar(Material, Diametr) Then Exit Function

    Dim kSposob As Double
    If Sposob = NaStanke Then
        kSposob = 1
    ElseIf Sposob = RuchnoyDrelyu Then
        kSposob = 1.5
    Else
        Exit Function
    End If
    
    Dim kMaterial As Double
    kMaterial = RaschotKoeffitsientaNaMaterial(CInt(Material))
    
    Dim kGlukhoye As Double
    kGlukhoye = IIf(Glukhoye, 1.2, 1)
    
    Dim tSverl As Double
    Dim i As Long
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = GruppaMateriala And Normativ(i, 2) >= Diametr And Normativ(i, 3) >= GlubinaOtverstiya Then
            tSverl = Normativ(i, 4)
            Exit For
        End If

    Next

    SverleniyeSlesar = kMaterial * kSposob * kGlukhoye * tSverl

End Function

Function LmaxSverleniyeSlesar(Material As EnumMaterialy, Diametr As Double) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_sverleniye")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)
    
    If Diametr > D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA Or GruppaMateriala = 0 Then Exit Function
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = GruppaMateriala And Normativ(i, 2) >= Diametr Then
            If Normativ(i, 3) > Lmax Then
                Lmax = Normativ(i, 3)
                If PoslednyayaStrokaVMassive(Normativ, i, 2) = True Then Exit For
            End If
        End If
    Next
   
    LmaxSverleniyeSlesar = Lmax
    
End Function

Private Function OpredelitGpuppuMateriala(Material As EnumMaterialy) As Integer
    
    Dim Gruppa As Integer
    
    If Material = ALUMINIYEVYYE_SPLAVY Then
        Gruppa = 1
    ElseIf Material = STAL_LEGIROVANNAYA Or Material = STAL_NERZHAVEYUSHCHAYA Or Material = STAL_UGLERODISTAYA Or Material = TITANOVYYE_SPLAVY Then
        Gruppa = 3
    ElseIf Material = MEDNYYE_SPLAVY Then
        Gruppa = 2
    End If
    
    OpredelitGpuppuMateriala = Gruppa
    
End Function

 Function NarezaniyeRezbyMetchikomSlesar(Material As EnumMaterialy, Diametr As Double, GlubinaRezby As Double, _
    Optional ByVal Kvalitet As Integer = 7, Optional Glukhoye As Boolean = False) As Double
    'Карта 13. Обработка отверстий

    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_narezaniye_rezby")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)

    If Diametr > D_MAX_REZBA_METCHIKOM Or Kvalitet < 7 Or GruppaMateriala = 0 Then Exit Function
    
    If GlubinaRezby > LmaxNarezaniyeRezbyMetchikomSlesar(Material, Diametr) Then Exit Function
    
    Dim kMaterial As Double
    kMaterial = RaschotKoeffitsientaNaMaterial(CInt(Material))
    
    Dim kGlukhoye As Double
    kGlukhoye = IIf(Glukhoye, 1.2, 1)
    
    Dim i As Long
    Dim tRezb As Double
    If Kvalitet > 7 Then Kvalitet = 9
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = GruppaMateriala And _
            Normativ(i, 2) >= Diametr And _
            Normativ(i, 3) = Kvalitet And _
            Normativ(i, 4) >= GlubinaRezby Then
            
            tRezb = Normativ(i, 5)
            Exit For
        End If
    Next
    
    NarezaniyeRezbyMetchikomSlesar = kMaterial * kGlukhoye * tRezb
    
End Function

Function LmaxNarezaniyeRezbyMetchikomSlesar(Material As EnumMaterialy, Diametr As Double) As Double
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_narezaniye_rezby")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)

    If Diametr > D_MAX_REZBA_METCHIKOM Or GruppaMateriala = 0 Then Exit Function
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
    
        If Normativ(i, 1) = GruppaMateriala And Normativ(i, 2) >= Diametr Then
            If Normativ(i, 4) > Lmax Then
                Lmax = Normativ(i, 4)
                If PoslednyayaStrokaVMassive(Normativ, i, 3) = True Then Exit For
            End If
        End If

    Next
   
    LmaxNarezaniyeRezbyMetchikomSlesar = Lmax
    
End Function

 Function RazvertyvaniyeSlesar(Material As EnumMaterialy, Diametr As Double, GlubinaOtverstiya As Double, _
    Optional ByVal Kvalitet As Integer = 7, Optional Glukhoye As Boolean = False) As Double
    'Карта 13. Обработка отверстий
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_razvertyvaniye")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)
    
    If Diametr > D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA Or Kvalitet < 7 Or GruppaMateriala = 0 Then Exit Function
    
    If GlubinaOtverstiya > LmaxRazvertyvaniyeSlesar(Material, Diametr) Then Exit Function
    
    Dim kMaterial As Double
    kMaterial = RaschotKoeffitsientaNaMaterial(CInt(Material))
    
    Dim kGlukhoye As Double
    kGlukhoye = IIf(Glukhoye, 1.2, 1)

    Dim i As Long
    Dim tRazvert As Double
    If Kvalitet > 7 Then Kvalitet = 9
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = GruppaMateriala And Normativ(i, 2) >= Diametr And Normativ(i, 3) = Kvalitet And Normativ(i, 4) >= GlubinaOtverstiya Then
            tRazvert = Normativ(i, 5)
            Exit For
        End If
    Next
    
    RazvertyvaniyeSlesar = kMaterial * kGlukhoye * tRazvert
    
End Function

Function LmaxRazvertyvaniyeSlesar(Material As EnumMaterialy, Diametr As Double) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_razvertyvaniye")
    
    Dim GruppaMateriala As Integer
    GruppaMateriala = OpredelitGpuppuMateriala(Material)

    If Diametr > D_MAX_SVERLENIYE_TOCHNAYA_MEKHANIKA Or GruppaMateriala = 0 Then Exit Function
    
    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) = GruppaMateriala And Normativ(i, 2) >= Diametr Then
            If Normativ(i, 4) > Lmax Then
                Lmax = Normativ(i, 4)
                If PoslednyayaStrokaVMassive(Normativ, i, 3) = True Then Exit For
            End If
        End If
    Next
    
    LmaxRazvertyvaniyeSlesar = Lmax
    
End Function

 Function KalibrovaniyeVnutrenneyRezby(Diametr As Double, GlubinaRezby As Double, Optional Glukhoye As Boolean = False) As Double
    'Карта 13. Обработка отверстий
    
    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_kalibrovaniye_rezby")
    
    If Diametr > D_MAX_REZBA_METCHIKOM Then Exit Function
    
    If GlubinaRezby > LmaxKalibrovaniyeVnutrenneyRezby(Diametr) Then Exit Function

    Dim i As Long
    Dim tKalibr As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) >= Diametr And Normativ(i, 2) >= GlubinaRezby Then
            tKalibr = Normativ(i, 3)
            Exit For
        End If
    Next
    
    Dim kGlukhoye As Double
    kGlukhoye = IIf(Glukhoye, 1.2, 1)
    
    KalibrovaniyeVnutrenneyRezby = kGlukhoye * tKalibr
    
End Function

Function LmaxKalibrovaniyeVnutrenneyRezby(Diametr As Double) As Double

    Dim Normativ
    Normativ = ZagruzitNormativ("tochnaya_kalibrovaniye_rezby")

    If Diametr > D_MAX_REZBA_METCHIKOM Then Exit Function

    Dim i As Long
    Dim Lmax As Double
    For i = 1 To UBound(Normativ)
        If Normativ(i, 1) >= Diametr Then
            If Normativ(i, 2) > Lmax Then
                Lmax = Normativ(i, 2)
                If PoslednyayaStrokaVMassive(Normativ, i, 1) = True Then Exit For
            End If
        End If
    Next
   
    LmaxKalibrovaniyeVnutrenneyRezby = Lmax
    
End Function

 Function SverleniyePlusRezba(Material As EnumMaterialy, Diametr As Double, GlubinaOtverstiya As Double, GlubinaRezby As Double, _
    Optional Kvalitet As Integer = 7, Optional Glukhoye As Boolean = False, Optional Sposob As EnumSposobySverleniya = RuchnoyDrelyu) As Double
    
    If Diametr > D_MAX_REZBA_METCHIKOM Or Kvalitet < 7 Then Exit Function
    
    Dim dSverl As Double
    dSverl = OpredelitDiametrSverleniyaPodRezbu(Diametr)

    Dim tSverl As Double
    tSverl = SverleniyeSlesar(CInt(Material), dSverl, GlubinaOtverstiya, Sposob, Glukhoye)
    
    If Not tSverl > 0 Then Exit Function
    
    Dim tRezba As Double
    tRezba = NarezaniyeRezbyMetchikomSlesar(CInt(Material), Diametr, GlubinaRezby, Kvalitet, Glukhoye)
    
    SverleniyePlusRezba = tSverl + tRezba
    
End Function
 
 Function LmaxSverleniyePlusRezba(Material As EnumMaterialy, Diametr As Double) As Integer
 
    Dim dSverl As Double
    dSverl = OpredelitDiametrSverleniyaPodRezbu(Diametr)
    
    LmaxSverleniyePlusRezba = WorksheetFunction.Min(LmaxSverleniyeSlesar(Material, dSverl), LmaxNarezaniyeRezbyMetchikomSlesar(Material, Diametr))
    
 End Function
Private Function OpredelitDiametrSverleniyaPodRezbu(DiametrRezby) As Double
    
    If DiametrRezby > D_MAX_REZBA_METCHIKOM Then Exit Function
 
    Dim dSverl As Double
    Select Case DiametrRezby
        Case Is <= 1.2: dSverl = 1
        Case Is <= 2:   dSverl = 1.5
        Case Is <= 3:   dSverl = 2
        Case Is <= 4:   dSverl = 3
        Case Is <= 5:   dSverl = 4
        Case Is <= 8:   dSverl = 6
        Case Is <= 10:  dSverl = 8
        Case Else:      dSverl = 10
    End Select
    
    OpredelitDiametrSverleniyaPodRezbu = dSverl
    
End Function

 Function SverleniyePlusRazvertyvaniye(Material As EnumMaterialy, Diametr As Double, GlubinaOtverstiya As Double, _
    Optional Glukhoye As Boolean = False, Optional Sposob As EnumSposobySverleniya = RuchnoyDrelyu, Optional Kvalitet As Integer = 7)
    
    SverleniyePlusRazvertyvaniye = SverleniyeSlesar(CInt(Material), Diametr, GlubinaOtverstiya, Sposob, Glukhoye) + _
        RazvertyvaniyeSlesar(Material, Diametr, GlubinaOtverstiya, Kvalitet, Glukhoye)

End Function

 Private Function ProverkaVkhodnykhParametrov_Obshchiye(Material As Integer, Diametr As Double, Glubina As Double, Lmax As Integer, Dmax As Integer) As Boolean
    
    ProverkaVkhodnykhParametrov_Obshchiye = True
    If Diametr <= 0 Or Glubina <= 0 Then
        ProverkaVkhodnykhParametrov_Obshchiye = False
    
    ElseIf Diametr > Dmax Then
        ProverkaVkhodnykhParametrov_Obshchiye = False
    
    Else
        If Glubina > Lmax Then
            ProverkaVkhodnykhParametrov_Obshchiye = False
        End If
    End If
    
End Function

Private Function RaschotKoeffitsientaNaMaterial(Material As EnumMaterialy) As Double
    
    Dim kMaterial As Double
    
    If Material = ALUMINIYEVYYE_SPLAVY Or Material = MEDNYYE_SPLAVY Or Material = STAL_UGLERODISTAYA Then
        kMaterial = 1
    ElseIf Material = STAL_LEGIROVANNAYA Then
        kMaterial = 1.1
    ElseIf Material = STAL_NERZHAVEYUSHCHAYA Then
        kMaterial = 1.5
    ElseIf Material = TITANOVYYE_SPLAVY Then
        kMaterial = 1.6
    Else
        Exit Function
    End If
    
    RaschotKoeffitsientaNaMaterial = kMaterial
    
End Function




